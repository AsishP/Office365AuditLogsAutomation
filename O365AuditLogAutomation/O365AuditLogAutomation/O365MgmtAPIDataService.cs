using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using Newtonsoft.Json;
using System.IO;
using Microsoft.Extensions.Logging;
using System.Linq;
using Microsoft.WindowsAzure.Storage; 
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.WindowsAzure.Storage.File;
using System.Text;
using CsvHelper;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Net;

namespace O365AuditLogAutomation
{
    /// <summary>
    /// Service Layer class to do all the heavy lifting
    /// </summary>
    class O365MgmtAPIDataService
    {
        private string accessToken;
        private string AuditLogDataTableName { get; set; }
        private string AuditLogAnalyticsTableName { get; set; }
        private string AuditLogCSVExportFileName { get; set; }
        private string AuditLogStorageConnectionString{ get; set; }
        private DateTime AuditLogDateTime { get; set; }
        private string SPUserName { get; set; }
        private string SPUserPassword { get; set; }

        private ILogger log;

        private string AuditLogPartitionKeyPrefix { get; set; }
        private string AuditLogOpspartitionKeyPrefix { get; set; }

        /// <summary>
        /// Intialize the service class with start up values
        /// </summary>
        /// <param name="token"></param>
        /// <param name="funclog"></param>
        public O365MgmtAPIDataService(string token, ILogger funclog)
        {
            try
            {
                accessToken = token;
                log = funclog;
                AuditLogStorageConnectionString = System.Environment.GetEnvironmentVariable("StorageConnectionString");
                AuditLogDataTableName = System.Environment.GetEnvironmentVariable("AuditLogDataTable");
                AuditLogAnalyticsTableName = System.Environment.GetEnvironmentVariable("AuditLogAnalyticsTable");
                AuditLogCSVExportFileName = System.Environment.GetEnvironmentVariable("AuditLogCSVExportLocation");
                SPUserName = System.Environment.GetEnvironmentVariable("SPUserName");
                SPUserPassword = System.Environment.GetEnvironmentVariable("SPUserPassword");
                AuditLogPartitionKeyPrefix = System.Environment.GetEnvironmentVariable("AuditLogDataTablePrefix");
                AuditLogOpspartitionKeyPrefix = System.Environment.GetEnvironmentVariable("AuditLogOpsTablePrefix");
                AuditLogDateTime = getAESTTime();
                //// For Testing only 
                //AuditLogDateTime = DateTime.Parse("2018-12-10T07:00");
                log.LogInformation($" Setting Audit Log Time to {getAuditLongDateTimeString(AuditLogDateTime)}");


                bool isServiceReady = AuditLogStorageConnectionString != "" && AuditLogDataTableName != "" 
                                  && AuditLogAnalyticsTableName != "" && AuditLogCSVExportFileName != ""
                                  && SPUserName != "" && SPUserPassword != ""
                                  && AuditLogPartitionKeyPrefix != "" && AuditLogOpspartitionKeyPrefix != "";

                if (!isServiceReady)
                    throw new Exception("Some of the environment variables are not initialised. Please fix them before proceeding");
 
            }
            catch(Exception ex)
            {
                log.LogError($"Couldn't initialize the data service. Error message - {ex.Message}");
            }
        }

        /// <summary>
        /// Fetch the initial data batch with signatures of detail calls
        /// </summary>
        /// <param name="SPServiceUrl"></param>
        /// <param name="urlParameters"></param>
        /// <returns></returns>
        public AuditInitialDataObject getAuditInitalData(string SPServiceUrl, string urlParameters)
        {
            AuditInitialDataObject auditInitialDataObj = new AuditInitialDataObject();
            try
            {
                List<AuditInitialReport> auditInitialReports = new List<AuditInitialReport>();
                // **** Call the Http Client Service ****
                HttpClient client = new HttpClient();
                client.BaseAddress = new Uri(SPServiceUrl);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken.ToString());
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                // List data response.
                HttpResponseMessage response = client.GetAsync(urlParameters, HttpCompletionOption.ResponseContentRead).Result; // Blocking call!
                if (response.IsSuccessStatusCode)
                {
                    // Parse the response body. Blocking!
                    Stream dataObjects = response.Content.ReadAsStreamAsync().Result;
                    StreamReader reader = new StreamReader(dataObjects);
                    string responseObj = reader.ReadToEnd();
                    auditInitialReports = JsonConvert.DeserializeObject<List<AuditInitialReport>>(responseObj);
                    IEnumerable<string> values;
                    
                    //// used to loop through across pages of initial data api calls
                    if (response.Headers.TryGetValues("NextPageUri", out values))
                    {
                        auditInitialDataObj.AuditNextPageUri = values.First();
                        auditInitialDataObj.AuditInitialDataObj = auditInitialReports;
                    }
                    else
                    {
                        auditInitialDataObj.AuditNextPageUri = "";
                        auditInitialDataObj.AuditInitialDataObj = auditInitialReports;
                    }
                }
                else
                {
                    log.LogError($"{(int)response.StatusCode} ({response.ReasonPhrase})");
                }
            }
            catch(Exception ex)
            {
                log.LogError($"Error while fetching initial Audit Data. Error message - {ex.Message}");
            }

            return auditInitialDataObj;
        }

        /// <summary>
        /// Fetch the detail log data from the URI got from the Initial Data Call
        /// </summary>
        /// <param name="SPServiceUrl"></param>
        /// <returns></returns>
        public List<AuditDetailedReport> getAuditDetailData(string SPServiceUrl)
        {
            List<AuditDetailedReport> auditDetailData = new List<AuditDetailedReport>();
            try
            {
                int retries = 0;
                bool success = false;
                // **** Call the Http Client Service ****
                HttpClient client = new HttpClient();
                string urlParameters = "";
                client.BaseAddress = new Uri(SPServiceUrl);

                // Add an Accept header for JSON format.
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken.ToString());
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                while (!success && retries <= 4)
                {
                    // List data response.
                    HttpResponseMessage response = client.GetAsync(urlParameters, HttpCompletionOption.ResponseContentRead).Result; // Blocking call!

                    //// In case you didn't get a response, try again as the next call might have dropped in between
                    if (response.IsSuccessStatusCode)
                    {
                        success = true;
                        // Parse the response body. Blocking!
                        Stream dataObjects = response.Content.ReadAsStreamAsync().Result;
                        StreamReader reader = new StreamReader(dataObjects);
                        string responseObj = reader.ReadToEnd();
                        auditDetailData = JsonConvert.DeserializeObject<List<AuditDetailedReport>>(responseObj);
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.BadRequest)
                    {
                        success = true;
                        log.LogError($" Bad Request {(int)response.StatusCode} ({response.ReasonPhrase}). Existing Retries");
                    }
                    else
                    {
                        log.LogError($" Response Failed. Error Message {(int)response.StatusCode} ({response.ReasonPhrase}). Retrying again");
                        log.LogWarning($"Trying Again. Retry {retries}");
                    }
                    //log.LogInformation(reader.ReadToEnd());
                    retries++;
                };
            }
            catch(Exception ex)
            {
                log.LogError($"Error while getting Detailed Audit Data. Error message - {ex.Message}");
            }
            
            return auditDetailData;
        }

        /// <summary>
        /// Fetching the additional properties for Audit log data
        /// </summary>
        /// <param name="auditDetailedReport"></param>
        /// <returns></returns>
        public AuditDetailedReport mapOrUpdateProperties(AuditDetailedReport auditDetailedReport)
        {
            log.LogInformation($"Mapping and Updating properties for Audit record {auditDetailedReport.Id} ");
            try
            {
                if (auditDetailedReport.UserType != null)
                    auditDetailedReport.UserTypeName = getTypeData("UserType", auditDetailedReport.UserType);  
            }
            catch(Exception ex)
            {
                log.LogError($"Error while fetching user type or record type information. Error {ex.Message} for Audit Detail report Id " + auditDetailedReport.Id);
                auditDetailedReport.UserTypeName = "Not Found";
            }

            try
            {
                if (auditDetailedReport.RecordType != null)
                    auditDetailedReport.RecordTypeName = getTypeData("RecordType", auditDetailedReport.RecordType);
            }
            catch (Exception ex)
            {
                log.LogError($"Error while fetching user type or record type information. Error {ex.Message} for Audit Detail report Id " + auditDetailedReport.Id);
                auditDetailedReport.RecordTypeName = "Not Found";
            }

            //// Get the SharePoint related information (Optional - Only required if needed SharePoint related infomrmation)
            try
            {
                if (auditDetailedReport.SiteUrl != null && !auditDetailedReport.SiteUrl.ToLower().Contains("my"))
                {
                    SecureString secpass = new SecureString();
                    foreach (char charpass in SPUserPassword)
                    {
                        secpass.AppendChar(charpass);
                    }

                    if (auditDetailedReport.Site != null && auditDetailedReport.Site != "")
                    {
                        using (var contextSite = new ClientContext(auditDetailedReport.SiteUrl))
                        {
                            contextSite.Credentials = new SharePointOnlineCredentials(SPUserName, secpass);
                            contextSite.Load(contextSite.Web, w => w.Title, w => w.Lists);
                            contextSite.ExecuteQueryAsync().GetAwaiter().GetResult();
                            auditDetailedReport.WebSiteName = contextSite.Web.Title;

                            if (auditDetailedReport.ListId != null && auditDetailedReport.ListId != "")
                            {
                                try
                                {
                                    Guid listId = new Guid(auditDetailedReport.ListId);
                                    List listAudited = contextSite.Web.Lists.GetById(listId);
                                    contextSite.Load(listAudited, l => l.Title);
                                    contextSite.ExecuteQueryAsync().GetAwaiter().GetResult();
                                    auditDetailedReport.ListName = listAudited.Title;
                                }
                                catch (Exception ex)
                                {
                                    log.LogError($"Error while fetching SharePoint List Information. Error {ex.Message} for Audit Detail report Id {auditDetailedReport.Id} for Site url {auditDetailedReport.SiteUrl} ");
                                    auditDetailedReport.ListName = "Not Found or Access Denied";

                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error while fetching SharePoint Information. Error {ex.Message} for Audit Detail report Id {auditDetailedReport.Id} for Site url {auditDetailedReport.SiteUrl} ");
                auditDetailedReport.WebSiteName = "Not Found or Access Denied";
            }

            return auditDetailedReport;
        }

        /// <summary>
        /// Insert the Audit detail data now into Azure Table
        /// </summary>
        /// <param name="auditDetailedReports"></param>
        /// <param name="auditLogDataAnalytics"></param>
        /// <returns></returns>
        public AuditLogAnalyticsDataInfo addDatatoAzureStore(List<AuditDetailedReport> auditDetailedReports, AuditLogAnalyticsDataInfo auditLogDataAnalytics)
        {
            try
            {
                List<string> operations = getAuditOperations();
                List<string> IncludedOperations = new List<string>();
                List<string> ExcludedOperations = new List<string>();
                int includedRecords = 0;
                int excludedRecords = 0;

                // Parse the connection string and return a reference to the storage account.
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(AuditLogStorageConnectionString);

                // Create the table client.
                CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

                // Create the CloudTable object that represents the "Audit Log" table.
                CloudTable auditDataLogTable = tableClient.GetTableReference(AuditLogDataTableName);

                //// Add the Audit log information to Azure Table -- Primary Table to capture Audit log data
                foreach (AuditDetailedReport auditDetailReport in auditDetailedReports)
                {
                    if(operations.Contains(auditDetailReport.Operation, StringComparer.OrdinalIgnoreCase))
                    {
                        includedRecords++;
                        string auditlogpartitionPrefix = AuditLogPartitionKeyPrefix + "_" + getAuditShortDateStringFromUTC(auditDetailReport.CreationTime);
                        string auditlogrowprefix = auditDetailReport.Id;
                        O365AuditLogData auditLogTableData = new O365AuditLogData(auditlogpartitionPrefix, auditlogrowprefix);
                        O365MgmtApiEntities<AuditDetailedReport, O365AuditLogData>.Copy(auditDetailReport, auditLogTableData);
                        auditLogTableData.CreationTime = auditLogTableData.CreationTime.ToLocalTime();
                        // Create the TableOperation object that inserts the customer entity.
                        TableOperation auditLoginsertOperation = TableOperation.InsertOrMerge(auditLogTableData);

                        log.LogInformation($"Inserting Operation - {auditDetailReport.Operation} Id - {auditDetailReport.Id} ");

                        // Execute the insert operation.
                        TableResult result = auditDataLogTable.ExecuteAsync(auditLoginsertOperation).GetAwaiter().GetResult();
                        if (result.HttpStatusCode.ToEnum<HttpStatusCode>() == HttpStatusCode.BadGateway || result.HttpStatusCode.ToEnum<HttpStatusCode>() == HttpStatusCode.BadRequest)
                        {
                            throw new Exception($"Table Update failed for Audit Log data entry. Captured error {result.HttpStatusCode} for Operation - { auditDetailReport.Operation },  Id {auditDetailReport.Id} ");
                        }

                        if (!IncludedOperations.Contains(auditDetailReport.Operation, StringComparer.OrdinalIgnoreCase))
                            IncludedOperations.Add(auditDetailReport.Operation);
                    }
                    else
                    {
                        excludedRecords++;
                        if (!ExcludedOperations.Contains(auditDetailReport.Operation, StringComparer.OrdinalIgnoreCase))
                            ExcludedOperations.Add(auditDetailReport.Operation);
                        log.LogInformation($"Record not entered. Operation - { auditDetailReport.Operation },  Id {auditDetailReport.Id} ");
                    }
                }


                //// Capture the audit log report analytics for each run into the table for future reference -- Secondary Table for capturing analytics of each run
                log.LogInformation($"\r\n Included Records count {includedRecords} and Excluded Records count {excludedRecords}");

                auditLogDataAnalytics.ExcludedOperations = string.Join(",", ExcludedOperations.ToArray());
                auditLogDataAnalytics.IncludedOperations = string.Join(",", IncludedOperations.ToArray());
                auditLogDataAnalytics.IncludedRecords = includedRecords;
                auditLogDataAnalytics.ExcludedRecords = excludedRecords;
                auditLogDataAnalytics.TotalRecords = auditDetailedReports.Count;
                auditLogDataAnalytics.LogOperationStatus = "Completed";
                auditLogDataAnalytics.LogOperationSuccessful = true;
            }
            catch(Exception ex)
            {
                log.LogError(ex.Message);
                auditLogDataAnalytics.LogOperationStatus = "Failed";
                auditLogDataAnalytics.LogOperationSuccessful = false;
            }
            return auditLogDataAnalytics;
        }

        /// <summary>
        /// Initializing the Audit log class
        /// </summary>
        /// <returns></returns>
        public AuditLogAnalyticsDataInfo getInitialAnalyticsInfo()
        {
            DateTime startHour = AuditLogDateTime.AddHours(-3);
            DateTime endHour = AuditLogDateTime.AddHours(-2);

            return new AuditLogAnalyticsDataInfo
            {
                PartitionKey = AuditLogOpspartitionKeyPrefix,
                RowPrefix = getAuditLongDateTimeString(AuditLogDateTime),
                StartHour = startHour,
                EndHour = endHour,
                ExcludedOperations = "",
                IncludedOperations = "",
                IncludedRecords = 0,
                ExcludedRecords = 0,
                TotalRecords = 0,
                RunDate = getAuditShortDateString(AuditLogDateTime),
                RunHour = AuditLogDateTime.ToString("HH:mm"),
                RunFrequency = "Current",
                LogOperationStatus =  "Started",
                LogOperationSuccessful = false
             };
        }

        /// <summary>
        /// Update existing audit log data after addition
        /// </summary>
        /// <param name="auditLogAnalyticsData"></param>
        /// <returns></returns>
        public bool updateAnalyticsDataToTable(AuditLogAnalyticsDataInfo auditLogAnalyticsData)
        {
            // Parse the connection string and return a reference to the storage account.
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(AuditLogStorageConnectionString);

            // Create the table client.
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

            // Create the CloudTable object that represents the "Audit Log" table.
            CloudTable auditDataLogTable = tableClient.GetTableReference(AuditLogDataTableName);

            // Create the CloudTable object that represents the "Audit Log" table.
            CloudTable auditOpsStatusTable = tableClient.GetTableReference(AuditLogAnalyticsTableName);
            O365AuditLogOperations auditOpstableData = new O365AuditLogOperations(auditLogAnalyticsData.PartitionKey, auditLogAnalyticsData.RowPrefix);
            auditOpstableData.ExcludedOperations = auditLogAnalyticsData.ExcludedOperations;
            auditOpstableData.IncludedOperations = auditLogAnalyticsData.IncludedOperations;
            auditOpstableData.IncludedRecords = auditLogAnalyticsData.IncludedRecords;
            auditOpstableData.ExcludedRecords = auditLogAnalyticsData.ExcludedRecords;
            auditOpstableData.TotalRecords = auditLogAnalyticsData.TotalRecords;
            auditOpstableData.RunDate = auditLogAnalyticsData.RunDate;
            auditOpstableData.StartHour = auditLogAnalyticsData.StartHour;
            auditOpstableData.EndHour = auditLogAnalyticsData.EndHour;
            auditOpstableData.RunHour = auditLogAnalyticsData.RunHour;
            auditOpstableData.RunFrequency = auditLogAnalyticsData.RunFrequency;
            auditOpstableData.LogOperationStatus = auditLogAnalyticsData.LogOperationStatus;
            auditOpstableData.LogOperationSuccessful = auditLogAnalyticsData.LogOperationSuccessful;
            TableOperation auditOpsinsertOperation = TableOperation.InsertOrMerge(auditOpstableData);

            log.LogInformation($"Inserting Operations Data");

            // Execute the insert operation.
            TableResult result = auditOpsStatusTable.ExecuteAsync(auditOpsinsertOperation).GetAwaiter().GetResult();
            if (result.HttpStatusCode.ToEnum<HttpStatusCode>() == HttpStatusCode.BadGateway || result.HttpStatusCode.ToEnum<HttpStatusCode>() == HttpStatusCode.BadRequest)
            {
                throw new Exception($"Table Update failed for Audit Log data entry. Captured error {result.HttpStatusCode} for Audit Log Analytics entry ");
            }
            return true;
        }

        #region "Internal Methods"
        private List<string> getAuditOperations()
        {
            List<string> auditOperations = new List<string>();
            StreamReader streamReader = new StreamReader(new MemoryStream(AuditLogResources.AuditOperations), Encoding.UTF8);
            string jsonString = streamReader.ReadToEnd();
            JObject auditJSON = JObject.Parse(jsonString);
            JArray Operations = (JArray)auditJSON["Operations"];
            foreach (string operation in Operations)
            {
                auditOperations.Add(operation);
            }
            return auditOperations;
        }

        private string getTypeData(string Type, string valueToCheck)
        {
            StreamReader streamReader = new StreamReader(new MemoryStream(AuditLogResources.MappingProperties), Encoding.UTF8);
            string jsonString = streamReader.ReadToEnd();
            JObject mappingJson = JObject.Parse(jsonString);
            Dictionary<string, string> typeDict = mappingJson.Value<JObject>(Type).Properties().ToDictionary(k => k.Name, v => v.Value.ToString());
            string TypeValue = "";
            if (typeDict.TryGetValue(valueToCheck, out TypeValue))
                return TypeValue;
            else
                return "";
        }

        public string getAuditShortDateString(DateTime date)
        {
            return date.ToString("yyyy_MM_dd");
        }

        public string getAuditShortDateStringFromUTC(DateTime date)
        {
            TimeZoneInfo aestTimeZone = TimeZoneInfo.FindSystemTimeZoneById("AUS Eastern Standard Time");
            return TimeZoneInfo.ConvertTimeFromUtc(date, aestTimeZone).ToString("yyyy_MM_dd");
        }

        public DateTime getAESTTime()
        {
            TimeZoneInfo aestTimeZone = TimeZoneInfo.FindSystemTimeZoneById("AUS Eastern Standard Time");
            return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, aestTimeZone);
                     
        }

        public string getAuditLongDateTimeString(DateTime date)
        {
            return date.ToString("yyyy_MM_ddTHH_mm");
        }
        #endregion
    }
}
