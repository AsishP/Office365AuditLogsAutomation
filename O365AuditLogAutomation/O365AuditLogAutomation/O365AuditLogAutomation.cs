using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Extensions.Logging;

/// <summary>
/// Main Audit Log Data Controller class
/// </summary>
namespace O365AuditLogAutomation
{
    public static class O365AuditLogAutomation
    {
        [FunctionName("O365AuditLogAutomation")]

        //// Set the timer in the timer trigger to run - More Info at here - https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer#c-example
        //// Note: this Function App is using Azure Funciton 1.x because of SharePoint CSOM dependency but can be used on Azure Function 2.x
        public static void Run([TimerTrigger("0 0 */1 * * *")]TimerInfo myTimer, ILogger log, TraceWriter logWriter)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            logWriter.Info($"C# Timer trigger function executed at: {DateTime.Now}");
            #region "Variables"

            //// Set the Variables. For local debugging, use localsettings.json (see below) and for Azure Function, use Configuration in Platform Features
            ///  for eg  - local.settings.json
            ///  Note : Please replace all the strings with < > with the open bracket and close brackets
            //{
            //    "IsEncrypted": false,
            //      "Values": {
            //        "AzureWebJobsStorage": "UseDevelopmentStorage=true",
            //        "AzureWebJobsDashboard": "UseDevelopmentStorage=true",
            //        "StorageConnectionString": "<connection string for Azure storage account in .net>",
            //         "FUNCTIONS_WORKER_RUNTIME": "dotnet",
            //         "AuditLogDataTable": "<Azure Table Name>",
            //         "AuditLogAnalyticsTable": "<O365 Audit log capture table for this function>",
            //          "AuditLogCSVExportLocation": "<Azure>",
            //          "SPUserName": "<SharePoint Admin User Name>",
            //          "SPUserPassword": "<SharePoint Admin Password>",
            //          "AuditLogDataTablePrefix": "<Partition Key Prefix for easy retrival>",
            //           "AuditLogOpsTablePrefix": "<Analytics Partition Key Prefic for easy retrival>",
            //            "TenantID": "<Tenant ID for the App>",
            //             "AuditLogAuthUrl": "https://manage.office.com",
            //             "AzureADAppID": "<Azure App ID>",
            //             "AzureADAppSecret": "<Azure App secret>"
            //      }
            //}

            string TenantID = System.Environment.GetEnvironmentVariable("TenantID");
            string authString = "https://login.windows.net/" + TenantID;
            string SPServiceUrl = "https://manage.office.com/api/v1.0/" + TenantID + "/activity/feed/subscriptions/content";
            string authUrl = System.Environment.GetEnvironmentVariable("AuditLogAuthUrl");
            string clientId = System.Environment.GetEnvironmentVariable("AzureADAppID");
            string clientSecret = System.Environment.GetEnvironmentVariable("AzureADAppSecret");
            #endregion

            try
            {
                //// **** Get the Authentication Token ****
                var authenticationContext = new AuthenticationContext(authString, false);
                //// Config for OAuth client credentials 
                ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
                AuthenticationResult authenticationResult = null;
                Task runTask = Task.Run(async () => authenticationResult = await authenticationContext.AcquireTokenAsync(authUrl, clientCred));
                runTask.Wait();
                string token = authenticationResult.AccessToken;

                O365MgmtAPIDataService dataService = new O365MgmtAPIDataService(token, log);
                AuditLogAnalyticsDataInfo auditLogAnalyticsDataInfo = dataService.getInitialAnalyticsInfo();

                if(dataService.updateAnalyticsDataToTable(auditLogAnalyticsDataInfo))
                {
                    //// Get the time zone of the destination tenant. The date and time used by Audit log service is UTC format.
                    TimeZoneInfo aestTimeZone = TimeZoneInfo.FindSystemTimeZoneById("AUS Eastern Standard Time");
                    DateTime startHourUTC = TimeZoneInfo.ConvertTimeToUtc(auditLogAnalyticsDataInfo.StartHour, aestTimeZone);
                    DateTime endHourUTC = TimeZoneInfo.ConvertTimeToUtc(auditLogAnalyticsDataInfo.EndHour, aestTimeZone);
                    string startDateString = startHourUTC.ToUniversalTime().ToString("yyyy-MM-ddTHH:00");
                    string endDateString = endHourUTC.ToString("yyyy-MM-ddTHH:00");

                    //// ****************** Step 1: Start the audit log gathering process *****************
                    log.LogInformation($"getting Data from {startDateString} to {endDateString}");
                    //// Here I am fetching SharePoint events.
                    string urlParameters = $"?contentType=Audit.SharePoint&startTime={startDateString}&endTime={endDateString}";
                    
                    //// Initialize the audit log data information
                    AuditInitialDataObject auditInitialDataObject = new AuditInitialDataObject();
                    List<AuditDetailedReport> auditDetailReportsFinal = new List<AuditDetailedReport>();

                    //// Loop through the detail URI information provided by the initial data call till there is no next page to be read
                    do
                    {
                        auditInitialDataObject = dataService.getAuditInitalData(SPServiceUrl, urlParameters);
                        if (auditInitialDataObject.AuditNextPageUri != "")
                            urlParameters = "?" + auditInitialDataObject.AuditNextPageUri.Split('?')[1];
                        List<AuditInitialReport> auditInitialReports = auditInitialDataObject.AuditInitialDataObj;
                        
                        //// set batch size = 200. Note : Above 500 the speed decreases drastically as per my tests.
                        int maxCalls = 200;
                        int count = 0;
                        
                        //// **************** Step 2: For each of the calls call the detailed data fetch api in batches. Here the batch = 200 ***********************
                        Parallel.ForEach(auditInitialReports, new ParallelOptions { MaxDegreeOfParallelism = maxCalls }, (auditInitialReport) =>
                        {
                            int loopCount = count++;
                            log.LogInformation("Looking at request " + loopCount);
                            List<AuditDetailedReport> auditDetailReports = dataService.getAuditDetailData(auditInitialReport.ContentUri);
                            log.LogInformation("Got Audit Detail Reports of " + auditDetailReports.Count + " for loop number " + loopCount);
                            foreach (AuditDetailedReport auditDetailReport in auditDetailReports)
                            {
                                auditDetailReportsFinal.Add(auditDetailReport);
                            }
                        });
                    } while (auditInitialDataObject.AuditNextPageUri != "");
                    log.LogInformation("Final Audit Detail Reports" + auditDetailReportsFinal.Count);

                    //// *************** Step 3 : Update additional properties to the Audit log data ***************************
                    int maxAuditUpdateCalls = 200;
                    Parallel.ForEach(auditDetailReportsFinal, new ParallelOptions { MaxDegreeOfParallelism = maxAuditUpdateCalls }, (auditDetailReport) =>
                    {
                        auditDetailReport = dataService.mapOrUpdateProperties(auditDetailReport);
                    });

                    //// **************** Step 4 : Add Data to Azure Table ****************
                    auditLogAnalyticsDataInfo = dataService.addDatatoAzureStore(auditDetailReportsFinal, auditLogAnalyticsDataInfo);

                    //// **************** Step 5 : Update the Report analytics for each Audit log run *******************
                    dataService.updateAnalyticsDataToTable(auditLogAnalyticsDataInfo);

                }
            }
            catch (Exception ex)
            {
                log.LogError(ex.Message);
            }
        }
    }
}
