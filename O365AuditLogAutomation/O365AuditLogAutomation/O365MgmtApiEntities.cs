using System;
using System.Collections.Generic;
using Microsoft.WindowsAzure.Storage.Table;

/// <summary>
/// Entity classes with class signatures
/// </summary>
namespace O365AuditLogAutomation
{
    /// <summary>
    /// Dynamic copy class to copy properties to destination class based on source class. Cool isn't it :)
    /// </summary>
    /// <typeparam name="TParent"></typeparam>
    /// <typeparam name="TChild"></typeparam>
    class O365MgmtApiEntities<TParent, TChild> where TParent : class
                                            where TChild : class
    {
        public static void Copy(TParent parent, TChild child)
        {
            var parentProperties = parent.GetType().GetProperties();
            var childProperties = child.GetType().GetProperties();
            
            foreach (var parentProperty in parentProperties)
            {
                foreach (var childProperty in childProperties)
                {
                    if (parentProperty.Name == childProperty.Name && parentProperty.PropertyType == childProperty.PropertyType)
                    {
                        childProperty.SetValue(child, parentProperty.GetValue(parent));
                        break;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Analytics info class for tracking Audit log operation done by this Azure Function
    /// </summary>
    public class AuditLogAnalyticsDataInfo
    {
       public string PartitionKey { get; set; }
       public string RowPrefix { get; set; }
       public DateTime StartHour { get; set; }
       public DateTime EndHour { get; set; }
       public string ExcludedOperations { get; set; }
       public string IncludedOperations { get; set; }
       public int IncludedRecords { get; set; }
       public int ExcludedRecords { get; set; }
       public int TotalRecords { get; set; }
       public string RunDate { get; set; }
       public string RunHour { get; set; }
       public string RunFrequency { get; set; }
       public string LogOperationStatus { get; set; }
       public bool LogOperationSuccessful { get; set; }
    }

    /// <summary>
    /// Entity class for the initial Audit call
    /// </summary>
    public class AuditInitialDataObject
    {
        public List<AuditInitialReport> AuditInitialDataObj { get; set; }
        public string AuditNextPageUri { get; set; }
    }

    /// <summary>
    /// Initial Audit data class
    /// </summary>
    public class AuditInitialReport
    {
        public string ContentUri { get; set; }
        public string ContentId { get; set; }
        public string ContentType { get; set; }
        public string ContentCreated { get; set; }
        public string ContentExpiration { get; set; }
    }

    /// <summary>
    /// Audit data Detail class with all the actual tracking information
    /// </summary>
    public class AuditDetailedReport
    {
        public DateTime CreationTime { get; set; }
        public string Id { get; set; }
        public string Operation { get; set; }
        public string Workload { get; set; }
        public string ObjectId { get; set; }
        public string UserType { get; set; }
        public string UserTypeName { get; set; }
        public string RecordType { get; set; }
        public string RecordTypeName { get; set; }
        public string UserId { get; set; }
        public string EventSource { get; set; }
        public string SiteUrl { get; set; }
        public string Site { get; set; }
        public string WebId { get; set; }
        public string WebSiteName { get; set; }
        public string ListId { get; set; }
        public string ListName { get; set; }
        public string ListItemUniqueId { get; set; }
        public string ItemName { get; set; }
        public string ItemType { get; set; }
        public string SourceFileExtension { get; set; }
        public string SourceFileName { get; set; }
        public string SourceRelativeUrl { get; set; }
        public string UserAgent { get; set; }
        public string EventData { get; set; }
        public string UniqueSharingId { get; set; }
        public string OrganizationId { get; set; }
        public string UserKey { get; set; }
        public string ClientIP { get; set; }
        public string CorrelationId { get; set; }
    }

    /// <summary>
    /// Entity calss for Azure Table storage
    /// </summary>
    public class O365AuditLogData : TableEntity
    {
        public O365AuditLogData(string partitionKey, string rowkey)
        {
            this.PartitionKey = partitionKey;
            this.RowKey = rowkey;
        }

        public O365AuditLogData(){ }

        public DateTime CreationTime { get; set; }
        public string Id { get; set; }
        public string Operation { get; set; }
        public string Workload { get; set; }
        public string ObjectId { get; set; }
        public string UserType { get; set; }
        public string UserTypeName { get; set; }
        public string RecordType { get; set; }
        public string RecordTypeName { get; set; }
        public string UserId { get; set; }
        public string EventSource { get; set; }
        public string SiteUrl { get; set; }
        public string Site { get; set; }
        public string WebId { get; set; }
        public string WebSiteName { get; set; }
        public string ListId { get; set; }
        public string ListName { get; set; }
        public string ListItemUniqueId { get; set; }
        public string ItemName { get; set; }
        public string ItemType { get; set; }
        public string SourceFileExtension { get; set; }
        public string SourceFileName { get; set; }
        public string SourceRelativeUrl { get; set; }
        public string UserAgent { get; set; }
        public string EventData { get; set; }
        public string TargetUserOrGroupType { get; set; }
        public string TargetUserOrGroupName { get; set; }
        public string TargetExtUserName { get; set; }
        public string UniqueSharingId { get; set; }
        public string OrganizationId { get; set; }
        public string UserKey { get; set; }
        public string ClientIP { get; set; }
        public string CorrelationId { get; set; }
    }

    /// <summary>
    /// Audit Log reporting Analytics class and operations
    /// </summary>
    public class O365AuditLogOperations : TableEntity
    {
        public O365AuditLogOperations(string partitionKey, string rowkey)
        {
            this.PartitionKey = partitionKey;
            this.RowKey = rowkey;
        }
        public O365AuditLogOperations() { }

        public DateTime StartHour { get; set; }
        public DateTime EndHour { get; set; }
        public string ExcludedOperations { get; set; }
        public string IncludedOperations { get; set; }
        public int IncludedRecords { get; set; }
        public int ExcludedRecords { get; set; }
        public int TotalRecords { get; set; }
        public string RunDate { get; set; }
        public string RunHour { get; set; }
        public string RunFrequency { get; set; }
        public string LogOperationStatus { get; set; }
        public bool LogOperationSuccessful { get; set; }
    }


    public class AuditOperations
    {
        Array Operations { get; set; }
    }
}
