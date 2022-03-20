using System;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class EventTypeDto
    {
        public int EventTypeId { get; set; }
        public string EventTypeName { get; set; }
        public string EventTypeNameEn { get; set; }
        public string EventMode { get; set; }
        public string EventModeName { get; set; }
        public string EventModeNameEn { get; set; }
        public Nullable<int> AreaId { get; set; }
        public string AreaCode { get; set; }
        public string AreaName { get; set; }
        public string AreaNameEn { get; set; }
        public Nullable<decimal> ApprovalMaxAmt { get; set; }
        public bool ShowStatus { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}