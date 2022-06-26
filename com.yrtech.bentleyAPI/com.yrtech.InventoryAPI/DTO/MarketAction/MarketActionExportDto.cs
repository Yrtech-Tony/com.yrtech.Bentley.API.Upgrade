using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionExportDto
    {
        public int MarketActionId { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public string ActionName { get; set; }
        public string ActionCode { get; set; }
        public string AreaName { get; set; }
        public int? ExpectLeadsCount { get; set; }
        public decimal? ActivityBudget { get; set; }
        public Nullable<int> EventTypeId { get; set; }
        public string EventTypeName { get; set; }
        public string EventTypeNameEn { get; set; }
        public string MarketActionStatusCode { get; set; }
        public string MarketActionStatusName { get; set; }
        public string MarketActionStatusNameEn { get; set; }
        public string MarketActionTargetModelCode { get; set; }
        public string MarketActionTargetModelName { get; set; }
        public string MarketActionTargetModelNameEn { get; set; }
        public Nullable<System.DateTime> StartDate { get; set; }
        public Nullable<System.DateTime> EndDate { get; set; }
        public string ActionPlace { get; set; }
        public string EventModeName { get; set; }
        public string DTTApproveStatus_Plan { get; set; }
        public string DTTApproveStatus_Report { get; set; }
        public string KeyVisionApprovalName { get; set; }
        public Nullable<bool> ExpenseAccount { get; set; }
       public MarketActionBefore4Weeks MarketActionBefore4Weeks{ get; set; }
        public decimal? ActualExpenseSum { get; set; }

       public MarketActionAfter7  MarketActionAfter7 { get; set; }
       public MarketActionLeadsCountDto LeadsCount { get; set; }



    }
}