using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionLeadsCountDto
    {

        public Nullable<int> LeadOwnerCount { get; set; }
        public Nullable<int> LeadPCCount { get; set; }
        public Nullable<int> TestDriverOwnerCount { get; set; }
        public Nullable<int> TestDriverPCCount { get; set; }
        public Nullable<int> ActualOrderOwnerCount { get; set; }
        public Nullable<int> ActualOrderPCCount { get; set; }
        public Nullable<decimal> ExpenseTotalAmt { get; set; }
    }
}