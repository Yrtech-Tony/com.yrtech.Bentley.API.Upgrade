using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionAfter7ActualExpenseDto
    {
        public int MarketActionId { get; set; }
        public int SeqNO { get; set; }
        public string Item { get; set; }
        public string Descs { get; set; }
        public Nullable<decimal> UnitPrice { get; set; }
        public Nullable<decimal> Counts { get; set; }
        public Nullable<decimal> Total { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}