using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionBefore4WeeksCoopFundDto
    {
        public int MarketActionId { get; set; }
        public int SeqNO { get; set; }
        public string CoopFundCode { get; set; }
        public Nullable<decimal> CoopFundAmt { get; set; }
        public Nullable<bool> CoopFund_DMFChk { get; set; }
        public string CoopFundDesc { get; set; }
        public Nullable<System.DateTime> StartDate { get; set; }
        public Nullable<System.DateTime> EndDate { get; set; }
        public Nullable<int> TotalDays { get; set; }
        public Nullable<decimal> AmtPerDay { get; set; }
        public string CoopFundTypeDesc { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }

    }
}