using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class ExpenseAccountDto
    {
        public int ExpenseAccountId { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public Nullable<int> DMFItemId { get; set; }
        public string DMFItemName { get; set; }
        public string DMFItemNameEn { get; set; }
        public Nullable<int> MarketActionId { get; set; }
        public string ActionName{get;set;}
        public string ExpenseAmt { get; set; }
        public string ApplyStatus { get; set; }
        public string ApprovalReason { get; set; }
        public string ReplyStatus { get; set; }
        public string ReplyReason { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
    }
}