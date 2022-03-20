using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MonthSaleDto
    {
        public int MonthSaleId { get; set; }
        public string YearMonth { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public string ActualSaleCount { get; set; }
        public string ActualSaleAmt { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}