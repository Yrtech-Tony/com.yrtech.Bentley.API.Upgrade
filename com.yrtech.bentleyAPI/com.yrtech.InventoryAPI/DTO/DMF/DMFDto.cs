using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class DMFDto
    {
        public int ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public string Quarters { get; set; }
        public int ActualMonthSaleCount { get; set; }
        public decimal? ActualMonthSaleAmt { get; set; }
        public decimal? ActualAmt { get; set; }
        public decimal? DiffAmt { get; set; }

    }
}