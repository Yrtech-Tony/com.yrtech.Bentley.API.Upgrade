using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class DMFDetailDto
    {
        public int DMFDetailId { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public string DMFItemName { get; set; }
        public string DMFItemNameEn { get; set; }
        public Nullable<int> DMFItemId { get; set; }
        public string Budget { get; set; }
        public string AcutalAmt { get; set; }
        public string Remark { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
    }
}