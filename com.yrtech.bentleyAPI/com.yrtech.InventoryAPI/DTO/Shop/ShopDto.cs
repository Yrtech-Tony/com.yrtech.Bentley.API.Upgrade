using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.DTO
{
    public class ShopDto
    {
        public int ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public string City { get; set; }
        public string Email { get; set; }
        public Nullable<decimal> Budget { get; set; }
        public Nullable<decimal> Balance { get; set; }
        public Nullable<int> AreaId { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
        public string AreaCode { get; set; }
        public string AreaName { get; set; }
        public string AreaNameEn { get; set; }
    }
}