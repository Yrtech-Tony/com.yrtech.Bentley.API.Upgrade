using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class AccountDto
    {
        public int UserId { get; set; }
        public string AccountId { get; set;  }
        public string AccountName { get; set; }
        public string AccountNameEn { get; set; }
        public string Password { get; set; }
        public string TelNO { get; set; }
        public string Email { get; set; }
        public string RoleTypeCode { get; set; }
        public List<Area> AreaList { get; set; }
        public List<Shop> ShopList { get; set; }
    }
}