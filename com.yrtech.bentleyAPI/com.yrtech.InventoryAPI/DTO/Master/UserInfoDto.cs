using System;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class UserInfoDto
    {
        public int UserId { get; set; }
        public string AccountId { get; set; }
        public string Password { get; set; }
        public string OldPassword { get; set; }
        public string AccountName { get; set; }
        public string AccountNameEn { get; set; }
        public string TelNO { get; set; }
        public string Email { get; set; }
        public string RoleTypeCode { get; set; }
        public string RoleTypeName { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public Nullable<int> AreaId { get; set; }
        public string AreaName { get; set; }
        public string AreaNameEn { get; set; }
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}