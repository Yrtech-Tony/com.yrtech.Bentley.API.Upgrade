using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class AnswerDto
    {
        public long AnswerId { get; set; }
        public int ProjectId { get; set; }
        public int ShopId { get; set; }
        public string ShopCode { get; set; }
        public string ShopName { get; set; }
        public string CheckCode { get; set; }
        public Nullable<int> CheckTypeId { get; set; }
        public string CheckTypeName { get; set; }
        public string Remark { get; set; }
        public string OtherProperty { get; set; }
        public Nullable<decimal> Score { get; set; }
        public string AddCheck { get; set; }
        public string ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
        public string InUserID { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public List<AnswerPhotoDto> answerPhotoList { get; set; }
    }
}