using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionAfter2LeadsReportDto
    {
        public int MarketActionId { get; set; }
        public string ActionName { get; set; }//活动名称
        public int ShopId { get; set; }
        public string ShopName { get; set; }
        public string ShopNameEn { get; set; }
        public int SeqNO { get; set; }
        public string CustomerName { get; set; }//顾客姓名
        public string TelNO { get; set; } //顾客电话
        public string BPNO { get; set; }//BPNO
        public Nullable<bool> OwnerCheck { get; set; }//是否车主
        public string OwnerCheckName { get; set; }//是否车主显示:是或者否
        public Nullable<bool> TestDriverCheck { get; set; }//是否试驾
        public string TestDriverCheckName { get; set; }//显示是否
        public Nullable<bool> LeadsCheck { get; set; }//是否为线索
        public string LeadsCheckName { get; set; }//显示是否
        public string InterestedModel { get; set; } // 感兴趣车型代码
        public string InterestedModelName { get; set; }// 感兴趣车型名称
        public string InterestedModelNameEn { get; set; }// 感兴趣车型名称
        public Nullable<bool> DealCheck { get; set; }//是否成交
        public string DealCheckName { get; set; }//显示是否
        public string DealModel { get; set; }//成交车型代码
        public string DealModelName { get; set; }//成交车型名称
        public string DealModelNameEn { get; set; }//成交车型名称
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}