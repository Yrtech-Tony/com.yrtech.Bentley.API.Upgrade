using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.DTO
{
    public class DTTApproveDto
    {
        public int DTTApproveId { get; set; }
        public string DTTType { get; set; }// 类型：活动计划：1，活动报告：2
        public Nullable<int> MarketActionId { get; set; }
        public string DTTApproveCode { get; set; }//审批状态代码：1 ：待审批，2：通过，3：修改
        public string DTTApproveName { get; set; }//审批状态显示文字：1 ：待审批，2：通过，3：修改
        public string DTTApproveNameEn { get; set; }//审批状态显示英文：1 ：待审批，2：通过，3：修改
        public string DTTApproveDesc { get; set; }// 审批意见
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> InUserId { get; set; }
    }
}