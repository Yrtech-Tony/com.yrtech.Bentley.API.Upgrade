using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionStatusCountDto
    {

        public Nullable<int> Before4WeeksNotCommit { get; set; }//未提交活动计划
        public Nullable<int> Before4WeeksWaitForChange { get; set; }// 未通过审批活动计划
        public Nullable<int> After7NotCommit { get; set; }//未提交活动报告
        public Nullable<int> After7WaitForChange { get; set; }//未通过审批活动报告
    }
}