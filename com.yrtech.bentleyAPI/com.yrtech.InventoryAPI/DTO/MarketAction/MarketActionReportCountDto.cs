using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionReportCountDto
    {

        public Nullable<int> PlanBugetUnCommit { get; set; }//未提交活动计划市场花费
        public Nullable<int> PlanCoopFundUnCommit { get; set; }// 未提交活动计划市场基金
        public Nullable<int> LeadsUnCommit { get; set; }//未提交线索报告
        public Nullable<int> ReportBugetUnCommit { get; set; }//未提交活动报告市场花费
        public Nullable<int> ReportCoopFundUnCommit { get; set; }//未提交活动报告市场基金
    }
}