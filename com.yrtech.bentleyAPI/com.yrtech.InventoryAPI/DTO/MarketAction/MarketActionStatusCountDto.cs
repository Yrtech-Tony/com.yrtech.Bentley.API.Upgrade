using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionStatusCountDto
    {
        public Nullable<int> MarketActionCount { get; set; } // 活动总数
        public Nullable<int> Plan_CommitCount { get; set; }// 活动计划_已提交
        public Nullable<int> Plan_4WeekNotCommit { get; set; } // 未提交_4周内
        public Nullable<int> Plan_2WeekNotCommit { get; set; } // 未提交_2周内
        public Nullable<int> Report_CommitCount { get; set; }// 活动报告_已提交
        public Nullable<int> Report_2WeekNotCommit { get; set; } // 未提交 2周内
        public Nullable<int> Report_1WeekNotCommit { get; set; } // 未提交 1周内
        public Nullable<decimal> Plan_CommitCountRate { get; set; }  // 活动计划_已提交 百分比
        public Nullable<decimal> Plan_4WeekNotCommitRate { get; set; } // 未提交_4周内 百分比
        public Nullable<decimal> Plan_2WeekNotCommitRate { get; set; } // 未提交_2周内 百分比
        public Nullable<decimal> Report_CommitCountRate { get; set; } // 活动报告_已提交 百分比
        public Nullable<decimal> Report_2WeekNotCommitRate { get; set; }// 未提交 2周内 百分比
        public Nullable<decimal> Report_1WeekNotCommitRate { get; set; }// 未提交 1周内 百分比
    }
}