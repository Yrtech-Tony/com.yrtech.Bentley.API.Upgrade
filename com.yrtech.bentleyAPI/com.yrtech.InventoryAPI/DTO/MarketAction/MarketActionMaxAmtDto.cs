using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionMaxAmtDto
    {

        public decimal MarketActionBudgetMax { get; set; }//市场活动预算历史最大值
        public decimal Before4WeeksBudgetMax { get; set; }// 活动计划预算历史最大值
        public decimal Before4WeeksDMFSumMax { get; set; }//活动计划市场基金金额合计历史最大值
        public decimal After7BudgetMax { get; set; }// 活动报告预算历史最大值
        public decimal After7DMFSumMax { get; set; }//活动报告市场基金金额合计历史最大值
    }
}