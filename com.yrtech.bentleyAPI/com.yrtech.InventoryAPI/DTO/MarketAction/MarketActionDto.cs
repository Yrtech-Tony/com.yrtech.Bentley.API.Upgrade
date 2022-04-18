using System;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionDto
    {
        public int MarketActionId { get; set; }
        public Nullable<int> ShopId { get; set; }
        public string ShopCode { get; set; }//经销商代码
        public string ShopName { get; set; }//经销商名称
        public string ShopNameEn { get; set; }//经销商名称-英文
        public string ActionName { get; set; }// 活动名称
        public string ActionCode { get; set; }// 活动代码
        public int? ExpectLeadsCount { get; set; }// 预期线索数量
        public decimal? ActivityBudget { get; set; }// 活动预算
       // public decimal? ActivityBudgetMax{ get; set; }// 活动预算
        public Nullable<int> EventTypeId { get; set; }// 活动类型Id
        public string EventTypeName { get; set; }// 活动类型名称
        public string EventTypeNameEn { get; set; }// 活动类型名称-英文
        public Nullable<int> EventModeId { get; set; }// 活动方式Id
        public string EventModeName { get; set; }// 活动方式名称
        public string EventModeNameEn { get; set; }// 活动方式名称-英文
        public string MarketActionStatusCode { get; set; }// 活动状态代码
        public string MarketActionStatusName { get; set; }// 活动状态名称
        public string MarketActionStatusNameEn { get; set; }// 活动状态名称-英文
        public string MarketActionTargetModelCode { get; set; }// 主推车型代码
        public string MarketActionTargetModelName { get; set; }// 主推车型名称
        public string MarketActionTargetModelNameEn { get; set; }// 主推车型名称-英文
        public Nullable<System.DateTime> StartDate { get; set; }// 开始时间
        public Nullable<System.DateTime> EndDate { get; set; }//结束时间
        public string ActionPlace { get; set; }// 活动场地
        public Nullable<bool> ExpenseAccount { get; set; }// 是否费用报销
        // 市场活动的提交状态参考如下返回值
        //Commited：已提交，UnCommitTime：未到时间，UnCommit：显示进度
        public string Before4Weeks { get; set; }// 活动计划状态
        public string After2Days { get; set; }// 线索报告状态
        public string After7Days { get; set; }// 活动报告状态
        public Nullable<int> InUserId { get; set; }
        public Nullable<System.DateTime> InDateTime { get; set; }
        public Nullable<int> ModifyUserId { get; set; }
        public Nullable<System.DateTime> ModifyDateTime { get; set; }
    }
}