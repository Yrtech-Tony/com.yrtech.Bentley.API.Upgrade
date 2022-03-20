using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionBefore4WeeksMainDto
    {
        public int MarketActionId { get; set; }
        public string TarketModelCode { get; set; }// 主推车型
        public MarketActionBefore4Weeks MarketActionBefore4Weeks { get; set; } // 活动计划信息
        public List<MarketActionBefore4WeeksActivityProcess> ActivityProcess { get; set; } // 活动流程信息
        public List<MarketActionPic> MarketActionBefore4WeeksPicList_OffLine { get; set; } // 线下照片
        public List<MarketActionPic> MarketActionBefore4WeeksPicList_OnLine { get; set; }// 线上照片

    }
}