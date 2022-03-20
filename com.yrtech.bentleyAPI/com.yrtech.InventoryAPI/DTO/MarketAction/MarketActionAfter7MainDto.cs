using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionAfter7MainDto
    {
        public int MarketActionId { get; set; }
        public MarketActionAfter7 MarketActionAfter7 { get; set; }
        public MarketActionLeadsCountDto LeadsCount { get; set; }
        public List<MarketActionPic> MarketActionAfter7PicList_OffLine { get; set; }
        public List<MarketActionPic> MarketActionAfter7PicList_OnLine { get; set; }
        public List<MarketActionAfter7ActualProcess> ActualProcess { get; set; }
    }
}