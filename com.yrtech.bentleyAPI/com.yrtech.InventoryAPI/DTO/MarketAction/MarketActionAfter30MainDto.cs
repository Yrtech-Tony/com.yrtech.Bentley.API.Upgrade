using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionAfter30MainDto
    {
        public int MarketActionId { get; set; }
        public List<MarketActionAfter2LeadsReport> LeadsReportList { get; set; }      
    }
}