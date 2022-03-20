using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionBefore3MainDto
    {
        public int MarketActionId { get; set; }
        public Nullable<decimal> BugetDetailSumAmt { get; set; }
        public List<MarketActionBefore3BugetDetailDto> BugetDetailListDto { get; set; }
       
    }
}