using System;
using System.Collections.Generic;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class MarketActionStatusCountDto
    {

        public Nullable<int> Before4WeeksCount { get; set; }
        public Nullable<int> Before3Count { get; set; }
        public Nullable<int> After2Count { get; set; }
        public Nullable<int> After7Count { get; set; }
        public Nullable<int> After30Count { get; set; }
    }
}