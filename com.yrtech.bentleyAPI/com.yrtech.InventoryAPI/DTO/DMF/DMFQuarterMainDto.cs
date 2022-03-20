using com.yrtech.bentley.DAL;
using System;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class DMFQuarterMainDto
    {
       public List<DMFDto> DMFQuarterList { get; set; }
       public List<DMFDetailDto> DMFDetailList { get; set; }
    }
}