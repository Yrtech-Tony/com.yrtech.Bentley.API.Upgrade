using System;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class OtherPropertyTypeDto
    {
        public int ProjectId { get; set; }
        public string OtherType { get; set; }
    }
}