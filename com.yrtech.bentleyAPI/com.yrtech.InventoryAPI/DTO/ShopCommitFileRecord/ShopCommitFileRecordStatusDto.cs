using System;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class ShopCommitFileRecordStatusDto
    {
        public int ShopId { get; set; }
        public int FileId { get; set; }
        public int FileCount { get; set; }
    }
}