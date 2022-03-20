using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class ShopCommitFileRecordListDto
    {
        public List<ShopDto> ShopList { get; set; }
        public List<CommitFile> CommitFileList { get; set; }
        public List<ShopCommitFileRecordStatusDto> ShopCommitFileRecordStatusList { get; set; }
    }
}