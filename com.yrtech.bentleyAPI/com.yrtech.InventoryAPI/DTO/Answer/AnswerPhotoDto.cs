using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.DTO
{
    [Serializable]
    public class AnswerPhotoDto
    {
        public int AnswerId { get; set; }
        public int PhotoId { get; set; }
        public int ProjectId { get; set; }
        public int ShopId { get; set; }
        public string CheckCode { get; set; }
        public string PhotoNameServer  { get; set; }
        public string PhotoName { get; set; }
        public string PhotoUrl { get; set; }
        
    }
}