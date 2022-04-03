using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.Common
{
    public class APIResult
    {
        public bool Status { get; set; }
        public string Body { get; set; }

        public static APIResult OK(string body)
        {
            return new APIResult()
            {
                Body = body,
                Status = true
            };
        }
        public static APIResult ERROR(string msg)
        {
            return new APIResult()
            {
                Body = msg,
                Status = false
            };
        }
    }
}