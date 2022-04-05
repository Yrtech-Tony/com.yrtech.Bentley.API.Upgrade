using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.Common
{
    public class PPTException : Exception
    {
        public PPTException(string message)
            : base(message)
        {
        }
    }
}