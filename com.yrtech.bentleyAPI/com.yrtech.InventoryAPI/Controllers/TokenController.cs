using System.Web.Http;
using com.yrtech.InventoryAPI.Service;
using com.yrtech.InventoryAPI.Common;
using System.Collections.Generic;
using System;
using com.yrtech.bentley;
using com.yrtech.InventoryAPI.DTO;
using com.yrtech.bentley.DAL;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace com.yrtech.InventoryAPI.Controllers
{
    [RoutePrefix("token/api")]
    public class TokenController : BaseController
    {
        [HttpGet]
        [Route("Token/TokenCreate")]
        public APIResult TokenCreate(string encryptString)
        {
            try
            {
               
               // return new APIResult() { Status = true, Body = TokenHelper.EncryptDES(encryptString) };
                return new APIResult() { Status = true, Body = encryptString };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Token/TokenCheck")]
        public APIResult TokenCheck(TokenDto token)
        {
            try
            { 

                return new APIResult() { Status = true, Body = TokenHelper.DecryptDES(token.TokenString) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
    }
    public class TokenDto
    {
        public string EncryptString { get; set; }
        public string TokenString { get; set; }
    }
}
