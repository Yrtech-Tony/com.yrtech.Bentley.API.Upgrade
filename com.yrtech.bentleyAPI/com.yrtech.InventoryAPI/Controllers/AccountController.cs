using System;
using System.Collections.Generic;
using System.Web.Http;
using com.yrtech.InventoryAPI.Common;
using com.yrtech.InventoryAPI.Service;
using com.yrtech.InventoryAPI.DTO;
using com.yrtech.bentley;

namespace com.yrtech.InventoryAPI.Controllers
{
    [RoutePrefix("bentley/api")]
    public class AccountController : BaseController
    {
        AccountService accountService = new AccountService();

        [HttpPost]
        [Route("Account/Login")]
        public APIResult Login(AccountDto account)
        {
            try
            {
                List<AccountDto> accountlist = accountService.Login(account.AccountId, account.Password);
                if (accountlist != null && accountlist.Count != 0)
                {
                    string roleTypeCode = accountlist[0].RoleTypeCode;
                    string userId = accountlist[0].UserId.ToString();
                    accountlist[0].AreaList = accountService.GetAreaByRole(userId, roleTypeCode);
                    accountlist[0].ShopList = accountService.GetShopByRole(userId, roleTypeCode);
                    return new APIResult() { Status = true, Body = CommonHelper.Encode(accountlist) };
                }
                else
                {
                    return new APIResult() { Status = false, Body = "用户不存在密码不匹配或账号已过期" };
                }
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
      
    }
}
