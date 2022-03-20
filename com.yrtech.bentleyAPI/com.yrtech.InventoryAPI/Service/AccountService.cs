using com.yrtech.InventoryAPI.DTO;
using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace com.yrtech.InventoryAPI.Service
{
    public class AccountService
    {
       Bentley db = new Bentley();
        /// <summary>
        /// 登录
        /// </summary>
        /// <param name="accountId"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public List<AccountDto> Login(string accountId, string password)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@AccountId", accountId),
                                                       new SqlParameter("@Password",password)};
            Type t = typeof(AccountDto);
            string sql = @"SELECT TOP 1 UserId,AccountId,AccountName,AccountNameEn,RoleTypeCode,Email FROM UserInfo
                            WHERE AccountId = @AccountId AND [Password] = @Password";
            return db.Database.SqlQuery(t, sql, para).Cast<AccountDto>().ToList();
        }
        public List<Shop> GetShopByRole(string userId, string roleTypeCode)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@UserId", userId),
                                                       new SqlParameter("@RoleTypeCode",roleTypeCode)};
            Type t = typeof(Shop);
            string sql = "";
            if (roleTypeCode.ToUpper() == "SYSADMIN" || roleTypeCode.ToUpper() == "MARKET" || roleTypeCode.ToUpper() == "BMC")
            {
                sql += @"SELECT * 
                        FROM Shop A";
            }
            else if (roleTypeCode.ToUpper() == "AREA")
            {
                sql += @"SELECT A.* 
                        FROM Shop A INNER JOIN UserInfo B ON A.AreaId = B.AreaId
                        WHERE B.UserId = @UserId";
            }
            else if (roleTypeCode.ToUpper() == "SHOP")
            {
                sql += @"SELECT A.* 
                        FROM Shop A INNER JOIN UserInfo B ON A.ShopId = B.ShopId
                        WHERE B.UserId = @UserId";
            }
            sql += @" ORDER BY ShopName";
            return db.Database.SqlQuery(t, sql, para).Cast<Shop>().ToList();
        }

        public List<Area> GetAreaByRole(string userId, string roleTypeCode)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@UserId", userId),
                                                       new SqlParameter("@RoleTypeCode",roleTypeCode)};
            Type t = typeof(Area);
            string sql = "";
            if (roleTypeCode.ToUpper() == "SYSADMIN" || roleTypeCode.ToUpper() == "MARKET"|| roleTypeCode.ToUpper()=="BMC")
            {
                sql += @"SELECT * 
                        FROM Area A";
            }
            else if (roleTypeCode.ToUpper() == "AREA")
            {
                sql += @"SELECT A.* 
                        FROM Area A INNER JOIN UserInfo B ON A.AreaId = B.AreaId
                        WHERE B.UserId = @UserId";
            }
            else if (roleTypeCode.ToUpper() == "SHOP")
            {
                sql += @"SELECT C.* 
                        FROM Shop A INNER JOIN UserInfo B ON A.ShopId = B.ShopId
                                    INNER JOIN Area C ON A.AreaId = C.AreaId
                        WHERE B.UserId = @UserId";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<Area>().ToList();
        }

    }
}