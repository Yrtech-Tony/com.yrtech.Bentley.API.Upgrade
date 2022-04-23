using com.yrtech.InventoryAPI.DTO;
using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace com.yrtech.InventoryAPI.Service
{
    public class MasterService
    {
        Bentley db = new Bentley();
        #region Shop
        public List<ShopDto> ShopSearch(string shopId,string shopCode,string shopName,string shopNameEn)
        {
            if (shopId == null) shopId = "";
            if (shopCode == null) shopCode = "";
            if (shopName == null) shopName = "";
            if (shopNameEn == null) shopNameEn = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@ShopCode", shopCode),
                                                    new SqlParameter("@ShopName", shopName),
                                                    new SqlParameter("@ShopNameEn", shopNameEn)};
            Type t = typeof(ShopDto);
            string sql = "";
             sql = @"SELECT A.*,B.AreaCode,B.AreaName,B.AreaNameEn,B.AreaId
                    FROM Shop A LEFT JOIN Area B ON A.AreaId = B.AreaId
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(shopCode))
            {
                sql += " AND ShopCode = @ShopCode";
            }
            if (!string.IsNullOrEmpty(shopName))
            {
                sql += " AND ShopName = @ShopName";
            }
            if (!string.IsNullOrEmpty(shopNameEn))
            {
                sql += " AND ShopNameEn = @ShopNameEn";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<ShopDto>().ToList();
        }
        public Shop ShopSave(Shop shop)
        {
            Shop findOne = db.Shop.Where(x => (x.ShopId ==shop.ShopId)).FirstOrDefault();
            if (findOne == null)
            {
                shop.InDateTime = DateTime.Now;
                shop.ModifyDateTime = DateTime.Now;
                db.Shop.Add(shop);
            }
            else
            {
                findOne.ShopCode = shop.ShopCode;
                findOne.ShopName = shop.ShopName;
                findOne.ShopNameEn = shop.ShopNameEn;
                findOne.AreaId = shop.AreaId;
                findOne.Budget = shop.Budget;
                findOne.Balance = shop.Balance;
                findOne.City = shop.City;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = shop.ModifyUserId;
                shop = findOne;
            }
            db.SaveChanges();
            return shop;
        }
        public void ShopDelete(string shopId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ShopId", shopId) };
            string sql = @"DELETE Shop WHERE ShopId = @ShopId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region Area
        public List<Area> AreaSearch(string areaId, string areaName, string areaNameEn)
        {
            if (areaId == null) areaId = "";
            if (areaName == null) areaName = "";
            if (areaNameEn == null) areaNameEn = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@AreaId", areaId),
                                                    new SqlParameter("@AreaName", areaName),
                                                    new SqlParameter("@AreaNameEn", areaNameEn)};
            Type t = typeof(Area);
            string sql = "";
            sql = @"SELECT A.* 
                    FROM Area A 
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(areaId))
            {
                sql += " AND AreaId = @AreaId";
            }
            if (!string.IsNullOrEmpty(areaName))
            {
                sql += " AND AreaName = @AreaName";
            }
            if (!string.IsNullOrEmpty(areaNameEn))
            {
                sql += " AND AreaNameEn = @AreaNameEn";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<Area>().ToList();
        }
        public Area AreaSave(Area area)
        {
            Area findOne = db.Area.Where(x => (x.AreaId == area.AreaId)).FirstOrDefault();
            if (findOne == null)
            {
                area.InDateTime = DateTime.Now;
                area.ModifyDateTime = DateTime.Now;
                db.Area.Add(area);
            }
            else
            {
                findOne.AreaCode = area.AreaCode;
                findOne.AreaName = area.AreaName;
                findOne.AreaNameEn = area.AreaNameEn;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = area.ModifyUserId;
                findOne.DTTEmail = area.DTTEmail;
                area = findOne;
            }
            db.SaveChanges();
            return area;
        }
        public void AreaDelete(string areaId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@AreaId", areaId) };
            string sql = @"DELETE Area WHERE AreaId = @AreaId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region HiddenCode
        public List<HiddenCode> HiddenCodeSearch(string hiddenCodeGroup, string hiddenCode,string hiddenCodeName)
        {
            if (hiddenCodeGroup == null) hiddenCodeGroup = "";
            if (hiddenCode == null) hiddenCode = "";
            if (hiddenCodeName == null) hiddenCodeName = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@HiddenCodeGroup", hiddenCodeGroup),
                                                    new SqlParameter("@HiddenCodeId", hiddenCode),
                                                    new SqlParameter("@HiddenCodeName", hiddenCodeName)};
            Type t = typeof(HiddenCode);
            string sql = "";
            sql = @"SELECT * FROM [HiddenCode] WHERE 1=1";
            if (!string.IsNullOrEmpty(hiddenCodeGroup))
            {
                sql += " AND HiddenCodeGroup = @HiddenCodeGroup";
            }
            if (!string.IsNullOrEmpty(hiddenCode))
            {
                sql += " AND HiddenCodeId = @HiddenCodeId";
            }
            if (!string.IsNullOrEmpty(hiddenCodeName))
            {
                sql += " AND HiddenCodeName = @HiddenCodeName";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<HiddenCode>().ToList();
        }
        #endregion
        #region EventType
        public List<EventTypeDto> EventTypeSearch(string eventTypeId, string eventTypeName,string eventTypeNameEn,bool? showStatus)
        {
            if (eventTypeId == null) eventTypeId = "";
            if (eventTypeName == null) eventTypeName = "";
            if (eventTypeNameEn == null) eventTypeNameEn = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@EventTypeId", eventTypeId),
                                                    new SqlParameter("@EventTypeName", eventTypeName),
                                                    new SqlParameter("@EventTypeNameEn", eventTypeNameEn)};
            Type t = typeof(EventTypeDto);
            string sql = "";
            sql = @"SELECT A.*,B.HiddenCodeName AS EventModeName,B.HiddenCodeNameEn AS EventModeNameEn,C.AreaCode,C.AreaName,C.AreaNameEn 
                    FROM EventType A INNER JOIN HiddenCode B ON A.EventMode = B.HiddenCodeId AND B.HiddenCodeGroup = 'EventMode' 
                                     LEFT JOIN Area C ON A.AreaId = C.AreaId
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(eventTypeId))
            {
                sql += " AND EventTypeId = @EventTypeId";
            }
            if (!string.IsNullOrEmpty(eventTypeName))
            {
                sql += " AND EventTypeName = @EventTypeName";
            }
            if (!string.IsNullOrEmpty(eventTypeNameEn))
            {
                sql += " AND EventTypeNameEn = @EventTypeNameEn";
            }
            if (showStatus.HasValue)
            {
                para = para.Concat(new SqlParameter[] { new SqlParameter("@ShowStatus", showStatus) }).ToArray();
                sql += " AND ShowStatus = @ShowStatus";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<EventTypeDto>().ToList();
        }
        public EventType EventTypeSave(EventType eventType)
        {

            EventType findOne = db.EventType.Where(x => (x.EventTypeId == eventType.EventTypeId)).FirstOrDefault();
            if (findOne == null)
            {
                eventType.InDateTime = DateTime.Now;
                eventType.ModifyDateTime = DateTime.Now;
                db.EventType.Add(eventType);
            }
            else
            {
                findOne.ApprovalMaxAmt = eventType.ApprovalMaxAmt;
                findOne.AreaId = eventType.AreaId;
                findOne.EventMode = eventType.EventMode;
                findOne.EventTypeName = eventType.EventTypeName;
                findOne.EventTypeNameEn = eventType.EventTypeNameEn;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = eventType.ModifyUserId;
                findOne.ShowStatus = eventType.ShowStatus;
                eventType = findOne;
            }
            db.SaveChanges();
            return eventType;
        }
        public void EventTypeDelete(string eventTypeId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@EventTypeId", eventTypeId) };
            string sql = @"DELETE EventType WHERE EventTypeId = @EventTypeId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region UserInfo
        public List<UserInfoDto> UserInfoSearch(string userId, string accountId, string accountName,string shopCode,string shopName,string email,string shopId,string areaId)
        {
            if (userId == null) userId = "";
            if (accountId == null) accountId = "";
            if (accountName == null) accountName = "";
            if (shopCode == null) shopCode = "";
            if (shopName == null) shopName = "";
            if (email == null) email = "";
            if (shopId == null) shopId = "";
            if (areaId == null) areaId = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@UserId", userId),
                                                    new SqlParameter("@AccountId", accountId),
                                                    new SqlParameter("@AccountName", accountName),
                                                    new SqlParameter("@ShopCode", shopCode),
                                                    new SqlParameter("@ShopName", shopName),
                                                    new SqlParameter("@Email", email),
                                                    new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@AreaId", areaId)};
            Type t = typeof(UserInfoDto);
            string sql = "";
            sql = @"SELECT [UserId]
                    ,[AccountId]
                    ,[Password]
                    ,[AccountName]
                    ,[AccountNameEn]
                    ,[TelNO]
                    ,[Email]
                    ,RoleTypeCode
                    ,CASE WHEN [RoleTypeCode] = 'SYSADMIN' THEN '管理员'
			            WHEN [RoleTypeCode] = 'AREA' THEN '区域经理'
			            WHEN [RoleTypeCode] = 'BMC' THEN 'BMC'
			            WHEN [RoleTypeCode] = 'Shop' THEN '经销商'
			            ELSE '' END AS RoleTypeName
                     ,CASE WHEN [RoleTypeCode] IN ('SYSADMIN','BMC','AREA') THEN 0
		                WHEN [RoleTypeCode] IN ('Shop') THEN A.ShopId
		                ELSE '' END AS ShopId
                    ,CASE WHEN [RoleTypeCode] IN ('SYSADMIN','BMC','AREA') THEN ''
		                WHEN [RoleTypeCode] IN ('Shop') THEN B.ShopName
		                ELSE '' END AS ShopName
                    ,CASE WHEN [RoleTypeCode] IN ('SYSADMIN','BMC','AREA') THEN ''
		                WHEN [RoleTypeCode] IN ('Shop') THEN B.ShopNameEn
		                ELSE '' END AS ShopNameEn
	                ,CASE WHEN [RoleTypeCode] IN ('SYSADMIN','BMC') THEN ''
		                WHEN [RoleTypeCode] IN ('Shop') 
		                THEN (SELECT TOP 1 AreaName FROM Shop X INNER JOIN Area Y ON X.AreaId = Y.AreaId AND X.ShopId = B.ShopId) 
		                WHEN [RoleTypeCode] IN ('AREA') THEN C.AreaName
		                ELSE '' END AS AreaName
		            ,CASE WHEN [RoleTypeCode] IN ('SYSADMIN','BMC') THEN ''
		                WHEN [RoleTypeCode] IN ('Shop') 
		                THEN (SELECT TOP 1 AreaNameEn FROM Shop X INNER JOIN Area Y ON X.AreaId = Y.AreaId AND X.ShopId = B.ShopId) 
		                WHEN [RoleTypeCode] IN ('AREA') THEN C.AreaNameEn
		                ELSE '' END AS AreaNameEn
                    ,A.[AreaId]
                    ,A.[InUserId]
                    ,A.[InDateTime]
                    ,A.[ModifyUserId]
                    ,A.[ModifyDateTime]
                FROM UserInfo A LEFT JOIN Shop B ON A.ShopId = B.ShopId
				                LEFT JOIN Area C ON A.AreaId = C.AreaId
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(userId))
            {
                sql += " AND UserId = @UserId";
            }
            if (!string.IsNullOrEmpty(accountId))
            {
                sql += " AND AccountId = @AccountId";
            }
            if (!string.IsNullOrEmpty(accountName))
            {
                sql += " AND AccountName = @AccountName";
            }
            if (!string.IsNullOrEmpty(shopCode))
            {
                sql += " AND ShopCode = @ShopCode";
            }
            if (!string.IsNullOrEmpty(shopName))
            {
                sql += " AND ShopName = @ShopName";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId AND A.RoleTypeCode = 'SHOP'";
            }
            if (!string.IsNullOrEmpty(areaId))
            {
                sql += " AND A.AreaId = @AreaId AND A.RoleTypeCode = 'AREA'";
            }
            if (!string.IsNullOrEmpty(email))
            {
                sql += " AND Email  LIKE '%'+@Email+'%'";
            }
            sql += " ORDER BY AccountId";
            return db.Database.SqlQuery(t, sql, para).Cast<UserInfoDto>().ToList();
        }

        public UserInfo UserInfoSave(UserInfo userInfo)
        {
            UserInfo findOne = db.UserInfo.Where(x => (x.UserId == userInfo.UserId)).FirstOrDefault();
            if (findOne == null)
            {
                userInfo.InDateTime = DateTime.Now;
                userInfo.ModifyDateTime = DateTime.Now;
                db.UserInfo.Add(userInfo);
            }
            else
            {
                findOne.AccountId = userInfo.AccountId;
                findOne.AccountName = userInfo.AccountName;
                findOne.AccountNameEn = userInfo.AccountNameEn;
                findOne.AreaId = userInfo.AreaId;
                findOne.Email = userInfo.Email;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = userInfo.ModifyUserId;
                findOne.Password = userInfo.Password;
                findOne.RoleTypeCode = userInfo.RoleTypeCode;
                findOne.ShopId = userInfo.ShopId;
                findOne.TelNO = userInfo.TelNO;
                userInfo = findOne;
            }
            db.SaveChanges();
            return userInfo;
        }
        public void UserInfoPasswordChange(string password,string userId)
        {
            int userIdInt = Convert.ToInt32(userId);
            UserInfo findOne = db.UserInfo.Where(x => (x.UserId == userIdInt)).FirstOrDefault();
            //if (findOne == null)
            //{
            //    userInfo.InDateTime = DateTime.Now;
            //    userInfo.ModifyDateTime = DateTime.Now;
            //    db.UserInfo.Add(userInfo);
            //}
            //else
            //{
            findOne.Password = password;
            findOne.ModifyDateTime = DateTime.Now;
            //}
            db.SaveChanges();
        }
        public void UserInfoDelete(string userId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@UserId", userId) };
            string sql = @"DELETE UserInfo WHERE UserId = @UserId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region CoopFundType
        public List<CoopFundType> CoopFundTypeSearch(string coopFundTypeId, string coopFundTypeCode,string coopFundTypeName, string coopFundTypeNameEn, bool? showChk,string modeType)
        {
            if (coopFundTypeId == null) coopFundTypeId = "";
            if (coopFundTypeCode == null) coopFundTypeCode = "";
            if (coopFundTypeName == null) coopFundTypeName = "";
            if (coopFundTypeNameEn == null) coopFundTypeNameEn = "";
            if (modeType == null) modeType = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@CoopFundTypeId", coopFundTypeId),
                                                    new SqlParameter("@CoopFundTypeName", coopFundTypeName),
                                                    new SqlParameter("@CoopFundTypeCode", coopFundTypeCode),
                                                    new SqlParameter("@CoopFundTypeNameEn", coopFundTypeNameEn),
                                                    new SqlParameter("@ModeType", modeType)};
            Type t = typeof(CoopFundType);
            string sql = "";
            sql = @"SELECT *
                    FROM CoopFundType 
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(coopFundTypeId))
            {
                sql += " AND CoopFundTypeId = @CoopFundTypeId";
            }
            if (!string.IsNullOrEmpty(coopFundTypeCode))
            {
                sql += " AND CoopFundTypeCode = @CoopFundTypeCode";
            }
            if (!string.IsNullOrEmpty(coopFundTypeName))
            {
                sql += " AND CoopFundTypeName = @CoopFundTypeName";
            }
            if (!string.IsNullOrEmpty(coopFundTypeNameEn))
            {
                sql += " AND CoopFundTypeNameEn = @CoopFundTypeNameEn";
            }
            if (!string.IsNullOrEmpty(modeType))
            {
                sql += " AND ModeType = @ModeType";
            }
            if (showChk.HasValue)
            {
                para = para.Concat(new SqlParameter[] { new SqlParameter("@ShowChk", showChk) }).ToArray();
                sql += " AND ShowChk = @ShowChk";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<CoopFundType>().ToList();
        }
        public CoopFundType CoopFundTypeSave(CoopFundType coopFundType)
        {

            CoopFundType findOne = db.CoopFundType.Where(x => (x.CoopFundTypeId == coopFundType.CoopFundTypeId)).FirstOrDefault();
            if (findOne == null)
            {
                coopFundType.InDateTime = DateTime.Now;
                coopFundType.ModifyDateTime = DateTime.Now;
                db.CoopFundType.Add(coopFundType);
            }
            else
            {
                findOne.CoopFundTypeName = coopFundType.CoopFundTypeName;
                findOne.CoopFundTypeNameEn = coopFundType.CoopFundTypeNameEn;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = coopFundType.ModifyUserId;
                findOne.ShowChk = coopFundType.ShowChk;
                coopFundType = findOne;
            }
            db.SaveChanges();
            return coopFundType;
        }
        public void CoopFundTypeDelete(string coopFundTypeId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@CoopFundTypeId", coopFundTypeId) };
            string sql = @"DELETE CoopFundType WHERE CoopFundTypeId = @CoopFundTypeId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion 

    }
}