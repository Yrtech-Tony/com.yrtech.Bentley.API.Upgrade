using com.yrtech.InventoryAPI.DTO;
using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace com.yrtech.InventoryAPI.Service
{
    public class MarketActionService
    {
        Bentley db = new Bentley();
        #region Common
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        public List<MarketActionDto> MarketActionSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, bool? expenseAccountChk)
        {
            if (actionName == null) actionName = "";
            if (year == null) year = "";
            if (month == null) month = "";
            if (marketActionStatusCode == null) marketActionStatusCode = "";
            if (shopId == null) shopId = "";
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ActionName", actionName),
                                                        new SqlParameter("@Year", year),
                                                        new SqlParameter("@Month", month),
                                                        new SqlParameter("@MarketActionStatusCode", marketActionStatusCode),
                                                        new SqlParameter("@ShopId", shopId),
                                                        new SqlParameter("@EventTypeId", eventTypeId)};
            Type t = typeof(MarketActionDto);
            string sql = "";
            sql += @"SELECT A.MarketActionId,A.ShopId,B.ShopCode,B.ShopName,B.ShopNameEn,A.ActionCode,A.ActionName
		                    ,A.EventTypeId,C.EventTypeName,C.EventTypeNameEn
                            ,(SELECT AreaName FROM Area WHERE AreaId = B.AreaId) AS AreaName
		                    ,(SELECT CAST(EventMode AS INT) FROM EventType WHERE EventTypeId = A.EventTypeId) AS EventModeId
		                    ,(SELECT HiddenCodeName FROM EventType X INNER JOIN HiddenCode Y ON  Y.HiddenCodeGroup='EventMode' AND X.EventMode = Y.HiddenCodeId 
											        WHERE X.EventTypeId =A.EventTypeId ) AS EventModeName
						   ,(SELECT HiddenCodeNameEn FROM EventType X INNER JOIN HiddenCode Y ON  Y.HiddenCodeGroup='EventMode'  AND X.EventMode = Y.HiddenCodeId 
											        WHERE X.EventTypeId =A.EventTypeId ) AS EventModeNameEn
		                    ,A.ActivityBudget,A.ExpectLeadsCount,A.StartDate,A.EndDate,A.ActionPlace
                            ,A.ExpenseAccount,A.InUserId,A.InDateTime,A.ModifyUserId,A.ModifyDateTime
		                    ,A.MarketActionStatusCode,D.HiddenCodeName AS MarketActionStatusName,D.HiddenCodeNameEn AS MarketActionStatusNameEn
		                    ,A.MarketActionTargetModelCode,E.HiddenCodeName AS MarketActionTargetModelName,E.HiddenCodeNameEn AS MarketActionTargetModelNameEn
		                    ,CASE 
                                  WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=1 AND DTTApproveCode=2) 
			                      THEN 'Approved'
                                  WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=1 AND DTTApproveCode=3) 
			                      THEN 'WaitForChange'
                                  WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=1 AND DTTApproveCode=1) 
			                      THEN 'Commited'
			                      WHEN  GETDATE()>DATEADD(DD,-28,A.StartDate) AND EXISTS(SELECT 1 FROM MarketActionBefore4Weeks WHERE MarketActionId = A.MarketActionId)
			                      THEN (SELECT CAST(ISNULL(ProcessPercent,0) AS VARCHAR) FROM MarketActionBefore4Weeks WHERE MarketActionId = A.MarketActionId) 
                                  WHEN GETDATE()>DATEADD(DD,-28,A.StartDate)AND NOT EXISTS(SELECT 1 FROM MarketActionBefore4Weeks WHERE MarketActionId = A.MarketActionId)
                                  THEN '0.00'
			                      ELSE 'UnCommit'
	                        END AS 	Before4Weeks
	                       
	                        ,CASE WHEN EXISTS(SELECT 1 FROM MarketActionAfter2LeadsReport WHERE MarketActionId = A.MarketActionId) 
			                      THEN 'Commited'
			                      WHEN NOT EXISTS(SELECT 1 FROM MarketActionAfter2LeadsReport WHERE MarketActionId = A.MarketActionId) 
				                       AND GETDATE()>DATEADD(DD,7,A.StartDate)
			                      THEN 'UnCommitTime'
			                      ELSE 'UnCommit'
	                        END AS 	After2Days
	                         ,CASE WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=2 AND DTTApproveCode=2) 
			                      THEN 'Approved'
                                  WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=2 AND DTTApproveCode=3) 
			                      THEN 'WaitForChange'
                                  WHEN EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=2 AND DTTApproveCode=1) 
			                      THEN 'Commited'
			                      WHEN  GETDATE()>DATEADD(DD,14,A.StartDate) AND EXISTS(SELECT 1 FROM MarketActionAfter7 WHERE MarketActionId = A.MarketActionId)
			                      THEN (SELECT CAST(ISNULL(ProcessPercent,0) AS VARCHAR) FROM MarketActionAfter7 WHERE MarketActionId = A.MarketActionId) 
                                  WHEN  GETDATE()>DATEADD(DD,14,A.StartDate) AND NOT EXISTS(SELECT 1 FROM MarketActionAfter7 WHERE MarketActionId = A.MarketActionId)
                                  THEN '0.00 '
			                      ELSE 'UnCommit'
	                        END AS 	After7Days	
	                       
                    FROM MarketAction A LEFT JOIN Shop B ON A.ShopId = B.ShopId
					                    LEFT JOIN EventType C ON A.EventTypeId = C.EventTypeId
					                    LEFT JOIN HiddenCode D ON A.MarketActionStatusCode = D.HiddenCodeId AND D.HiddenCodeGroup = 'MarketActionStatus'
					                    LEFT JOIN HiddenCode E ON A.MarketActionTargetModelCode  = E.HiddenCodeId AND E.HiddenCodeGroup = 'TargetModels'
                    WHERE 1=1  ";
            if (!string.IsNullOrEmpty(actionName))
            {
                sql += " AND A.ActionName LIKE '%'+@ActionName+'%'";
            }
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(A.StartDate)= @Year";
            }
            if (!string.IsNullOrEmpty(month))
            {
                sql += " AND Month(A.StartDate)= @Month";
            }
            if (!string.IsNullOrEmpty(marketActionStatusCode))
            {
                sql += " AND A.MarketActionStatusCode =@MarketActionStatusCode";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId =@ShopId";
            }
            if (!string.IsNullOrEmpty(eventTypeId))
            {
                sql += " AND A.EventTypeId =@EventTypeId";
            }
            if (expenseAccountChk.HasValue)
            {
                para = para.Concat(new SqlParameter[] { new SqlParameter("@ExpenseAccountChk", expenseAccountChk) }).ToArray();
                sql += " AND A.ExpenseAccount = @ExpenseAccountChk";
            }
            sql += " ORDER BY A.StartDate DESC";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionDto>().ToList();
        }
        public List<MarketActionPlanDto> MarketActionPlanSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId)
        {
            if (actionName == null) actionName = "";
            if (year == null) year = "";
            if (month == null) month = "";
            if (marketActionStatusCode == null) marketActionStatusCode = "";
            if (shopId == null) shopId = "";
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ActionName", actionName),
                                                        new SqlParameter("@Year", year),
                                                        new SqlParameter("@Month", month),
                                                        new SqlParameter("@MarketActionStatusCode", marketActionStatusCode),
                                                        new SqlParameter("@ShopId", shopId),
                                                        new SqlParameter("@EventTypeId", eventTypeId)};
            Type t = typeof(MarketActionPlanDto);
            string sql = "";
            sql += @"SELECT A.MarketActionId,A.ShopId,B.ShopCode,B.ShopName,
                            ISNULL((SELECT TOP 1 AreaName FROM Area WHERE AreaId = B.AreaId),'') AS AreaName,
                            A.ActionCode,A.ActionName
		                    ,A.EventTypeId,C.EventTypeName,
                            ISNULL((SELECT TOP 1 HiddenCodeName FROM HiddenCode WHERE HiddenCodeGroup = 'EventMode' AND HiddenCodeId = C.EventMode),'') AS EventModeName
                            ,A.ActivityBudget,A.ExpectLeadsCount,A.StartDate,A.EndDate
                            ,CASE WHEN Month(A.StartDate) IN (1,2,3) THEN 'Q1'
                                  WHEN Month(A.StartDate) IN (4,5,6) THEN 'Q2'
                                  WHEN Month(A.StartDate) IN (7,8,9) THEN 'Q3'
                                  WHEN Month(A.StartDate) IN (10,11,12) THEN 'Q4'
                             ELSE ''
                             END AS Quarter
                            ,CASE WHEN A.ExpenseAccount=1 THEN 'Y' ELSE '' END AS ExpenseAccount
		                    ,A.MarketActionStatusCode,D.HiddenCodeName AS MarketActionStatusName,D.HiddenCodeNameEn AS MarketActionStatusNameEn
		                    ,A.MarketActionTargetModelCode,E.HiddenCodeName AS MarketActionTargetModelName,E.HiddenCodeNameEn AS MarketActionTargetModelNameEn
                    FROM MarketAction A LEFT JOIN Shop B ON A.ShopId = B.ShopId
					                    LEFT JOIN EventType C ON A.EventTypeId = C.EventTypeId
					                    LEFT JOIN HiddenCode D ON A.MarketActionStatusCode = D.HiddenCodeId AND D.HiddenCodeGroup = 'MarketActionStatus'
					                    LEFT JOIN HiddenCode E ON A.MarketActionTargetModelCode  = E.HiddenCodeId AND E.HiddenCodeGroup = 'TargetModels'
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(actionName))
            {
                sql += " AND A.ActionName LIKE '%'+@ActionName+'%'";
            }
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(A.StartDate)= @Year";
            }
            if (!string.IsNullOrEmpty(month))
            {
                sql += " AND Month(A.StartDate)= @Month";
            }
            if (!string.IsNullOrEmpty(marketActionStatusCode))
            {
                sql += " AND A.MarketActionStatusCode =@MarketActionStatusCode";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId =@ShopId";
            }
            if (!string.IsNullOrEmpty(eventTypeId))
            {
                sql += " AND A.EventTypeId =@EventTypeId";
            }
            sql += " ORDER BY A.StartDate DESC";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionPlanDto>().ToList();
        }
        // 查询未取消的市场活动
        public List<MarketAction> MarketActionNotCancelSearch(string eventTypeId)
        {
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@EventTypeId", eventTypeId) };
            Type t = typeof(MarketAction);
            string sql = "";
            sql += @"SELECT *
                    FROM MarketAction A 
                    WHERE A.MarketActionStatusCode<>2 AND  A.ExpenseAccount=1
                    AND EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=2 AND DTTApproveCode = 2)";

            if (!string.IsNullOrEmpty(eventTypeId))
            {
                sql += " AND A.EventTypeId =@EventTypeId";
            }
            sql += " ORDER BY A.StartDate DESC";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketAction>().ToList();
        }
        // 查询所有市场活动预算金额最大值
        public decimal MarketActionBudgetMaxSearch(string shopId)
        {
            if (shopId == null) shopId = "";
             SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ShopId", shopId) };
            Type t = typeof(decimal);
            string sql = "";
            sql += @"SELECT ISNULL(Max(ActivityBudget),0) 
                    FROM MarketAction A 
                    WHERE A.MarketActionStatusCode<>2 ";
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND ShopId = @ShopId";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<decimal>().FirstOrDefault();
        }
        public List<MarketActionDto> MarketActionSearchById(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionDto);
            string sql = "";
            sql += @"SELECT A.MarketActionId,A.ShopId,B.ShopCode,B.ShopName,B.ShopNameEn,A.ActionCode,A.ActionName
		                    ,A.EventTypeId,C.EventTypeName,C.EventTypeNameEn
		                    ,(SELECT CAST(EventMode AS INT) FROM EventType WHERE EventTypeId = A.EventTypeId) AS EventModeId
		                    ,(SELECT HiddenCodeName FROM EventType X INNER JOIN HiddenCode Y ON  Y.HiddenCodeGroup='EventMode' AND X.EventMode = Y.HiddenCodeId 
											        WHERE X.EventTypeId =A.EventTypeId ) AS EventModeName
						   ,(SELECT HiddenCodeNameEn FROM EventType X INNER JOIN HiddenCode Y ON  Y.HiddenCodeGroup='EventMode'  AND X.EventMode = Y.HiddenCodeId 
											        WHERE X.EventTypeId =A.EventTypeId ) AS EventModeNameEn
		                    ,A.ActivityBudget,A.ExpectLeadsCount,A.StartDate,A.EndDate,A.ActionPlace
                            ,A.ExpenseAccount,A.InUserId,A.InDateTime,A.ModifyUserId,A.ModifyDateTime
		                    ,A.MarketActionStatusCode,D.HiddenCodeName AS MarketActionStatusName,D.HiddenCodeNameEn AS MarketActionStatusNameEn
		                    ,A.MarketActionTargetModelCode,E.HiddenCodeName AS MarketActionTargetModelName,E.HiddenCodeNameEn AS MarketActionTargetModelNameEn

                    FROM MarketAction A LEFT JOIN Shop B ON A.ShopId = B.ShopId
					                    LEFT JOIN EventType C ON A.EventTypeId = C.EventTypeId
					                    LEFT JOIN HiddenCode D ON A.MarketActionStatusCode = D.HiddenCodeId AND D.HiddenCodeGroup = 'MarketActionStatus'
					                    LEFT JOIN HiddenCode E ON A.MarketActionTargetModelCode  = E.HiddenCodeId AND E.HiddenCodeGroup = 'TargetModels'
                    WHERE 1=1 AND MarketActionId = @MarketActionId";

            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionDto>().ToList();
        }
        public void MarketActionSave(MarketAction marketAction)
        {
            MarketAction findOne = db.MarketAction.Where(x => (x.MarketActionId == marketAction.MarketActionId)).FirstOrDefault();
            if (findOne == null)
            {
                marketAction.InDateTime = DateTime.Now;
                marketAction.ModifyDateTime = DateTime.Now;
                db.MarketAction.Add(marketAction);
            }
            else
            {
                findOne.ActionCode = marketAction.ActionCode;
                findOne.ActionName = marketAction.ActionName;
                findOne.ActionPlace = marketAction.ActionPlace;
                findOne.ActivityBudget = marketAction.ActivityBudget;
                findOne.ExpectLeadsCount = marketAction.ExpectLeadsCount;
                findOne.EndDate = marketAction.EndDate;
                findOne.StartDate = marketAction.StartDate;
                findOne.EventTypeId = marketAction.EventTypeId;
                findOne.ExpenseAccount = marketAction.ExpenseAccount;
                findOne.MarketActionStatusCode = marketAction.MarketActionStatusCode;
                findOne.MarketActionTargetModelCode = marketAction.MarketActionTargetModelCode;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = marketAction.ModifyUserId;
                findOne.ShopId = marketAction.ShopId;
            }

            db.SaveChanges();
        }
        public void MarketActionDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"
                        DELETE MarketActionBefore4Weeks WHERE MarketActionId = @MarketActionId 
                        DELETE MarketActionBefore4WeeksActivityProcess WHERE MarketActionId = @MarketActionId 
                        DELETE MarketActionAfter2LeadsReport WHERE MarketActionId = @MarketActionId 
                        DELETE MarketActionAfter7 WHERE MarketActionId = @MarketActionId  
                        DELETE MarketActionAfter7ActualExpense WHERE MarketActionId = @MarketActionId 
                        DELETE MarketActionAfter7ActualProcess WHERE MarketActionId = @MarketActionId 
                        DELETE MarketAction WHERE MarketActionId = @MarketActionId 
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #region 市场活动照片
        public List<MarketActionPic> MarketActionPicSearch(string marketActionId, string picType)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId),
                                                        new SqlParameter("@PicType", picType)};
            Type t = typeof(MarketActionPic);
            string sql = "";
            sql += @"SELECT A.* 
                    FROM [MarketActionPic] A 
                    WHERE  PicType Like @PicType+'%' AND MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionPic>().ToList();
        }
        public void MarketActionPicSave(MarketActionPic marketActionPic)
        {
            MarketActionPic findOneMax = db.MarketActionPic.Where(x => (x.MarketActionId == marketActionPic.MarketActionId && x.PicType == marketActionPic.PicType)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionPic.SeqNO = 1;
            }
            else
            {
                marketActionPic.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionPic.InDateTime = DateTime.Now;
            db.MarketActionPic.Add(marketActionPic);
            db.SaveChanges();
        }
        public void MarketActionPicDelete(string marketActionId, string picType, string seqNO)
        {
            if (picType == null) picType = "";
            if (seqNO == null) seqNO = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId)
                                                        ,new SqlParameter("@PicType", picType)
                                                        ,new SqlParameter("@SeqNO", seqNO) };
            string sql = @"DELETE MarketActionPic WHERE MarketActionId = @MarketActionId 
                        ";
            if (!string.IsNullOrEmpty(picType))
            {
                sql += @" AND PicType LIKE @PicType+'%' ";
            }
            if (!string.IsNullOrEmpty(seqNO))
            {
                sql += @" AND SeqNO = @SeqNO";
            }

            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region Before 4 weeks
        public List<MarketActionBefore4Weeks> MarketActionBefore4WeeksSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionBefore4Weeks);
            string sql = "";
            sql += @"SELECT *
                    FROM [MarketActionBefore4Weeks] A 
                    WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionBefore4Weeks>().ToList();
        }
        public void MarketActionBefore4WeeksSave(MarketActionBefore4Weeks marketActionBefore4Weeks)
        {
            MarketActionBefore4Weeks findOne = db.MarketActionBefore4Weeks.Where(x => (x.MarketActionId == marketActionBefore4Weeks.MarketActionId)).FirstOrDefault();
            if (findOne == null)
            {
                marketActionBefore4Weeks.InDateTime = DateTime.Now;
                marketActionBefore4Weeks.ModifyDateTime = DateTime.Now;
                db.MarketActionBefore4Weeks.Add(marketActionBefore4Weeks);
            }
            else
            {
                findOne.ActivityBackground = marketActionBefore4Weeks.ActivityBackground;
                findOne.ActivityDesc = marketActionBefore4Weeks.ActivityDesc;
                findOne.ActivityObjective = marketActionBefore4Weeks.ActivityObjective;
                findOne.CoopFundSumAmt = marketActionBefore4Weeks.CoopFundSumAmt;
                findOne.People_DCPIDCount = marketActionBefore4Weeks.People_DCPIDCount;
                findOne.People_InvitationCarOwnerCount = marketActionBefore4Weeks.People_InvitationCarOwnerCount;
                findOne.People_InvitationDepositorCount = marketActionBefore4Weeks.People_InvitationDepositorCount;
                findOne.People_InvitationOtherCount = marketActionBefore4Weeks.People_InvitationOtherCount;
                findOne.People_InvitationPotentialCount = marketActionBefore4Weeks.People_InvitationPotentialCount;
                //findOne.People_InvitationTotalCount = marketActionBefore4Weeks.People_InvitationTotalCount;
                findOne.People_NewLeadsThisYearCount = marketActionBefore4Weeks.People_NewLeadsThisYearCount;
                if (marketActionBefore4Weeks.People_InvitationCarOwnerCount == null) marketActionBefore4Weeks.People_InvitationCarOwnerCount = 0;
                if (marketActionBefore4Weeks.People_InvitationDepositorCount == null) marketActionBefore4Weeks.People_InvitationDepositorCount = 0;
                if (marketActionBefore4Weeks.People_InvitationPotentialCount == null) marketActionBefore4Weeks.People_InvitationPotentialCount = 0;
                if (marketActionBefore4Weeks.People_InvitationOtherCount == null) marketActionBefore4Weeks.People_InvitationOtherCount = 0;
                findOne.People_InvitationTotalCount = marketActionBefore4Weeks.People_InvitationCarOwnerCount
                                                       + marketActionBefore4Weeks.People_InvitationDepositorCount
                                                       + marketActionBefore4Weeks.People_InvitationPotentialCount
                                                       + marketActionBefore4Weeks.People_InvitationOtherCount;
                findOne.People_ParticipantsCount = marketActionBefore4Weeks.People_ParticipantsCount;
                findOne.ProcessPercent = marketActionBefore4Weeks.ProcessPercent;
                findOne.Vehide_Model = marketActionBefore4Weeks.Vehide_Model;
                findOne.Vehide_Qty = marketActionBefore4Weeks.Vehide_Qty;
                findOne.Vehide_Usage = marketActionBefore4Weeks.Vehide_Usage;
                findOne.Platform_ExposureForm = marketActionBefore4Weeks.Platform_ExposureForm;
                findOne.Platform_Media = marketActionBefore4Weeks.Platform_Media;
                findOne.PerformPlan = marketActionBefore4Weeks.PerformPlan;
                findOne.PhotographerIntro = marketActionBefore4Weeks.PhotographerIntro;
                findOne.KeyVisionApprovalCode = marketActionBefore4Weeks.KeyVisionApprovalCode;
                findOne.KeyVisionApprovalDesc = marketActionBefore4Weeks.KeyVisionApprovalDesc;
                findOne.KeyVisionDesc = marketActionBefore4Weeks.KeyVisionDesc;
                if (marketActionBefore4Weeks.KeyVisionPic != "https://yrsurvey.oss-cn-beijing.aliyuncs.com/Bentley/fail2.png")
                    findOne.KeyVisionPic = marketActionBefore4Weeks.KeyVisionPic;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = marketActionBefore4Weeks.ModifyUserId;
                findOne.PlatformReason = marketActionBefore4Weeks.PlatformReason;
            }

            db.SaveChanges();
        }
        public List<MarketActionBefore4WeeksActivityProcess> MarketActionBefore4WeeksActivityProcessSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionBefore4WeeksActivityProcess);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionBefore4WeeksActivityProcess] WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionBefore4WeeksActivityProcess>().ToList();
        }
        public void MarketActionBefore4WeeksActivityProcessSave(MarketActionBefore4WeeksActivityProcess marketActionBefore4WeeksActivityProcess)
        {
            //if (marketActionBefore4WeeksActivityProcess.SeqNO == 0)
            //{
            MarketActionBefore4WeeksActivityProcess findOneMax = db.MarketActionBefore4WeeksActivityProcess.Where(x => (x.MarketActionId == marketActionBefore4WeeksActivityProcess.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionBefore4WeeksActivityProcess.SeqNO = 1;
            }
            else
            {
                marketActionBefore4WeeksActivityProcess.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionBefore4WeeksActivityProcess.InDateTime = DateTime.Now;
            marketActionBefore4WeeksActivityProcess.ModifyDateTime = DateTime.Now;
            db.MarketActionBefore4WeeksActivityProcess.Add(marketActionBefore4WeeksActivityProcess);

            //}
            //else
            //{
            //MarketActionBefore4WeeksActivityProcess findOne = db.MarketActionBefore4WeeksActivityProcess.Where(x => (x.MarketActionId == marketActionBefore4WeeksActivityProcess.MarketActionId && x.SeqNO == marketActionBefore4WeeksActivityProcess.SeqNO)).FirstOrDefault();
            //findOne.ActivityDateTime = marketActionBefore4WeeksActivityProcess.ActivityDateTime;
            //findOne.Contents = marketActionBefore4WeeksActivityProcess.Contents;
            //findOne.Item = marketActionBefore4WeeksActivityProcess.Item;
            //findOne.ModifyDateTime = DateTime.Now;
            //findOne.ModifyUserId = marketActionBefore4WeeksActivityProcess.ModifyUserId;
            //findOne.Remark = marketActionBefore4WeeksActivityProcess.Remark;
            //}
            db.SaveChanges();
        }
        public void MarketActionBefore4WeeksActivityProcessDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionBefore4WeeksActivityProcess WHERE MarketActionId = @MarketActionId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        public List<MarketActionBefore4WeeksCoopFund> MarketActionBefore4WeeksCoopFundSearch(string marketActionId,string coopFundTypeCode)
        {
            if (marketActionId == null) marketActionId = "";
            if (coopFundTypeCode == null) coopFundTypeCode = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId), new SqlParameter("@CoopFundCode", coopFundTypeCode) };
            Type t = typeof(MarketActionBefore4WeeksCoopFund);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionBefore4WeeksCoopFund] WHERE MarketActionId = @MarketActionId";
            if (!string.IsNullOrEmpty(coopFundTypeCode))
            {
                sql += " AND CoopFundCode = @CoopFundCode";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionBefore4WeeksCoopFund>().ToList();
        }
        public MarketActionBefore4WeeksCoopFund MarketActionBefore4WeeksCoopFundSave(MarketActionBefore4WeeksCoopFund marketActionBefore4WeeksCoopFund)
        {
            MarketActionBefore4WeeksCoopFund findOneMax = db.MarketActionBefore4WeeksCoopFund.Where(x => (x.MarketActionId == marketActionBefore4WeeksCoopFund.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionBefore4WeeksCoopFund.SeqNO = 1;
            }
            else
            {
                marketActionBefore4WeeksCoopFund.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionBefore4WeeksCoopFund.InDateTime = DateTime.Now;
            marketActionBefore4WeeksCoopFund.ModifyDateTime = DateTime.Now;
            db.MarketActionBefore4WeeksCoopFund.Add(marketActionBefore4WeeksCoopFund);

            db.SaveChanges();
            return marketActionBefore4WeeksCoopFund;
        }
        public void MarketActionBefore4WeeksCoopFundDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionBefore4WeeksCoopFund WHERE MarketActionId = @MarketActionId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        public List<MarketActionBefore4WeeksHandOverArrangement> MarketActionBefore4WeeksHandOverArrangementSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionBefore4WeeksHandOverArrangement);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionBefore4WeeksHandOverArrangement] WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionBefore4WeeksHandOverArrangement>().ToList();
        }
        public void MarketActionBefore4WeeksHandOverArrangementSave(MarketActionBefore4WeeksHandOverArrangement marketActionBefore4WeeksHandOverArrangement)
        {
            MarketActionBefore4WeeksHandOverArrangement findOneMax = db.MarketActionBefore4WeeksHandOverArrangement.Where(x => (x.MarketActionId == marketActionBefore4WeeksHandOverArrangement.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionBefore4WeeksHandOverArrangement.SeqNO = 1;
            }
            else
            {
                marketActionBefore4WeeksHandOverArrangement.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionBefore4WeeksHandOverArrangement.InDateTime = DateTime.Now;
            marketActionBefore4WeeksHandOverArrangement.ModifyDateTime = DateTime.Now;
            db.MarketActionBefore4WeeksHandOverArrangement.Add(marketActionBefore4WeeksHandOverArrangement);

            db.SaveChanges();
        }
        public void MarketActionBefore4WeeksHandOverArrangementDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionBefore4WeeksHandOverArrangement WHERE MarketActionId = @MarketActionId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        // 活动预算合计的计算
        public decimal? MarketActionBefore4WeeksTotalBudgetAmt(string marketActionId)
        {
            decimal? totalBudgetAmt = 0;
            List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFundList = MarketActionBefore4WeeksCoopFundSearch(marketActionId,"");
            foreach (MarketActionBefore4WeeksCoopFund marketActionBefore4WeeksCoopFund in marketActionBefore4WeeksCoopFundList)
            {
                totalBudgetAmt += marketActionBefore4WeeksCoopFund.CoopFundAmt == null ? 0 : marketActionBefore4WeeksCoopFund.CoopFundAmt;
            }
            return totalBudgetAmt;
        }
        #endregion
        #region two days after
        public List<MarketActionAfter2LeadsReportDto> MarketActionAfter2LeadsReportSearch(string marketActionId, string year)
        {
            if (marketActionId == null) marketActionId = "";
            if (year == null) year = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId),new SqlParameter("@Year", year) };
            Type t = typeof(MarketActionAfter2LeadsReportDto);
            string sql = "";
            sql += @"SELECT A.*,B.ActionName,B.ShopId,C.ShopName,C.ShopNameEn,F.AreaName,D.HiddenCodeName AS InterestedModelName,D.HiddenCodeNameEn AS InterestedModelNameEn
                    ,E.HiddenCodeName AS DealModelName,E.HiddenCodeNameEn AS DealModelNameEn
                   , CASE WHEN DCPCheck=1 THEN '是' ELSE '否' END AS DCPCheckName
                    ,CASE WHEN LeadsCheck=1 THEN '是' ELSE '否' END AS LeadsCheckName
                    ,CASE WHEN DealCheck=1 THEN '是' ELSE '否' END AS DealCheckName
                    FROM [MarketActionAfter2LeadsReport] A  INNER JOIN MarketAction B ON A.MarketActionId = B.MarketActionId
                                                            INNER JOIN Shop C ON B.ShopId = C.ShopId
                                                            INNER JOIN Area F ON C.AreaId = F.AreaId
                                                            LEFT JOIN HiddenCode D ON A.InterestedModel = D.HiddenCodeId AND D.HiddenCodeGroup = 'TargetModels'
                                                            LEFT JOIN HiddenCode E ON A.DealModel = E.HiddenCodeId AND E.HiddenCodeGroup = 'TargetModels'
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(marketActionId))
            {
                sql += " AND A.MarketActionId = @MarketActionId";

            }
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(B.StartDate) = @Year";

            }
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter2LeadsReportDto>().ToList();
        }
        public MarketActionAfter2LeadsReport MarketActionAfter2LeadsReportSave(MarketActionAfter2LeadsReport marketActionAfter2LeadsReport)
        {
            if (marketActionAfter2LeadsReport.SeqNO == 0)
            {
                MarketActionAfter2LeadsReport findOneMax = db.MarketActionAfter2LeadsReport.Where(x => (x.MarketActionId == marketActionAfter2LeadsReport.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
                if (findOneMax == null)
                {
                    marketActionAfter2LeadsReport.SeqNO = 1;
                }
                else
                {
                    marketActionAfter2LeadsReport.SeqNO = findOneMax.SeqNO + 1;
                }
                marketActionAfter2LeadsReport.InDateTime = DateTime.Now;
                marketActionAfter2LeadsReport.ModifyDateTime = DateTime.Now;
                db.MarketActionAfter2LeadsReport.Add(marketActionAfter2LeadsReport);

            }
            else
            {
                MarketActionAfter2LeadsReport findOne = db.MarketActionAfter2LeadsReport.Where(x => (x.MarketActionId == marketActionAfter2LeadsReport.MarketActionId && x.SeqNO == marketActionAfter2LeadsReport.SeqNO)).FirstOrDefault();
                if (findOne == null)
                {
                    marketActionAfter2LeadsReport.InDateTime = DateTime.Now;
                    marketActionAfter2LeadsReport.ModifyDateTime = DateTime.Now;
                    db.MarketActionAfter2LeadsReport.Add(marketActionAfter2LeadsReport);
                }
                else
                {
                    findOne.BPNO = marketActionAfter2LeadsReport.BPNO;
                    findOne.CustomerName = marketActionAfter2LeadsReport.CustomerName;
                    //findOne.TelNO = marketActionAfter2LeadsReport.TelNO;
                    findOne.DealCheck = marketActionAfter2LeadsReport.DealCheck;
                    findOne.DealModel = marketActionAfter2LeadsReport.DealModel;
                    findOne.InterestedModel = marketActionAfter2LeadsReport.InterestedModel;
                    findOne.LeadsCheck = marketActionAfter2LeadsReport.LeadsCheck;
                    findOne.ModifyDateTime = DateTime.Now;
                    findOne.ModifyUserId = marketActionAfter2LeadsReport.ModifyUserId;
                    findOne.DCPCheck = marketActionAfter2LeadsReport.DCPCheck;
                   // findOne.OwnerCheck = marketActionAfter2LeadsReport.OwnerCheck;
                   // findOne.TestDriverCheck = marketActionAfter2LeadsReport.TestDriverCheck;
                }
            }
            db.SaveChanges();
            return marketActionAfter2LeadsReport;
        }
        public void MarketActionAfter2LeadsReportDelete(string marketActionId, string seqNO)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId), new SqlParameter("@SeqNO", seqNO) };
            string sql = @"DELETE MarketActionAfter2LeadsReport WHERE MarketActionId = @MarketActionId AND SeqNO = @SeqNO
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        public List<MarketActionLeadsCountDto> MarketActionLeadsCountSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionLeadsCountDto);
            string sql = "";
            sql += @"SELECT * FROM
                    (SELECT
                            ISNULL(SUM(CASE WHEN OwnerCheck= 1 AND LeadsCheck = 1 THEN 1 ELSE 0 END),0) AS LeadOwnerCount,
                            ISNULL(SUM(CASE WHEN OwnerCheck <> 1 AND LeadsCheck = 1 THEN 1 ELSE 0 END), 0) AS LeadPCCount,
                            ISNULL(SUM(CASE WHEN OwnerCheck = 1 AND TestDriverCheck = 1 THEN 1 ELSE 0 END), 0) AS TestDriverOwnerCount,
                            ISNULL(SUM(CASE WHEN OwnerCheck <> 1 AND TestDriverCheck = 1 THEN 1 ELSE 0 END), 0) AS TestDriverPCCount,
                            ISNULL(SUM(CASE WHEN OwnerCheck = 1 AND DealCheck = 1 THEN 1 ELSE 0 END), 0) AS ActualOrderOwnerCount,
                            ISNULL(SUM(CASE WHEN OwnerCheck <> 1 AND DealCheck = 1 THEN 1 ELSE 0 END), 0) AS ActualOrderPCCount
                    FROM MarketActionAfter2LeadsReport A 
                    WHERE   A.MarketActionId = @MarketActionId) X INNER JOIN 
                    (SELECT ISNULL(SUM(ISNULL(UnitPrice,0)*ISNULL(Counts,0)),0) AS ExpenseTotalAmt 
                    FROM MarketActionAfter7ActualExpense A WHERE A.MarketActionId = @MarketActionId) Y ON 1=1";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionLeadsCountDto>().ToList();
        }
        #endregion
        #region Seven days after
        public List<MarketActionAfter7> MarketActionAfter7Search(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionAfter7);
            string sql = "";
            sql += @"SELECT A.* 
                    FROM [MarketActionAfter7] A 
                    WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter7>().ToList();
        }
        public void MarketActionAfter7Save(MarketActionAfter7 marketActionAfter7)
        {
            MarketActionAfter7 findOne = db.MarketActionAfter7.Where(x => (x.MarketActionId == marketActionAfter7.MarketActionId)).FirstOrDefault();
            if (findOne == null)
            {
                marketActionAfter7.InDateTime = DateTime.Now;
                marketActionAfter7.ModifyDateTime = DateTime.Now;
                db.MarketActionAfter7.Add(marketActionAfter7);
            }
            else
            {
                findOne.CoopFundSumAmt = marketActionAfter7.CoopFundSumAmt;
                findOne.CustomerFeedback = marketActionAfter7.CustomerFeedback;
                findOne.HightLights = marketActionAfter7.HightLights;
                findOne.ImproveArea = marketActionAfter7.ImproveArea;
                findOne.MarketSaleTeamAdvice = marketActionAfter7.MarketSaleTeamAdvice;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = marketActionAfter7.ModifyUserId;
                findOne.People_ActualArrivalCount = marketActionAfter7.People_ActualArrivalCount;
                findOne.People_ActualCarOwnerCount = marketActionAfter7.People_ActualCarOwnerCount;
                findOne.People_ActualDepositorCount = marketActionAfter7.People_ActualDepositorCount;
                findOne.People_ActualPotentialCount = marketActionAfter7.People_ActualPotentialCount;
                findOne.People_DCPIDCount = marketActionAfter7.People_DCPIDCount;
                findOne.People_NewLeadsThsYearCount = marketActionAfter7.People_NewLeadsThsYearCount;
                findOne.People_OthersCount = marketActionAfter7.People_OthersCount;
                findOne.People_ParticipantsCount = marketActionAfter7.People_ParticipantsCount;
                findOne.People_NewOrderCount = marketActionAfter7.People_NewOrderCount;
                findOne.ProcessPercent = marketActionAfter7.ProcessPercent;
                findOne.Platform_Media = marketActionAfter7.Platform_Media;
                findOne.Platform_ExposuerForm = marketActionAfter7.Platform_ExposuerForm;
            }

            db.SaveChanges();
        }
        public List<MarketActionAfter7ActualExpenseDto> MarketActionAfter7ActualExpenseSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionAfter7ActualExpenseDto);
            string sql = "";
            sql += @"SELECT A.*,ISNULL(A.UnitPrice,0)*ISNULL(Counts,0) AS Total  FROM [MarketActionAfter7ActualExpense] A WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter7ActualExpenseDto>().ToList();
        }
        //public void MarketActionAfter7ActualExpenseSave(MarketActionAfter7ActualExpense marketActionAfter7ActualExpense)
        //{
        //    //if (marketActionAfter7ActualExpense.SeqNO == 0)
        //    //{
        //    MarketActionAfter7ActualExpense findOneMax = db.MarketActionAfter7ActualExpense.Where(x => (x.MarketActionId == marketActionAfter7ActualExpense.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
        //    if (findOneMax == null)
        //    {
        //        marketActionAfter7ActualExpense.SeqNO = 1;
        //    }
        //    else
        //    {
        //        marketActionAfter7ActualExpense.SeqNO = findOneMax.SeqNO + 1;
        //    }
        //    marketActionAfter7ActualExpense.InDateTime = DateTime.Now;
        //    marketActionAfter7ActualExpense.ModifyDateTime = DateTime.Now;
        //    db.MarketActionAfter7ActualExpense.Add(marketActionAfter7ActualExpense);
        //    db.SaveChanges();
        //}
        public void MarketActionAfter7ActualExpenseDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionAfter7ActualExpense WHERE MarketActionId = @MarketActionId 
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }

        public List<MarketActionAfter7ActualProcess> MarketActionAfter7ActualProcessSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionAfter7ActualProcess);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionAfter7ActualProcess] WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter7ActualProcess>().ToList();
        }
        public void MarketActionAfter7ActualProcessSave(MarketActionAfter7ActualProcess marketActionAfter7ActualProcess)
        {
            //if (marketActionAfter7ActualProcess.SeqNO == 0)
            //{
            MarketActionAfter7ActualProcess findOneMax = db.MarketActionAfter7ActualProcess.Where(x => (x.MarketActionId == marketActionAfter7ActualProcess.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionAfter7ActualProcess.SeqNO = 1;
            }
            else
            {
                marketActionAfter7ActualProcess.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionAfter7ActualProcess.InDateTime = DateTime.Now;
            marketActionAfter7ActualProcess.ModifyDateTime = DateTime.Now;
            db.MarketActionAfter7ActualProcess.Add(marketActionAfter7ActualProcess);

            //}
            //else
            //{
            //    MarketActionAfter7ActualProcess findOne = db.MarketActionAfter7ActualProcess.Where(x => (x.MarketActionId == marketActionAfter7ActualProcess.MarketActionId && x.SeqNO == marketActionAfter7ActualProcess.SeqNO)).FirstOrDefault();
            //    findOne.ActivityDateTime = marketActionAfter7ActualProcess.ActivityDateTime;
            //    findOne.Contents = marketActionAfter7ActualProcess.Contents;
            //    findOne.Item = marketActionAfter7ActualProcess.Item;
            //    findOne.ModifyDateTime = DateTime.Now;
            //    findOne.ModifyUserId = marketActionAfter7ActualProcess.ModifyUserId;
            //    findOne.Remark = marketActionAfter7ActualProcess.Remark;

            //}
            db.SaveChanges();
        }
        public void MarketActionAfter7ActualProcessDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionAfter7ActualProcess WHERE MarketActionId = @MarketActionId 
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        public List<MarketActionAfter7CoopFund> MarketActionAfter7CoopFundSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionAfter7CoopFund);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionAfter7CoopFund] WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter7CoopFund>().ToList();
        }
        public MarketActionAfter7CoopFund MarketActionAfter7CoopFundSave(MarketActionAfter7CoopFund marketActionAfter7CoopFund)
        {
            MarketActionAfter7CoopFund findOneMax = db.MarketActionAfter7CoopFund.Where(x => (x.MarketActionId == marketActionAfter7CoopFund.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionAfter7CoopFund.SeqNO = 1;
            }
            else
            {
                marketActionAfter7CoopFund.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionAfter7CoopFund.InDateTime = DateTime.Now;
            marketActionAfter7CoopFund.ModifyDateTime = DateTime.Now;
            db.MarketActionAfter7CoopFund.Add(marketActionAfter7CoopFund);

            db.SaveChanges();
            return marketActionAfter7CoopFund;
        }
        public void MarketActionAfter7CoopFundDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionAfter7CoopFund WHERE MarketActionId = @MarketActionId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        // 活动预算合计的计算
        public decimal? MarketActionAfter7TotalBudgetAmt(string marketActionId)
        {
            decimal? totalBudgetAmt = 0;
            List<MarketActionAfter7CoopFund> marketActionAfter7CoopFundList = MarketActionAfter7CoopFundSearch(marketActionId);
            foreach (MarketActionAfter7CoopFund marketActionAfter7CoopFund in marketActionAfter7CoopFundList)
            {
                totalBudgetAmt += marketActionAfter7CoopFund.CoopFundAmt == null ? 0 : marketActionAfter7CoopFund.CoopFundAmt;
            }
            return totalBudgetAmt;
        }
        public List<MarketActionAfter7HandOverArrangement> MarketActionAfter7HandOverArrangementSearch(string marketActionId)
        {
            if (marketActionId == null) marketActionId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            Type t = typeof(MarketActionAfter7HandOverArrangement);
            string sql = "";
            sql += @"SELECT *  FROM [MarketActionAfter7HandOverArrangement] WHERE MarketActionId = @MarketActionId";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionAfter7HandOverArrangement>().ToList();
        }
        public void MarketActionAfter7HandOverArrangementSave(MarketActionAfter7HandOverArrangement marketActionAfter7HandOverArrangement)
        {
            MarketActionAfter7HandOverArrangement findOneMax = db.MarketActionAfter7HandOverArrangement.Where(x => (x.MarketActionId == marketActionAfter7HandOverArrangement.MarketActionId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
            if (findOneMax == null)
            {
                marketActionAfter7HandOverArrangement.SeqNO = 1;
            }
            else
            {
                marketActionAfter7HandOverArrangement.SeqNO = findOneMax.SeqNO + 1;
            }
            marketActionAfter7HandOverArrangement.InDateTime = DateTime.Now;
            marketActionAfter7HandOverArrangement.ModifyDateTime = DateTime.Now;
            db.MarketActionAfter7HandOverArrangement.Add(marketActionAfter7HandOverArrangement);

            db.SaveChanges();
        }
        public void MarketActionAfter7HandOverArrangementDelete(string marketActionId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MarketActionId", marketActionId) };
            string sql = @"DELETE MarketActionAfter7HandOverArrangement WHERE MarketActionId = @MarketActionId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region 总览
        // 市场活动和交车仪式统计
        public List<MarketActionStatusCountDto> MarketActionStatusCountSearch(string year, string eventTypeId, List<Shop> roleTypeShop)
        {
            if (year == null) year = "";
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@Year", year) };
            Type t = typeof(MarketActionStatusCountDto);
            string sql = "";
            sql += @"SELECT ISNULL(SUM(Before4WeeksNotCommit),0) AS Before4WeeksNotCommit 
	                       ,ISNULL(SUM(Before4WeeksWaitForChange),0) AS Before4WeeksWaitForChange
	                       ,ISNULL(SUM(After7NotCommit),0) AS After7NotCommit
	                       ,ISNULL(SUM(After7WaitForChange),0) AS After7WaitForChange
                    FROM (
                             SELECT 
                            CASE WHEN GETDATE()>DATEADD(DD,-28,StartDate) AND NOT EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=1)
				                            THEN 1
				                            ELSE 0
			                            END AS Before4WeeksNotCommit-- 到时间还未提交
                            , CASE WHEN  EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType =1 AND DTTApproveCode<>'2')
				                            THEN 1
				                            ELSE 0
			                            END AS Before4WeeksWaitForChange -- 提交后还未通过(包括待审批和待修改)
                            , CASE WHEN GETDATE()>DATEADD(DD,14,StartDate) AND NOT EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType=2)
				                            THEN 1
				                            ELSE 0
			                            END AS After7NotCommit
                            , CASE WHEN  EXISTS(SELECT 1 FROM DTTApprove WHERE MarketActionId = A.MarketActionId AND DTTType =2 AND DTTApproveCode<>'2')
				                            THEN 1
				                            ELSE 0
			                            END AS After7WaitForChange
                            FROM MarketAction A WHERE 1=1 AND A.MarketActionStatusCode<>2 ";
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(A.StartDate) = @Year";
            }
            if (eventTypeId == "99")
            {
                sql += " AND A.EventTypeId = 99 ";
            }
            else
            {
                sql += " AND A.EventTypeId <> 99 ";
            }
            if (roleTypeShop != null && roleTypeShop.Count > 0)
            {
                sql += " AND A.ShopId IN( ";
                foreach (Shop shop in roleTypeShop)
                {
                    if (roleTypeShop.IndexOf(shop) == roleTypeShop.Count - 1)
                    {
                        sql += shop.ShopId;
                    }
                    else
                    {
                        sql += shop.ShopId + ",";
                    }
                }
                sql += ")";
            }
            sql += " ) B";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionStatusCountDto>().ToList();
        }
        // 报告
        public List<MarketActionReportCountDto> MarketActionReportCountSearch(string year, string eventTypeId, List<Shop> roleTypeShop)
        {
            if (year == null) year = "";
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@Year", year) };
            Type t = typeof(MarketActionReportCountDto);
            string sql = "";
            sql += @"SELECT ISNULL(SUM(PlanBugetUnCommit),0) AS PlanBugetUnCommit 
	                       ,ISNULL(SUM(PlanCoopFundUnCommit),0) AS PlanCoopFundUnCommit
	                       ,ISNULL(SUM(LeadsUnCommit),0) AS LeadsUnCommit
	                       ,ISNULL(SUM(ReportBugetUnCommit),0) AS ReportBugetUnCommit
	                       ,ISNULL(SUM(ReportCoopFundUnCommit),0) AS ReportCoopFundUnCommit
                    FROM (
                            SELECT 
                            CASE WHEN GETDATE()>DATEADD(DD,-28,StartDate) AND NOT EXISTS(SELECT 1 FROM MarketActionBefore4WeeksCoopFund WHERE MarketActionId = A.MarketActionId )
				                            THEN 1
				                            ELSE 0
			                            END AS PlanBugetUnCommit -- 到时间还填写预算费用
                            , CASE WHEN GETDATE()>DATEADD(DD,-28,StartDate) AND B.CoopFundSumAmt IS NULL
				                            THEN 1
				                            ELSE 0
			                            END AS PlanCoopFundUnCommit -- 到时间还未填写市场基金金额合计
                            , CASE WHEN  GETDATE()>DATEADD(DD,14,StartDate) AND NOT EXISTS(SELECT 1 FROM MarketActionAfter2LeadsReport WHERE MarketActionId = A.MarketActionId )
				                            THEN 1
				                            ELSE 0
			                            END AS LeadsUnCommit --到时间还未填写线索报告
			                            
			                 ,CASE WHEN GETDATE()>DATEADD(DD,14,StartDate) AND NOT EXISTS(SELECT 1 FROM MarketActionAfter7CoopFund WHERE MarketActionId = A.MarketActionId )
				                            THEN 1
				                            ELSE 0
			                            END AS ReportBugetUnCommit -- 到时间还填写预算费用
                            , CASE WHEN GETDATE()>DATEADD(DD,14,StartDate) AND B.CoopFundSumAmt IS NULL
				                            THEN 1
				                            ELSE 0
			                            END AS ReportCoopFundUnCommit -- 到时间还未填写市场基金金额合计
                            FROM MarketAction A LEFT JOIN  MarketActionBefore4Weeks B ON  A.MarketActionId = B.MarketActionId
												LEFT JOIN MarketActionAfter7 C  ON  A.MarketActionId = C.MarketActionId
                            WHERE 1=1 AND A.MarketActionStatusCode<>2 ";
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(A.StartDate) = @Year";
            }
            if (eventTypeId == "99")
            {
                sql += " AND A.EventTypeId = 99 ";
            }
            else
            {
                sql += " AND A.EventTypeId <> 99 ";
            }
            if (roleTypeShop != null && roleTypeShop.Count > 0)
            {
                sql += " AND A.ShopId IN( ";
                foreach (Shop shop in roleTypeShop)
                {
                    if (roleTypeShop.IndexOf(shop) == roleTypeShop.Count - 1)
                    {
                        sql += shop.ShopId;
                    }
                    else
                    {
                        sql += shop.ShopId + ",";
                    }
                }
                sql += ")";
            }
            sql += " ) B";
            return db.Database.SqlQuery(t, sql, para).Cast<MarketActionReportCountDto>().ToList();
        }

        #endregion
    }
}