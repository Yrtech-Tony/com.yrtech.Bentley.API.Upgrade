using com.yrtech.InventoryAPI.DTO;
using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using com.yrtech.InventoryAPI.Common;

namespace com.yrtech.InventoryAPI.Service
{
    public class DMFService
    {
        Bentley db = new Bentley();
        #region DMFItem
        public List<DMFItem> DMFItemSearch(string dmfItemId, string dmfItemName, string dmfItemNameEn, bool? expenseAccountChk, bool? publishChk)
        {
            if (dmfItemId == null) dmfItemId = "";
            if (dmfItemName == null) dmfItemName = "";
            if (dmfItemNameEn == null) dmfItemNameEn = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DMFItemId", dmfItemId),
                                                    new SqlParameter("@DMFItemName", dmfItemName),
                                                    new SqlParameter("@DMFItemNameEn", dmfItemNameEn)};

            Type t = typeof(DMFItem);

            string sql = "";
            sql = @"SELECT A.* 
                    FROM DMFItem A 
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(dmfItemId))
            {
                sql += " AND DMFItemId = @DMFItemId";
            }
            if (!string.IsNullOrEmpty(dmfItemName))
            {
                sql += " AND DMFItemName = @DMFItemName";
            }
            if (!string.IsNullOrEmpty(dmfItemNameEn))
            {
                sql += " AND DMFItemNameEn = @DMFItemNameEn";
            }
            if (expenseAccountChk.HasValue)
            {
                para = para.Concat(new SqlParameter[] { new SqlParameter("@ExpenseAccountChk", expenseAccountChk) }).ToArray();
                sql += " AND ExpenseAccountChk = @ExpenseAccountChk";
            }
            if (publishChk.HasValue)
            {
                para = para.Concat(new SqlParameter[] { new SqlParameter("@PublishChk", publishChk) }).ToArray();
                sql += " AND PublishChk = @PublishChk";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<DMFItem>().ToList();
        }
        public DMFItem DMFItemSave(DMFItem dmfItem)
        {
            DMFItem findOne = db.DMFItem.Where(x => (x.DMFItemId == dmfItem.DMFItemId)).FirstOrDefault();
            if (findOne == null)
            {
                dmfItem.InDateTime = DateTime.Now;
                dmfItem.ModifyDateTime = DateTime.Now;
                db.DMFItem.Add(dmfItem);
            }
            else
            {
                findOne.DMFItemName = dmfItem.DMFItemName;
                findOne.DMFItemNameEn = dmfItem.DMFItemNameEn;
                findOne.DMFItemRemark = dmfItem.DMFItemRemark;
                findOne.ExpenseAccountChk = dmfItem.ExpenseAccountChk;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = dmfItem.ModifyUserId;
                findOne.PublishChk = dmfItem.PublishChk;

                dmfItem = findOne;
            }
            db.SaveChanges();
            return dmfItem;
        }
        public void DMFItemDelete(string dmfItemId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DMFItemId", dmfItemId) };
            string sql = @"DELETE DMFItem WHERE DMFItemId = @DMFItemId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region DMFDetail
        public string DMFDetailDecryptSearch(string dmfDetailId, string shopId, string dmfItemId)
        {
            if (dmfDetailId == null) dmfDetailId = "";
            if (shopId == null) shopId = "";
            if (dmfItemId == null) dmfItemId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DMFDetailId", dmfDetailId),
                                                    new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@DMFItemId", dmfItemId)};

            Type t = typeof(DMFDetail);

            string sql = "";
            sql = @" SELECT * FROM DMFDetail WHERE 1=1 ";
            if (!string.IsNullOrEmpty(dmfDetailId))
            {
                sql += " AND DMFDetailId = @DMFDetailId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(dmfItemId))
            {
                sql += " AND DMFItemId = @DMFItemId";
            }
            List<DMFDetail> list = db.Database.SqlQuery(t, sql, para).Cast<DMFDetail>().ToList();

            string sqlReturn = @"CREATE TABLE #DMFDetail(
                                                        [DMFDetailId][int]  NOT NULL,
                                                        [ShopId] [int] NULL,
	                                                    [DMFItemId] [int] NULL,
	                                                    [Budget] [decimal](19, 2) NULL,
	                                                    [AcutalAmt] [decimal](19, 2) NULL,
	                                                    [Remark] [nvarchar](max) NULL
                                                     )";

            foreach (DMFDetail dmfDetail in list)
            {
                sqlReturn += " INSERT INTO #DMFDetail VALUES(";
                sqlReturn += dmfDetail.DMFDetailId + ",'";
                if (dmfDetail.ShopId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += dmfDetail.ShopId + "','";
                }
                if (dmfDetail.DMFItemId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += dmfDetail.DMFItemId + "','";
                }
                if (string.IsNullOrEmpty(dmfDetail.Budget)||string.IsNullOrEmpty(TokenHelper.DecryptDES(dmfDetail.Budget)))
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += TokenHelper.DecryptDES(dmfDetail.Budget) + "','";
                }
                if (string.IsNullOrEmpty(dmfDetail.AcutalAmt)||string.IsNullOrEmpty(TokenHelper.DecryptDES(dmfDetail.AcutalAmt)))
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += TokenHelper.DecryptDES(dmfDetail.AcutalAmt) + "','";
                }
                if (dmfDetail.Remark == null)
                {
                    sqlReturn += "null')";
                }
                else
                {
                    sqlReturn += dmfDetail.Remark + "')";
                }
            }

            return sqlReturn;
        }
        public List<DMFDetailDto> DMFDetailSearch(string dmfDetailId, string shopId, string dmfItemId, string dmfItemName)
        {
            if (dmfDetailId == null) dmfDetailId = "";
            if (shopId == null) shopId = "";
            if (dmfItemId == null) dmfItemId = "";
            if (dmfItemName == null) dmfItemName = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DMFDetailId", dmfDetailId),
                                                    new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@DMFItemId", dmfItemId),
                                                    new SqlParameter("@DMFItemName", dmfItemName)};

            Type t = typeof(DMFDetailDto);

            string sql = "";
            sql += DMFDetailDecryptSearch(dmfDetailId, shopId, dmfItemId);
            sql += ExpenseAccountDecryptSearch("", shopId, dmfItemId, "");
            sql += MonthSaleDesryptSearch("", shopId, "");

            sql += @"  SELECT * FROM (
                        SELECT ISNULL(D.DMFDetailId,0) AS DMFDetailId,A.ShopId,A.DMFItemId,CAST(ISNULL(D.Budget,0) AS NVARCHAR) AS Budget
							    ,CAST(A.AcutalAmt AS NVARCHAR) AS AcutalAmt,ISNULL(D.Remark,'') AS Remark
							    ,B.ShopCode,B.ShopName,B.ShopNameEn,C.DMFItemName,C.DMFItemNameEn 
					    FROM (
							    SELECT ShopId,DMFItemId,ISNULL(SUM(ExpenseAmt),0) AS AcutalAmt
							    FROM #ExpenseAccount  
							    WHERE ApplyStatus='通过'
							    GROUP BY ShopId,DMFItemId) A INNER JOIN Shop B ON A.ShopId = B.ShopId
														    INNER JOIN DMFItem C ON A.DMFItemId = C.DMFItemId
														LEFT JOIN #DMFDetail D ON A.ShopId = D.ShopId AND A.DMFItemId = D.DMFItemId	
                    UNION ALL
					 
					SELECT A.[DMFDetailId],A.[ShopId],A.[DMFItemId],CAST(A.[Budget] AS NVARCHAR) AS Budget,CAST(ISNULL(A.[AcutalAmt],0) AS NVARCHAR) AS AcutalAmt
                           ,ISNULL(A.[Remark],'') AS Remark,B.ShopCode,B.ShopName,B.ShopNameEn,C.DMFItemName,C.DMFItemNameEn 
                    FROM #DMFDetail A INNER JOIN Shop B ON A.ShopId = B.ShopId
                                    INNER JOIN DMFItem C ON A.DMFItemId = C.DMFItemId
                    WHERE NOT EXISTS(SELECT 1 FROM #ExpenseAccount WHERE ShopId = A.ShopId AND DMFItemId = A.DMFItemId AND  ApplyStatus='通过')
                    ) X WHERE 1=1";
            if (!string.IsNullOrEmpty(dmfDetailId))
            {
                sql += " AND DMFDetailId = @DMFDetailId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND X.ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(dmfItemId))
            {
                sql += " AND X.DMFItemId = @DMFItemId";
            }
            if (!string.IsNullOrEmpty(dmfItemName))
            {
                sql += " AND X.dmfItemName LIKE '%'+ @DMFItemName+ '%'";
            }
            sql += " DROP TABLE #ExpenseAccount";
            sql += " DROP TABLE #MonthSale";
            sql += " DROP TABLE #DMFDetail";
            return db.Database.SqlQuery(t, sql, para).Cast<DMFDetailDto>().ToList();
        }
        public DMFDetail DMFDetailSave(DMFDetail dmfDetail)
        {
            DMFDetail findOne = db.DMFDetail.Where(x => (x.DMFDetailId == dmfDetail.DMFDetailId)).FirstOrDefault();
            if (findOne == null)
            {
                dmfDetail.InDateTime = DateTime.Now;
                dmfDetail.ModifyDateTime = DateTime.Now;
                db.DMFDetail.Add(dmfDetail);
            }
            else
            {
                findOne.AcutalAmt = dmfDetail.AcutalAmt;
                findOne.Budget = dmfDetail.Budget;
                findOne.DMFItemId = dmfDetail.DMFItemId;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = dmfDetail.ModifyUserId;
                findOne.Remark = dmfDetail.Remark;
                findOne.ShopId = dmfDetail.ShopId;

                dmfDetail = findOne;
            }
            db.SaveChanges();
            return dmfDetail;
        }
        public void DMFDetailDelete(string dmfDetailId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DMFDetailId", dmfDetailId) };
            string sql = @"DELETE DMFDetail WHERE DMFDetailId = @DMFDetailId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region DMF
        public List<DMFDto> DMFSearch(string shopId)
        {
            if (shopId == null) shopId = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ShopId", shopId) };

            Type t = typeof(DMFDto);

            string sql = "";
            sql += MonthSaleDesryptSearch("", shopId, "");
            sql += @"                 
                SELECT DISTINCT A.ShopId,A.ShopCode,A.ShopName,A.ShopNameEn
			                ,ISNULL(B.ActualMonthSaleCount,0) AS ActualMonthSaleCount
			                ,ISNULL(B.ActualMonthSaleAmt,0) AS ActualMonthSaleAmt
                FROM Shop A LEFT JOIN
                            (SELECT ShopId
		                            ,ISNULL(SUM(ActualSaleCount),0) AS ActualMonthSaleCount
		                            ,ISNULL(SUM(ActualSaleAmt),0) AS ActualMonthSaleAmt
                            FROM #MonthSale GROUP BY ShopId) B ON A.ShopId = B.ShopId
               WHERE 1=1";
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            List<DMFDto> dmfList = db.Database.SqlQuery(t, sql, para).Cast<DMFDto>().ToList();

            // 实际费用和差额赋值
            var detailList = DMFDetailSearch("", shopId, "", "").GroupBy(x => new { x.ShopId, x.ShopName, x.ShopNameEn }).Select(y => new
            {
                ShopId = y.Key.ShopId,
                ShopName = y.Key.ShopName,
                ShopNameEn = y.Key.ShopNameEn,
                ActualAmt = y.Sum(x => Convert.ToDecimal(x.AcutalAmt))
            }).ToList();

            foreach (DMFDto dmf in dmfList)
            {
                foreach (var dmfDetail in detailList)
                {
                    if (dmf.ShopId == dmfDetail.ShopId)
                    {
                        dmf.ActualAmt = dmfDetail.ActualAmt;
                        dmf.DiffAmt = dmf.ActualMonthSaleAmt - dmfDetail.ActualAmt;
                    }
                }
            }
            return dmfList;
        }
        public List<DMFDto> DMFQuarterSearch(string shopId)
        {
            if (shopId == null) shopId = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ShopId", shopId) };

            Type t = typeof(DMFDto);

            string sql = "";
            sql += MonthSaleDesryptSearch("", shopId, "");
            sql += @"                 
                SELECT A.ShopId,A.ShopName,A.ShopNameEn,Y.Quarters
	                ,ISNULL(SUM(Y.ActualSaleCount),0) AS ActualMonthSaleCount
	                ,ISNULL(SUM(Y.ActualSaleAmt),0) AS ActualMonthSaleAmt
                FROM Shop A INNER JOIN 
			                (SELECT ShopId,CASE WHEN MonthStr IN ('01','1','02','2','3','03') THEN 'Q1'
								                WHEN MonthStr IN ('04','4','05','5','06','6') THEN 'Q2'
								                WHEN MonthStr IN ('07','7','08','8','09','9') THEN 'Q3'
								                ELSE 'Q4' END AS Quarters
				                  ,ActualSaleAmt,ActualSaleCount
			                FROM 
				                (SELECT ShopId
					                ,CASE WHEN LEN(YearMonth)=7 THEN LEFT(YearMonth,2)
						                   WHEN LEN(YearMonth)=6 THEN LEFT(YearMonth,1)
						                   ELSE '' END AS MonthStr
					                ,ActualSaleAmt,ActualSaleCount
				                FROM #MonthSale) X) Y ON A.ShopId = Y.ShopId
                WHERE 1=1 
               ";
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            sql += " GROUP BY A.ShopId,A.ShopName,A.ShopNameEn,Y.Quarters";
            return db.Database.SqlQuery(t, sql, para).Cast<DMFDto>().ToList();
        }
        #endregion
        #region ExpenseAccount
        public string ExpenseAccountDecryptSearch(string expenseAccountId, string shopId, string dmfItemId, string marketActionId)
        {
            if (expenseAccountId == null) expenseAccountId = "";
            if (shopId == null) shopId = "";
            if (dmfItemId == null) dmfItemId = "";
            if (marketActionId == null) marketActionId = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ExpenseAccountId", expenseAccountId),
                                                    new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@DMFItemId", dmfItemId),
                                                    new SqlParameter("@MarketActionId", marketActionId)};
            Type t = typeof(ExpenseAccount);
            string sql = "";
            sql = @"SELECT * FROM ExpenseAccount A
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(expenseAccountId))
            {
                sql += " AND A.ExpenseAccountId = @ExpenseAccountId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(dmfItemId))
            {
                sql += " AND A.DMFItemId = @DMFItemId";
            }
            if (!string.IsNullOrEmpty(marketActionId))
            {
                sql += " AND A.MarketActionId = @MarketActionId";
            }

            List<ExpenseAccount> list = db.Database.SqlQuery(t, sql, para).Cast<ExpenseAccount>().ToList();
            string sqlReturn = @"CREATE TABLE #ExpenseAccount(
	                                            [ExpenseAccountId] [int]  NOT NULL,
	                                            [ShopId] [int] NULL,
	                                            [DMFItemId] [int] NULL,
	                                            [MarketActionId] [int] NULL,
	                                            [ExpenseAmt] [decimal](19, 2) NULL,
	                                            [ApplyStatus] [nvarchar](50) NULL,
	                                            [ApprovalReason] [nvarchar](max) NULL,
	                                            [ReplyStatus] [nvarchar](50) NULL,
	                                            [ReplyReason] [nvarchar](max) NULL )";

            foreach (ExpenseAccount expenseAccount in list)
            {
                sqlReturn += " INSERT INTO #ExpenseAccount VALUES('";
                sqlReturn += expenseAccount.ExpenseAccountId + "','";
                if (expenseAccount.ShopId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += expenseAccount.ShopId + "','";
                }
                if (expenseAccount.DMFItemId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += expenseAccount.DMFItemId + "','";
                }
                if (expenseAccount.MarketActionId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += expenseAccount.MarketActionId + "','";
                }
                if (string.IsNullOrEmpty(expenseAccount.ExpenseAmt) ||
                    string.IsNullOrEmpty(TokenHelper.DecryptDES(expenseAccount.ExpenseAmt)))
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += TokenHelper.DecryptDES(expenseAccount.ExpenseAmt) + "','";
                }
                if (expenseAccount.ApplyStatus == null)
                {
                    sqlReturn += "null','";
                }
                else
                {
                    sqlReturn += expenseAccount.ApplyStatus + "','";
                }
                if (expenseAccount.ApprovalReason == null)
                {
                    sqlReturn += "null','";
                }
                else
                {
                    sqlReturn += expenseAccount.ApprovalReason + "','";
                }
                if (expenseAccount.ReplyStatus == null)
                {
                    sqlReturn += "null','";
                }
                else
                {
                    sqlReturn += expenseAccount.ReplyStatus + "','";
                }
                if (expenseAccount.ReplyReason == null)
                {
                    sqlReturn += "null')";
                }
                else
                {
                    sqlReturn += expenseAccount.ReplyReason + "')";
                }
            }
            return sqlReturn;
        }
        public List<ExpenseAccountDto> ExpenseAccountSearch(string expenseAccountId, string shopId, string dmfItemId, string marketActionId)
        {
            if (expenseAccountId == null) expenseAccountId = "";
            if (shopId == null) shopId = "";
            if (dmfItemId == null) dmfItemId = "";
            if (marketActionId == null) marketActionId = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ExpenseAccountId", expenseAccountId),
                                                    new SqlParameter("@ShopId", shopId),
                                                    new SqlParameter("@DMFItemId", dmfItemId),
                                                    new SqlParameter("@MarketActionId", marketActionId)};
            Type t = typeof(ExpenseAccountDto);
            string sql = "";
            sql = @"SELECT DISTINCT A.ExpenseAccountId,A.DMFItemId,A.ShopId,A.MarketActionId,A.ExpenseAmt,A.ApprovalReason
                       ,A.ReplyReason
                       ,CASE WHEN A.ApplyStatus IS NOT NULL AND A.ApplyStatus<>'' 
                            THEN A.ApplyStatus
                            WHEN  EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode = 1)
                                 AND EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode = 2)
                            THEN '待审批'
                            ELSE '未提交'  
                        END AS ApplyStatus
                     ,CASE WHEN A.ReplyStatus IS NOT NULL AND A.ReplyStatus<>'' 
                            THEN A.ReplyStatus
                            WHEN  EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode = 3)
                                 AND EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode = 4)
                                AND EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode = 5)
                            THEN '待审批'
                            ELSE '未提交'  
                        END AS ReplyStatus,
                    B.ShopName,B.ShopNameEn,C.DMFItemName,C.DMFItemNameEn,D.ActionName,D.EventTypeId,
                    (SELECT EventTypeName FROM EventType WHERE EventTypeId = D.EventTypeId) AS EventTypeName,
                    (SELECT EventTypeNameEn FROM EventType WHERE EventTypeId = D.EventTypeId) AS EventTypeNameEn
                    FROM ExpenseAccount A INNER JOIN Shop B ON A.ShopId = B.ShopId
                                            LEFT JOIN DMFItem C ON A.DMFItemId = C.DMFItemId
                                            LEFT JOIN MarketAction D ON A.MarketActionId = D.MarketActionId
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(expenseAccountId))
            {
                sql += " AND A.ExpenseAccountId = @ExpenseAccountId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(dmfItemId))
            {
                sql += " AND A.DMFItemId = @DMFItemId";
            }
            if (!string.IsNullOrEmpty(marketActionId))
            {
                sql += " AND A.MarketActionId = @MarketActionId";
            }

            return db.Database.SqlQuery(t, sql, para).Cast<ExpenseAccountDto>().ToList();
        }
        public ExpenseAccount ExpenseAccountSave(ExpenseAccount expenseAccount)
        {
            if (expenseAccount.ApplyStatus == "未提交")
            {
                expenseAccount.ApplyStatus = "";
            }
            if (expenseAccount.ReplyStatus == "未提交")
            {
                expenseAccount.ReplyStatus = "";
            }
            ExpenseAccount findOne = db.ExpenseAccount.Where(x => (x.ExpenseAccountId == expenseAccount.ExpenseAccountId)).FirstOrDefault();
            if (findOne == null)
            {
                expenseAccount.InDateTime = DateTime.Now;
                expenseAccount.ModifyDateTime = DateTime.Now;
                db.ExpenseAccount.Add(expenseAccount);
            }
            else
            {
                findOne.ApplyStatus = expenseAccount.ApplyStatus;
                findOne.ApprovalReason = expenseAccount.ApprovalReason;
                findOne.DMFItemId = expenseAccount.DMFItemId;
                findOne.ExpenseAmt = expenseAccount.ExpenseAmt;
                findOne.MarketActionId = expenseAccount.MarketActionId;
                findOne.ReplyReason = expenseAccount.ReplyReason;
                findOne.ReplyStatus = expenseAccount.ReplyStatus;
                findOne.ShopId = expenseAccount.ShopId;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = expenseAccount.ModifyUserId;
                expenseAccount = findOne;
            }
            db.SaveChanges();

            return expenseAccount;
        }
        public void ExpenseAccountDelete(string expenseAccountId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ExpenseAccountId", expenseAccountId) };
            string sql = @"DELETE ExpenseAccount WHERE ExpenseAccountId = @ExpenseAccountId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        public List<ExpenseAccountFile> ExpenseAccountFileSearch(string expenseAccountId, string seqNO, string fileType)
        {
            if (expenseAccountId == null) expenseAccountId = "";
            if (seqNO == null) seqNO = "";
            if (fileType == null) fileType = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ExpenseAccountId", expenseAccountId),
                                                    new SqlParameter("@SeqNO", seqNO),
                                                    new SqlParameter("@FileType", fileType)};
            Type t = typeof(ExpenseAccountFile);
            string sql = "";
            sql = @"SELECT A.*  FROM dbo.ExpenseAccountFile A WHERE 1=1 AND ExpenseAccountId = @ExpenseAccountId";
            //if (!string.IsNullOrEmpty(expenseAccountId))
            //{
            //    sql += " AND ExpenseAccountId = @ExpenseAccountId";
            //}
            if (!string.IsNullOrEmpty(seqNO))
            {
                sql += " AND SeqNO = @SeqNO";
            }
            if (!string.IsNullOrEmpty(fileType))
            {
                sql += " AND FileTypeCode = @FileType";
            }

            return db.Database.SqlQuery(t, sql, para).Cast<ExpenseAccountFile>().ToList();
        }
        public void ExpenseAccountFileSave(ExpenseAccountFile expenseAccountFile)
        {
            if (expenseAccountFile.SeqNO == 0)
            {
                ExpenseAccountFile findOneMax = db.ExpenseAccountFile.Where(x => (x.ExpenseAccountId == expenseAccountFile.ExpenseAccountId)).OrderByDescending(x => x.SeqNO).FirstOrDefault();
                if (findOneMax == null)
                {
                    expenseAccountFile.SeqNO = 1;
                }
                else
                {
                    expenseAccountFile.SeqNO = findOneMax.SeqNO + 1;
                }
                expenseAccountFile.InDateTime = DateTime.Now;
                expenseAccountFile.ModifyDateTime = DateTime.Now;
                db.ExpenseAccountFile.Add(expenseAccountFile);

            }
            else
            {

            }
            db.SaveChanges();
        }
        public void ExpenseAccountFileDelete(string expenseAccountId, string seqNO)
        {
            if (seqNO == null) seqNO = "";
             SqlParameter[] para = new SqlParameter[] { new SqlParameter("@ExpenseAccountId", expenseAccountId),
                                                        new SqlParameter("@SeqNO", seqNO)};
            string sql = @"DELETE ExpenseAccountFile WHERE ExpenseAccountId = @ExpenseAccountId 
                        ";
            if (!string.IsNullOrEmpty(seqNO))
            {
                sql += " AND SeqNO = @SeqNO ";
            }
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region MonthSale
        public string MonthSaleDesryptSearch(string monthSaleId, string shopId, string yearMonth)
        {
            if (monthSaleId == null) monthSaleId = "";
            if (shopId == null) shopId = "";
            if (yearMonth == null) yearMonth = "";


            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MonthSaleId", monthSaleId),
                                                    new SqlParameter("@ShopId", shopId),
                                                 new SqlParameter("@YearMonth", yearMonth)};

            Type t = typeof(MonthSale);

            string sql = "";
            sql = @"SELECT * 
                    FROM MonthSale A
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(monthSaleId))
            {
                sql += " AND MonthSaleId = @MonthSaleId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(yearMonth))
            {
                sql += " AND A.YearMonth = @YearMonth";
            }
            List<MonthSale> list = db.Database.SqlQuery(t, sql, para).Cast<MonthSale>().ToList();
            string sqlReturn = @"CREATE TABLE #MonthSale(
	                                                [MonthSaleId] [int]  NOT NULL,
	                                                [YearMonth] [nvarchar](500) NULL,
	                                                [ShopId] [int] NULL,
	                                                [ActualSaleCount] [int] NULL,
	                                                [ActualSaleAmt] [decimal](19, 2) NULL)";

            foreach (MonthSale monthSale in list)
            {
                sqlReturn += " INSERT INTO #MonthSale VALUES(";
                sqlReturn += monthSale.MonthSaleId + ",'";
                if (monthSale.YearMonth == null)
                {
                    sqlReturn += "null','";
                }
                else
                {
                    sqlReturn += monthSale.YearMonth + "','";
                }
                if (monthSale.ShopId == null)
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += monthSale.ShopId + "','";
                }
                if (string.IsNullOrEmpty(monthSale.ActualSaleCount)||string.IsNullOrEmpty(TokenHelper.DecryptDES(monthSale.ActualSaleCount)))
                {
                    sqlReturn += "0','";
                }
                else
                {
                    sqlReturn += TokenHelper.DecryptDES(monthSale.ActualSaleCount) + "','";
                }
                if (string.IsNullOrEmpty(monthSale.ActualSaleAmt)||string.IsNullOrEmpty(TokenHelper.DecryptDES(monthSale.ActualSaleAmt)))
                {
                    sqlReturn += "0')";

                }
                else
                {
                    sqlReturn += TokenHelper.DecryptDES(monthSale.ActualSaleAmt) + "')";
                }
            }
            return sqlReturn;
        }
        public List<MonthSaleDto> MonthSaleSearch(string monthSaleId, string shopId, string yearMonth)
        {
            if (monthSaleId == null) monthSaleId = "";
            if (shopId == null) shopId = "";
            if (yearMonth == null) yearMonth = "";


            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MonthSaleId", monthSaleId),
                                                    new SqlParameter("@ShopId", shopId),
                                                 new SqlParameter("@YearMonth", yearMonth)};

            Type t = typeof(MonthSaleDto);

            string sql = "";
            sql = @"SELECT A.*,B.ShopCode,B.ShopName,B.ShopNameEn
                    FROM MonthSale A INNER JOIN Shop B ON A.ShopId = B.ShopId
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(monthSaleId))
            {
                sql += " AND MonthSaleId = @MonthSaleId";
            }
            if (!string.IsNullOrEmpty(shopId))
            {
                sql += " AND A.ShopId = @ShopId";
            }
            if (!string.IsNullOrEmpty(yearMonth))
            {
                sql += " AND A.YearMonth = @YearMonth";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<MonthSaleDto>().ToList();
        }
        public void MonthSaleSave(MonthSale monthSale)
        {
            MonthSale findOne = db.MonthSale.Where(x => (x.MonthSaleId == monthSale.MonthSaleId)).FirstOrDefault();
            if (findOne == null)
            {
                monthSale.InDateTime = DateTime.Now;
                monthSale.ModifyDateTime = DateTime.Now;
                db.MonthSale.Add(monthSale);
            }
            else
            {
                findOne.ActualSaleAmt = monthSale.ActualSaleAmt;
                findOne.ActualSaleCount = monthSale.ActualSaleCount;
                findOne.ShopId = monthSale.ShopId;
                findOne.YearMonth = monthSale.YearMonth;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = monthSale.ModifyUserId;
            }
            db.SaveChanges();
        }
        public void MonthSaleDelete(string monthSaleId)
        {
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@MonthSaleId", monthSaleId) };
            string sql = @"DELETE MonthSale WHERE MonthSaleId = @MonthSaleId
                        ";
            db.Database.ExecuteSqlCommand(sql, para);
        }
        #endregion
        #region 总览
        public List<ExpenseAccountStatusCountDto> ExpenseAccountStatusCountSearch(string year, string eventTypeId, List<Shop> roleTypeShop)
        {
            if (year == null) year = "";
            if (eventTypeId == null) eventTypeId = "";

            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@Year", year) };
            Type t = typeof(ExpenseAccountStatusCountDto);
            string sql = "";
            sql += @"SELECT ISNULL(SUM(ExpenseAccount1),0) AS ExpenseAccount1 
	                       ,ISNULL(SUM(ExpenseAccount2),0) AS ExpenseAccount2
	                       ,ISNULL(SUM(ExpenseAccount3),0) AS ExpenseAccount3
	                       ,ISNULL(SUM(ExpenseAccount4),0) AS ExpenseAccount4
	                       ,ISNULL(SUM(ExpenseAccount5),0) AS ExpenseAccount5
	                       ,ISNULL(SUM(ExpenseAccount9),0) AS ExpenseAccount9
                    FROM (
                            SELECT 
                            CASE WHEN NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode=1)
				                            THEN 1
				                            ELSE 0
			                            END AS ExpenseAccount1
                            , CASE WHEN NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode =2)
				                            THEN 1
				                            ELSE 0
			                            END AS ExpenseAccount2
                            , CASE WHEN NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode =3)
				                            THEN 1
				                            ELSE 0
			                            END AS ExpenseAccount3
                            , CASE WHEN  NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode =4)
				                            THEN 1
				                            ELSE 0
			                            END AS ExpenseAccount4
                             , CASE WHEN  NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode =5)
											THEN 1
											ELSE 0
										END AS ExpenseAccount5
                             , CASE WHEN  NOT EXISTS(SELECT 1 FROM ExpenseAccountFile WHERE ExpenseAccountId = A.ExpenseAccountId AND FileTypeCode =9)
											THEN 1
											ELSE 0
										END AS ExpenseAccount9
                            FROM ExpenseAccount A WHERE 1=1  ";
            if (!string.IsNullOrEmpty(year))
            {
                sql += " AND Year(A.StartDate) = @Year";
            }
            if (eventTypeId == "99")
            {
                sql += " AND EXISTS(SELECT 1 FROM MarketAction WHERE MarketActionId = A.MarketActionId AND EventTypeId = 99) ";
            }
            else
            {
                sql += " AND EXISTS(SELECT 1 FROM MarketAction WHERE MarketActionId = A.MarketActionId AND EventTypeId <> 99 )";
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
            return db.Database.SqlQuery(t, sql, para).Cast<ExpenseAccountStatusCountDto>().ToList();
        }
        #endregion


    }
}