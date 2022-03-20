using com.yrtech.InventoryAPI.DTO;
using System;
using com.yrtech.bentley.DAL;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace com.yrtech.InventoryAPI.Service
{
    public class ApproveService
    {
        Bentley db = new Bentley();
        #region DTTApprove
        public List<DTTApproveDto> DTTApproveSearch(string dttApproveId,string marketActionId,string dttType,string dttApproveCode)
        {
            if (dttApproveId == null) dttApproveId = "";
            if (marketActionId == null) marketActionId = "";
            if (dttType == null) dttType = "";
            if (dttApproveCode == null) dttApproveCode = "";
            SqlParameter[] para = new SqlParameter[] { new SqlParameter("@DTTApproveId", dttApproveId),
                                                    new SqlParameter("@MarketActionId", marketActionId),
                                                    new SqlParameter("@DTTType", dttType),
                                                    new SqlParameter("@DTTApproveCode", dttApproveCode)};
            Type t = typeof(DTTApproveDto);
            string sql = "";
             sql = @"SELECT A.*,B.HiddenCodeName,B.HiddenCodeNameEn 
                    FROM DTTApprove A LEFT JOIN HiddenCode B ON A.DTTApproveCode = HiddenCodeId AND B.HiddenCodeGroup = 'KeyVisionApproval'
                    WHERE 1=1";
            if (!string.IsNullOrEmpty(dttApproveId))
            {
                sql += " AND A.DTTApproveId = @DTTApproveId";
            }
            if (!string.IsNullOrEmpty(marketActionId))
            {
                sql += " AND A.MarketActionId = @MarketActionId";
            }
            if (!string.IsNullOrEmpty(dttType))
            {
                sql += " AND A.DTTType = @DTTType";
            }
            if (!string.IsNullOrEmpty(dttApproveCode))
            {
                sql += " AND DTTApproveCode = @DTTApproveCode";
            }
            return db.Database.SqlQuery(t, sql, para).Cast<DTTApproveDto>().ToList();
        }
        public void DTTApproveSave(DTTApprove dttApprove)
        {
            DTTApprove findOne = db.DTTApprove.Where(x => (x.MarketActionId == dttApprove.MarketActionId&& x.DTTType == dttApprove.DTTType)).FirstOrDefault();
            if (findOne == null)
            {
                dttApprove.InDateTime = DateTime.Now;
                dttApprove.ModifyDateTime = DateTime.Now;
                db.DTTApprove.Add(dttApprove);
            }
            else
            {
                findOne.DTTApproveCode = dttApprove.DTTApproveCode;
                findOne.DTTApproveDesc = dttApprove.DTTApproveDesc;
                findOne.ModifyDateTime = DateTime.Now;
                findOne.ModifyUserId = dttApprove.ModifyUserId;
            }
            db.SaveChanges();
        }
        #endregion

    }
}