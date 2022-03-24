using System.Web.Http;
using System.Linq;
using com.yrtech.InventoryAPI.Service;
using com.yrtech.InventoryAPI.Common;
using System.Collections.Generic;
using System;
using com.yrtech.InventoryAPI.Controllers;
using com.yrtech.InventoryAPI.DTO;
using com.yrtech.bentley.DAL;
using System.Web.Configuration;

namespace com.yrtech.SurveyAPI.Controllers
{

    [RoutePrefix("bentley/api")]
    public class AnswerController : BaseController
    {
        CommitFileService commitFileService = new CommitFileService();
        MasterService masterService = new MasterService();
        MarketActionService marketActionService = new MarketActionService();
        AccountService accountService = new AccountService();
        DMFService dmfService = new DMFService();
        ExcelDataService excelDataService = new ExcelDataService();
        ApproveService approveService = new ApproveService();

        #region CommitFile
        [HttpGet]
        [Route("CommitFile/ShopCommitFileRecordStatusSearch")]
        public APIResult ShopCommitFileRecordStatusSearch(string year, string shopId, string userId, string roleTypeCode)
        {
            try
            {
                ShopCommitFileRecordListDto shopCommitFileRecordList = new ShopCommitFileRecordListDto();
                shopCommitFileRecordList.ShopCommitFileRecordStatusList = commitFileService.ShopCommitFileRecordStatusSearch(year, shopId);

                // 按照权限查询显示经销商信息
                List<Shop> roleTypeShopList = accountService.GetShopByRole(userId, roleTypeCode);
                List<ShopDto> shopListTemp = masterService.ShopSearch(shopId, "", "", "");
                List<ShopDto> shopList = new List<ShopDto>();
                foreach (ShopDto shopdto in shopListTemp)
                {
                    foreach (Shop shop in roleTypeShopList)
                    {
                        if (shopdto.ShopId == shop.ShopId)
                        {
                            shopList.Add(shopdto);
                        }
                    }
                }

                shopCommitFileRecordList.ShopList = shopList;
                shopCommitFileRecordList.CommitFileList = commitFileService.CommitFileSearch(year);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(shopCommitFileRecordList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

        [HttpGet]
        [Route("CommitFile/ShopCommitFileRecordSearch")]
        public APIResult ShopCommitFileRecordSearch(string shopId, string fileId)
        {
            try
            {
                List<ShopCommitFileRecord> shopCommitFileRecordList = commitFileService.ShopCommitFileRecordSearch(shopId, fileId);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(shopCommitFileRecordList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("CommitFile/ShopCommitFileRecordSave")]
        public APIResult ShopCommitFileRecordSave(ShopCommitFileRecord shopCommitFileRecord)
        {
            try
            {
                commitFileService.ShopCommitFileRecordSave(shopCommitFileRecord);
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("CommitFile/ShopCommitFileRecordDelete")]
        public APIResult ShopCommitFileRecordDelete([FromBody]UploadData upload)
        {
            try
            {
                List<ShopCommitFileRecord> list = CommonHelper.DecodeString<List<ShopCommitFileRecord>>(upload.ListJson);
                foreach (ShopCommitFileRecord record in list)
                {
                    commitFileService.ShopCommitFileRecordDelete(record.ShopId.ToString(), record.FileId.ToString(), record.SeqNO.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region MarketAction
        [HttpGet]
        [Route("MarketAction/MarketActionSearch")]
        public APIResult MarketActionSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, bool? expenseAccountChk, string userId, string roleTypeCode)
        {
            try
            {

                List<MarketActionDto> marketActionListTemp = marketActionService.MarketActionSearch(actionName, year, month, marketActionStatusCode, shopId, eventTypeId, expenseAccountChk);
                List<Shop> roleTypeShopList = accountService.GetShopByRole(userId, roleTypeCode);
                List<MarketActionDto> marketActionList = new List<MarketActionDto>();

                foreach (MarketActionDto marketActionDto in marketActionListTemp)
                {
                    foreach (Shop shop in roleTypeShopList)
                    {
                        if (marketActionDto.ShopId == shop.ShopId)
                        {
                            marketActionList.Add(marketActionDto);
                        }
                    }
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/MarketActionNotCancelSearch")]
        public APIResult MarketActionNotCancelSearch(string eventTypeId, string userId, string roleTypeCode)
        {
            try
            {

                List<MarketAction> marketActionListTemp = marketActionService.MarketActionNotCancelSearch(eventTypeId);
                List<Shop> roleTypeShopList = accountService.GetShopByRole(userId, roleTypeCode);
                List<MarketAction> marketActionList = new List<MarketAction>();

                foreach (MarketAction marketAction in marketActionListTemp)
                {
                    foreach (Shop shop in roleTypeShopList)
                    {
                        if (marketAction.ShopId == shop.ShopId)
                        {
                            marketActionList.Add(marketAction);
                        }
                    }
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/MarketActionSearchById")]
        public APIResult MarketActionSearchById(string marketActionId)
        {
            try
            {

                List<MarketActionDto> marketActionList = marketActionService.MarketActionSearchById(marketActionId);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/MarketActionExport")]
        public APIResult MarketActionExportSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, bool? expenseAccountChk, string userId, string roleTypeCode)
        {
            try
            {
                string filePath = excelDataService.MarketActionExport(actionName, year, month, marketActionStatusCode, shopId, eventTypeId, expenseAccountChk, userId, roleTypeCode);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };

            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/MarketActionPlanExport")]
        public APIResult MarketActionExportPlanSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, string userId, string roleTypeCode)
        {
            try
            {
                string filePath = excelDataService.MarketActionPlanExport(actionName, year, month, marketActionStatusCode, shopId, eventTypeId, userId, roleTypeCode);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };

            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/MarketActionSave")]
        public APIResult MarketActionSave(MarketAction marketAction)
        {
            try
            {
                marketActionService.MarketActionSave(marketAction);
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("MarketAction/MarketActionDelete")]
        public APIResult MarketActionDelete(UploadData upload)
        {
            try
            {
                List<MarketAction> list = CommonHelper.DecodeString<List<MarketAction>>(upload.ListJson);
                foreach (MarketAction marketAction in list)
                {
                    marketActionService.MarketActionDelete(marketAction.MarketActionId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("MarketAction/MarketActionAllLeadsReportExport")]
        public APIResult MarketActionAllLeadsReportExport(string year, string userId, string roleTypeCode)
        {
            try
            {
                string filePath = excelDataService.MarketActionAllLeadsReportExport(year, userId, roleTypeCode);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/MarketActionPicDelete")]
        public APIResult MarketActionPicDelete(UploadData upload)
        {
            try
            {
                List<MarketActionPic> list = CommonHelper.DecodeString<List<MarketActionPic>>(upload.ListJson);
                foreach (MarketActionPic marketActionPic in list)
                {
                    marketActionService.MarketActionPicDelete(marketActionPic.MarketActionId.ToString(), marketActionPic.PicType, marketActionPic.SeqNO.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #region Before4Weeks
        [HttpGet]
        [Route("MarketAction/MarketActionBefore4WeeksSearch")]
        public APIResult MarketActionBefore4WeeksSearch(string marketActionId)
        {
            try
            {
                MarketActionBefore4WeeksMainDto marketActionBefore4WeeksMainDto = new MarketActionBefore4WeeksMainDto();
                // 需要绑定的输入信息
                List<MarketActionBefore4Weeks> marketActionBefore4WeeksList = marketActionService.MarketActionBefore4WeeksSearch(marketActionId);
                List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFundList = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
                if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                {
                    // 如果活动模式为线下活动，预算的总金额=市场基金的合计
                    marketActionBefore4WeeksList[0].TotalBudgetAmt = marketActionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId, marketActionBefore4WeeksList[0].TotalBudgetAmt);
                    marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks = marketActionBefore4WeeksList[0];
                }
                marketActionBefore4WeeksMainDto.ActivityProcess = marketActionService.MarketActionBefore4WeeksActivityProcessSearch(marketActionId);
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund = marketActionBefore4WeeksCoopFundList;
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine = marketActionService.MarketActionPicSearch(marketActionId, "MPF");
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine = marketActionService.MarketActionPicSearch(marketActionId, "MPN");
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionBefore4WeeksMainDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/MarketActionBefore4WeeksSave")]
        public APIResult MarketActionBefore4WeeksSave(UploadData upload)
        {
            try
            {
                MarketActionBefore4WeeksMainDto marketActionBefore4WeeksMainDto = CommonHelper.DecodeString<MarketActionBefore4WeeksMainDto>(upload.ListJson);
                // 更新主推车型
                List<MarketActionDto> marketActionList = marketActionService.MarketActionSearchById(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                foreach (MarketActionDto marketDto in marketActionList)
                {
                    MarketAction market = new MarketAction();
                    market.ActionCode = marketDto.ActionCode;
                    market.ActionName = marketDto.ActionName;
                    market.ActionPlace = marketDto.ActionPlace;
                    market.ActivityBudget = marketDto.ActivityBudget;
                    market.EndDate = marketDto.EndDate;
                    market.EventTypeId = marketDto.EventTypeId;
                    market.ExpectLeadsCount = marketDto.ExpectLeadsCount;
                    market.ExpenseAccount = marketDto.ExpenseAccount;
                    market.MarketActionId = marketDto.MarketActionId;
                    market.MarketActionStatusCode = marketDto.MarketActionStatusCode;
                    market.MarketActionTargetModelCode = marketActionBefore4WeeksMainDto.TarketModelCode;//更新主推车型
                    market.ShopId = marketDto.ShopId;
                    market.StartDate = marketDto.StartDate;
                    marketActionService.MarketActionSave(market);
                }
                // marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks.KeyVisionPic = UploadBase64Pic("", marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks.KeyVisionPic);
                marketActionService.MarketActionBefore4WeeksSave(marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks);
                // 先全部删除活动流程，然后统一再保存
                marketActionService.MarketActionBefore4WeeksActivityProcessDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                foreach (MarketActionBefore4WeeksActivityProcess process in marketActionBefore4WeeksMainDto.ActivityProcess)
                {
                    marketActionService.MarketActionBefore4WeeksActivityProcessSave(process);
                }
                // 先全部删除市场基金申请，然后统一再保存
                marketActionService.MarketActionBefore4WeeksCoopFundDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                foreach (MarketActionBefore4WeeksCoopFund coopFund in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund)
                {
                    marketActionService.MarketActionBefore4WeeksCoopFundSave(coopFund);
                }
                //保存线上的照片
                foreach (MarketActionPic marketActionPic in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine)
                {
                    marketActionService.MarketActionPicSave(marketActionPic);
                }
                //保存线下的照片
                foreach (MarketActionPic marketActionPic in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine)
                {
                    marketActionService.MarketActionPicSave(marketActionPic);
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("MarketAction/KeyVisionSendEmailToBMC")]
        public APIResult KeyVisionSendEmailToBMC(string marketActionId)
        {
            string marketactionName = "";
            List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
            List<ShopDto> shop = new List<ShopDto>();
            if (marketAction != null && marketAction.Count > 0)
            {
                marketactionName = marketAction[0].ActionName;
                shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
            }
            try
            {
                CommonHelper.log("开始调用" + marketActionId + "-" + shop[0].ShopName + "-" + marketactionName);
                SendEmail(WebConfigurationManager.AppSettings["KeyVisionEmail_To"], WebConfigurationManager.AppSettings["KeyVisionEmail_CC"]
                        , "主视觉画面审批", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的画面审核已提交，请审核", "", "");
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                CommonHelper.log("邮件异常" + marketActionId + "-" + shop[0].ShopName + "-" + marketactionName + "-" + ex.Message.ToString());
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/KeyVisionSendEmailToShop")]
        public APIResult KeyVisionSendEmailToShop(string marketActionId)
        {
            try
            {
                string marketactionName = "";
                List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                List<ShopDto> shop = new List<ShopDto>();
                List<UserInfoDto> userinfo = new List<UserInfoDto>();
                if (marketAction != null && marketAction.Count > 0)
                {
                    marketactionName = marketAction[0].ActionName;
                    shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                    userinfo = masterService.UserInfoSearch("", "", shop[0].ShopName.ToString(), "", "", "");
                }
                // 发送给经销商时抄送给自己，以备查看
                SendEmail(userinfo[0].Email, "keyvisionApproval@163.com", "主视觉审批修改意见", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的画面审核意见已更新,请登陆DMN系统查看，并按要求完成更新", "", "");
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                //CommonHelper.log(ex.Message.ToString());
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/DMFApplyEmail")]
        public APIResult DMFApplyEmail(string marketActionId)
        {
            try
            {
                string marketactionName = "";
                List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                List<ShopDto> shop = new List<ShopDto>();
                List<UserInfoDto> userinfo = new List<UserInfoDto>();
                if (marketAction != null && marketAction.Count > 0)
                {
                    marketactionName = marketAction[0].ActionName;
                    shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                    userinfo = masterService.UserInfoSearch("", "", shop[0].ShopName.ToString(), "", "", "");
                }
                // 发送给经销商时抄送给自己，以备查看
                SendEmail("71443365@qq.com", "keyvisionApproval@163.com", "市场基金申请邮件", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的市场基金申请邮件", "", "");
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                //CommonHelper.log(ex.Message.ToString());
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/DTTApproveEmail")]
        public APIResult DTTApproveEmail(string marketActionId)
        {
            try
            {
                string marketactionName = "";
                List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                List<ShopDto> shop = new List<ShopDto>();
                List<UserInfoDto> userinfo = new List<UserInfoDto>();
                if (marketAction != null && marketAction.Count > 0)
                {
                    marketactionName = marketAction[0].ActionName;
                    shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                    userinfo = masterService.UserInfoSearch("", "", shop[0].ShopName.ToString(), "", "", "");
                }
                // 发送给经销商时抄送给自己，以备查看
                SendEmail("71443365@qq.com", "keyvisionApproval@163.com", "市场基金申请邮件", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的市场基金申请邮件", "", "");
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                //CommonHelper.log(ex.Message.ToString());
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        #endregion
        #region After2
        [HttpGet]
        [Route("MarketAction/MarketActionAfter2LeadsReportSearch")]
        public APIResult MarketActionAfter2LeadsReportSearch(string marketActionId)
        {
            try
            {
                List<MarketActionAfter2LeadsReportDto> marketAfterLeadsReportList = marketActionService.MarketActionAfter2LeadsReportSearch(marketActionId, "");
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketAfterLeadsReportList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/MarketActionAfter2LeadsReportImport")]
        public APIResult MarketActionAfter2LeadsReportImport(UploadData upload)
        {
            try
            {
                List<MarketActionAfter2LeadsReportDto> list = CommonHelper.DecodeString<List<MarketActionAfter2LeadsReportDto>>(upload.ListJson);
                foreach (MarketActionAfter2LeadsReportDto leadsReportDto in list)
                {
                    MarketActionAfter2LeadsReport leadsReport = new MarketActionAfter2LeadsReport();
                    leadsReport.BPNO = leadsReportDto.BPNO;
                    leadsReport.CustomerName = leadsReportDto.CustomerName;
                    if (leadsReportDto.DealCheckName == "是")
                    { leadsReport.DealCheck = true; }
                    else
                    {
                        leadsReport.DealCheck = false;
                    }
                    if (!string.IsNullOrEmpty(leadsReportDto.DealModelName))
                    {
                        List<HiddenCode> hiddenCodeList = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.DealModelName);
                        if (hiddenCodeList != null && hiddenCodeList.Count > 0)
                        {
                            leadsReport.DealModel = hiddenCodeList[0].HiddenCodeId;
                        }
                    }
                    if (!string.IsNullOrEmpty(leadsReportDto.InterestedModelName))
                    {
                        List<HiddenCode> hiddenCodeList_Insterested = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.InterestedModelName);
                        if (hiddenCodeList_Insterested != null && hiddenCodeList_Insterested.Count > 0)
                        {
                            leadsReport.InterestedModel = hiddenCodeList_Insterested[0].HiddenCodeId;
                        }
                    }
                    leadsReport.InUserId = leadsReportDto.InUserId;
                    if (leadsReportDto.LeadsCheckName == "是")
                    { leadsReport.LeadsCheck = true; }
                    else
                    {
                        leadsReport.LeadsCheck = false;
                    }
                    leadsReport.MarketActionId = leadsReportDto.MarketActionId;
                    leadsReport.ModifyDateTime = DateTime.Now;
                    leadsReport.ModifyUserId = leadsReportDto.ModifyUserId;
                    if (leadsReportDto.OwnerCheckName == "是")
                    { leadsReport.OwnerCheck = true; }
                    else
                    {
                        leadsReport.OwnerCheck = false;
                    }
                    leadsReport.TelNO = leadsReportDto.TelNO;
                    if (leadsReportDto.TestDriverCheckName == "是")
                    { leadsReport.TestDriverCheck = true; }
                    else
                    {
                        leadsReport.TestDriverCheck = false;
                    }
                    marketActionService.MarketActionAfter2LeadsReportSave(leadsReport);
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("MarketAction/MarketActionAfter2LeadsReportImportServer")]
        public APIResult MarketActionAfter2LeadsReportImportServer(string marketActionId, string userId, string path)
        {
            try
            {
                List<MarketActionAfter2LeadsReportDto> list = excelDataService.LeadsReportImport(path);
                foreach (MarketActionAfter2LeadsReportDto leadsReportDto in list)
                {
                    MarketActionAfter2LeadsReport leadsReport = new MarketActionAfter2LeadsReport();
                    leadsReport.BPNO = leadsReportDto.BPNO;
                    leadsReport.CustomerName = leadsReportDto.CustomerName;
                    if (leadsReportDto.DealCheckName == "是")
                    { leadsReport.DealCheck = true; }
                    else
                    {
                        leadsReport.DealCheck = false;
                    }
                    if (!string.IsNullOrEmpty(leadsReportDto.DealModelName))
                    {
                        List<HiddenCode> hiddenCodeList = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.DealModelName);
                        if (hiddenCodeList != null && hiddenCodeList.Count > 0)
                        {
                            leadsReport.DealModel = hiddenCodeList[0].HiddenCodeId;
                        }
                    }
                    if (!string.IsNullOrEmpty(leadsReportDto.InterestedModelName))
                    {
                        List<HiddenCode> hiddenCodeList_Insterested = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.InterestedModelName);
                        if (hiddenCodeList_Insterested != null && hiddenCodeList_Insterested.Count > 0)
                        {
                            leadsReport.InterestedModel = hiddenCodeList_Insterested[0].HiddenCodeId;
                        }
                    }
                    leadsReport.InUserId = Convert.ToInt32(userId);
                    if (leadsReportDto.LeadsCheckName == "是")
                    { leadsReport.LeadsCheck = true; }
                    else
                    {
                        leadsReport.LeadsCheck = false;
                    }
                    leadsReport.MarketActionId = Convert.ToInt32(marketActionId);
                    leadsReport.ModifyDateTime = DateTime.Now;
                    leadsReport.ModifyUserId = Convert.ToInt32(userId);
                    if (leadsReportDto.OwnerCheckName == "是")
                    { leadsReport.OwnerCheck = true; }
                    else
                    {
                        leadsReport.OwnerCheck = false;
                    }
                    leadsReport.TelNO = leadsReportDto.TelNO;
                    if (leadsReportDto.TestDriverCheckName == "是")
                    { leadsReport.TestDriverCheck = true; }
                    else
                    {
                        leadsReport.TestDriverCheck = false;
                    }
                    marketActionService.MarketActionAfter2LeadsReportSave(leadsReport);
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("MarketAction/MarketActionAfter2LeadsReportSave")]
        public APIResult MarketActionAfter2LeadsReportSave(MarketActionAfter2LeadsReport marketActionAfter2LeadsReport)
        {
            try
            {
                // List<MarketActionAfter2LeadsReport> list = CommonHelper.DecodeString<List<MarketActionAfter2LeadsReport>>(upload.ListJson);

                //foreach (MarketActionAfter2LeadsReport leadsReport in list)
                //{
                marketActionAfter2LeadsReport = marketActionService.MarketActionAfter2LeadsReportSave(marketActionAfter2LeadsReport);
                //}
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionAfter2LeadsReport) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("MarketAction/MarketActionAfter2LeadsReportDelete")]
        public APIResult MarketActionAfter2LeadsReportDelete(UploadData upload)
        {
            try
            {
                List<MarketActionAfter2LeadsReport> list = CommonHelper.DecodeString<List<MarketActionAfter2LeadsReport>>(upload.ListJson);
                foreach (MarketActionAfter2LeadsReport leadsReport in list)
                {
                    marketActionService.MarketActionAfter2LeadsReportDelete(leadsReport.MarketActionId.ToString(), leadsReport.SeqNO.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }

        [HttpGet]
        [Route("MarketAction/MarketActionAfter2LeadsReportExport")]
        public APIResult MarketActionAfter2LeadsReportExport(string marketActionId)
        {
            try
            {
                string filePath = excelDataService.MarketActionAfter2LeadsReportExport(marketActionId);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        #endregion
        #region After7
        [HttpGet]
        [Route("MarketAction/MarketActionAfter7Search")]
        public APIResult MarketActionAfter7Search(string marketActionId)
        {
            try
            {
                MarketActionAfter7MainDto marketActionAfter7MainDto = new MarketActionAfter7MainDto();
                // 活动报告填写信息
                List<MarketActionAfter7> marketActionAfter7List = marketActionService.MarketActionAfter7Search(marketActionId);
                // 活动计划查询信息，用于显示从活动计划关联过来的信息
                List<MarketActionBefore4Weeks> marketActionBefore4WeeksList = marketActionService.MarketActionBefore4WeeksSearch(marketActionId);
                if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                {
                    marketActionBefore4WeeksList[0].TotalBudgetAmt = marketActionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId, marketActionBefore4WeeksList[0].TotalBudgetAmt);
                    marketActionAfter7MainDto.MarketActionBefore4Weeks = marketActionBefore4WeeksList[0];
                }
                if (marketActionAfter7List != null && marketActionAfter7List.Count > 0)
                {
                    marketActionAfter7MainDto.MarketActionAfter7 = marketActionAfter7List[0];
                }
                marketActionAfter7MainDto.ActualProcess = marketActionService.MarketActionAfter7ActualProcessSearch(marketActionId);
                marketActionAfter7MainDto.MarketActionAfter7CoopFund = marketActionService.MarketActionAfter7CoopFundSearch(marketActionId);
                //List<MarketActionLeadsCountDto> marketActionLeadsCountList = marketActionService.MarketActionLeadsCountSearch(marketActionId);
                //if (marketActionLeadsCountList != null && marketActionLeadsCountList.Count > 0)
                //{
                //    marketActionAfter7MainDto.LeadsCount = marketActionLeadsCountList[0];
                //}
                marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine = marketActionService.MarketActionPicSearch(marketActionId, "MRF");
                marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine = marketActionService.MarketActionPicSearch(marketActionId, "MRN");
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionAfter7MainDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/MarketActionAfter7Save")]
        public APIResult MarketActionAfter7Save(UploadData upload)
        {
            try
            {
                MarketActionAfter7MainDto marketActionAfter7MainDto = CommonHelper.DecodeString<MarketActionAfter7MainDto>(upload.ListJson);
                marketActionService.MarketActionAfter7Save(marketActionAfter7MainDto.MarketActionAfter7);

                // 先删除再全部保存
                marketActionService.MarketActionAfter7ActualProcessDelete(marketActionAfter7MainDto.MarketActionId.ToString());
                foreach (MarketActionAfter7ActualProcess process in marketActionAfter7MainDto.ActualProcess)
                {
                    marketActionService.MarketActionAfter7ActualProcessSave(process);
                }
                // 先删除再全部保存
                marketActionService.MarketActionAfter7CoopFundDelete(marketActionAfter7MainDto.MarketActionId.ToString());
                foreach (MarketActionAfter7CoopFund marketActionAfter7CoopFund in marketActionAfter7MainDto.MarketActionAfter7CoopFund)
                {
                    marketActionService.MarketActionAfter7CoopFundSave(marketActionAfter7CoopFund);
                }
                //保存线上的照片
                foreach (MarketActionPic marketActionPic in marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine)
                {
                    marketActionService.MarketActionPicSave(marketActionPic);
                }
                //保存线下的照片
                foreach (MarketActionPic marketActionPic in marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine)
                {
                    marketActionService.MarketActionPicSave(marketActionPic);
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        //[HttpPost]
        //[Route("MarketAction/MarketActionAfter7ActualExpenseDelete")]
        //public APIResult MarketActionAfter7ActualExpenseDelete(UploadData upload)
        //{
        //    try
        //    {
        //        List<MarketActionAfter7ActualExpense> list = CommonHelper.DecodeString<List<MarketActionAfter7ActualExpense>>(upload.ListJson);
        //        foreach (MarketActionAfter7ActualExpense expense in list)
        //        {
        //            marketActionService.MarketActionAfter7ActualExpenseDelete(expense.MarketActionId.ToString(), expense.SeqNO.ToString());
        //        }
        //        return new APIResult() { Status = true, Body = "" };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new APIResult() { Status = false, Body = ex.Message.ToString() };
        //    }

        //}
        //[HttpPost]
        //[Route("MarketAction/MarketActionAfter7ActualProcessDelete")]
        //public APIResult MarketActionAfter7ActualProcessDelete(UploadData upload)
        //{
        //    try
        //    {
        //        List<MarketActionAfter7ActualProcess> list = CommonHelper.DecodeString<List<MarketActionAfter7ActualProcess>>(upload.ListJson);
        //        foreach (MarketActionAfter7ActualProcess process in list)
        //        {
        //            marketActionService.MarketActionAfter7ActualProcessDelete(process.MarketActionId.ToString(), process.SeqNO.ToString());
        //        }
        //        return new APIResult() { Status = true, Body = "" };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new APIResult() { Status = false, Body = ex.Message.ToString() };
        //    }

        //}
        #endregion
        #region 总览
        [HttpGet]
        [Route("MarketAction/MarketActionStatusCountSearch")]
        public APIResult MarketActionStatusCountSearch(string year, string eventTypeId, string userId, string roleTypeCode)
        {
            try
            {
                List<MarketActionStatusCountDto> marketActionStatusCountListDto = marketActionService.MarketActionStatusCountSearch(year, eventTypeId, accountService.GetShopByRole(userId, roleTypeCode));
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionStatusCountListDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        #endregion
        #region DTTApprove
        [HttpGet]
        [Route("MarketAction/DTTApproveSearch")]
        public APIResult DTTApproveSearch(string dttApproveId, string marketActionId, string dttType, string dttApproveCode)
        {
            try
            {
                List<DTTApproveDto> dttApproveList = approveService.DTTApproveSearch(dttApproveId, marketActionId, dttType, dttApproveCode);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(dttApproveList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("MarketAction/DTTApproveSave")]
        public APIResult DTTApproveSave(DTTApprove dttApprove)
        {
            try
            {
                approveService.DTTApproveSave(dttApprove);
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #endregion
        #region DMFItem
        [HttpGet]
        [Route("DMF/DMFItemSearch")]
        public APIResult DMFItemSearch(string dmfItemId, string dmfItemName, string dmfItemNameEn, bool? expenseAccountChk, bool? publishChk)
        {
            try
            {
                List<DMFItem> dmfItemList = dmfService.DMFItemSearch(dmfItemId, dmfItemName, dmfItemNameEn, expenseAccountChk, publishChk);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfItemList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

        [HttpPost]
        [Route("DMF/DMFItemSave")]
        public APIResult DMFItemSave(DMFItem dmfItem)
        {
            try
            {
                dmfItem = dmfService.DMFItemSave(dmfItem);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfItem) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/DMFItemDelete")]
        public APIResult DMFItemDelete(UploadData upload)
        {
            try
            {
                List<DMFItem> list = CommonHelper.DecodeString<List<DMFItem>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (DMFItem dfmItem in list)
                {
                    dmfService.DMFItemDelete(dfmItem.DMFItemId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region DMF
        [HttpGet]
        [Route("DMF/DMFSearch")]
        public APIResult DMFSearch(string shopId, string userId, string roleTypeCode)
        {
            try
            {
                List<Shop> roleTypeShopList = accountService.GetShopByRole(userId, roleTypeCode);
                List<DMFDto> dmfList = new List<DMFDto>();
                List<DMFDto> dmfListTemp = dmfService.DMFSearch(shopId);

                foreach (DMFDto dmfDto in dmfListTemp)
                {
                    foreach (Shop shop in roleTypeShopList)
                    {
                        if (dmfDto.ShopId == shop.ShopId)
                        {
                            dmfList.Add(dmfDto);

                        }
                    }
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("DMF/DMFQuarterSearch")]
        public APIResult DMFQuarterSearch(string shopId)
        {
            try
            {
                DMFQuarterMainDto dmfQuarterMainDto = new DMFQuarterMainDto();
                // 季度
                List<DMFDto> dmfList = dmfService.DMFSearch(shopId);
                List<DMFDto> dmfQuarterList = dmfService.DMFQuarterSearch(shopId);
                foreach (DMFDto quarter in dmfQuarterList)
                {
                    foreach (DMFDto dmf in dmfList)
                    {
                        if (quarter.ShopId == dmf.ShopId)
                        {
                            quarter.ActualAmt = dmf.ActualAmt;
                            quarter.DiffAmt = dmf.DiffAmt;
                        }
                    }
                }

                dmfQuarterMainDto.DMFQuarterList = dmfQuarterList;
                dmfQuarterMainDto.DMFDetailList = dmfService.DMFDetailSearch("", shopId, "", ""); ;
                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfQuarterMainDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        #endregion
        #region DMFDetail
        [HttpGet]
        [Route("DMF/DMFDetailSearch")]
        public APIResult DMFDetailSearch(string dmfDetailId, string shopId, string dmfItemId, string dmfItemName)
        {
            try
            {
                List<DMFDetailDto> dmfDetailList = dmfService.DMFDetailSearch(dmfDetailId, shopId, dmfItemId, dmfItemName);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfDetailList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

        [HttpPost]
        [Route("DMF/DMFDetailSave")]
        public APIResult DMFDetailSave(DMFDetail dmfDetail)
        {
            try
            {
                List<DMFDetailDto> detailList = dmfService.DMFDetailSearch("", dmfDetail.ShopId.ToString(), dmfDetail.DMFItemId.ToString(), "");
                if (detailList != null && detailList.Count != 0 && detailList[0].DMFDetailId != dmfDetail.DMFDetailId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,同一经销商不能添加重复项目" };
                }
                //勾选了费用报销的项目在费用报销申请，不在此处添加。暂时注释
                //List<DMFItem> itemList = dmfService.DMFItemSearch(dmfDetail.DMFItemId.ToString(), "", "", null, null);
                //if (itemList != null && itemList.Count > 0 && itemList[0].ExpenseAccountChk == false)
                //{
                //    return new APIResult() { Status = false, Body = "保存失败,不能添加费用报销项目" };
                //}
                dmfDetail = dmfService.DMFDetailSave(dmfDetail);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(dmfDetail) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/DMFDetailDelete")]
        public APIResult DMFDetailDelete(UploadData upload)
        {
            try
            {
                List<DMFDetail> list = CommonHelper.DecodeString<List<DMFDetail>>(upload.ListJson);
                foreach (DMFDetail dmfDetail in list)
                {
                    dmfService.DMFDetailDelete(dmfDetail.DMFDetailId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("DMF/DMFDetailExport")]
        public APIResult DMFDetailExport(string shopId, string dmfItemName, string userId, string roleTypeCode)
        {
            try
            {
                string filePath = excelDataService.DMFDetailExport(shopId, dmfItemName, userId, roleTypeCode);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };

            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("DMF/DMFDetailImport")]
        public APIResult DMFDetailImport(UploadData upload)
        {
            try
            {
                List<DMFDetailDto> list = CommonHelper.DecodeString<List<DMFDetailDto>>(upload.ListJson);
                foreach (DMFDetailDto dmfDetailDto in list)
                {
                    DMFDetail dmfDetail = new DMFDetail();
                    List<ShopDto> shopList = masterService.ShopSearch("", "", dmfDetailDto.ShopName, "");
                    if (shopList != null && shopList.Count > 0)
                    {
                        dmfDetail.ShopId = shopList[0].ShopId;
                    }
                    List<DMFItem> dmfItemList = dmfService.DMFItemSearch("", dmfDetailDto.DMFItemName, "", null, null);
                    if (dmfItemList != null && dmfItemList.Count > 0)
                    {
                        dmfDetail.DMFItemId = dmfItemList[0].DMFItemId;
                    }
                    List<DMFDetailDto> detailList = dmfService.DMFDetailSearch("", dmfDetail.ShopId.ToString(), dmfDetail.DMFItemId.ToString(), "");

                    dmfDetail.AcutalAmt = dmfDetailDto.AcutalAmt;
                    dmfDetail.Budget = dmfDetailDto.Budget;
                    dmfDetail.InUserId = dmfDetailDto.InUserId;
                    dmfDetail.ModifyUserId = dmfDetailDto.ModifyUserId;
                    dmfDetail.Remark = dmfDetailDto.Remark;
                    if (detailList != null && detailList.Count != 0 && detailList[0].DMFDetailId != dmfDetail.DMFDetailId)
                    {
                        return new APIResult() { Status = false, Body = "保存失败,同一经销商不能添加重复项目" };
                    }
                    dmfService.DMFDetailSave(dmfDetail);

                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }

        [HttpGet]
        [Route("DMF/DMFDetailImportServer")]
        public APIResult DMFDetailImportServer(string userId, string ossPath)
        {
            try
            {
                List<DMFDetailDto> list = excelDataService.DMFDetailImport(ossPath);

                foreach (DMFDetailDto dmfDetail in list)
                {
                    foreach (DMFDetailDto dmfDetail1 in list)
                    {
                        if (dmfDetail != dmfDetail1 && dmfDetail.ShopName == dmfDetail1.ShopName && dmfDetail.DMFItemName == dmfDetail1.DMFItemName)
                        {
                            return new APIResult() { Status = false, Body = "导入失败,存在经销商名称和市场基金项目重复的数据，请检查文件" };
                        }
                    }
                }
                // 验证数据库和excel里面是否有重复数据
                foreach (DMFDetailDto dmfDetailDto in list)
                {
                    List<ShopDto> shopList = masterService.ShopSearch("", "", dmfDetailDto.ShopName.Trim(), "");
                    if (shopList == null
                        || shopList.Count == 0)
                    {
                        return new APIResult() { Status = false, Body = "导入失败,文件中存在系统中未登记的经销商，请检查文件" };
                    }
                    if (shopList != null && shopList.Count > 0)
                    {
                        dmfDetailDto.ShopId = shopList[0].ShopId;
                    }
                    List<DMFItem> dmfItemList = dmfService.DMFItemSearch("", dmfDetailDto.DMFItemName.Trim(), "", null, null);
                    if (dmfItemList == null || dmfItemList.Count == 0)
                    {
                        return new APIResult() { Status = false, Body = "导入失败,文件中存在系统中未登记的市场基金项目，请检查文件" };
                    }
                    if (dmfItemList != null && dmfItemList.Count > 0)
                    {
                        dmfDetailDto.DMFItemId = dmfItemList[0].DMFItemId;
                    }
                    List<DMFDetailDto> dmfDetailList = dmfService.DMFDetailSearch("", dmfDetailDto.ShopId.ToString(), dmfDetailDto.DMFItemId.ToString(), "");
                    if (dmfDetailList != null && dmfDetailList.Count != 0 && dmfDetailDto.DMFDetailId != dmfDetailList[0].DMFDetailId)
                    {
                        return new APIResult() { Status = false, Body = "导入失败,文件中存在和系统重复的数据(经销商和市场基金项目同时重复)，请检查文件" };
                    }
                }
                foreach (DMFDetailDto dmfDetailDto in list)
                {
                    DMFDetail dmfDetail = new DMFDetail();
                    List<ShopDto> shopList = masterService.ShopSearch("", "", dmfDetailDto.ShopName, "");
                    if (shopList != null && shopList.Count > 0)
                    {
                        dmfDetail.ShopId = shopList[0].ShopId;
                    }
                    List<DMFItem> dmfItemList = dmfService.DMFItemSearch("", dmfDetailDto.DMFItemName, "", null, null);
                    if (dmfItemList != null && dmfItemList.Count > 0)
                    {
                        dmfDetail.DMFItemId = dmfItemList[0].DMFItemId;
                    }
                    dmfDetail.AcutalAmt = dmfDetailDto.AcutalAmt;
                    dmfDetail.Budget = dmfDetailDto.Budget;
                    dmfDetail.InUserId = Convert.ToInt32(userId);
                    dmfDetail.ModifyUserId = Convert.ToInt32(userId);
                    dmfDetail.Remark = dmfDetailDto.Remark;
                    dmfService.DMFDetailSave(dmfDetail);

                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region ExpenseAccount
        [HttpGet]
        [Route("DMF/ExpenseAccountSearch")]
        public APIResult ExpenseAccountSearch(string expenseAccountId, string shopId, string dmfItemId, string marketActionId, string userId, string roleTypeCode)
        {
            try
            {
                List<Shop> roleTypeShopList = accountService.GetShopByRole(userId, roleTypeCode);
                List<ExpenseAccountDto> expenseAccountList = new List<ExpenseAccountDto>();
                List<ExpenseAccountDto> expenseAccountListTemp = dmfService.ExpenseAccountSearch(expenseAccountId, shopId, dmfItemId, marketActionId);

                foreach (ExpenseAccountDto expenseAccountDto in expenseAccountListTemp)
                {
                    foreach (Shop shop in roleTypeShopList)
                    {
                        if (expenseAccountDto.ShopId == shop.ShopId)
                        {
                            expenseAccountDto.ExpenseAmt = TokenHelper.DecryptDES(expenseAccountDto.ExpenseAmt);
                            expenseAccountList.Add(expenseAccountDto);

                        }
                    }
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(expenseAccountList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

        [HttpPost]
        [Route("DMF/ExpenseAccountSave")]
        public APIResult ExpenseAccountSave(ExpenseAccount expenseAccount)
        {
            try
            {
                // 把活动报告的市场基金金额自动赋值到费用报销
                expenseAccount.ExpenseAmt = TokenHelper.EncryptDES(marketActionService.MarketActionAfter7Search(expenseAccount.MarketActionId.ToString())[0].CoopFundSumAmt.ToString());
                //保存费用报销
                expenseAccount = dmfService.ExpenseAccountSave(expenseAccount);
                /*活动报告的报价单，合同，发票，报价单自动赋值到费用报销,查询该活动是否已经有报销的附件.
                 * 如果已经有报销的附件，说明已经关联过，不再进行管理.
                 * 如果不存在附件，说明还没有关联，自动把活动计划和活动报告的附件关联过来*/
                List<ExpenseAccountFile> expenseAccountFileList= dmfService.ExpenseAccountFileSearch(expenseAccount.ExpenseAccountId.ToString(), "", "");
                if (expenseAccountFileList == null || expenseAccountFileList.Count == 0)
                {
                    List<MarketActionPic>marketActionPicList = new List<MarketActionPic>();
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPF01"));//活动计划报价单
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPF13"));//活动计划PPT
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF01"));//活动报告报价单
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF02")); //活动报告合同
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF03")); //活动报告发票
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF04")); //活动报告邮件截图
                    marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF15")); //活动报告PPT
                    foreach (MarketActionPic marketActionPic in marketActionPicList)
                    {
                        ExpenseAccountFile expenseAccountFile = new ExpenseAccountFile();
                        expenseAccountFile.ExpenseAccountId = expenseAccount.ExpenseAccountId;
                        expenseAccountFile.SeqNO = 0;
                        expenseAccountFile.FileName = marketActionPic.PicName;
                        if (marketActionPic.PicType == "MPF01")
                        {
                            expenseAccountFile.FileTypeCode = "1";
                        }
                        else if (marketActionPic.PicType == "MPF13")
                        {
                            expenseAccountFile.FileTypeCode = "7";
                        }
                        else if (marketActionPic.PicType == "MRF01") {
                            expenseAccountFile.FileTypeCode = "3";
                        }
                        else if (marketActionPic.PicType == "MRF02")
                        {
                            expenseAccountFile.FileTypeCode = "4";
                        }
                        else if (marketActionPic.PicType == "MRF03")
                        {
                            expenseAccountFile.FileTypeCode = "5";
                        }
                        else if (marketActionPic.PicType == "MRF04")
                        {
                            expenseAccountFile.FileTypeCode = "9";
                        }
                        else if (marketActionPic.PicType == "MRF15")
                        {
                            expenseAccountFile.FileTypeCode = "8";
                        }
                        expenseAccountFile.FileUrl = marketActionPic.PicPath;
                        expenseAccountFile.InUserId = expenseAccount.InUserId;
                        expenseAccountFile.InDateTime = expenseAccount.InDateTime;
                        expenseAccountFile.ModifyUserId = expenseAccount.ModifyUserId;
                        expenseAccountFile.ModifyDateTime = expenseAccount.ModifyDateTime;
                        dmfService.ExpenseAccountFileSave(expenseAccountFile);
                    }
                }
                
                return new APIResult() { Status = true, Body = CommonHelper.Encode(expenseAccount) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/ExpenseAccountDelete")]
        public APIResult ExpenseAccountDelete(UploadData upload)
        {
            try
            {
                List<ExpenseAccount> list = CommonHelper.DecodeString<List<ExpenseAccount>>(upload.ListJson);
                // 需要确认什么条件下不能删除
                foreach (ExpenseAccount expenseAccount in list)
                {
                    dmfService.ExpenseAccountDelete(expenseAccount.ExpenseAccountId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("DMF/ExpenseAccountExport")]
        public APIResult ExpenseAccountExport(string shopId, string userId, string roleTypeCode)
        {
            try
            {
                string filePath = excelDataService.ExpenseAccountExport(shopId, userId, roleTypeCode);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };

            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("DMF/ExpenseAccountFileSearch")]
        public APIResult ExpenseAccountFileSearch(string expenseAccountId, string seqNO, string fileType)
        {
            try
            {
                List<ExpenseAccountFile> expenseAccountFileList = dmfService.ExpenseAccountFileSearch(expenseAccountId, seqNO, fileType);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(expenseAccountFileList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

        [HttpPost]
        [Route("DMF/ExpenseAccountFileSave")]
        public APIResult ExpenseAccountFileSave(ExpenseAccountFile expenseAccount)
        {
            try
            {
                dmfService.ExpenseAccountFileSave(expenseAccount);
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/ExpenseAccountFileDelete")]
        public APIResult ExpenseAccountFileDelete(UploadData upload)
        {
            try
            {
                List<ExpenseAccountFile> list = CommonHelper.DecodeString<List<ExpenseAccountFile>>(upload.ListJson);
                foreach (ExpenseAccountFile expenseAccountFile in list)
                {
                    dmfService.ExpenseAccountFileDelete(expenseAccountFile.ExpenseAccountId.ToString(), expenseAccountFile.SeqNO.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region MonthSale
        [HttpGet]
        [Route("DMF/MonthSaleSearch")]
        public APIResult MonthSaleSearch(string monthSaleId, string shopId)
        {
            try
            {
                List<MonthSaleDto> monthSaleList = dmfService.MonthSaleSearch(monthSaleId, shopId, "");
                foreach (MonthSaleDto monthSale in monthSaleList)
                {
                    monthSale.ActualSaleAmt = TokenHelper.DecryptDES(monthSale.ActualSaleAmt);
                    monthSale.ActualSaleCount = TokenHelper.DecryptDES(monthSale.ActualSaleCount);
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(monthSaleList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("DMF/MonthSaleSave")]
        public APIResult MonthSaleSave(MonthSale monthSale)
        {
            try
            {
                List<MonthSaleDto> monthSaleList = dmfService.MonthSaleSearch("", monthSale.ShopId.ToString(), monthSale.YearMonth);
                if (monthSaleList != null && monthSaleList.Count != 0 && monthSaleList[0].MonthSaleId != monthSale.MonthSaleId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,同一经销商年月不能重复" };
                }
                dmfService.MonthSaleSave(monthSale);
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/MonthSaleImport")]
        public APIResult MonthSaleImport(UploadData upload)
        {
            try
            {
                List<MonthSaleDto> list = CommonHelper.DecodeString<List<MonthSaleDto>>(upload.ListJson);
                foreach (MonthSaleDto monthSaleDto in list)
                {
                    foreach (MonthSaleDto monthSaleDto1 in list)
                    {
                        if (monthSaleDto != monthSaleDto1 && monthSaleDto.ShopName == monthSaleDto1.ShopName && monthSaleDto.YearMonth == monthSaleDto1.YearMonth)
                        {
                            return new APIResult() { Status = false, Body = "导入失败,存在经销商名称及年月重复数据，请检查文件" };
                        }
                    }
                }
                // 暂时不验证，如果系统中已经存在的进行更新
                //foreach (MonthSaleDto monthSaleDto in list)
                //{
                //    List<ShopDto> shopList = masterService.ShopSearch("", "", monthSaleDto.ShopName, "");
                //    if (shopList != null && shopList.Count > 0)
                //    {
                //        monthSaleDto.ShopId = shopList[0].ShopId;
                //    }
                //    List<MonthSaleDto> monthSaleList = dmfService.MonthSaleSearch("", monthSaleDto.ShopId.ToString(), monthSaleDto.YearMonth);
                //    if (monthSaleList != null && monthSaleList.Count != 0 && monthSaleDto.MonthSaleId != monthSaleList[0].MonthSaleId)
                //    {
                //        return new APIResult() { Status = false, Body = "导入失败,同一经销商年月不能重复，请检查文件" };
                //    }
                //}
                foreach (MonthSaleDto monthSaleDto in list)
                {
                    MonthSale monthSale = new MonthSale();
                    List<ShopDto> shopList = masterService.ShopSearch("", "", monthSaleDto.ShopName, "");
                    if (shopList != null && shopList.Count > 0)
                    {
                        monthSale.ShopId = shopList[0].ShopId;
                    }
                    monthSale.ActualSaleAmt = monthSaleDto.ActualSaleAmt;
                    monthSale.ActualSaleCount = monthSaleDto.ActualSaleCount;
                    monthSale.InUserId = monthSaleDto.InUserId;
                    monthSale.ModifyUserId = monthSaleDto.ModifyUserId;
                    monthSale.YearMonth = monthSaleDto.YearMonth;
                    dmfService.MonthSaleSave(monthSale);

                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("DMF/MonthSaleImportServer")]
        public APIResult MonthSaleImportServer(string userId, string ossPath)
        {
            try
            {
                List<MonthSaleDto> list = excelDataService.MonthSaleImport(ossPath);
                foreach (MonthSaleDto monthSaleDto in list)
                {
                    foreach (MonthSaleDto monthSaleDto1 in list)
                    {
                        if (monthSaleDto != monthSaleDto1 && monthSaleDto.ShopName == monthSaleDto1.ShopName && monthSaleDto.YearMonth == monthSaleDto1.YearMonth)
                        {
                            return new APIResult() { Status = false, Body = "导入失败,存在经销商名称及年月重复数据，请检查文件" };
                        }
                    }
                }
                // 验证数据库和excel里面是否有重复数据
                foreach (MonthSaleDto monthSaleDto in list)
                {
                    List<ShopDto> shopList = masterService.ShopSearch("", "", monthSaleDto.ShopName.Trim(), "");
                    if (shopList == null || shopList.Count == 0)
                    {
                        return new APIResult() { Status = false, Body = "导入失败,文件中存在系统中未登记的经销商，请检查文件" };
                    }
                    if (shopList != null && shopList.Count > 0)
                    {
                        monthSaleDto.ShopId = shopList[0].ShopId;
                    }
                    List<MonthSaleDto> monthSaleList = dmfService.MonthSaleSearch("", monthSaleDto.ShopId.ToString(), monthSaleDto.YearMonth);
                    if (monthSaleList != null && monthSaleList.Count != 0 && monthSaleDto.MonthSaleId != monthSaleList[0].MonthSaleId)
                    {
                        return new APIResult() { Status = false, Body = "导入失败,文件中存在和系统重复的数据(经销商和年月同时重复)，请检查文件" };
                    }
                }
                foreach (MonthSaleDto monthSaleDto in list)
                {
                    MonthSale monthSale = new MonthSale();
                    List<ShopDto> shopList = masterService.ShopSearch("", "", monthSaleDto.ShopName, "");
                    if (shopList != null && shopList.Count > 0)
                    {
                        monthSale.ShopId = shopList[0].ShopId;
                    }
                    monthSale.ActualSaleAmt = monthSaleDto.ActualSaleAmt;
                    monthSale.ActualSaleCount = monthSaleDto.ActualSaleCount;
                    monthSale.InUserId = monthSaleDto.InUserId;
                    monthSale.ModifyUserId = monthSaleDto.ModifyUserId;
                    monthSale.YearMonth = monthSaleDto.YearMonth;
                    dmfService.MonthSaleSave(monthSale);

                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("DMF/MonthSaleDelete")]
        public APIResult MonthSaleDelete(UploadData upload)
        {
            try
            {
                List<MonthSale> list = CommonHelper.DecodeString<List<MonthSale>>(upload.ListJson);
                foreach (MonthSale monthSale in list)
                {
                    dmfService.MonthSaleDelete(monthSale.MonthSaleId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion

    }
}
