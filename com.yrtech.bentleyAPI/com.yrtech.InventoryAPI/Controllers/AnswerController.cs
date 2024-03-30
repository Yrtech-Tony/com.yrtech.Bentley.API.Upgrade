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
using System.IO;

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
        PPTService service = new PPTService();

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
        public APIResult MarketActionSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, bool? expenseAccountChk, string userId, string roleTypeCode,string areaId="")
        {
            try
            {

                List<MarketActionDto> marketActionListTemp = marketActionService.MarketActionSearch(actionName, year, month, marketActionStatusCode, shopId, eventTypeId, expenseAccountChk, areaId);
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
        public APIResult MarketActionExportSearch(string actionName, string year, string month, string marketActionStatusCode, string shopId, string eventTypeId, bool? expenseAccountChk, string userId, string roleTypeCode,string areaId="")
        {
            try
            {
                string filePath = excelDataService.MarketActionExport(actionName, year, month, marketActionStatusCode, shopId, eventTypeId, expenseAccountChk, userId, roleTypeCode,areaId);

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
        [HttpGet]
        [Route("MarketAction/MarketActionBudgetMaxSearch")]
        public APIResult MarketActionBudgetMaxSearch(string shopId = "")
        {
            try
            {
                MarketActionMaxAmtDto max = new MarketActionMaxAmtDto();
                List<MarketActionMaxAmtDto>  maxList = marketActionService.MarketActionMaxSearch(shopId);
                if (maxList != null && maxList.Count > 0)
                {
                    max = maxList[0];
                }
                return new APIResult() { Status = true, Body = CommonHelper.Encode(max) };
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
        [HttpGet]
        [Route("MarketAction/MarketActionPicSearch")]
        public APIResult MarketActionPicSearch(string marketActionId, string picType)
        {
            try
            {
                List<MarketActionPic> marketActionPicList = marketActionService.MarketActionPicSearch(marketActionId, picType);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionPicList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpGet]
        [Route("MarketAction/CreatePPT")]
        public APIResult CreatePPT(string marketActionId, string type, string userId)
        {

            try
            {
                CommonHelper.log("Fun:CreatePPT:  " + DateTime.Now.ToString() + "--" + marketActionId.ToString()+"---"+userId.ToString());
                string path = "";
                string picType = "";
                switch (type)
                {
                    case "MPF":
                        CommonHelper.log("开始调用");
                        path = service.GetActionPlanPPT(marketActionId);
                        CommonHelper.log("调用完成" + path);
                        picType = "MPF13";
                        break;
                    case "MPN":
                        path = service.GetActionPlanOnlinePPT(marketActionId);
                        picType = "MPN11";
                        break;
                    case "MRF":
                        path = service.GetActionReportPPT(marketActionId);
                        picType = "MRF15";
                        break;
                    case "MRN":
                        path = service.GetActionReportOnlinePPT(marketActionId);
                        picType = "MRN11";
                        break;
                    case "MPH":
                        path = service.GetHandOverPlatPPT(marketActionId);
                        picType = "MPH11";
                        break;
                    case "MRH":
                        path = service.GetHandOverReportPPT(marketActionId);
                        picType = "MRH11";
                        break;
                    default:
                        path = service.GetActionPlanPPT(marketActionId);
                        picType = "MPF13";
                        break;
                }

                if (!string.IsNullOrEmpty(path))
                {
                    string fileName = Path.GetFileName(path);
                    Stream stream = new FileStream(path, FileMode.Open);
                    string ossFile = @"Bentley/" + WebConfigurationManager.AppSettings["Year"] + @"/MarketAction/" + fileName;

                    OSSClientHelper.UploadOSSFile(ossFile, stream, stream.Length);
                    MarketActionService actionService = new MarketActionService();
                    MarketActionPic marketActionPic = new MarketActionPic()
                    {
                        MarketActionId = Convert.ToInt32(marketActionId),
                        InUserId = Convert.ToInt32(userId),
                        PicDesc = fileName,
                        PicName = fileName,
                        PicPath = ossFile,
                        PicType = picType,
                    };
                    // 插入之前先把之前的删除
                    actionService.MarketActionPicDelete(marketActionPic.MarketActionId.ToString(), marketActionPic.PicType, "");
                    actionService.MarketActionPicSave(marketActionPic);

                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.ToString() };
            }
        }

        [HttpGet]
        [Route("MarketAction/GetPPTContent")]
        public APIResult GetPPTContent(string file, string slide)
        {
            try
            {
                string content = service.GetContent(file, slide);
                return new APIResult() { Status = true, Body = content };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.ToString() };
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
                List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFundList = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId, "");
                List<MarketActionBefore4WeeksCoopFundDto> marketActionBefore4WeeksCoopFundDtoList = new List<MarketActionBefore4WeeksCoopFundDto>();
                if (marketActionBefore4WeeksCoopFundList != null && marketActionBefore4WeeksCoopFundList.Count > 0)
                {
                    foreach (MarketActionBefore4WeeksCoopFund marketActionBefore4WeeksCoopFund in marketActionBefore4WeeksCoopFundList)
                    {
                        MarketActionBefore4WeeksCoopFundDto marketActionBefore4WeeksCoopFundDto = new MarketActionBefore4WeeksCoopFundDto();
                        marketActionBefore4WeeksCoopFundDto.AmtPerDay = marketActionBefore4WeeksCoopFund.AmtPerDay;
                        marketActionBefore4WeeksCoopFundDto.CoopFundAmt = marketActionBefore4WeeksCoopFund.CoopFundAmt;
                        marketActionBefore4WeeksCoopFundDto.CoopFundCode = marketActionBefore4WeeksCoopFund.CoopFundCode;
                        marketActionBefore4WeeksCoopFundDto.CoopFundDesc = marketActionBefore4WeeksCoopFund.CoopFundDesc;
                        List<CoopFundType> coopFundType = new List<CoopFundType>();
                        coopFundType = masterService.CoopFundTypeSearch("", marketActionBefore4WeeksCoopFund.CoopFundCode, "", "", null, "");
                        if (coopFundType != null && coopFundType.Count > 0)
                        {
                            marketActionBefore4WeeksCoopFundDto.CoopFundTypeDesc = coopFundType[0].CoopFundTypeDesc;
                        }
                        marketActionBefore4WeeksCoopFundDto.CoopFund_DMFChk = marketActionBefore4WeeksCoopFund.CoopFund_DMFChk;
                        marketActionBefore4WeeksCoopFundDto.EndDate = marketActionBefore4WeeksCoopFund.EndDate;
                        marketActionBefore4WeeksCoopFundDto.InDateTime = marketActionBefore4WeeksCoopFund.InDateTime;
                        marketActionBefore4WeeksCoopFundDto.InUserId = marketActionBefore4WeeksCoopFund.InUserId;
                        marketActionBefore4WeeksCoopFundDto.MarketActionId = marketActionBefore4WeeksCoopFund.MarketActionId;
                        marketActionBefore4WeeksCoopFundDto.ModifyDateTime = marketActionBefore4WeeksCoopFund.ModifyDateTime;
                        marketActionBefore4WeeksCoopFundDto.ModifyUserId = marketActionBefore4WeeksCoopFund.ModifyUserId;
                        marketActionBefore4WeeksCoopFundDto.SeqNO = marketActionBefore4WeeksCoopFund.SeqNO;
                        marketActionBefore4WeeksCoopFundDto.StartDate = marketActionBefore4WeeksCoopFund.StartDate;
                        marketActionBefore4WeeksCoopFundDto.TotalDays = marketActionBefore4WeeksCoopFund.TotalDays;
                        marketActionBefore4WeeksCoopFundDtoList.Add(marketActionBefore4WeeksCoopFundDto);
                    }
                }
                if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                {
                    marketActionBefore4WeeksList[0].TotalBudgetAmt = marketActionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                    marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks = marketActionBefore4WeeksList[0];
                }
                marketActionBefore4WeeksMainDto.ActivityProcess = marketActionService.MarketActionBefore4WeeksActivityProcessSearch(marketActionId);
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksHandOverArrangement = marketActionService.MarketActionBefore4WeeksHandOverArrangementSearch(marketActionId);
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund = marketActionBefore4WeeksCoopFundDtoList;
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine = marketActionService.MarketActionPicSearch(marketActionId, "MPF");
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine = marketActionService.MarketActionPicSearch(marketActionId, "MPN");
                marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_Handover = marketActionService.MarketActionPicSearch(marketActionId, "MPH");
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
                //  marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks.KeyVisionPic = UploadBase64Pic("", marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks.KeyVisionPic);
                marketActionService.MarketActionBefore4WeeksSave(marketActionBefore4WeeksMainDto.MarketActionBefore4Weeks);
                CommonHelper.log("Fun:MarketActionBefore4WeeksSave:  "+DateTime.Now.ToString()+"--"+ marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                // 先全部删除活动流程，然后统一再保存
                marketActionService.MarketActionBefore4WeeksActivityProcessDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                if (marketActionBefore4WeeksMainDto.ActivityProcess != null && marketActionBefore4WeeksMainDto.ActivityProcess.Count > 0)
                {
                    foreach (MarketActionBefore4WeeksActivityProcess process in marketActionBefore4WeeksMainDto.ActivityProcess)
                    {
                        marketActionService.MarketActionBefore4WeeksActivityProcessSave(process);
                    }
                }
                // 先全部删除市场基金申请，然后统一再保存,

                marketActionService.MarketActionBefore4WeeksCoopFundDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                if (marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund != null && marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund.Count > 0)
                {
                    foreach (MarketActionBefore4WeeksCoopFundDto coopFundDto in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksCoopFund)
                    {
                        MarketActionBefore4WeeksCoopFund coopFund = new MarketActionBefore4WeeksCoopFund();
                        coopFund.AmtPerDay = coopFundDto.AmtPerDay;
                        coopFund.CoopFundAmt = coopFundDto.CoopFundAmt;
                        coopFund.CoopFundCode = coopFundDto.CoopFundCode;
                        coopFund.CoopFundDesc = coopFundDto.CoopFundDesc;
                        coopFund.CoopFund_DMFChk = coopFundDto.CoopFund_DMFChk;
                        coopFund.EndDate = coopFundDto.EndDate;
                        coopFund.InDateTime = coopFundDto.InDateTime;
                        coopFund.InUserId = coopFundDto.InUserId;
                        coopFund.MarketActionId = coopFundDto.MarketActionId;
                        coopFund.ModifyDateTime = coopFundDto.ModifyDateTime;
                        coopFund.ModifyUserId = coopFundDto.ModifyUserId;
                        coopFund.SeqNO = coopFundDto.SeqNO;
                        coopFund.StartDate = coopFundDto.StartDate;
                        coopFund.TotalDays = coopFundDto.TotalDays;
                        marketActionService.MarketActionBefore4WeeksCoopFundSave(coopFund);
                    }
                }
                // 先全部删除交车仪式流程，然后统一再保存
                marketActionService.MarketActionBefore4WeeksHandOverArrangementDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString());
                if (marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksHandOverArrangement != null && marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksHandOverArrangement.Count > 0)
                {
                    foreach (MarketActionBefore4WeeksHandOverArrangement handOverArrangement in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksHandOverArrangement)
                    {
                        marketActionService.MarketActionBefore4WeeksHandOverArrangementSave(handOverArrangement);
                    }
                }

                // 删除线上照片
                marketActionService.MarketActionPicDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString(), "MPN", "");
                //保存线上的照片
                if (marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine != null && marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OnLine)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                //保存线下的照片
                marketActionService.MarketActionPicDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString(), "MPF", "");
                if (marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine != null && marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_OffLine)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                //保存交车仪式的照片
                marketActionService.MarketActionPicDelete(marketActionBefore4WeeksMainDto.MarketActionId.ToString(), "MPH", "");
                if (marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_Handover != null && marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_Handover.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionBefore4WeeksMainDto.MarketActionBefore4WeeksPicList_Handover)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                // 更新主推车型和活动预算
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
                    
                    // 调整到审核通过才去更新金额
                 //   market.ActivityBudget = marketActionService.MarketActionBefore4WeeksTotalBudgetAmt(marketDto.MarketActionId.ToString()); 
                    market.ShopId = marketDto.ShopId;
                    market.StartDate = marketDto.StartDate;
                    marketActionService.MarketActionSave(market);
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("MarketAction/MarketActionBefore4WeeksCoopFundSave")]
        public APIResult MarketActionBefore4WeeksCoopFundSave(MarketActionBefore4WeeksCoopFund marketActionBefore4WeeksCoopFund)
        {
            try
            {
                CommonHelper.log("Fun:MarketActionBefore4WeeksCoopFundSave:  " + DateTime.Now.ToString() + "--" + marketActionBefore4WeeksCoopFund.MarketActionId.ToString());
                if (marketActionBefore4WeeksCoopFund != null && marketActionBefore4WeeksCoopFund.StartDate != null && marketActionBefore4WeeksCoopFund.EndDate != null)
                {
                    DateTime start = Convert.ToDateTime(Convert.ToDateTime(marketActionBefore4WeeksCoopFund.StartDate).ToShortDateString());
                    DateTime end = Convert.ToDateTime(Convert.ToDateTime(marketActionBefore4WeeksCoopFund.EndDate).ToShortDateString());
                    TimeSpan sp = end.Subtract(start);
                    marketActionBefore4WeeksCoopFund.TotalDays = sp.Days + 1;
                }
                if (marketActionBefore4WeeksCoopFund != null && marketActionBefore4WeeksCoopFund.TotalDays != null && marketActionBefore4WeeksCoopFund.TotalDays != 0)
                {
                    marketActionBefore4WeeksCoopFund.CoopFundAmt = marketActionBefore4WeeksCoopFund.CoopFundAmt == null ? 0 : marketActionBefore4WeeksCoopFund.CoopFundAmt;
                    marketActionBefore4WeeksCoopFund.AmtPerDay = marketActionBefore4WeeksCoopFund.CoopFundAmt / marketActionBefore4WeeksCoopFund.TotalDays;
                }
                //marketActionBefore4WeeksCoopFund = marketActionService.MarketActionBefore4WeeksCoopFundSave(marketActionBefore4WeeksCoopFund);
                MarketActionBefore4WeeksCoopFundDto marketActionBefore4WeeksCoopFundDto = new MarketActionBefore4WeeksCoopFundDto();
                marketActionBefore4WeeksCoopFundDto.AmtPerDay = marketActionBefore4WeeksCoopFund.AmtPerDay;
                marketActionBefore4WeeksCoopFundDto.CoopFundAmt = marketActionBefore4WeeksCoopFund.CoopFundAmt;
                marketActionBefore4WeeksCoopFundDto.CoopFundCode = marketActionBefore4WeeksCoopFund.CoopFundCode;
                marketActionBefore4WeeksCoopFundDto.CoopFundDesc = marketActionBefore4WeeksCoopFund.CoopFundDesc;
                List<CoopFundType> coopFundType = new List<CoopFundType>();
                coopFundType = masterService.CoopFundTypeSearch("", marketActionBefore4WeeksCoopFund.CoopFundCode, "", "", null, "");
                if (coopFundType != null && coopFundType.Count > 0)
                {
                    marketActionBefore4WeeksCoopFundDto.CoopFundTypeDesc = coopFundType[0].CoopFundTypeDesc;
                }
                marketActionBefore4WeeksCoopFundDto.CoopFund_DMFChk = marketActionBefore4WeeksCoopFund.CoopFund_DMFChk;
                marketActionBefore4WeeksCoopFundDto.EndDate = marketActionBefore4WeeksCoopFund.EndDate;
                marketActionBefore4WeeksCoopFundDto.InDateTime = marketActionBefore4WeeksCoopFund.InDateTime;
                marketActionBefore4WeeksCoopFundDto.InUserId = marketActionBefore4WeeksCoopFund.InUserId;
                marketActionBefore4WeeksCoopFundDto.MarketActionId = marketActionBefore4WeeksCoopFund.MarketActionId;
                marketActionBefore4WeeksCoopFundDto.ModifyDateTime = marketActionBefore4WeeksCoopFund.ModifyDateTime;
                marketActionBefore4WeeksCoopFundDto.ModifyUserId = marketActionBefore4WeeksCoopFund.ModifyUserId;
                marketActionBefore4WeeksCoopFundDto.SeqNO = marketActionBefore4WeeksCoopFund.SeqNO;
                marketActionBefore4WeeksCoopFundDto.StartDate = marketActionBefore4WeeksCoopFund.StartDate;
                marketActionBefore4WeeksCoopFundDto.TotalDays = marketActionBefore4WeeksCoopFund.TotalDays;

                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionBefore4WeeksCoopFundDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region 发送邮件
        [HttpGet]
        [Route("MarketAction/KeyVisionSendEmailToBMC")]
        public APIResult KeyVisionSendEmailToBMC(string marketActionId,string keyVisionPic)
        {
            string keyvisionpicOld = "";
            //bool keyVisionSendToBMCChk = false;
            List<MarketActionBefore4Weeks> marketActionBefore4WeeksList = marketActionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
            {
                keyvisionpicOld = marketActionBefore4WeeksList[0].KeyVisionPicOld;
                //keyVisionSendToBMCChk = marketActionBefore4WeeksList[0].KeyVisionSendToBMCChk == null? false:Convert.ToBoolean(marketActionBefore4WeeksList[0].KeyVisionSendToBMCChk); 
            }
            if (!string.IsNullOrEmpty(keyVisionPic)&&keyVisionPic != keyvisionpicOld)
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
                    //CommonHelper.log("开始调用" + marketActionId + "-" + shop[0].ShopName + "-" + marketactionName);
                    SendEmail(WebConfigurationManager.AppSettings["KeyVisionEmail_To"], WebConfigurationManager.AppSettings["KeyVisionEmail_CC"]
                            , "主视觉画面审批", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的画面审核已提交，请审核", "", "");
                    // 发完邮件更新发送状态和时间
                    if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                    {
                        marketActionBefore4WeeksList[0].KeyVisionSendToBMCChk = true;
                        marketActionBefore4WeeksList[0].KeyVisionPicOld = keyVisionPic;
                        marketActionBefore4WeeksList[0].KeyVisionSendToBMCDateTime = DateTime.Now;
                        marketActionService.MarketActionBefore4WeeksSave(marketActionBefore4WeeksList[0]);
                    }
                }
                catch (Exception ex)
                {
                    CommonHelper.log("邮件异常" + marketActionId + "-" + shop[0].ShopName + "-" + marketactionName + "-" + ex.Message.ToString());
                    return new APIResult() { Status = false, Body = ex.Message.ToString() };
                }
            }
            return new APIResult() { Status = true, Body = "" };
        }
        [HttpGet]
        [Route("MarketAction/KeyVisionSendEmailToShop")]
        public APIResult KeyVisionSendEmailToShop(string marketActionId,string keyVisionApprovalCode,string keyVisionApprovalDesc)
        {
            try
            {
               // string keyVisionApprovalCodeDB = "";
                //string keyVisionApprovalDescDB = "";
               // List<MarketActionBefore4Weeks> marketActionBefore4WeeksList = marketActionService.MarketActionBefore4WeeksSearch(marketActionId);
                //if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                //{
                //    keyVisionApprovalCodeDB = marketActionBefore4WeeksList[0].KeyVisionApprovalCode;
                //    keyVisionApprovalDescDB = marketActionBefore4WeeksList[0].KeyVisionApprovalDesc;
                //}
                //if ((keyVisionApprovalCodeDB != keyVisionApprovalCode || keyVisionApprovalDesc != keyVisionApprovalDescDB)
                //    && (keyVisionApprovalCode=="2"|| keyVisionApprovalCode=="3"))
                if(keyVisionApprovalCode == "2" || keyVisionApprovalCode == "3")
                {
                    string marketactionName = "";
                    List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                    List<ShopDto> shop = new List<ShopDto>();
                    List<UserInfoDto> userinfo = new List<UserInfoDto>();
                    if (marketAction != null && marketAction.Count > 0)
                    {
                        marketactionName = marketAction[0].ActionName;
                        shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                        userinfo = masterService.UserInfoSearch("", "", "", "", "", "", marketAction[0].ShopId.ToString(), "", "");
                    }
                    // 发送给经销商时抄送给自己，以备查看
                    SendEmail(userinfo[0].DTTEmail, "keyvisionApproval@163.com", "主视觉审批修改意见", "宾利经销商【" + shop[0].ShopName + "】的市场活动【" + marketactionName + "】的画面审核意见已更新,请登陆DMN系统查看，并按要求完成更新", "", "");
                    // 发完邮件更新发送状态和时间
                   // if (marketActionBefore4WeeksList != null && marketActionBefore4WeeksList.Count > 0)
                    //{
                       // marketActionBefore4WeeksList[0].KeyVisionSendToShopChk = true;
                       // marketActionBefore4WeeksList[0].KeyVisionSendToShopDateTime = DateTime.Now;
                        //如果是审批已经是修改的话，把经销商发送邮件的状态更新为未发送
                        //if (keyVisionApprovalCode == "3")
                 
                        //    marketActionBefore4WeeksList[0].KeyVisionSendToBMCChk = false;
                        //    marketActionBefore4WeeksList[0].KeyVisionPicOld = "";
                        //}
                       // marketActionService.MarketActionBefore4WeeksSave(marketActionBefore4WeeksList[0]);
                   // }
                }
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
        public APIResult DMFApplyEmail(string marketActionId, string type)
        {
            try
            {
                string marketactionName = "";
                string marketactionId = "";
                // string eventMode = "";
                //string fileName = "";
                string title = "";
                string content = "";
                List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                List<ShopDto> shop = new List<ShopDto>();
                List<Area> area = new List<Area>();
                List<UserInfoDto> userinfo = new List<UserInfoDto>();
                if (marketAction != null && marketAction.Count > 0)
                {
                    marketactionName = marketAction[0].ActionName;
                    marketactionId = marketAction[0].MarketActionId.ToString();
                    //if (marketAction[0].EventModeId == 1) { eventMode = "线上"; }
                    //else { eventMode = "线下"; }
                    shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                    //userinfo_area = masterService.UserInfoSearch("", "", "", "", "", "", "", shop[0].AreaId.ToString());
                    if (shop != null && shop.Count > 0)
                    {
                        area = masterService.AreaSearch(shop[0].AreaId.ToString(), "", "");
                    }
                    //userinfo = masterService.UserInfoSearch("", "", "", "", "", "", marketAction[0].ShopId.ToString(), "");
                    //fileName = marketactionId + "-" + shop[0].ShopName + "-" + "市场活动-活动计划" + eventMode + "-" + marketactionName;
                    if (type == "MP") { type = "市场活动计划"; }
                    else if (type == "MR") { type = "市场活动报告"; }
                    else if (type == "HP") { type = "交车仪式计划"; }
                    else if (type == "HR") { type = "交车仪式报告"; }
                    title = "【DMN】请审批" + marketactionId.ToString() + "-" + marketactionName + "-" + type;
                    content = "尊敬的区域负责人，" + "<br/>";
                    content += "您所在区域的经销商市场活动模块有新增提交材料，请您尽快登录DMN进行审核" + "<br/>";
                    content += "变动信息为:" + shop[0].ShopName + "-" + marketActionId.ToString() + "-" + marketactionName + "-" + type + "<br/>";
                    content += "DMN市场行动智能助理" + "<br/>";
                    content += "邮件由系统自动发送如有问题请联系区域负责同事";
                }
                // 发送给经销商时抄送给自己，以备查看
                //SendEmail(WebConfigurationManager.AppSettings["DTTApprove_To"], WebConfigurationManager.AppSettings["DTTApprove_CC"]
                //        , title, content, "", "");
                SendEmail(area[0].DTTEmail,WebConfigurationManager.AppSettings["AllEmail_CC"], title, content, "", "");
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
        public APIResult DTTApproveEmail(string marketActionId, string type)
        {
            try
            {
                string marketactionName = "";
                string dttType = "";
                string dttApproveCode = "";
                if (type == "MP") { type = "市场活动计划"; dttType = "1"; }
                else if (type == "MR") { type = "市场活动报告"; dttType = "2"; }
                else if (type == "HP") { type = "交车仪式计划"; dttType = "1"; }
                else if (type == "HR") { type = "交车仪式报告"; dttType = "2"; }
                List<DTTApproveDto> dttAproveList = approveService.DTTApproveSearch("", marketActionId, dttType, "");
                if (dttAproveList != null && dttAproveList.Count > 0)
                {
                    dttApproveCode = dttAproveList[0].DTTApproveCode;
                }
                string title = "";
                string content = "";
                List<MarketActionDto> marketAction = marketActionService.MarketActionSearchById(marketActionId);
                List<ShopDto> shop = new List<ShopDto>();
                List<UserInfoDto> userinfo_shop = new List<UserInfoDto>();
                List<UserInfoDto> userinfo_area = new List<UserInfoDto>();
                if (marketAction != null && marketAction.Count > 0)
                {
                    marketactionName = marketAction[0].ActionName;
                    marketActionId = marketAction[0].MarketActionId.ToString();
                    shop = masterService.ShopSearch(marketAction[0].ShopId.ToString(), "", "", "");
                    userinfo_shop = masterService.UserInfoSearch("", "", "", "", "", "", marketAction[0].ShopId.ToString(), "","");
                    userinfo_area = masterService.UserInfoSearch("", "", "", "", "", "", "", shop[0].AreaId.ToString(),"");
                    if (dttApproveCode == "2")
                    {
                        title = "【DMN】" + marketActionId.ToString() + "-" + marketactionName + "-" + type;
                        content = "尊敬的经销商市场经理，" + "<br/>";
                        content += "您在DMN填报的" + marketActionId.ToString() + marketactionName + type + "初审已通过，请您知悉。" + "<br/>";
                        content += "此外还提醒您，发送邮件申请至BMC区域经理邮箱并抄送德勤区域同事，以确保市场基金正常审批。" + "<br/>";
                        content += "顺颂商祺" + "<br/>";
                        content += "DMN市场行动智能助理" + "<br/>";
                        content += "邮件由系统自动发送如有问题请联系区域负责同事。";
                    }
                    else if (dttApproveCode == "3")
                    {
                        title = "【DMN】请及时修改" + marketActionId.ToString() + "-" + marketactionName + "-" + type;
                        content = "尊敬的经销商市场经理，" + "<br/>";
                        content += "您在DMN填报的" + marketActionId.ToString() + marketactionName + type + "初审未通过，请您知悉。" + "<br/>";
                        content += "请您尽快登录DMN，查看审批意见，及时修改，重新提交。感谢支持！" + "<br/>";
                        content += "顺颂商祺" + "<br/>";
                        content += "DMN市场行动智能助理" + "<br/>";
                        content += "邮件由系统自动发送如有问题请联系区域负责同事。";
                    }
                }
                if (dttApproveCode == "2" || dttApproveCode == "3")
                {
                    // 发送给经销商时抄送给自己，以备查看,不抄送区域经理了
                    //SendEmail(userinfo_shop[0].DTTEmail, userinfo_area[0].Email + "," + WebConfigurationManager.AppSettings["AllEmail_CC"], title, content, "", "");
                    SendEmail(userinfo_shop[0].DTTEmail,WebConfigurationManager.AppSettings["AllEmail_CC"], title, content, "", "");
                }
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
                    if (leadsReportDto.DCPCheckName == "是")
                    { leadsReport.DCPCheck = true; }
                    else
                    {
                        leadsReport.DCPCheck = false;
                    }
                    //if (leadsReportDto.OwnerCheckName == "是")
                    //{ leadsReport.OwnerCheck = true; }
                    //else
                    //{
                    //    leadsReport.OwnerCheck = false;
                    //}
                    //leadsReport.TelNO = leadsReportDto.TelNO;
                    //if (leadsReportDto.TestDriverCheckName == "是")
                    //{ leadsReport.TestDriverCheck = true; }
                    //else
                    //{
                    //    leadsReport.TestDriverCheck = false;
                    //}
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
                if (list == null || list.Count == 0)
                {
                    return new APIResult() { Status = false, Body = "无数据，请填写完整再上传" };
                }
                foreach (MarketActionAfter2LeadsReportDto leadsReportDto in list)
                {
                    if (string.IsNullOrEmpty(leadsReportDto.CustomerName))
                    {
                        return new APIResult() { Status = false, Body = "客户姓名不能为空，请填写完整" };
                    }
                    if (string.IsNullOrEmpty(leadsReportDto.BPNO))
                    {
                        return new APIResult() { Status = false, Body = "DCPID不能为空，请填写完整" };
                    }
                    if (string.IsNullOrEmpty(leadsReportDto.DCPCheckName))
                    {
                        return new APIResult() { Status = false, Body = "活动前是否已有DCPID不能为空，请填写完整" };
                    }
                    if (string.IsNullOrEmpty(leadsReportDto.LeadsCheckName))
                    {
                        return new APIResult() { Status = false, Body = "是否为线索不能为空，请填写完整" };
                    }
                    if (string.IsNullOrEmpty(leadsReportDto.DealCheckName))
                    {
                        return new APIResult() { Status = false, Body = "是否成交不能为空，请填写完整" };
                    }
                    if (string.IsNullOrEmpty(leadsReportDto.InterestedModelName))
                    {
                        return new APIResult() { Status = false, Body = "感兴趣车型不能为空，请填写完整" };
                    }
                    //if (string.IsNullOrEmpty(leadsReportDto.DealCheckName))
                    //{
                    //    return new APIResult() { Status = false, Body = "成交车型不能为空，请填写完整" };
                    //}
                    List<HiddenCode> hiddenCodeList_InterestedMode = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.InterestedModelName.Trim());
                    if (hiddenCodeList_InterestedMode == null
                       || hiddenCodeList_InterestedMode.Count == 0)
                    {
                        return new APIResult() { Status = false, Body = "请填写正确的感兴趣车型" };
                    }
                    if (!string.IsNullOrEmpty(leadsReportDto.DealModelName))
                    {
                        List<HiddenCode> hiddenCodeList_DealMode = masterService.HiddenCodeSearch("TargetModels", "", leadsReportDto.DealModelName.Trim());
                        if (hiddenCodeList_DealMode == null
                           || hiddenCodeList_DealMode.Count == 0)
                        {
                            return new APIResult() { Status = false, Body = "请填写正确的成交车型" };
                        }
                    }
                }
                foreach (MarketActionAfter2LeadsReportDto leadsReportDto in list)
                {
                    MarketActionAfter2LeadsReport leadsReport = new MarketActionAfter2LeadsReport();
                    leadsReport.BPNO = leadsReportDto.BPNO;
                    leadsReport.CustomerName = leadsReportDto.CustomerName;
                    if (leadsReportDto.DealCheckName == "是")
                    {
                        leadsReport.DealCheck = true;
                    }
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
                    if (leadsReportDto.DCPCheckName == "是")
                    { leadsReport.DCPCheck = true; }
                    else
                    {
                        leadsReport.DCPCheck = false;
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

                if (marketActionAfter7List != null && marketActionAfter7List.Count > 0)
                {
                    marketActionAfter7List[0].TotalBudgetAmt = marketActionService.MarketActionAfter7TotalBudgetAmt(marketActionId.ToString());
                    marketActionAfter7MainDto.MarketActionAfter7 = marketActionAfter7List[0];
                }
                // 转换为DTO，需要显示填写指引和活动计划的信息
                List<MarketActionAfter7CoopFund> marketActionAfter7CoopFundList = marketActionService.MarketActionAfter7CoopFundSearch(marketActionId);
                List<MarketActionAfter7CoopFundDto> marketActionAfterCoopFundDtoList = new List<MarketActionAfter7CoopFundDto>();
                decimal coopFundAmt_Catering = 0;
                if (marketActionAfter7CoopFundList != null && marketActionAfter7CoopFundList.Count > 0)
                {
                    foreach (MarketActionAfter7CoopFund marketActionAfter7CoopFund in marketActionAfter7CoopFundList)
                    {
                        MarketActionAfter7CoopFundDto marketActionAfter7CoopFundDto = new MarketActionAfter7CoopFundDto();
                        marketActionAfter7CoopFundDto.AmtPerDay = marketActionAfter7CoopFund.AmtPerDay;
                        marketActionAfter7CoopFundDto.CoopFundAmt = marketActionAfter7CoopFund.CoopFundAmt;
                        marketActionAfter7CoopFundDto.CoopFundCode = marketActionAfter7CoopFund.CoopFundCode;
                        marketActionAfter7CoopFundDto.CoopFundDesc = marketActionAfter7CoopFund.CoopFundDesc;
                        // 绑定指引说明
                        List<CoopFundType> coopFundType = new List<CoopFundType>();
                        coopFundType = masterService.CoopFundTypeSearch("", marketActionAfter7CoopFund.CoopFundCode, "", "", null, "");
                        if (coopFundType != null && coopFundType.Count > 0)
                        {
                            marketActionAfter7CoopFundDto.CoopFundTypeDesc = coopFundType[0].CoopFundTypeDesc;
                        }
                        // 绑定预算金额
                        if (marketActionAfter7CoopFund.CoopFundCode == "Catering")
                        {
                            List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFund_Food = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), "Catering_Food");
                            List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFund_Drink = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), "Catering_Drink");
                            if (marketActionBefore4WeeksCoopFund_Food != null && marketActionBefore4WeeksCoopFund_Food.Count > 0)
                            {
                                if (marketActionBefore4WeeksCoopFund_Food[0].CoopFundAmt != null)
                                    coopFundAmt_Catering += Convert.ToDecimal(marketActionBefore4WeeksCoopFund_Food[0].CoopFundAmt);
                            }
                            if (marketActionBefore4WeeksCoopFund_Drink != null && marketActionBefore4WeeksCoopFund_Drink.Count > 0)
                            {
                                if (marketActionBefore4WeeksCoopFund_Drink[0].CoopFundAmt != null)
                                    coopFundAmt_Catering += Convert.ToDecimal(marketActionBefore4WeeksCoopFund_Drink[0].CoopFundAmt);
                            }
                            marketActionAfter7CoopFundDto.CoopFundAmt_Budget = coopFundAmt_Catering;
                        }
                        else
                        {
                            List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFundList = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), marketActionAfter7CoopFund.CoopFundCode);
                            if (marketActionBefore4WeeksCoopFundList != null && marketActionBefore4WeeksCoopFundList.Count > 0)
                            {
                                marketActionAfter7CoopFundDto.CoopFundAmt_Budget = marketActionBefore4WeeksCoopFundList[0].CoopFundAmt;
                            }
                        }
                        marketActionAfter7CoopFundDto.CoopFund_DMFChk = marketActionAfter7CoopFund.CoopFund_DMFChk;
                        marketActionAfter7CoopFundDto.EndDate = marketActionAfter7CoopFund.EndDate;
                        marketActionAfter7CoopFundDto.InDateTime = marketActionAfter7CoopFund.InDateTime;
                        marketActionAfter7CoopFundDto.InUserId = marketActionAfter7CoopFund.InUserId;
                        marketActionAfter7CoopFundDto.MarketActionId = marketActionAfter7CoopFund.MarketActionId;
                        marketActionAfter7CoopFundDto.ModifyDateTime = marketActionAfter7CoopFund.ModifyDateTime;
                        marketActionAfter7CoopFundDto.ModifyUserId = marketActionAfter7CoopFund.ModifyUserId;
                        marketActionAfter7CoopFundDto.SeqNO = marketActionAfter7CoopFund.SeqNO;
                        marketActionAfter7CoopFundDto.StartDate = marketActionAfter7CoopFund.StartDate;
                        marketActionAfter7CoopFundDto.TotalDays = marketActionAfter7CoopFund.TotalDays;
                        marketActionAfterCoopFundDtoList.Add(marketActionAfter7CoopFundDto);
                    }
                }
                marketActionAfter7MainDto.ActualProcess = marketActionService.MarketActionAfter7ActualProcessSearch(marketActionId);
                marketActionAfter7MainDto.MarketActionAfter7CoopFund = marketActionAfterCoopFundDtoList;
                marketActionAfter7MainDto.MarketActionAfter7HandOverArrangement = marketActionService.MarketActionAfter7HandOverArrangementSearch(marketActionId);
                marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine = marketActionService.MarketActionPicSearch(marketActionId, "MRF");
                marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine = marketActionService.MarketActionPicSearch(marketActionId, "MRN");
                marketActionAfter7MainDto.MarketActionAfter7PicList_HandOver = marketActionService.MarketActionPicSearch(marketActionId, "MRH");
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
                CommonHelper.log("Fun:MarketActionAfter7Save:  " + DateTime.Now.ToString() + "--" + marketActionAfter7MainDto.MarketActionId.ToString());
                // 先删除再全部保存
                marketActionService.MarketActionAfter7ActualProcessDelete(marketActionAfter7MainDto.MarketActionId.ToString());
                if (marketActionAfter7MainDto.ActualProcess != null && marketActionAfter7MainDto.ActualProcess.Count > 0)
                {
                    foreach (MarketActionAfter7ActualProcess process in marketActionAfter7MainDto.ActualProcess)
                    {
                        marketActionService.MarketActionAfter7ActualProcessSave(process);
                    }
                }
                // 先删除再全部保存
                marketActionService.MarketActionAfter7CoopFundDelete(marketActionAfter7MainDto.MarketActionId.ToString());
                if (marketActionAfter7MainDto.MarketActionAfter7CoopFund != null && marketActionAfter7MainDto.MarketActionAfter7CoopFund.Count > 0)
                {
                    foreach (MarketActionAfter7CoopFundDto coopFundDto in marketActionAfter7MainDto.MarketActionAfter7CoopFund)
                    {
                        MarketActionAfter7CoopFund coopFund = new MarketActionAfter7CoopFund();
                        coopFund.AmtPerDay = coopFundDto.AmtPerDay;
                        coopFund.CoopFundAmt = coopFundDto.CoopFundAmt;
                        coopFund.CoopFundCode = coopFundDto.CoopFundCode;
                        coopFund.CoopFundDesc = coopFundDto.CoopFundDesc;
                        coopFund.CoopFund_DMFChk = coopFundDto.CoopFund_DMFChk;
                        coopFund.EndDate = coopFundDto.EndDate;
                        coopFund.InDateTime = coopFundDto.InDateTime;
                        coopFund.InUserId = coopFundDto.InUserId;
                        coopFund.MarketActionId = coopFundDto.MarketActionId;
                        coopFund.ModifyDateTime = coopFundDto.ModifyDateTime;
                        coopFund.ModifyUserId = coopFundDto.ModifyUserId;
                        coopFund.SeqNO = coopFundDto.SeqNO;
                        coopFund.StartDate = coopFundDto.StartDate;
                        coopFund.TotalDays = coopFundDto.TotalDays;
                        marketActionService.MarketActionAfter7CoopFundSave(coopFund);
                    }
                }
                // 先删除再全部保存
                marketActionService.MarketActionAfter7HandOverArrangementDelete(marketActionAfter7MainDto.MarketActionId.ToString());
                if (marketActionAfter7MainDto.MarketActionAfter7HandOverArrangement != null && marketActionAfter7MainDto.MarketActionAfter7HandOverArrangement.Count > 0)
                {
                    foreach (MarketActionAfter7HandOverArrangement marketActionAfter7HandOverArrangement in marketActionAfter7MainDto.MarketActionAfter7HandOverArrangement)
                    {
                        marketActionService.MarketActionAfter7HandOverArrangementSave(marketActionAfter7HandOverArrangement);
                    }
                }
                //保存线上的照片
                marketActionService.MarketActionPicDelete(marketActionAfter7MainDto.MarketActionId.ToString(), "MRN", "");
                if (marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine != null && marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionAfter7MainDto.MarketActionAfter7PicList_OnLine)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                //保存线下的照片
                marketActionService.MarketActionPicDelete(marketActionAfter7MainDto.MarketActionId.ToString(), "MRF", "");
                if (marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine != null && marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionAfter7MainDto.MarketActionAfter7PicList_OffLine)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                //保存交车仪式照片
                marketActionService.MarketActionPicDelete(marketActionAfter7MainDto.MarketActionId.ToString(), "MRH", "");
                if (marketActionAfter7MainDto.MarketActionAfter7PicList_HandOver != null && marketActionAfter7MainDto.MarketActionAfter7PicList_HandOver.Count > 0)
                {
                    foreach (MarketActionPic marketActionPic in marketActionAfter7MainDto.MarketActionAfter7PicList_HandOver)
                    {
                        marketActionService.MarketActionPicSave(marketActionPic);
                    }
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("MarketAction/MarketActionAfter7CoopFundSave")]
        public APIResult MarketActionAfter7CoopFundSave(MarketActionAfter7CoopFund marketActionAfter7CoopFund)
        {
            try
            {
                CommonHelper.log("Fun:MarketActionAfter7CoopFundSave:  " + DateTime.Now.ToString() + "--" + marketActionAfter7CoopFund.MarketActionId.ToString());
                if (marketActionAfter7CoopFund != null && marketActionAfter7CoopFund.StartDate != null && marketActionAfter7CoopFund.EndDate != null)
                {
                    DateTime start = Convert.ToDateTime(Convert.ToDateTime(marketActionAfter7CoopFund.StartDate).ToShortDateString());
                    DateTime end = Convert.ToDateTime(Convert.ToDateTime(marketActionAfter7CoopFund.EndDate).ToShortDateString());
                    TimeSpan sp = end.Subtract(start);
                    marketActionAfter7CoopFund.TotalDays = sp.Days+1;
                }
                if (marketActionAfter7CoopFund != null && marketActionAfter7CoopFund.TotalDays != null && marketActionAfter7CoopFund.TotalDays != 0)
                {
                    marketActionAfter7CoopFund.CoopFundAmt = marketActionAfter7CoopFund.CoopFundAmt == null ? 0 : marketActionAfter7CoopFund.CoopFundAmt;
                    marketActionAfter7CoopFund.AmtPerDay = marketActionAfter7CoopFund.CoopFundAmt / marketActionAfter7CoopFund.TotalDays;
                }
                //marketActionAfter7CoopFund = marketActionService.MarketActionAfter7CoopFundSave(marketActionAfter7CoopFund);
                MarketActionAfter7CoopFundDto marketActionAfter7CoopFundDto = new MarketActionAfter7CoopFundDto();
                marketActionAfter7CoopFundDto.AmtPerDay = marketActionAfter7CoopFund.AmtPerDay;
                marketActionAfter7CoopFundDto.CoopFundAmt = marketActionAfter7CoopFund.CoopFundAmt;
                marketActionAfter7CoopFundDto.CoopFundCode = marketActionAfter7CoopFund.CoopFundCode;
                marketActionAfter7CoopFundDto.CoopFundDesc = marketActionAfter7CoopFund.CoopFundDesc;
                // 费用填写指引
                List<CoopFundType> coopFundType = new List<CoopFundType>();
                coopFundType = masterService.CoopFundTypeSearch("", marketActionAfter7CoopFund.CoopFundCode, "", "", null, "");
                if (coopFundType != null && coopFundType.Count > 0)
                {
                    marketActionAfter7CoopFundDto.CoopFundTypeDesc = coopFundType[0].CoopFundTypeDesc;
                }
                // 预算金额
                decimal coopFundAmt_Catering = 0;
                // 绑定预算金额
                if (marketActionAfter7CoopFund.CoopFundCode == "Catering")
                {
                    List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFund_Food = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), "Catering_Food");
                    List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFund_Drink = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), "Catering_Drink");
                    if (marketActionBefore4WeeksCoopFund_Food != null && marketActionBefore4WeeksCoopFund_Food.Count > 0)
                    {
                        if (marketActionBefore4WeeksCoopFund_Food[0].CoopFundAmt != null)
                            coopFundAmt_Catering += Convert.ToDecimal(marketActionBefore4WeeksCoopFund_Food[0].CoopFundAmt);
                    }
                    if (marketActionBefore4WeeksCoopFund_Drink != null && marketActionBefore4WeeksCoopFund_Drink.Count > 0)
                    {
                        if (marketActionBefore4WeeksCoopFund_Drink[0].CoopFundAmt != null)
                            coopFundAmt_Catering += Convert.ToDecimal(marketActionBefore4WeeksCoopFund_Drink[0].CoopFundAmt);
                    }
                    marketActionAfter7CoopFundDto.CoopFundAmt_Budget = coopFundAmt_Catering;
                }
                else
                {
                    List<MarketActionBefore4WeeksCoopFund> marketActionBefore4WeeksCoopFundList = marketActionService.MarketActionBefore4WeeksCoopFundSearch(marketActionAfter7CoopFund.MarketActionId.ToString(), marketActionAfter7CoopFund.CoopFundCode);
                    if (marketActionBefore4WeeksCoopFundList != null && marketActionBefore4WeeksCoopFundList.Count > 0)
                    {
                        marketActionAfter7CoopFundDto.CoopFundAmt_Budget = marketActionBefore4WeeksCoopFundList[0].CoopFundAmt;
                    }
                }
                marketActionAfter7CoopFundDto.CoopFund_DMFChk = marketActionAfter7CoopFund.CoopFund_DMFChk;
                marketActionAfter7CoopFundDto.EndDate = marketActionAfter7CoopFund.EndDate;
                marketActionAfter7CoopFundDto.InDateTime = marketActionAfter7CoopFund.InDateTime;
                marketActionAfter7CoopFundDto.InUserId = marketActionAfter7CoopFund.InUserId;
                marketActionAfter7CoopFundDto.MarketActionId = marketActionAfter7CoopFund.MarketActionId;
                marketActionAfter7CoopFundDto.ModifyDateTime = marketActionAfter7CoopFund.ModifyDateTime;
                marketActionAfter7CoopFundDto.ModifyUserId = marketActionAfter7CoopFund.ModifyUserId;
                marketActionAfter7CoopFundDto.SeqNO = marketActionAfter7CoopFund.SeqNO;
                marketActionAfter7CoopFundDto.StartDate = marketActionAfter7CoopFund.StartDate;
                marketActionAfter7CoopFundDto.TotalDays = marketActionAfter7CoopFund.TotalDays;

                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionAfter7CoopFundDto) };
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
        public APIResult MarketActionStatusCountSearch(string year, string eventModeId, string userId, string roleTypeCode)
        {
            try
            {
                List<MarketActionStatusCountDto> marketActionStatusCountListDto = marketActionService.MarketActionStatusCountSearch(year, eventModeId, accountService.GetShopByRole(userId, roleTypeCode));
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionStatusCountListDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/MarketActionReportCountSearch")]
        public APIResult MarketActionReportCountSearch(string year, string eventTypeId, string userId, string roleTypeCode)
        {
            try
            {
                List<MarketActionReportCountDto> marketActionStatusCountListDto = marketActionService.MarketActionReportCountSearch(year, eventTypeId, accountService.GetShopByRole(userId, roleTypeCode));
                
                return new APIResult() { Status = true, Body = CommonHelper.Encode(marketActionStatusCountListDto) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("MarketAction/ExpenseAccountStatusCountSearch")]
        public APIResult ExpenseAccountStatusCountSearch(string year, string eventTypeId, string userId, string roleTypeCode)
        {
            try
            {
                List<ExpenseAccountStatusCountDto> expenseAccountStatusCountList = dmfService.ExpenseAccountStatusCountSearch(year, eventTypeId, accountService.GetShopByRole(userId, roleTypeCode));
                
                return new APIResult() { Status = true, Body = CommonHelper.Encode(expenseAccountStatusCountList) };
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
                if (dttApprove.DTTApproveCode == "2")  //DTT审批通过时， 更新市场活动的活动预算
                {
                   
                    List<MarketActionDto> marketActionList = marketActionService.MarketActionSearchById(dttApprove.MarketActionId.ToString());
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
                        market.MarketActionTargetModelCode = marketDto.MarketActionTargetModelCode;
                        // 调整到审核通过才去更新预算
                        market.ActivityBudget = marketActionService.MarketActionBefore4WeeksTotalBudgetAmt(marketDto.MarketActionId.ToString());
                        market.ShopId = marketDto.ShopId;
                        market.StartDate = marketDto.StartDate;
                        marketActionService.MarketActionSave(market);
                    }
                }
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
                // 先验证传过来的数据是否选择了市场活动，
                #region
                if (expenseAccount.MarketActionId != null && expenseAccount.MarketActionId != 0)
                {
                    // 如果传过来的数据包含市场活动，根据费用Id去系统里面查询数据
                    List<ExpenseAccountDto> expenseAccountList = dmfService.ExpenseAccountSearch(expenseAccount.ExpenseAccountId.ToString(), "", "", "");
                    /*如下3种情况自动同步活动报告的金额，同时把附件同步过来
                     * 1.没有查询到数据，即第一次添加这条费用.
                     * 2.查询到数据，市场活动Id不存在,即添加过费用报销，但是没有选择过市场活动
                     * 3.查询到数据，市场活动Id存在，但是和传过来的市场活动Id不同，即重新选择了市场活动
                    */
                    if ((expenseAccountList == null || expenseAccountList.Count == 0)
                        || (expenseAccountList != null && expenseAccountList.Count > 0 && (expenseAccountList[0].MarketActionId == null || expenseAccountList[0].MarketActionId == 0))
                        || (expenseAccountList != null && expenseAccountList.Count > 0 && expenseAccount.MarketActionId != expenseAccountList[0].MarketActionId))
                    {
                        // 把活动报告的市场基金金额自动赋值到费用报销
                        List<MarketActionAfter7> marketActionAfter7List = marketActionService.MarketActionAfter7Search(expenseAccount.MarketActionId.ToString());
                        decimal? expenseAmt = 0;
                        if (marketActionAfter7List != null && marketActionAfter7List.Count > 0)
                        {
                            expenseAmt = marketActionAfter7List[0].CoopFundSumAmt;
                        }
                        expenseAccount.ExpenseAmt = TokenHelper.EncryptDES(expenseAmt.ToString());
                        //保存费用报销
                        expenseAccount = dmfService.ExpenseAccountSave(expenseAccount);
                        /*
                         * 把附件关联过来
                         */
                        List<MarketActionPic> marketActionPicList = new List<MarketActionPic>();
                        // 活动计划-报价单
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPF01"));//活动计划报价单-线下
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPN01"));//活动计划报价单-线上
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPH01"));//活动计划报价单-交车仪式
                                                                                                                                                  
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPF20"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPN02"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPH02"));
                        // 活动计划PPT
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPF13"));//活动计划PPT
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPN11"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MPH11"));
                        // 活动报告-报价单
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF01"));//活动报告报价单
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRN01"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRH01"));
                        // 活动报告-合同
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF02"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRN02"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRH02"));
                        // 活动报告-发票
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF03"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRN03"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRH03"));
                        // 活动报告-邮件截图
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF04"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRN04"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRH04"));
                        // 活动报告-ppt
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRF15"));//活动计划PPT
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRN11"));
                        marketActionPicList.AddRange(marketActionService.MarketActionPicSearch(expenseAccount.MarketActionId.ToString(), "MRH11"));

                        // 保存附件之前，先把已有的附件删除，即更换活动时，先把之前的附件删除
                        dmfService.ExpenseAccountFileDelete(expenseAccount.ExpenseAccountId.ToString(), "");
                        foreach (MarketActionPic marketActionPic in marketActionPicList)
                        {
                            ExpenseAccountFile expenseAccountFile = new ExpenseAccountFile();
                            expenseAccountFile.ExpenseAccountId = expenseAccount.ExpenseAccountId;
                            expenseAccountFile.SeqNO = 0;
                            expenseAccountFile.FileName = marketActionPic.PicName;
                            if (marketActionPic.PicType == "MPF01" || marketActionPic.PicType == "MPN01" || marketActionPic.PicType == "MPH01")
                            {
                                expenseAccountFile.FileTypeCode = "1";
                            }
                            if (marketActionPic.PicType == "MPF20" || marketActionPic.PicType == "MPN02" || marketActionPic.PicType == "MPH02")
                            {
                                expenseAccountFile.FileTypeCode = "2";
                            }
                            else if (marketActionPic.PicType == "MPF13" || marketActionPic.PicType == "MPN11" || marketActionPic.PicType == "MPH11")
                            {
                                expenseAccountFile.FileTypeCode = "7";
                            }
                            else if (marketActionPic.PicType == "MRF01" || marketActionPic.PicType == "MRN01" || marketActionPic.PicType == "MRH01")
                            {
                                expenseAccountFile.FileTypeCode = "5";
                            }
                            else if (marketActionPic.PicType == "MRF02" || marketActionPic.PicType == "MRN02" || marketActionPic.PicType == "MRH02")
                            {
                                expenseAccountFile.FileTypeCode = "3";
                            }
                            else if (marketActionPic.PicType == "MRF03" || marketActionPic.PicType == "MRN03" || marketActionPic.PicType == "MRH03")
                            {
                                expenseAccountFile.FileTypeCode = "4";
                            }
                            else if (marketActionPic.PicType == "MRF04" || marketActionPic.PicType == "MRN04" || marketActionPic.PicType == "MRH04")
                            {
                                expenseAccountFile.FileTypeCode = "9";
                            }
                            else if (marketActionPic.PicType == "MRF15" || marketActionPic.PicType == "MRN11" || marketActionPic.PicType == "MRH11")
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
                    // 3种情况以外，直接保存
                    else
                    {
                        expenseAccount = dmfService.ExpenseAccountSave(expenseAccount);
                    }
                }
                #endregion
                // 如果没有选择市场活动，直接保存
                else
                {
                    expenseAccount = dmfService.ExpenseAccountSave(expenseAccount);
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
        public APIResult MonthSaleSearch(string monthSaleId, string shopId,string yearMonth="")
        {
            try
            {
                List<MonthSaleDto> monthSaleList = dmfService.MonthSaleSearch(monthSaleId, shopId, yearMonth);
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

        // 手动删除数据使用
        [HttpGet]
        [Route("MarketAction/DeleteMarketFile")]
        public APIResult DeleteMarketFile()
        {
            try
            {
                List<MarketActionDto> marketList = marketActionService.MarketActionSearch("", "", "", "", "", "", null, "");
                List<MarketActionPic> picList = new List<MarketActionPic>();
                foreach (MarketActionDto marketAction in marketList)
                {
                    if(marketAction.MarketActionId<2854&& marketAction.MarketActionId>2774)
                    picList.AddRange(marketActionService.MarketActionPicSearch(marketAction.MarketActionId.ToString(), ""));
                }
                foreach (MarketActionPic pic in picList)
                {
                    OSSClientHelper.DeleteObject(pic.PicPath);
                    marketActionService.MarketActionPicDelete(pic.MarketActionId.ToString(),pic.PicType,pic.SeqNO.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }

    }
}
