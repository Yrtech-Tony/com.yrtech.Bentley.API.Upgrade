using System.Web.Http;
using com.yrtech.InventoryAPI.Service;
using com.yrtech.InventoryAPI.Common;
using System.Collections.Generic;
using System;
using com.yrtech.bentley;
using com.yrtech.InventoryAPI.DTO;
using com.yrtech.bentley.DAL;

namespace com.yrtech.InventoryAPI.Controllers
{
    [RoutePrefix("bentley/api")]
    public class MasterController : BaseController
    {
        MasterService masterService = new MasterService();
        ExcelDataService excelDataService = new ExcelDataService();
        #region Shop
        [HttpGet]
        [Route("Master/ShopSearch")]
        public APIResult ShopSearch(string shopId, string shopCode, string shopName, string shopNameEn)
        {
            try
            {
                List<ShopDto> shopList = masterService.ShopSearch(shopId, shopCode, shopName, shopNameEn);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(shopList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("Master/ShopSave")]
        public APIResult ShopSave(Shop shop)
        {
            try
            {
                List<ShopDto> shopList = masterService.ShopSearch("", "", shop.ShopName, "");
                if (shopList != null && shopList.Count != 0 && shopList[0].ShopId != shop.ShopId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,经销商名称重复" };
                }
                shop = masterService.ShopSave(shop);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(shop) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/ShopDelete")]
        public APIResult ShopDelete(UploadData upload)
        {
            try
            {
                List<Shop> list = CommonHelper.DecodeString<List<Shop>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (Shop shop in list)
                {
                    masterService.ShopDelete(shop.ShopId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region Area
        [HttpGet]
        [Route("Master/AreaSearch")]
        public APIResult AreaSearch(string areaId, string areaName, string areaNameEn)
        {
            try
            {
                List<Area> areaList = masterService.AreaSearch(areaId, areaName, areaNameEn);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(areaList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("Master/AreaSave")]
        public APIResult AreaSave(Area area)
        {
            try
            {
                List<Area> areaList = masterService.AreaSearch("", area.AreaName, "");
                if (areaList != null && areaList.Count != 0 && areaList[0].AreaId != area.AreaId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,区域中文名称重复" };
                }
                area = masterService.AreaSave(area);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(area) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/AreaDelete")]
        public APIResult AreaDelete(UploadData upload)
        {
            try
            {
                List<Area> list = CommonHelper.DecodeString<List<Area>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (Area area in list)
                {
                    masterService.AreaDelete(area.AreaId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region EventType
        [HttpGet]
        [Route("Master/EventTypeSearch")]
        public APIResult EventTypeSearch(string eventTypeId, string eventTypeName, string eventTypeNameEn, bool? showStatus)
        {
            try
            {
                List<EventTypeDto> eventTypeList = masterService.EventTypeSearch(eventTypeId, eventTypeName, eventTypeNameEn, showStatus);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(eventTypeList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("Master/EventTypeSave")]
        public APIResult EventTypeSave(EventType eventType)
        {
            try
            {
                List<EventTypeDto> eventTypeList = masterService.EventTypeSearch("", eventType.EventTypeName, "", null);
                if (eventTypeList != null && eventTypeList.Count != 0 && eventTypeList[0].EventTypeId != eventType.EventTypeId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,活动类型名称重复" };
                }
                eventType = masterService.EventTypeSave(eventType);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(eventType) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/EventTypeDelete")]
        public APIResult EventTypeDelete(UploadData upload)
        {
            try
            {
                List<EventType> list = CommonHelper.DecodeString<List<EventType>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (EventType eventType in list)
                {
                    masterService.EventTypeDelete(eventType.EventTypeId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region HiddenCode
        [HttpGet]
        [Route("Master/HiddenCodeSearch")]
        public APIResult HiddenCodeSearch(string hiddenCodeGroup, string hiddenCode)
        {
            try
            {
                List<HiddenCode> hiddenCodeList = masterService.HiddenCodeSearch(hiddenCodeGroup, hiddenCode, "");

                return new APIResult() { Status = true, Body = CommonHelper.Encode(hiddenCodeList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        #endregion
        #region UserInfo
        [HttpGet]
        [Route("Master/UserInfoSearch")]
        public APIResult UserInfoSearch(string userId, string accountId, string accountName, string shopCode, string shopName, string email)
        {
            try
            {
                List<UserInfoDto> userInfoList = masterService.UserInfoSearch(userId, accountId, accountName, shopCode, shopName, email,"","");
                foreach (UserInfoDto userinfo in userInfoList)
                {
                    userinfo.Password = TokenHelper.DecryptDES(userinfo.Password);
                }

                return new APIResult() { Status = true, Body = CommonHelper.Encode(userInfoList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpGet]
        [Route("Master/UserInfoExport")]
        public APIResult UserInfoExport()
        {
            try
            {
                string filePath = excelDataService.UserInfoExport();

                return new APIResult() { Status = true, Body = CommonHelper.Encode(new { FilePath = filePath }) };

            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("Master/UserInfoSave")]
        public APIResult UserInfoSave(UserInfo userInfo)
        {
            try
            {
                List<UserInfoDto> userInfoList = masterService.UserInfoSearch("", userInfo.AccountId, "", "", "", "","","");

                if (userInfoList != null && userInfoList.Count != 0 && userInfoList[0].UserId != userInfo.UserId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,账号重复" };
                }
                //List<UserInfoDto> userInfoList1 = masterService.UserInfoSearch("", "", userInfo.AccountName,"","","");
                //if (userInfoList1 != null && userInfoList1.Count != 0 && userInfoList1[0].UserId != userInfo.UserId)
                //{
                //    return new APIResult() { Status = false, Body = "保存失败,账号名称重复" };
                //}
                //userInfo.Password = TokenHelper.EncryptDES(userInfo.Password);
                userInfo = masterService.UserInfoSave(userInfo);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(userInfo) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/UserInfoPasswordChange")]
        public APIResult UserInfoPasswordChange(UserInfoDto userInfo)
        {
            try
            {
                List<UserInfoDto> userInfoList = masterService.UserInfoSearch(userInfo.UserId.ToString(), "", "", "", "", "","","");
                if (userInfoList == null || userInfoList.Count == 0)
                {
                    return new APIResult() { Status = false, Body = "该账号不存在，请确认账号" };
                }
                if (userInfoList[0].Password != userInfo.OldPassword)
                {
                    return new APIResult() { Status = false, Body = "原密码不正确，请确认密码" };
                }

                //List<UserInfoDto> userInfoList1 = masterService.UserInfoSearch("", "", userInfo.AccountName,"","","");
                //if (userInfoList1 != null && userInfoList1.Count != 0 && userInfoList1[0].UserId != userInfo.UserId)
                //{
                //    return new APIResult() { Status = false, Body = "保存失败,账号名称重复" };
                //}
                //userInfo.Password = TokenHelper.EncryptDES(userInfo.Password);
                masterService.UserInfoPasswordChange(userInfo.Password, userInfo.UserId.ToString());
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/UserInfoDelete")]
        public APIResult UserInfoDelete(UploadData upload)
        {
            try
            {
                List<UserInfo> list = CommonHelper.DecodeString<List<UserInfo>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (UserInfo userInfo in list)
                {
                    masterService.UserInfoDelete(userInfo.UserId.ToString());
                }
                return new APIResult() { Status = true, Body = "" };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        #endregion
        #region CoopFundType
        [HttpGet]
        [Route("Master/CoopFundTypeSearch")]
        public APIResult CoopFundTypeSearch(string coopFundId, string coopFundName, string coopFundNameEn, bool? showChk,string modeType="")
        {
            try
            {
                List<CoopFundType> coopFundTypeList = masterService.CoopFundTypeSearch(coopFundId, "",coopFundName, coopFundNameEn, showChk, modeType);

                return new APIResult() { Status = true, Body = CommonHelper.Encode(coopFundTypeList) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }
        }
        [HttpPost]
        [Route("Master/CoopFundTypeSave")]
        public APIResult CoopFundTypeSave(CoopFundType coopFundType)
        {
            try
            {
                List<CoopFundType> coopFundTypeList = masterService.CoopFundTypeSearch("","",coopFundType.CoopFundTypeName, "", null,"");
                if (coopFundTypeList != null && coopFundTypeList.Count != 0 && coopFundTypeList[0].CoopFundTypeId != coopFundType.CoopFundTypeId)
                {
                    return new APIResult() { Status = false, Body = "保存失败,市场基金申请费用类型名称重复" };
                }
                coopFundType = masterService.CoopFundTypeSave(coopFundType);
                return new APIResult() { Status = true, Body = CommonHelper.Encode(coopFundType) };
            }
            catch (Exception ex)
            {
                return new APIResult() { Status = false, Body = ex.Message.ToString() };
            }

        }
        [HttpPost]
        [Route("Master/CoopFundTypeDelete")]
        public APIResult CoopFundTypeDelete(UploadData upload)
        {
            try
            {
                List<CoopFundType> list = CommonHelper.DecodeString<List<CoopFundType>>(upload.ListJson);
                // 需要添加一个已经使用不能删除的验证。后期添加
                foreach (CoopFundType coopFundType in list)
                {
                    masterService.CoopFundTypeDelete(coopFundType.CoopFundTypeId.ToString());
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
