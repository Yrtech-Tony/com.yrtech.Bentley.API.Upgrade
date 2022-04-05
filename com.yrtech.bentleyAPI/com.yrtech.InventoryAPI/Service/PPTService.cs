using Aspose.Slides;
using com.yrtech.bentley.DAL;
using com.yrtech.InventoryAPI.Common;
using com.yrtech.InventoryAPI.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Hosting;
namespace com.yrtech.InventoryAPI.Service
{
    public class PPTService
    {
        /**
         *  VenueRetal 场地租赁
Setup 搭建
Catering_Food 餐饮_餐费
Catering_Drink 餐饮_酒水
PhotoGraphy 摄影
Performance 表演
MC 支持人
Hospitality 礼仪
Others 其他
Catering 餐饮
         * */
        string[] ActionPlanUnderBudgetTypes = { "VenueRetal", "PhotoGraphy", "Setup", "Performance", "Catering_Food", "MC", "Catering_Drink", "MC", "Others", "Hospitality" };
        string[] ActionReportUnderBudgetTypes = { "VenueRetal", "Setup", "PhotoGraphy", "Performance", "MC", "Hospitality", "Catering", "Others" };
        string[] HandOverReportUnderBudgetTypes  = { "VenueRetal", "Setup", "PhotoGraphy", "Catering", "Others" };
        string[] HandOverPlanUnderBudgetTypes = { "VenueRetal", "PhotoGraphy", "Setup", "Others", "Catering_Food","", "Catering_Drink","" };
        string[] ActionPlanOnlineBudgetTypes = { "BaiduKeyWords", "OnLineLeads", "MediaBuy" };
        string[] ActionReportOnlineBudgetTypes = { "BaiduKeyWords", "OnLineLeads", "MediaBuy" };

        public string GetContent(string file, string slide)
        {
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(file);
            ISlide thirdSlide = helper.GetSlide(int.Parse(slide));
            string content = "";
            for (int s = 0; s < thirdSlide.Shapes.Count; s++)
            {
                IShape shape = thirdSlide.Shapes[s];
                if (shape.GetType().Name == "Table")
                {
                    Aspose.Slides.Table table = (Aspose.Slides.Table)shape;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        for (int j = 0; j < table.Columns.Count; j++)
                        {
                            string cell = table[j, i].TextFrame.Text;
                            content += string.Format("cell({0},{1}) = {2}", i + 1, j + 1, cell);
                        }
                    }
                }
                if (shape.GetType().Name == "AutoShape")
                {
                    Aspose.Slides.AutoShape autoShape = (Aspose.Slides.AutoShape)shape;
                    content += autoShape.TextFrame.Text;
                }
            }

            return content;
        }
        /// <summary>
        /// 生成活动计划PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetActionPlanPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\PlanOffLine.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";
            //第二页 活动总览 Overview
            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
            if (lst.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape shape = helper.GetShape(secSlide, 2);
                helper.SaveTableCell(shape, 2, 2, actionName);
                helper.SaveTableCell(shape, 3, 2, date);
                helper.SaveTableCell(shape, 3, 5, place);

            }
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 2);
                helper.SaveTableCell(table1, 4, 3, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 5, 3, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 6, 3, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 7, 3, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                helper.SaveTableCell(table1, 4, 6, before4Weeks[0].Vehide_Usage);
                helper.SaveTableCell(table1, 6, 6, before4Weeks[0].Vehide_Model);
                helper.SaveTableCell(table1, 7, 6, IntNullabelToString(before4Weeks[0].Vehide_Qty));

                IShape table2 = helper.GetShape(secSlide, 5);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

            }
            //第3页 Event Budget 费用总览
            if (before4Weeks.Count > 0)
            {
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table3 = helper.GetShape(thirdSlide, 2);
                helper.SaveTableCell(table3, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table3, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));
            }

            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table3 = helper.GetShape(thirdSlide, 2);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionPlanUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 7 + index / 2;
                    int col = 3 + (index % 2) * 4;
                    helper.SaveTableCell(table3, row, col, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table3, row, col + 1, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table3, row, col + 2, item.CoopFundDesc);
                });
            }

            //第4页 Venue 场地简介
            List<MarketActionPic> MPF02Pics = actionService.MarketActionPicSearch(marketActionId, "MPF02");
            if (MPF02Pics.Count > 0)
            {
                //绑定场地实景照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF02Pics.ForEach(item =>
                {
                    int index = MPF02Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(4), pic);

            }
            //场地选择理由 TODO


            //第5页 Venue 场地简介 内部照片  平面图
            List<MarketActionPic> MPF03Pics = actionService.MarketActionPicSearch(marketActionId, "MPF03");
            if (MPF03Pics.Count > 0)
            {
                //绑定场地内部照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF03Pics.ForEach(item =>
                {
                    int index = MPF03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);

            }
            List<MarketActionPic> MPF04Pics = actionService.MarketActionPicSearch(marketActionId, "MPF04");
            if (MPF04Pics.Count > 0)
            {
                //绑定场地使用计划 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF04Pics.ForEach(item =>
                {
                    int index = MPF04Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            // 第6页 Brand Representation – KV 活动主视觉或背板设计 
            if (before4Weeks.Count > 0 && !string.IsNullOrEmpty(before4Weeks[0].KeyVisionPic))
            {
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
            }

            //第7页 Event Setup 场地布置
            List<MarketActionPic> MPF05Pics = actionService.MarketActionPicSearch(marketActionId, "MPF05");
            if (MPF05Pics.Count > 0)
            {
                //绑定场地搭建方案 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF05Pics.ForEach(item =>
                {
                    int index = MPF05Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }
            List<MarketActionPic> MPF06Pics = actionService.MarketActionPicSearch(marketActionId, "MPF06");
            if (MPF06Pics.Count > 0)
            {
                //绑定场地效果图 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF06Pics.ForEach(item =>
                {
                    int index = MPF06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }

            //第8页 Performance 表演
            List<MarketActionPic> MPF07Pics = actionService.MarketActionPicSearch(marketActionId, "MPF07");
            if (MPF07Pics.Count > 0)
            {
                //绑定表演计划 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF07Pics.ForEach(item =>
                {
                    int index = MPF07Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }
            List<MarketActionPic> MPF08Pics = actionService.MarketActionPicSearch(marketActionId, "MPF08");
            if (MPF08Pics.Count > 0)
            {
                //绑定表演方案 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF08Pics.ForEach(item =>
                {
                    int index = MPF08Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }

            //第9页 Photography 摄影摄像
            List<MarketActionPic> MPF09Pics = actionService.MarketActionPicSearch(marketActionId, "MPF09");
            if (MPF09Pics.Count > 0)
            {
                //绑定摄影师介绍 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF09Pics.ForEach(item =>
                {
                    int index = MPF09Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }
            List<MarketActionPic> MPF10Pics = actionService.MarketActionPicSearch(marketActionId, "MPF10");
            if (MPF10Pics.Count > 0)
            {
                //绑定摄影师作品 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF10Pics.ForEach(item =>
                {
                    int index = MPF10Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }

            //第10页 Hospitality 礼仪 others其他
            List<MarketActionPic> MPF11Pics = actionService.MarketActionPicSearch(marketActionId, "MPF11");
            if (MPF11Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF11Pics.ForEach(item =>
                {
                    int index = MPF11Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }
            List<MarketActionPic> MPF12Pics = actionService.MarketActionPicSearch(marketActionId, "MPF12");
            if (MPF12Pics.Count > 0)
            {
                //绑定其他 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF12Pics.ForEach(item =>
                {
                    int index = MPF12Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }
            //第11页 活动流程
            List<MarketActionBefore4WeeksActivityProcess> before4WeeksActivitys = actionService.MarketActionBefore4WeeksActivityProcessSearch(marketActionId);
            if (before4WeeksActivitys.Count > 0)
            {
                //绑定活动流程
                ISlide elevenSlide = helper.GetSlide(11);
                IShape table2 = helper.GetShape(elevenSlide, 2);
                before4WeeksActivitys.ForEach(item =>
                {
                    int index = before4WeeksActivitys.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.ActivityDateTime);
                    helper.SaveTableCell(table2, row, 2, item.Contents);
                    helper.SaveTableCell(table2, row, 3, item.Responsible);
                });
            }

            string dirPath = basePath + @"\Temp\";
            string path = dirPath + marketActionId.ToString()+"-"+shopName+"-市场活动-活动计划"+eventModeName+"-"+actionName+"-"+ DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }
        /// <summary>
        /// 生成活动计划报告PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetActionReportPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\ReportOffLine.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";
            //第二页 活动总览 Overview
            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
            //第2页 Overview 概述
            if (lst.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 2, 2, actionName);
                helper.SaveTableCell(table1, 2, 5, date);
                helper.SaveTableCell(table1, 3, 2, place);

            }
            List<MarketActionAfter7> actionAfter7 = actionService.MarketActionAfter7Search(marketActionId);
            if (actionAfter7.Count > 0)
            {
                actionAfter7[0].TotalBudgetAmt = actionService.MarketActionAfter7TotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 5, 3, IntNullabelToString(actionAfter7[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 3, IntNullabelToString(actionAfter7[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 3, GetCostPerLead(actionAfter7[0].TotalBudgetAmt, actionAfter7[0].People_NewLeadsThsYearCount));
                helper.SaveTableCell(table1, 8, 3, IntNullabelToString(actionAfter7[0].People_NewLeadsThsYearCount));

            }
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 5, 5, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 5, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 5, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 8, 5, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                IShape table2 = helper.GetShape(secSlide, 7);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

            }

            //第3页 线索报告
            List<MarketActionAfter2LeadsReportDto> after2LeadsReport = actionService.MarketActionAfter2LeadsReportSearch(marketActionId, "");
            if (after2LeadsReport.Count > 0)
            {
                //绑定线索报告
                ISlide fiveSlide = helper.GetSlide(3);
                IShape table2 = helper.GetShape(fiveSlide, 5);
                after2LeadsReport.ForEach(item =>
                {
                    int index = after2LeadsReport.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.CustomerName);
                    helper.SaveTableCell(table2, row, 2, item.TelNO);
                    helper.SaveTableCell(table2, row, 3, BoolNullabelToString(item.TestDriverCheck));
                    helper.SaveTableCell(table2, row, 4, item.InterestedModelName);
                    helper.SaveTableCell(table2, row, 5, BoolNullabelToString(item.DealCheck));
                    helper.SaveTableCell(table2, row, 6, item.DealCheckName);
                });
            }

            //第4页 Event Budget 费用总览 
            //Spending Overview 费用总计 实际
            if (actionAfter7.Count > 0)
            {
                //Event Budget 费用总览
                ISlide slide = helper.GetSlide(4);
                IShape table2 = helper.GetShape(slide, 5);
                helper.SaveTableCell(table2, 2, 2, DecimalNullabelToString(actionAfter7[0].TotalBudgetAmt));
                helper.SaveTableCell(table2, 3, 2, DecimalNullabelToString(actionAfter7[0].CoopFundSumAmt));
            }
            //Spending Overview 费用总计 预算
            if (before4Weeks.Count > 0)
            {
                //Event Budget 费用总览
                ISlide slide = helper.GetSlide(4);
                IShape table2 = helper.GetShape(slide, 5);
                helper.SaveTableCell(table2, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table2, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));
            }
            List<MarketActionAfter7CoopFund> after7CoopFund = actionService.MarketActionAfter7CoopFundSearch(marketActionId);
            if (after7CoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情 Actual Cost
                ISlide slide = helper.GetSlide(4);
                IShape table1 = helper.GetShape(slide, 6);
                after7CoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionReportUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 2 + index;
                    helper.SaveTableCell(table1, row, 2, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table1, row, 4, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table1, row, 5, item.CoopFundDesc);
                });
            }
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情  Plan Budget预算列
                ISlide slide = helper.GetSlide(4);
                IShape table1 = helper.GetShape(slide, 6);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionReportUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 2 + index;
                    helper.SaveTableCell(table1, row, 3, DecimalNullabelToString(item.CoopFundAmt));
                });
            }

            //第5页 活动流程
            List<MarketActionAfter7ActualProcess> after7ActualProcess = actionService.MarketActionAfter7ActualProcessSearch(marketActionId);
            if (after7ActualProcess.Count > 0)
            {
                //绑定活动流程
                ISlide fiveSlide = helper.GetSlide(5);
                IShape table2 = helper.GetShape(fiveSlide, 2);
                after7ActualProcess.ForEach(item =>
                {
                    int index = after7ActualProcess.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.ActivityDateTime);
                    helper.SaveTableCell(table2, row, 2, item.Process);
                });
            }

            //第6页 Venue 场地照片 
            List<MarketActionPic> MRF05Pics = actionService.MarketActionPicSearch(marketActionId, "MRF05");
            if (MRF05Pics.Count > 0)
            {
                //绑定场地实景 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF05Pics.ForEach(item =>
                {
                    int index = MRF05Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
            }
            List<MarketActionPic> MRF06Pics = actionService.MarketActionPicSearch(marketActionId, "MRF06");
            if (MRF06Pics.Count > 0)
            {
                //绑定场地内部 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF06Pics.ForEach(item =>
                {
                    int index = MRF06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
            }

            //第7页
            List<MarketActionPic> MRF08Pics = actionService.MarketActionPicSearch(marketActionId, "MRF08");
            if (MRF08Pics.Count > 0)
            {
                //绑定车辆 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF08Pics.ForEach(item =>
                {
                    int index = MRF08Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }
            //第8页 Brand Representation – KV 活动主视觉或背板设计
            if (before4Weeks.Count > 0 && !string.IsNullOrEmpty(before4Weeks[0].KeyVisionPic))
            {
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }
            //第9页
            List<MarketActionPic> MRF09Pics = actionService.MarketActionPicSearch(marketActionId, "MRF09");
            if (MRF09Pics.Count > 0)
            {
                //绑定场地布置 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF09Pics.ForEach(item =>
                {
                    int index = MRF09Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }

            //第10页 Performance and Sign-in 表演及现场互动
            List<MarketActionPic> MRF10Pics = actionService.MarketActionPicSearch(marketActionId, "MRF10");
            if (MRF10Pics.Count > 0)
            {
                //绑定表演 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF10Pics.ForEach(item =>
                {
                    int index = MRF10Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }

            List<MarketActionPic> MRF11Pics = actionService.MarketActionPicSearch(marketActionId, "MRF11");
            if (MRF11Pics.Count > 0)
            {
                //绑定现场互动 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF11Pics.ForEach(item =>
                {
                    int index = MRF11Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }

            //第11页 礼仪、其他工作人员
            List<MarketActionPic> MRF12Pics = actionService.MarketActionPicSearch(marketActionId, "MRF12");
            if (MRF12Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF12Pics.ForEach(item =>
                {
                    int index = MRF12Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(11), pic);
            }
            List<MarketActionPic> MRF13Pics = actionService.MarketActionPicSearch(marketActionId, "MRF13");
            if (MRF13Pics.Count > 0)
            {
                //绑定其他工作人员 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF13Pics.ForEach(item =>
                {
                    int index = MRF13Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(11), pic);
            }

            //第12页 Others- 其他活动
            List<MarketActionPic> MRF14Pics = actionService.MarketActionPicSearch(marketActionId, "MRF14");
            if (MRF14Pics.Count > 0)
            {
                //Others- 其他活动
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF14Pics.ForEach(item =>
                {
                    int index = MRF14Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(12), pic);
            }

            //13页  Reimbursement Materials 报销材料
            List<MarketActionPic> MRF02Pics = actionService.MarketActionPicSearch(marketActionId, "MRF02");
            if (MRF02Pics.Count > 0)
            {
                //绑定合同照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF02Pics.ForEach(item =>
                {
                    int index = MRF02Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;

                helper.AddPictureToSlide(helper.GetSlide(13), pic);
            }
            List<MarketActionPic> MRF01Pics = actionService.MarketActionPicSearch(marketActionId, "MRF01");
            if (MRF01Pics.Count > 0)
            {
                //绑定报价单照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF01Pics.ForEach(item =>
                {
                    int index = MRF01Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;
                pic.X += 290;
                helper.AddPictureToSlide(helper.GetSlide(13), pic);
            }
            List<MarketActionPic> MRF03Pics = actionService.MarketActionPicSearch(marketActionId, "MRF03");
            if (MRF03Pics.Count > 0)
            {
                //绑定发票照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF03Pics.ForEach(item =>
                {
                    int index = MRF03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;
                pic.X += 580;
                helper.AddPictureToSlide(helper.GetSlide(13), pic);
            }

            string dirPath = basePath + @"\Temp\";
            //string path = dirPath + "活动报告_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            string path = dirPath + marketActionId.ToString() + "-" + shopName + "-市场活动-活动报告" + eventModeName + "-" + actionName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        /// <summary>
        /// 生成线上活动计划PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetActionPlanOnlinePPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\PlanOnLine.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";
            
            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
    
            if (lst.Count > 0)
            {
                ISlide firstSlide = helper.GetSlide(1);
                IShape shape = helper.GetShape(firstSlide, 2);
                helper.SaveTableCell(shape, 2, 2, actionName);
                helper.SaveTableCell(shape, 3, 2, date);
                helper.SaveTableCell(shape, 3, 4, place);

            }
            //第1页 活动总览 Overview
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide firstSlide = helper.GetSlide(1);
                IShape table1 = helper.GetShape(firstSlide, 2);
                helper.SaveTableCell(table1, 4, 3, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 5, 3, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 6, 3, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                IShape table2 = helper.GetShape(firstSlide, 7);
                helper.SaveTableCell(table1, 2, 2, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table1, 2, 4, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));

            }
            //第1页 Event Budget 费用总览
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {

                ISlide firstSlide = helper.GetSlide(1);
                IShape table3 = helper.GetShape(firstSlide, 8);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionPlanOnlineBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 3 + index;
                    helper.SaveTableCell(table3, row, 2, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table3, row, 3, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table3, row, 4, DateTimeToString(item.StartDate));
                    helper.SaveTableCell(table3, row, 5, DateTimeToString(item.EndDate));
                    helper.SaveTableCell(table3, row, 6, IntNullabelToString(item.TotalDays));
                    helper.SaveTableCell(table3, row, 7, DecimalNullabelToString(item.AmtPerDay));
                });
            }

            //第2页  Platform 平台简介
            List<MarketActionPic> MPN03Pics = actionService.MarketActionPicSearch(marketActionId, "MPN03");
            if (MPN03Pics.Count > 0)
            {
                //绑定媒体平台截图
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPN03Pics.ForEach(item =>
                {
                    int index = MPN03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(2), pic);

            }
            //媒体平台选择理由

            //第3页  品牌曝光
            List<MarketActionPic> MPN04Pics = actionService.MarketActionPicSearch(marketActionId, "MPN04");
            if (MPN04Pics.Count > 0)
            {
                //绑定
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPN04Pics.ForEach(item =>
                {
                    int index = MPN04Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                helper.AddPictureToSlide(helper.GetSlide(3), pic);

            }

            //第4页  物料设计
            List<MarketActionPic> MPN05Pics = actionService.MarketActionPicSearch(marketActionId, "MPN05");
            if (MPN05Pics.Count > 0)
            {
                //绑定物料设计
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPN05Pics.ForEach(item =>
                {
                    int index = MPN05Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                helper.AddPictureToSlide(helper.GetSlide(4), pic);
            }

            //第5页 媒体投放示例
            List<MarketActionPic> MPN06Pics = actionService.MarketActionPicSearch(marketActionId, "MPN06");
            if (MPN06Pics.Count > 0)
            {
                //绑定媒体投放示例
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPN06Pics.ForEach(item =>
                {
                    int index = MPN06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            string dirPath = basePath + @"\Temp\";
            //string path = dirPath + "活动计划_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            string path = dirPath + marketActionId.ToString() + "-" + shopName + "-市场活动-活动计划" + eventModeName + "-" + actionName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        /// <summary>
        /// 生成线上活动报告PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetActionReportOnlinePPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\ReportOnLine.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";

            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
            //第1页 活动总览 Overview
         
            if (lst.Count > 0)
            {
                ISlide firstSlide = helper.GetSlide(1);
                IShape shape = helper.GetShape(firstSlide, 6);
                helper.SaveTableCell(shape, 2, 3, actionName);
            }

            //第1页 活动总览 Overview
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                ISlide firstSlide = helper.GetSlide(1);
                IShape shape = helper.GetShape(firstSlide, 6);
                helper.SaveTableCell(shape, 3, 3, before4Weeks[0].Platform_Media);
                helper.SaveTableCell(shape, 4, 3, before4Weeks[0].Platform_ExposureForm);
            }

            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide firstSlide = helper.GetSlide(1);
                IShape table1 = helper.GetShape(firstSlide, 6);
                helper.SaveTableCell(table1, 6, 5, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 5, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 8, 5, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));
            }

            List<MarketActionAfter7> actionAfter7 = actionService.MarketActionAfter7Search(marketActionId);
            if (actionAfter7.Count > 0)
            {
                actionAfter7[0].TotalBudgetAmt = actionService.MarketActionAfter7TotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide firstSlide = helper.GetSlide(1);
                IShape table1 = helper.GetShape(firstSlide, 6);
                helper.SaveTableCell(table1, 6, 4, IntNullabelToString(actionAfter7[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 4, GetCostPerLead(actionAfter7[0].TotalBudgetAmt, actionAfter7[0].People_NewLeadsThsYearCount));
                helper.SaveTableCell(table1, 8, 4, IntNullabelToString(actionAfter7[0].People_NewLeadsThsYearCount));
            }

            //第2页 Event Budget 费用总览
            if (before4Weeks.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 5);
                helper.SaveTableCell(table1, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table1, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));
            }
            if (actionAfter7.Count > 0)
            {
                actionAfter7[0].TotalBudgetAmt = actionService.MarketActionAfter7TotalBudgetAmt(marketActionId);
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 5);
                helper.SaveTableCell(table1, 2, 2, DecimalNullabelToString(actionAfter7[0].TotalBudgetAmt));
                helper.SaveTableCell(table1, 3, 2, DecimalNullabelToString(actionAfter7[0].CoopFundSumAmt));
            }
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {

                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionReportOnlineBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 3 + index;
                    helper.SaveTableCell(table1, row, 3, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table1, row, 4, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table1, row, 5, DateTimeToString(item.StartDate));
                    helper.SaveTableCell(table1, row, 6, DateTimeToString(item.EndDate));
                    helper.SaveTableCell(table1, row, 7, IntNullabelToString(item.TotalDays));
                    helper.SaveTableCell(table1, row, 9, DecimalNullabelToString(item.AmtPerDay));
                });
            }
            List<MarketActionAfter7CoopFund> after7CoopFunds = actionService.MarketActionAfter7CoopFundSearch(marketActionId);
            if (after7CoopFunds.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                after7CoopFunds.ForEach(item =>
                {
                    int index = Array.IndexOf(ActionReportOnlineBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 3 + index;
                    helper.SaveTableCell(table1, row, 2, DecimalNullabelToString(item.CoopFundAmt));
                });
            }


            //第3页 线索报告
            List<MarketActionAfter2LeadsReportDto> after2LeadsReport = actionService.MarketActionAfter2LeadsReportSearch(marketActionId, "2022");
            if (after2LeadsReport.Count > 0)
            {
                //绑定线索报告
                ISlide fiveSlide = helper.GetSlide(5);
                IShape table2 = helper.GetShape(fiveSlide, 2);
                after2LeadsReport.ForEach(item =>
                {
                    int index = after2LeadsReport.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.CustomerName);
                    helper.SaveTableCell(table2, row, 2, item.TelNO);
                    helper.SaveTableCell(table2, row, 3, BoolNullabelToString(item.TestDriverCheck));
                    helper.SaveTableCell(table2, row, 4, item.InterestedModelName);
                    helper.SaveTableCell(table2, row, 5, BoolNullabelToString(item.DealCheck));
                    helper.SaveTableCell(table2, row, 6, item.DealCheckName);
                });
            }


            //第4页 实际投放截图
            List<MarketActionPic> MRN05Pics = actionService.MarketActionPicSearch(marketActionId, "MRN05");
            if (MRN05Pics.Count > 0)
            {
                //实际投放截图 开始
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRN05Pics.ForEach(item =>
                {
                    int index = MRN05Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(4), pic);
            }
            List<MarketActionPic> MRN06Pics = actionService.MarketActionPicSearch(marketActionId, "MRN06");
            if (MRN06Pics.Count > 0)
            {
                //实际投放截图 结束
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRN06Pics.ForEach(item =>
                {
                    int index = MRN06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X += 400;
                helper.AddPictureToSlide(helper.GetSlide(4), pic);
            }


            //5页  Reimbursement Materials 报销材料
            List<MarketActionPic> MRF02Pics = actionService.MarketActionPicSearch(marketActionId, "MRF02");
            if (MRF02Pics.Count > 0)
            {
                //绑定合同照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF02Pics.ForEach(item =>
                {
                    int index = MRF02Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }
            List<MarketActionPic> MRF01Pics = actionService.MarketActionPicSearch(marketActionId, "MRF01");
            if (MRF01Pics.Count > 0)
            {
                //绑定报价单照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF01Pics.ForEach(item =>
                {
                    int index = MRF01Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;
                pic.X += 300;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }
            List<MarketActionPic> MRF03Pics = actionService.MarketActionPicSearch(marketActionId, "MRF03");
            if (MRF03Pics.Count > 0)
            {
                //绑定发票照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF03Pics.ForEach(item =>
                {
                    int index = MRF03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 250;
                pic.X += 600;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            string dirPath = basePath + @"\Temp\";
           // string path = dirPath + "活动报告_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            string path = dirPath + marketActionId.ToString() + "-" + shopName + "-市场活动-活动计划" + eventModeName + "-" + actionName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        /// <summary>
        /// 生成交车仪式计划PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetHandOverPlatPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\PlanHandOver.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";

            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
            //Overview 概述 第二页
            if (lst.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 2);

                helper.SaveTableCell(table1, 2, 2, actionName);
                //helper.SaveTableCell(table1, 3, 2, date);
                helper.SaveTableCell(table1, 3, 5, place);

            }
            //Overview 概述 第二页
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 2);
                helper.SaveTableCell(table1, 3, 3, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 4, 3, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 5, 3, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 6, 3, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                //helper.SaveTableCell(table1, 4, 6, before4Weeks[0].Vehide_Usage);
                //helper.SaveTableCell(table1, 6, 6, before4Weeks[0].Vehide_Model);
                //helper.SaveTableCell(table1, 7, 6, IntNullabelToString(before4Weeks[0].Vehide_Qty));

                IShape table2 = helper.GetShape(secSlide, 5);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

            }
            ////Event Budget 费用总览  ppt 第3页
            if (before4Weeks.Count > 0)
            {
                //Event Budget 费用总览
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table3 = helper.GetShape(thirdSlide, 2);
                helper.SaveTableCell(table3, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table3, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));
            }
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table3 = helper.GetShape(thirdSlide, 2);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(HandOverPlanUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 7 + index / 2;
                    int col = 3 + (index % 2) * 4;
                    helper.SaveTableCell(table3, row, col, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table3, row, col + 1, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table3, row, col + 2, item.CoopFundDesc);
                });
            }


            //ppt 第4页 本月交车仪式安排
            List<MarketActionBefore4WeeksHandOverArrangement> before4WeeksHandOvers = actionService.MarketActionBefore4WeeksHandOverArrangementSearch(marketActionId);
            if (before4WeeksHandOvers.Count > 0)
            {
                //绑定
                ISlide elevenSlide = helper.GetSlide(4);
                IShape table2 = helper.GetShape(elevenSlide, 2);
                before4WeeksHandOvers.ForEach(item =>
                {
                    int index = before4WeeksHandOvers.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 2, DateTimeToString(item.HandOverDate));
                    helper.SaveTableCell(table2, row, 3, item.Model);
                    helper.SaveTableCell(table2, row, 4, item.MainProcess);
                });
            }

            List<MarketActionPic> MPF02Pics = actionService.MarketActionPicSearch(marketActionId, "MPF02");
            if (MPF02Pics.Count > 0)
            {
                //绑定场地实景照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF02Pics.ForEach(item =>
                {
                    int index = MPF02Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);

            }

            //第5页 场地内部照片 场地实景照片
            List<MarketActionPic> MPF03Pics = actionService.MarketActionPicSearch(marketActionId, "MPH03");
            if (MPF03Pics.Count > 0)
            {
                //绑定场地内部照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF03Pics.ForEach(item =>
                {
                    int index = MPF03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);

            }

            List<MarketActionPic> MPF04Pics = actionService.MarketActionPicSearch(marketActionId, "MPH04");
            if (MPF04Pics.Count > 0)
            {
                //绑定场地实景照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF04Pics.ForEach(item =>
                {
                    int index = MPF04Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            // 第6页 Brand Representation – KV 活动主视觉或背板设计 
            if (before4Weeks.Count > 0 && !string.IsNullOrEmpty(before4Weeks[0].KeyVisionPic))
            {
                ISlide sixSlide = helper.GetSlide(7);
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(sixSlide, pic);
            }


            // 第7页 Event Setup 场地布置
            List<MarketActionPic> MPH06Pics = actionService.MarketActionPicSearch(marketActionId, "MPH06");
            if (MPH06Pics.Count > 0)
            {
                //绑定场地搭建方案 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPH06Pics.ForEach(item =>
                {
                    int index = MPH06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }
            List<MarketActionPic> MPH07Pics = actionService.MarketActionPicSearch(marketActionId, "MPH07");
            if (MPH07Pics.Count > 0)
            {
                //绑定场地效果图 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPH07Pics.ForEach(item =>
                {
                    int index = MPH07Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }
            // 第8页 Photography 摄影摄像
            List<MarketActionPic> MPH08Pics = actionService.MarketActionPicSearch(marketActionId, "MPH08");
            if (MPH08Pics.Count > 0)
            {
                //绑定摄影师介绍 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPH08Pics.ForEach(item =>
                {
                    int index = MPH08Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }
            List<MarketActionPic> MPH09Pics = actionService.MarketActionPicSearch(marketActionId, "MPH09");
            if (MPH09Pics.Count > 0)
            {
                //绑定摄影师作品 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPH09Pics.ForEach(item =>
                {
                    int index = MPH09Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }


            string dirPath = basePath + @"\Temp\";
           // string path = dirPath + "交车仪式计划_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            string path = dirPath + marketActionId.ToString() + "-" + shopName + "-交车仪式-活动计划-"  + actionName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        /// <summary>
        /// 生成交车仪式报告PPT
        /// </summary>
        /// <param name="marketActionId"></param>
        /// <returns></returns>
        public string GetHandOverReportPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            AsposePPTHelper helper = new AsposePPTHelper();
            helper.Open(basePath + @"template\ReportHandOver.pptx");

            MarketActionService actionService = new MarketActionService();
            string shopName = "";
            string eventModeName = "";
            string actionName = "";
            string date = "";
            string place = "";

            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst != null && lst.Count > 0)
            {
                shopName = lst[0].ShopName;
                eventModeName = lst[0].EventModeName;
                actionName = lst[0].ActionName;
                date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                place = lst[0].ActionPlace;
            }
 
            if (lst.Count > 0)
            {
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 2, 2, actionName);
                helper.SaveTableCell(table1, 2, 5, date);
                helper.SaveTableCell(table1, 3, 2, place);

            }
            //第二页  活动概述
            List<MarketActionAfter7> actionAfter7 = actionService.MarketActionAfter7Search(marketActionId);
            if (actionAfter7.Count > 0)
            {
                actionAfter7[0].TotalBudgetAmt = actionService.MarketActionAfter7TotalBudgetAmt(marketActionId);
                //活动总览 Overview
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 5, 3, IntNullabelToString(actionAfter7[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 3, IntNullabelToString(actionAfter7[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 3, GetCostPerLead(actionAfter7[0].TotalBudgetAmt, actionAfter7[0].People_NewLeadsThsYearCount));
                helper.SaveTableCell(table1, 8, 3, IntNullabelToString(actionAfter7[0].People_NewLeadsThsYearCount));

            }
            //第二页  活动概述 计划
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId);
                //活动总览 Overview 计划
                ISlide secSlide = helper.GetSlide(2);
                IShape table1 = helper.GetShape(secSlide, 6);
                helper.SaveTableCell(table1, 5, 5, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 5, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 5, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 8, 5, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                IShape table2 = helper.GetShape(secSlide, 7);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

            }
            //第二页  活动概述 actual实际情况
            List<MarketActionAfter7ActualProcess> after7ActualProcess = actionService.MarketActionAfter7ActualProcessSearch(marketActionId);
            if (after7ActualProcess.Count > 0)
            {
                //绑定活动流程
                ISlide sevenSlide = helper.GetSlide(7);
                IShape table2 = helper.GetShape(sevenSlide, 2);
                after7ActualProcess.ForEach(item =>
                {
                    int index = after7ActualProcess.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.ActivityDateTime);
                    helper.SaveTableCell(table2, row, 2, item.Process);
                });
            }

            //第3页 Event Budget 费用总览
            List<MarketActionAfter7CoopFund> after7CoopFund = actionService.MarketActionAfter7CoopFundSearch(marketActionId);
            if (after7CoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情 Actual Cost
                ISlide slide = helper.GetSlide(3);
                IShape table1 = helper.GetShape(slide, 6);
                after7CoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(HandOverReportUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 2 + index;
                    helper.SaveTableCell(table1, row, 2, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table1, row, 4, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table1, row, 5, item.CoopFundDesc);
                });
            }
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情  Plan Budget预算列
                ISlide slide = helper.GetSlide(3);
                IShape table1 = helper.GetShape(slide, 6);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(HandOverReportUnderBudgetTypes, item.CoopFundCode);
                    if (index < 0) return;
                    int row = 2 + index;
                    helper.SaveTableCell(table1, row, 3, DecimalNullabelToString(item.CoopFundAmt));
                });
            }
            //Event Budget 费用总览 实际
            if (actionAfter7.Count > 0)
            {
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table2 = helper.GetShape(thirdSlide, 5);
                helper.SaveTableCell(table2, 2, 2, DecimalNullabelToString(actionAfter7[0].TotalBudgetAmt));
                helper.SaveTableCell(table2, 3, 2, DecimalNullabelToString(actionAfter7[0].CoopFundSumAmt));
            }
            //Event Budget 费用总览 实际
            if (before4Weeks.Count > 0)
            {
                ISlide thirdSlide = helper.GetSlide(3);
                IShape table2 = helper.GetShape(thirdSlide, 5);
                helper.SaveTableCell(table2, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table2, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));
            }

            //第4页 本月交车仪式清单
            List<MarketActionAfter7HandOverArrangement> after7HandOverArrangement = actionService.MarketActionAfter7HandOverArrangementSearch(marketActionId);
            if (after7HandOverArrangement.Count > 0)
            {
                ISlide fourSlide = helper.GetSlide(4);
                IShape table2 = helper.GetShape(fourSlide, 2);
                after7HandOverArrangement.ForEach(item =>
                {
                    int index = after7HandOverArrangement.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 2, DateTimeToString(item.HandOverDate));
                    helper.SaveTableCell(table2, row, 3, item.Model);
                    helper.SaveTableCell(table2, row, 4, item.MainProcess);
                });
            }

            //第5页
            List<MarketActionPic> MRF05Pics = actionService.MarketActionPicSearch(marketActionId, "MRF05");
            if (MRF05Pics.Count > 0)
            {
                //绑定场地实景 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF05Pics.ForEach(item =>
                {
                    int index = MRF05Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }
            List<MarketActionPic> MRF06Pics = actionService.MarketActionPicSearch(marketActionId, "MRF06");
            if (MRF06Pics.Count > 0)
            {
                //绑定场地内部 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF06Pics.ForEach(item =>
                {
                    int index = MRF06Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X = pic.X + 450;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            //第6页
            List<MarketActionPic> MRF08Pics = actionService.MarketActionPicSearch(marketActionId, "MRF08");
            if (MRF08Pics.Count > 0)
            {
                //绑定车辆 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF08Pics.ForEach(item =>
                {
                    int index = MRF08Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
            }
            //ppt 第7页
            if (before4Weeks.Count > 0 && !string.IsNullOrEmpty(before4Weeks[0].KeyVisionPic))
            {
                //Brand Representation – KV 活动主视觉或背板设计  ppt 第7页
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(helper.GetSlide(7), pic);
            }
            //第8页
            List<MarketActionPic> MRF09Pics = actionService.MarketActionPicSearch(marketActionId, "MRF09");
            if (MRF09Pics.Count > 0)
            {
                //绑定场地布置 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF09Pics.ForEach(item =>
                {
                    int index = MRF09Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }

            //第9页
            List<MarketActionPic> MRF13Pics = actionService.MarketActionPicSearch(marketActionId, "MRF13");
            if (MRF13Pics.Count > 0)
            {
                //绑定其他活动 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF13Pics.ForEach(item =>
                {
                    int index = MRF13Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                helper.AddPictureToSlide(helper.GetSlide(19), pic);
            }

            //第10页
            List<MarketActionPic> MRF02Pics = actionService.MarketActionPicSearch(marketActionId, "MRF02");
            if (MRF02Pics.Count > 0)
            {
                //绑定合同照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF02Pics.ForEach(item =>
                {
                    int index = MRF02Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;

                helper.AddPictureToSlide(helper.GetSlide(10), pic);

            }

            List<MarketActionPic> MRF01Pics = actionService.MarketActionPicSearch(marketActionId, "MRF01");
            if (MRF01Pics.Count > 0)
            {
                //绑定报价单照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF01Pics.ForEach(item =>
                {
                    int index = MRF01Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X += 300;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }


            List<MarketActionPic> MRF03Pics = actionService.MarketActionPicSearch(marketActionId, "MRF03");
            if (MRF03Pics.Count > 0)
            {
                //绑定发票照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF03Pics.ForEach(item =>
                {
                    int index = MRF03Pics.IndexOf(item);
                    pic.Paths.Add(OSSClientHelper.OSS_BASE_URL + item.PicPath);
                });
                pic.Width = 400;
                pic.X += 600;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }


            string dirPath = basePath + @"\Temp\";
            //string path = dirPath + "交车报告计划_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            string path = dirPath + marketActionId.ToString() + "-" + shopName + "-交车仪式-活动报告-" + actionName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        private string GetCostPerLead(decimal? budgetAmt, decimal? count)
        {
            if (budgetAmt.HasValue && count.HasValue)
            {
                return (budgetAmt.Value / count.Value).ToString("0.00");
            }
            return "";
        }

        private string DateTimeToString(DateTime? time)
        {
            return time.HasValue ? time.Value.ToString("dd/MM/yyyy") : "";
        }

        private string IntNullabelToString(int? count)
        {
            return count.HasValue ? count.Value.ToString() : "";
        }

        private string DecimalNullabelToString(decimal? count)
        {
            return count.HasValue ? count.Value.ToString() : "";
        }

        private string BoolNullabelToString(bool? check)
        {
            return check.HasValue ? check.Value ? "是" : "否" : "否";
        }
    }
}