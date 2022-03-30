using com.yrtech.bentley.DAL;
using com.yrtech.InventoryAPI.Common;
using com.yrtech.InventoryAPI.DTO;
using Microsoft.Office.Interop.PowerPoint;
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
        string[] BudgetTypes = { "VenueRetal", "PhotoGraphy", "Setup", "Performance", "Catering_Food", "MC", "Catering_Drink", "MC", "Others", "Hospitality" };
        string[] ReportBudgetTypes = { "VenueRetal", "Setup", "PhotoGraphy", "Performance", "MC", "Hospitality", "Catering", "Others" };

        public string GetActionPlanPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            PPTHelper helper = new PPTHelper();
            helper.Open(basePath + @"template\2022 Dealer Coop fund event plan template-线下.pptx");

            MarketActionService actionService = new MarketActionService();
            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst.Count > 0)
            {
                Slide secSlide = helper.GetSlide(2);
                Shape shape = helper.GetShape(secSlide, 2);

                string actionName = lst[0].ActionName;
                string date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                string place = lst[0].ActionPlace;
                helper.SaveTableCell(shape, 2, 2, actionName);
                helper.SaveTableCell(shape, 3, 2, date);
                helper.SaveTableCell(shape, 3, 5, place);

            }
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId, before4Weeks[0].TotalBudgetAmt);
                //活动总览 Overview
                Slide secSlide = helper.GetSlide(2);
                Shape table1 = helper.GetShape(secSlide, 2);
                helper.SaveTableCell(table1, 4, 3, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 5, 3, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 6, 3, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 7, 3, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                helper.SaveTableCell(table1, 4, 6, before4Weeks[0].Vehide_Usage);
                helper.SaveTableCell(table1, 6, 6, before4Weeks[0].Vehide_Model);
                helper.SaveTableCell(table1, 7, 6, IntNullabelToString(before4Weeks[0].Vehide_Qty));

                Shape table2 = helper.GetShape(secSlide, 5);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

                //Event Budget 费用总览
                Slide thirdSlide = helper.GetSlide(3);
                Shape table3 = helper.GetShape(thirdSlide, 2);
                helper.SaveTableCell(table3, 2, 3, DecimalNullabelToString(before4Weeks[0].TotalBudgetAmt));
                helper.SaveTableCell(table3, 3, 3, DecimalNullabelToString(before4Weeks[0].CoopFundSumAmt));

                //Brand Representation – KV 活动主视觉或背板设计  ppt 第6页
                Slide sixSlide = helper.GetSlide(6);
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(sixSlide, pic);
            }
            List<MarketActionBefore4WeeksCoopFund> before4WeeksCoopFund = actionService.MarketActionBefore4WeeksCoopFundSearch(marketActionId);
            if (before4WeeksCoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情
                Slide thirdSlide = helper.GetSlide(3);
                Shape table3 = helper.GetShape(thirdSlide, 2);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(BudgetTypes, item.CoopFundCode);
                    int row = 7 + index / 2;
                    int col = 3 + (index % 2) * 4;
                    helper.SaveTableCell(table3, row, col, DecimalNullabelToString(item.CoopFundAmt));
                    helper.SaveTableCell(table3, row, col + 1, BoolNullabelToString(item.CoopFund_DMFChk));
                    helper.SaveTableCell(table3, row, col + 2, item.CoopFundDesc);
                });
            }

            List<MarketActionBefore4WeeksActivityProcess> before4WeeksActivitys = actionService.MarketActionBefore4WeeksActivityProcessSearch(marketActionId);
            if (before4WeeksActivitys.Count > 0)
            {
                //绑定活动流程
                Slide elevenSlide = helper.GetSlide(11);
                Shape table2 = helper.GetShape(elevenSlide, 2);
                before4WeeksActivitys.ForEach(item =>
                {
                    int index = before4WeeksActivitys.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.ActivityDateTime);
                    helper.SaveTableCell(table2, row, 2, item.Contents);
                    helper.SaveTableCell(table2, row, 3, item.Responsible);
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(4), pic);

            }

            List<MarketActionPic> MPF03Pics = actionService.MarketActionPicSearch(marketActionId, "MPF03");
            if (MPF03Pics.Count > 0)
            {
                //绑定场地内部照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF03Pics.ForEach(item =>
                {
                    int index = MPF03Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(5), pic);
            }

            List<MarketActionPic> MPF05Pics = actionService.MarketActionPicSearch(marketActionId, "MPF05");
            if (MPF05Pics.Count > 0)
            {
                //绑定场地搭建方案 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF05Pics.ForEach(item =>
                {
                    int index = MPF05Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(6), pic);
            }

            List<MarketActionPic> MPF07Pics = actionService.MarketActionPicSearch(marketActionId, "MPF07");
            if (MPF07Pics.Count > 0)
            {
                //绑定表演计划 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF07Pics.ForEach(item =>
                {
                    int index = MPF07Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }

            List<MarketActionPic> MPF09Pics = actionService.MarketActionPicSearch(marketActionId, "MPF09");
            if (MPF09Pics.Count > 0)
            {
                //绑定摄影师介绍 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF09Pics.ForEach(item =>
                {
                    int index = MPF09Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }

            List<MarketActionPic> MPF11Pics = actionService.MarketActionPicSearch(marketActionId, "MPF11");
            if (MPF11Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MPF11Pics.ForEach(item =>
                {
                    int index = MPF11Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }

            string dirPath = basePath + @"\Temp\";
            string path = dirPath + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
            helper.SaveAs(path); //保存ppt

            return path;
        }

        public string GetActionReportPPT(string marketActionId)
        {
            string basePath = HostingEnvironment.MapPath(@"~/");
            PPTHelper helper = new PPTHelper();
            helper.Open(basePath + @"template\2022 Dealer Coop fund event report template-线下.pptx");

            MarketActionService actionService = new MarketActionService();
            List<MarketActionDto> lst = actionService.MarketActionSearchById(marketActionId);
            if (lst.Count > 0)
            {
                Slide fourSlide = helper.GetSlide(4);
                Shape table1 = helper.GetShape(fourSlide, 6);

                string actionName = lst[0].ActionName;
                string date = DateTimeToString(lst[0].StartDate) + DateTimeToString(lst[0].EndDate);
                string place = lst[0].ActionPlace;
                helper.SaveTableCell(table1, 2, 2, actionName);
                helper.SaveTableCell(table1, 2, 5, date);
                helper.SaveTableCell(table1, 3, 2, place);

            }
            List<MarketActionAfter7> actionAfter7 = actionService.MarketActionAfter7Search(marketActionId);
            if (actionAfter7.Count > 0)
            {
                actionAfter7[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId, actionAfter7[0].TotalBudgetAmt);
                //活动总览 Overview
                Slide fourSlide = helper.GetSlide(4);
                Shape table1 = helper.GetShape(fourSlide, 6);
                helper.SaveTableCell(table1, 5, 3, IntNullabelToString(actionAfter7[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 3, IntNullabelToString(actionAfter7[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 3, GetCostPerLead(actionAfter7[0].TotalBudgetAmt, actionAfter7[0].People_NewLeadsThsYearCount));
                helper.SaveTableCell(table1, 8, 3, IntNullabelToString(actionAfter7[0].People_NewLeadsThsYearCount));

                //Event Budget 费用总览
                Slide thirdSlide = helper.GetSlide(6);
                Shape table2 = helper.GetShape(thirdSlide, 5);
                helper.SaveTableCell(table2, 2, 3, DecimalNullabelToString(actionAfter7[0].TotalBudgetAmt));
                helper.SaveTableCell(table2, 3, 3, DecimalNullabelToString(actionAfter7[0].CoopFundSumAmt));
            }
            List<MarketActionBefore4Weeks> before4Weeks = actionService.MarketActionBefore4WeeksSearch(marketActionId);
            if (before4Weeks.Count > 0)
            {
                before4Weeks[0].TotalBudgetAmt = actionService.MarketActionBefore4WeeksTotalBudgetAmt(marketActionId, before4Weeks[0].TotalBudgetAmt);
                //活动总览 Overview
                Slide fourSlide = helper.GetSlide(4);
                Shape table1 = helper.GetShape(fourSlide, 6);
                helper.SaveTableCell(table1, 5, 5, IntNullabelToString(before4Weeks[0].People_ParticipantsCount));
                helper.SaveTableCell(table1, 6, 5, IntNullabelToString(before4Weeks[0].People_DCPIDCount));
                helper.SaveTableCell(table1, 7, 5, GetCostPerLead(before4Weeks[0].TotalBudgetAmt, before4Weeks[0].People_NewLeadsThisYearCount));
                helper.SaveTableCell(table1, 8, 5, IntNullabelToString(before4Weeks[0].People_NewLeadsThisYearCount));

                Shape table2 = helper.GetShape(fourSlide, 7);
                helper.SaveTableCell(table2, 2, 1, IntNullabelToString(before4Weeks[0].People_InvitationTotalCount));
                helper.SaveTableCell(table2, 2, 2, IntNullabelToString(before4Weeks[0].People_InvitationCarOwnerCount));
                helper.SaveTableCell(table2, 2, 3, IntNullabelToString(before4Weeks[0].People_InvitationDepositorCount));
                helper.SaveTableCell(table2, 2, 4, IntNullabelToString(before4Weeks[0].People_InvitationPotentialCount));
                helper.SaveTableCell(table2, 2, 5, IntNullabelToString(before4Weeks[0].People_InvitationOtherCount));

                //Brand Representation – KV 活动主视觉或背板设计  ppt 第10页
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                pic.Paths.Add(before4Weeks[0].KeyVisionPic);
                helper.AddPictureToSlide(helper.GetSlide(10), pic);
            }

            List<MarketActionAfter7ActualProcess> after7ActualProcess = actionService.MarketActionAfter7ActualProcessSearch(marketActionId);
            if (after7ActualProcess.Count > 0)
            {
                //绑定活动流程
                Slide sevenSlide = helper.GetSlide(7);
                Shape table2 = helper.GetShape(sevenSlide, 2);
                after7ActualProcess.ForEach(item =>
                {
                    int index = after7ActualProcess.IndexOf(item);
                    int row = 2 + index;
                    helper.SaveTableCell(table2, row, 1, item.ActivityDateTime);
                    helper.SaveTableCell(table2, row, 2, item.Process);
                });
            }
            List<MarketActionAfter7CoopFund> after7CoopFund = actionService.MarketActionAfter7CoopFundSearch(marketActionId);
            if (after7CoopFund.Count > 0)
            {
                //Event Budget 费用总览 Budget Detail 费用详情 Actual Cost
                Slide sixSlide = helper.GetSlide(6);
                Shape table1 = helper.GetShape(sixSlide, 6);
                after7CoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ReportBudgetTypes, item.CoopFundCode);
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
                Slide sixSlide = helper.GetSlide(6);
                Shape table1 = helper.GetShape(sixSlide, 6);
                before4WeeksCoopFund.ForEach(item =>
                {
                    int index = Array.IndexOf(ReportBudgetTypes, item.CoopFundCode);
                    int row = 2 + index;
                    helper.SaveTableCell(table1, row, 3, DecimalNullabelToString(item.CoopFundAmt));
                });
            }


            List<MarketActionPic> MRF02Pics = actionService.MarketActionPicSearch(marketActionId, "MRF02");
            if (MRF02Pics.Count > 0)
            {
                //绑定合同照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF02Pics.ForEach(item =>
                {
                    int index = MRF02Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;

                helper.AddPictureToSlide(helper.GetSlide(15), pic);

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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X += 300;
                helper.AddPictureToSlide(helper.GetSlide(15), pic);
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X += 600;
                helper.AddPictureToSlide(helper.GetSlide(15), pic);
            }

            //第8页
            List<MarketActionPic> MRF05Pics = actionService.MarketActionPicSearch(marketActionId, "MRF05");
            if (MRF05Pics.Count > 0)
            {
                //绑定场地实景 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF05Pics.ForEach(item =>
                {
                    int index = MRF05Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(8), pic);
            }


            //第9页
            List<MarketActionPic> MRF08Pics = actionService.MarketActionPicSearch(marketActionId, "MRF08");
            if (MRF08Pics.Count > 0)
            {
                //绑定表演方案 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF08Pics.ForEach(item =>
                {
                    int index = MRF08Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }

            //第11页
            List<MarketActionPic> MRF09Pics = actionService.MarketActionPicSearch(marketActionId, "MRF09");
            if (MRF09Pics.Count > 0)
            {
                //绑定场地布置 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF09Pics.ForEach(item =>
                {
                    int index = MRF09Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(9), pic);
            }

            //第12页
            List<MarketActionPic> MRF10Pics = actionService.MarketActionPicSearch(marketActionId, "MRF10");
            if (MRF10Pics.Count > 0)
            {
                //绑定摄影师作品 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF10Pics.ForEach(item =>
                {
                    int index = MRF10Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(12), pic);
            }

            List<MarketActionPic> MRF11Pics = actionService.MarketActionPicSearch(marketActionId, "MRF11");
            if (MRF11Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF11Pics.ForEach(item =>
                {
                    int index = MRF11Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(12), pic);
            }

            //第13页
            List<MarketActionPic> MRF12Pics = actionService.MarketActionPicSearch(marketActionId, "MRF12");
            if (MRF12Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF12Pics.ForEach(item =>
                {
                    int index = MRF12Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(13), pic);
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
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                pic.X = pic.X + 400;
                helper.AddPictureToSlide(helper.GetSlide(13), pic);
            }

            //第14页
            List<MarketActionPic> MRF14Pics = actionService.MarketActionPicSearch(marketActionId, "MRF14");
            if (MRF14Pics.Count > 0)
            {
                //绑定礼仪 照片
                PicturePPTObject pic = new PicturePPTObject();
                pic.Paths = new List<string>();
                MRF14Pics.ForEach(item =>
                {
                    int index = MRF14Pics.IndexOf(item);
                    pic.Paths.Add(item.PicPath);
                });
                pic.Width = 350;
                helper.AddPictureToSlide(helper.GetSlide(14), pic);
            }

            string dirPath = basePath + @"\Temp\";
            string path = dirPath + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pptx";
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