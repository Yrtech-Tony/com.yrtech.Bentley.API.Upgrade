using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Threading;

namespace com.yrtech.InventoryAPI.Common
{
    public class PPTHelper
    {
        Application objApp = null;
        Presentation objPresSet = null;

        bool bOpenState = false;
        double pixperPoint = 0;
        double offsetx = 0;
        double offsety = 0;

        public bool BOpenState
        {
            set { bOpenState = value; }
            get { return bOpenState; }
        }

        /// <summary>
        /// 打开PPT文档
        /// </summary>
        /// <param name="filePath">PPT文件路径</param>
        public void Open(string filePath)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null)
            {
                return;
            }
            try
            {
                CommonHelper.log("进入try");
                Thread.Sleep(1000);
                
                objApp = new Application();
                CommonHelper.log("new");
                Thread.Sleep(1000);
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = objApp.Presentations.Open(filePath,MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                CommonHelper.log("打开");
                Thread.Sleep(1000);
                BOpenState = true;
            }
            catch (Exception ex)
            {
                if (this.objApp != null)
                    this.objApp.Quit();
                CommonHelper.log("打开ppt异常；file=" + filePath + "\n  error=" + ex.ToString());
                throw ex;
            }
        }
        public Slide GetSlide(int index)
        {
            if (objPresSet != null)
            {
                return objPresSet.Slides[index];
            }
            else
            {
                throw new PPTException("PPT 对象未初始化，使用前需要调用Open方法");
            }
        }

        public Microsoft.Office.Interop.PowerPoint.Shape GetShape(Slide slide, int index)
        {
            if (objPresSet != null)
            {
                return slide.Shapes[index];
            }
            else
            {
                throw new PPTException("PPT 对象未初始化，使用前需要调用Open方法");
            }
        }
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Next();
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Previous();
        }
        /// <summary>
        /// 对目标幻灯片指定位置插入图片 图片包含在幻灯片中
        /// </summary>
        /// <param name="slide">要插入图片的幻灯片</param>
        /// <param name="pic">图片信息（图片地址集合，位置，大小）  默认一行两个，需要计算</param>
        public void AddPictureToSlide(Slide slide, PicturePPTObject pic)
        {
            AddPictureToSlide(slide, pic, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }
        /// <summary>
        /// 对目标幻灯片指定位置插入图片 
        /// </summary>
        /// <param name="slide">要插入图片的幻灯片</param>
        /// <param name="pic">图片信息（图片地址集合，位置，大小）  默认一行两个，需要计算</param>
        public void AddLinkPictureToSlide(Slide slide, PicturePPTObject pic)
        {
            AddPictureToSlide(slide, pic, MsoTriState.msoTrue, MsoTriState.msoFalse);
        }

        /// <summary>
        /// 对目标幻灯片指定位置插入图片 
        /// </summary>
        private void AddPictureToSlide(Slide slide, PicturePPTObject pic, MsoTriState linkToFile, MsoTriState saveWithDoc)
        {
            if (pic.Paths.Count == 0)
            {
                throw new PPTException("没有图片地址！");
            }
            else if (pic.Paths.Count == 1)
            {
                slide.Shapes.AddPicture(pic.Paths[0], linkToFile, saveWithDoc, pic.X, pic.Y, pic.Width, pic.Height);
            }
            else
            {
                if (pic.Rows == 0 || pic.Cols == 0)
                {
                    if (pic.Paths.Count == 2)
                    {
                        if (pic.Width > pic.Height * 1.2)
                        {
                            pic.Rows = 1;
                            pic.Cols = 2;
                        }
                        else
                        {
                            pic.Rows = 2;
                            pic.Cols = 1;
                        }
                    }
                    else
                    {
                        int count = pic.Paths.Count;//Math.Min(pic.Paths.Count, 6);
                        pic.Cols = 2;
                        pic.Rows = count / 2 + (count % 2 == 0 ? 0 : 1);
                    }
                }


                float height = pic.Height / pic.Rows;
                float width = pic.Width / pic.Cols;
                for (int r = 0; r < pic.Rows; r++)
                {
                    for (int c = 0; c < pic.Cols; c++)
                    {
                        int index = r * pic.Cols + c;
                        if (index >= pic.Paths.Count) break;
                        slide.Shapes.AddPicture(pic.Paths[index], linkToFile, saveWithDoc, pic.X + c * width, pic.Y + r * height, width, height);
                    }
                }
            }

        }
        public void SaveTitle(Microsoft.Office.Interop.PowerPoint.Shape shape, string title)
        {
            if (shape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Text = title;
            }
        }

        public void SaveTableCell(Microsoft.Office.Interop.PowerPoint.Shape shape, int row, int col, string cell)
        {
            if (shape.HasTable == MsoTriState.msoTrue)
            {
                if (shape.Table != null)
                {
                    shape.Table.Cell(row, col).Shape.TextFrame.TextRange.Text = cell;
                }
            }
        }

        /// <summary>
        /// 保存PPT文档。
        /// </summary>
        public void Save()
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                this.objPresSet.Save();
                this.objPresSet.Close();
            }
            if (this.objApp != null)
                this.objApp.Quit();

            GC.Collect();
        }

        /// <summary>
        /// 保存PPT文档。
        /// </summary>
        public void SaveAs(string path)
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                this.objPresSet.SaveAs(path);
                this.objPresSet.Close();
            }
            if (this.objApp != null)
                this.objApp.Quit();

            GC.Collect();
        }
    }

}
