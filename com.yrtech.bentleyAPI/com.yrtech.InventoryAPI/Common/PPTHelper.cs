using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;
using OFFICECORE = Microsoft.Office.Core;

namespace com.yrtech.InventoryAPI.Common
{
    public class PPTHelper
    {
        POWERPOINT.Application objApp = null;
        POWERPOINT.Presentation objPresSet = null;

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
                objApp = new POWERPOINT.Application();
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);

                BOpenState = true;
            }
            catch (Exception ex)
            {
                this.objApp.Quit();
                throw new PPTException("PPT 初始化失败;" + ex.ToString());
            }
        }
        public POWERPOINT.Slide GetSlide(int index)
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

        public POWERPOINT.Shape GetShape(POWERPOINT.Slide slide, int index)
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
        public void AddPictureToSlide(POWERPOINT.Slide slide, PicturePPTObject pic)
        {
            AddPictureToSlide(slide, pic, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoTrue);
        }
        /// <summary>
        /// 对目标幻灯片指定位置插入图片 
        /// </summary>
        /// <param name="slide">要插入图片的幻灯片</param>
        /// <param name="pic">图片信息（图片地址集合，位置，大小）  默认一行两个，需要计算</param>
        public void AddLinkPictureToSlide(POWERPOINT.Slide slide, PicturePPTObject pic)
        {
            AddPictureToSlide(slide, pic, OFFICECORE.MsoTriState.msoTrue, OFFICECORE.MsoTriState.msoFalse);
        }

        /// <summary>
        /// 对目标幻灯片指定位置插入图片 
        /// </summary>
        private void AddPictureToSlide(POWERPOINT.Slide slide, PicturePPTObject pic, OFFICECORE.MsoTriState linkToFile, OFFICECORE.MsoTriState saveWithDoc)
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
        public void SaveTitle(POWERPOINT.Shape shape, string title)
        {
            if (shape.TextFrame.HasText == OFFICECORE.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Text = title;
            }
        }

        public void SaveTableCell(POWERPOINT.Shape shape, int row, int col, string cell)
        {
            if (shape.HasTable == OFFICECORE.MsoTriState.msoTrue)
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

    public class PicturePPTObject
    {
        List<string> paths;
        float x = 50;
        float y = 135;
        float width = 860;
        float height = 380;
        int rows;
        int cols;

        public int Rows
        {
            get { return rows; }
            set { rows = value; }
        }

        public int Cols
        {
            get { return cols; }
            set { cols = value; }
        }


        public float X
        {
            set { x = value; }
            get { return x; }
        }
        public float Y
        {
            set { y = value; }
            get { return y; }
        }
        public float Width
        {
            set { width = value; }
            get { return width; }
        }
        public float Height
        {
            set { height = value; }
            get { return height; }
        }
        public List<string> Paths
        {
            set { paths = value; }
            get { return paths; }
        }

    }

    public class PPTException : Exception
    {
        public PPTException(string message)
            : base(message)
        {
        }
    }

}
