using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace com.yrtech.InventoryAPI.Common
{
    public class AsposePPTHelper
    {
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
            if (this.objPresSet != null)
            {
                return;
            }
            try
            {
                //以非只读方式打开,方便操作结束后保存.
                objPresSet = new Aspose.Slides.Presentation(filePath);
                BOpenState = true;
            }
            catch (Exception ex)
            {
                if (this.objPresSet != null)
                    this.objPresSet.Dispose();
                CommonHelper.log("打开ppt异常；file=" + filePath + "\n  error=" + ex.ToString());
                throw ex;
            }
        }
        public ISlide GetSlide(int index)
        {
            if (objPresSet != null)
            {
                return objPresSet.Slides[index - 1];
            }
            else
            {
                throw new PPTException("PPT 对象未初始化，使用前需要调用Open方法");
            }
        }

        public IShape GetShape(ISlide slide, int index)
        {
            if (objPresSet != null)
            {
                return slide.Shapes[index-1];
            }
            else
            {
                throw new PPTException("PPT 对象未初始化，使用前需要调用Open方法");
            }
        }
        private static Image GetImage(string url)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                using (WebResponse response = request.GetResponse())
                {
                    return Image.FromStream(response.GetResponseStream());
                }
            }
            catch (Exception ex)
            {

            }
            return null;
        }
        /// <summary>
        /// 对目标幻灯片指定位置插入图片 
        /// </summary>
        public void AddPictureToSlide(ISlide slide, PicturePPTObject pic)
        {
            if (pic.Paths.Count == 0)
            {
                throw new PPTException("没有图片地址！");
            }
             List<Image> images = new List<Image>();
            foreach(string path in pic.Paths){
                Image image = GetImage(path);
                if(image!=null){
                    images.Add(image);
                }
            }

            if (images.Count == 1)
            {
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, pic.X, pic.Y, pic.Width, pic.Height, objPresSet.Images.AddImage(images[0]));
            }
            else
            {
                if (pic.Rows == 0 || pic.Cols == 0)
                {
                    if (images.Count == 2)
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
                        int count = Math.Min(pic.Paths.Count, 4);
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
                        if (index >= images.Count) break;

                        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, pic.X + c * width, pic.Y + r * height, width, height, objPresSet.Images.AddImage(images[index]));
                    }
                }
            }

        }
      
        public void SaveTableCell(IShape shape, int row, int col, string cell)
        {
            if (shape.GetType().Name == "Table")
            {
                Table table = (Table)shape;
                table[col-1, row-1].TextFrame.Text = cell == null?"":cell;
            }            
        }
        /// <summary>
        /// 写入文本框内容
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="value"></param>
        public void WriteTextFrame(AutoShape shape, string value)
        {
            if (shape!=null)
            {
                shape.TextFrame.Text = (value==null?"":value);
            }
        }

        /// <summary>
        /// 保存PPT文档。
        /// </summary>
        public void SaveAs(string path)
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                this.objPresSet.Save(path, Aspose.Slides.Export.SaveFormat.Pptx);
                this.objPresSet.Dispose();
            }

            GC.Collect();
        }
    }


}
