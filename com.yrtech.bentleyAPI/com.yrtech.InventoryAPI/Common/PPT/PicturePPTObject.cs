using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace com.yrtech.InventoryAPI.Common
{
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
}