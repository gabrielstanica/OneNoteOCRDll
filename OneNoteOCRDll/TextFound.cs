using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Collections.Generic;
using System.Drawing;

namespace OneNoteOCRDll
{

    /// <summary>
    /// 
    /// </summary>
    public class TextFound
    {
        /// <summary>
        /// 
        /// </summary>
        public RectangleF TextArea
        {
            get { return Rectangle.Round(_mArea); }
        }

        public string SearchText;

        public string ImageText;

        private RectangleF _mArea;

        public TextFound(float x, float y, float width, float height, string search, string text)
        {
            SearchText = search;
            ImageText = text;
            _mArea = new RectangleF(x, y, width, height);

    }
        public Point Center()
        {
            var center_x = TextArea.Location.X + (TextArea.Width / 2);
            var center_y = TextArea.Location.Y + (TextArea.Height / 2);
            return Point.Round(new PointF(center_x, center_y));
        }

    }
}
