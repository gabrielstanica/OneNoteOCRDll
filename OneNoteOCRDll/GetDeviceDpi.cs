using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace OneNoteOCRDll
{
    public class GetDeviceDpi
    {
        private Bitmap _imageCaptured { get; }

        public GetDeviceDpi(Image imageCreated)
        {
            _imageCaptured = new Bitmap(imageCreated);
        }

        public void TransformToPixels(float point, out float pixel)
        {
            using (Graphics g = Graphics.FromImage(_imageCaptured))
            {
                pixel = point * g.DpiX / 72;
            }
        }

    }
}
