using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNoteOCRDll;
using System.Drawing;
using System.Drawing.Imaging;

namespace OneNoteTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var ocr = new OneNoteOCR();
            var findText = new ActionNote();

            //var pixel = new GetDeviceDpi();
            try
            {
                ocr.Verify();
            }
            catch (Exception)
            {
                Console.WriteLine("you do not have OneNote 15 ");
                return;
            }
            if (args.Length == 0)
            {
                Console.WriteLine("please add argument = path to the image file");
                return;
            }
            //var path = @"C:\Users\gastanica\Desktop\first.png";
            var path = @"C:\Users\gastanica\Desktop\second.png";
            //var path = @"C:\Users\gastanica\Desktop\screen.jpg";
            //var path = @"C:\Users\gastanica\Desktop\2.3.png";
            //var path = @"C:\Users\gastanica\Desktop\test2.png";
            //var ocrText = ocr.RecognizeImage(args[0]);
            var ocrText = findText.FindText(path, "ATEUR");

            Console.ReadLine();
        }

        public static void CreateImage(Rectangle findArea, string imagePath)
        {
            var wantedRegion = new Rectangle(findArea.X, findArea.Y, findArea.Width, findArea.Height);

            var screenshotToBitmap = new Bitmap(imagePath);
            var cloneRectangle = screenshotToBitmap.Clone(wantedRegion, PixelFormat.DontCare);

            cloneRectangle.Save(@"C:\Users\gastanica\Desktop\rectangle.png", ImageFormat.Png);
            screenshotToBitmap.Dispose();
        }

    }
}
