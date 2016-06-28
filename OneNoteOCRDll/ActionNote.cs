using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OneNoteOCRDll
{
    public class ActionNote
    {
        private OneNoteOCR OneNote;
        private GetDeviceDpi GetDpi;

        public ActionNote()
        {
            OneNote = new OneNoteOCR();
        }

        public List<TextFound> FindText(string imagePath, string wantedText)
        {
            List<TextFound> result = new List<TextFound>();

            var foundItems = OneNote.RecognizeImage(imagePath);
            var xmlDocument = foundItems.Item1;
            var imageCreated = foundItems.Item2;

            GetDpi = new GetDeviceDpi(imageCreated);

            IEnumerable<XElement> tokenArray = null;
            var wantedTextLength = wantedText.Length;

            try
            {
                var textArray = xmlDocument.Descendants().First(t => t.Name.LocalName == "OCRText");
                var textValue = textArray.Value;
                tokenArray = xmlDocument.Descendants().Where(t => t.Name.LocalName == "OCRToken").ToList();

                foreach (var elementToken in tokenArray)
                {
                    bool checkExists = false;
                    var valueTokenPosition = int.Parse(elementToken.Attribute("startPos").Value);
                    var stillToSearch = textValue.Length - valueTokenPosition - wantedTextLength;

                    if (stillToSearch > 0)
                    {
                        var wantedSubstring = textValue.Substring(valueTokenPosition, wantedTextLength);
                        checkExists = wantedSubstring.Equals(wantedText);
                    }

                    if (checkExists)
                    {

                        float x, y, width, height;

                        var xPoint = float.Parse(elementToken.Attribute("x").Value);
                        var yPoint = float.Parse(elementToken.Attribute("y").Value);
                        var widthPoint = float.Parse(elementToken.Attribute("width").Value);
                        var heightPoint = float.Parse(elementToken.Attribute("height").Value);

                        GetDpi.TransformToPixels(xPoint, out x);
                        GetDpi.TransformToPixels(yPoint, out y);
                        GetDpi.TransformToPixels(widthPoint, out width);
                        GetDpi.TransformToPixels(heightPoint, out height);

                        result.Add(new TextFound(x, y, width, height, wantedText, textValue));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("Text not found");
            }
            return result;
        }

    }
}
