using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Collections.Generic;

namespace OneNoteOCRDll
{
    /// <summary>
    /// ocr with one node
    /// </summary>
    public class OneNoteOCR
    {
        public OneNoteOCR()
        {
            try
            {
                Verify();
            }
            catch (Exception)
            {
                Console.WriteLine("OneNote 15 must be installed for this user");
                return;
            }
        }

        /// <summary>
        /// verify one note exists on pc
        /// </summary>
        public void Verify()
        {
            var a = new Application();
            Marshal.ReleaseComObject(a);
            a = null;
            Thread.Sleep(2000);
        }

        /// <summary>
        /// Recognize text in image
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public Tuple<XDocument, Image> RecognizeImage(string imagePath)
        {
            var oneNoteApp = new Application();
            string sections;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsSections, out sections);
            var doc = XDocument.Parse(sections);
            var ns = doc.Root.Name.Namespace;
            var node = doc.Descendants(ns + "Section").First();
            var sectionId = node.Attribute("ID").Value;
            string pageId;
            oneNoteApp.CreateNewPage(sectionId, out pageId);
            var imageCreated = InsertImage(imagePath, pageId);
            //update the note page 
            Thread.Sleep(2000);
            var str = "";
            oneNoteApp.GetPageContent(pageId, out str, PageInfo.piBasic, XMLSchema.xsCurrent);
            doc = XDocument.Parse(str);
            oneNoteApp.DeleteHierarchy(pageId, deletePermanently: true);
            Marshal.ReleaseComObject(oneNoteApp);
            return new Tuple<XDocument, Image>(doc, imageCreated);

        }

        Image InsertImage(string pathImage, string existingPageId)
        {
            string strNamespace = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            string m_xmlImageContent =
                "<one:Image><one:Size width=\"{1}\" height=\"{2}\" isSetByUser=\"true\" /><one:Data>{0}</one:Data></one:Image>";
            string m_xmlNewOutline =
                "<?xml version=\"1.0\"?><one:Page xmlns:one=\"{2}\" ID=\"{1}\"><one:Title><one:OE><one:T><![CDATA[{3}]]></one:T></one:OE></one:Title>{0}</one:Page>";
            string pageToBeChange = "RecognizeImage" + DateTime.Now.ToString("yyyyMMddHHmmss");
            string fileString;

            Image baseScreenshot = Image.FromFile(pathImage, true);
            using (var bitmap = new Bitmap(baseScreenshot))
            {
                var stream = new MemoryStream();
                bitmap.Save(stream, ImageFormat.Png);
                fileString = Convert.ToBase64String(stream.ToArray());

                var onenoteApp = new Application();

                string imageXmlStr = string.Format(m_xmlImageContent, fileString, bitmap.Width, bitmap.Height);
                string pageChangesXml = string.Format(m_xmlNewOutline,
                    new object[] { imageXmlStr, existingPageId, strNamespace, pageToBeChange });

                onenoteApp.UpdatePageContent(pageChangesXml);
                Marshal.ReleaseComObject(onenoteApp);
                onenoteApp = null;
            }

            return baseScreenshot;
        }
    }
}
