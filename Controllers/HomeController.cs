using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GemBox.Document;

namespace Testing.Controllers
{
    public class HomeController : Controller
    {
        public HomeController() {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }

        public ActionResult Index()
        {            
            return View();
        }
        
        //[HttpPost]
        //[ValidateInput(false)]
        //public ActionResult CrearWord_OLD(string TxtHtml) {
        //    DocumentModel document = new DocumentModel();

        //    Section section = new Section(document);
        //    document.Sections.Add(section);
        //    int MaxParagraphFreeVersion = 20;
        //    foreach (string Linea in TxtHtml.Split(new string[] { "\r\n" }, StringSplitOptions.None))
        //    {

        //        Paragraph paragraph = new Paragraph(document);
        //        paragraph.ParagraphFormat.SpaceAfter = 0;
        //        paragraph.ParagraphFormat.SpaceBefore = 0;
        //        section.Blocks.Add(paragraph);
        //        Run run = new Run(document, Linea);
        //        paragraph.Inlines.Add(run);
        //        MaxParagraphFreeVersion--;
        //        if (MaxParagraphFreeVersion == 0)
        //            break;
        //    }
        //    MemoryStream ms;
        //    using (ms = new MemoryStream())
        //    {
        //        document.Save(ms, SaveOptions.DocxDefault);
        //    }
        //    return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Converted_{DateTime.Now.Ticks}.docx");
        //}

        [HttpPost]
        [ValidateInput(false)]
        public ActionResult CrearWord(string TxtHtml)
        {
            DocumentModel document = new DocumentModel();
            using (MemoryStream DocHtml = new MemoryStream())
            {
                var sw = new StreamWriter(DocHtml, System.Text.Encoding.UTF8);
                try
                {
                    sw.Write(TxtHtml);
                    sw.Flush();
                    DocHtml.Seek(0, SeekOrigin.Begin);
                    document = DocumentModel.Load(DocHtml, LoadOptions.HtmlDefault);
                    Section section = document.Sections[0];
                    PageSetup pageSetup = section.PageSetup;
                    PageMargins pageMargins = pageSetup.PageMargins;
                    pageMargins.Top = pageMargins.Bottom = pageMargins.Left = pageMargins.Right = 80;
                }
                finally
                {
                    sw.Dispose();
                }
            }
            MemoryStream ms;
            using (ms = new MemoryStream())
            {
                document.Save(ms, SaveOptions.DocxDefault);
            }
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Converted_{DateTime.Now.Ticks}.docx");
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}