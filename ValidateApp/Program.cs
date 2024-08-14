using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Docor;
using System.Collections.Immutable;
using System.Linq;
using ConsoleTables;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using System.Numerics;

namespace ValidateApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            string templatePath = "C:\\Users\\patri\\Desktop\\123\\base.docx";
            File.Copy(templatePath, "123.docx", true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open("123.docx", true))
            {
                var mainPart = doc.MainDocumentPart;
                var body = mainPart.Document.Body;
                var headparts = mainPart.HeaderParts.ToList();
                var newHeader=mainPart.AddNewPart<HeaderPart>();
                newHeader.Header= new Header(new Paragraph(new Run(new Text("new Header 0"))));
                var newHeader1=mainPart.AddNewPart<HeaderPart>();
                newHeader1.Header = new Header(new Paragraph(new Run(new Text("new Header 1"))));
                var newParaPr = new ParagraphProperties();
                var newSecPr = new SectionProperties();
                //newSecPr.RemoveAllChildren<HeaderReference>();
                var id = mainPart.GetIdOfPart(newHeader);
                var newHR = new HeaderReference() { Id = id, Type = HeaderFooterValues.Default };
                var id1=mainPart.GetIdOfPart(newHeader1);
                var newHR1=new HeaderReference() { Id=id1, Type = HeaderFooterValues.Default };
                newSecPr.AppendChild(newHR);

                body.Elements<Paragraph>().LastOrDefault().Descendants<ParagraphProperties>().FirstOrDefault().Append(new SectionProperties());
                var newPara = new Paragraph();
                //newParaPr.AppendChild(newSecPr);
                newPara.Append(new Run(new Text("11111111111")));
                newPara.Append(newParaPr);
                body.AppendChild(newPara);
                var last = body.Elements<SectionProperties>().FirstOrDefault();
                last.RemoveAllChildren<HeaderReference>();
                last.Append(newHR1);
                mainPart.Document.Save();
                
            }

        }

    }
        
    
}
