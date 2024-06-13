using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Docor;

namespace ValidateApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 输入模板文件路径
            string templatePath = "template.docx";
          
            
           
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                var mainPart = doc.MainDocumentPart;
                //var documentBody = mainPart.Document.Body;
                Operator.fillTextField(mainPart, 0, 3, 1, "450450450");
                //var tables = documentBody.Descendants<Table>().Take(10).ToList();

                //var rows = tables[0].Descendants<TableRow>().Take(10).ToList();
                //var cell = rows[3].Descendants<TableCell>().ToList()[1];
                //var x = cell.Descendants<FormFieldData>().FirstOrDefault();
                //var s = x.Descendants<DefaultTextBoxFormFieldString>().FirstOrDefault();
                //Console.WriteLine(s.Val);
                mainPart.Document.Save();
               
            }

            Console.WriteLine("done");
        }
    }
        
}
