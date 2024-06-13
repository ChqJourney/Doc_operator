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
                var documentBody = mainPart.Document.Body;
               
                    var tables = documentBody.Descendants<Table>().Take(10).ToList();
                   
                  var rows= tables[0].Descendants<TableRow>().Take(10).ToList();
                  foreach (var row in rows){

                    Console.WriteLine(row.InnerText);
                  }
                    
                    
               
            }
              

            Console.WriteLine("done");
        }
    }
        
}
