
using DocumentFormat.OpenXml.Packaging;

namespace DocParser
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length == 0) return;
            
            string templatePath = "Templates/iec60598_2_1i_2022-08-26.docx";
          
            var fieldsParser = new Fields();
            var rows = new Rows();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                if (doc.MainDocumentPart == null) throw new NullReferenceException("maindocumentpart of doc is null");
                switch (args[0])
                {
                    case "fields":
                        //fields collect
                        var fields = fieldsParser.CollectFields(doc.MainDocumentPart, action => Console.WriteLine(action));
                        break;
                    case "verdicts":
                        //report rows collect
                        var reportRows = rows.ParseRows(doc.MainDocumentPart, 7, 9, action => Console.WriteLine(action));
                        break;
                    case "Components":
                        break;
                    case "Measurements":
                        break;
                    default:
                        break;

                }
                
              
            }
        }
    }
}
