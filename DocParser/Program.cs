
using DocumentFormat.OpenXml.Packaging;

namespace DocParser
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string templatePath = "iec60598_2_1i_2022-08-26.docx";
          
            var fieldsParser = new Fields();
            var rows = new Rows();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                if (doc.MainDocumentPart == null) throw new NullReferenceException("maindocumentpart of doc is null");
                //fields collect
                var fields=fieldsParser.CollectFields(doc.MainDocumentPart,action=>Console.WriteLine(action));
                //report rows collect
                var reportRows = rows.ParseRows(doc.MainDocumentPart, 5, 9, action => Console.WriteLine(action));
            }
        }
    }
}
