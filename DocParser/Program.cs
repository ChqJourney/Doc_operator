
using DocumentFormat.OpenXml.Packaging;

namespace DocParser
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
#if DEBUG   
            string templatePath = "Templates/iec60598_2_1i_2022-08-26.docx";
#else
            string templatePath = args[0];
            
            
#endif
            var fieldsParser = new Fields();
            var rows = new Rows();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                if (doc.MainDocumentPart == null) throw new NullReferenceException("maindocumentpart of doc is null");
#if DEBUG
                switch(args[0])
#else
                switch (args[1])
#endif
                {
                    case "fields":
                        //fields collect
                        var fields = fieldsParser.CollectFields(doc.MainDocumentPart, action => Console.WriteLine(action));
                        break;
                    case "verdicts":
                        //compliance rows collect
#if DEBUG
                        var reportRows = rows.ParseRows(doc.MainDocumentPart, 4, 9, action => Console.WriteLine(action));
#else
                        var startTableIdx=int.Parse(args[2]);
                        var endTableIdx=int.Parse(args[3]);
                        var reportRows = rows.ParseRows(doc.MainDocumentPart, startTableIdx, endTableIdx, action => Console.WriteLine(action));
#endif
                        break;
                    case "components":
                        //component list rows collect
#if DEBUG
                        var grows = await rows.ParseGeneralRows(doc.MainDocumentPart, 17, action => Console.WriteLine(action));
#else
                        var tableIdx = int.Parse(args[2]);
                        var componentRows = await rows.ParseGeneralRows(doc.MainDocumentPart, tableIdx, action => Console.WriteLine(action));
#endif
                        break;
                    case "Measurements":
                        //Measurements table rows collect
#if DEBUG
                        for (int i = 10; i < 14; i++)
                        {
                            var mRows = await rows.ParseGeneralRows(doc.MainDocumentPart, i,action=>Console.WriteLine(action));
                        }
#else
                        var startIdx = int.Parse(args[2]);
                        var endIdx = int.Parse(args[3]);
                        for (int i = startIdx; i < endIdx; i++)
                        {
                            var mRows = await rows.ParseGeneralRows(doc.MainDocumentPart, i,action=>Console.WriteLine(action));
                        }
#endif
                        break;
                    case "Clear":
                        Console.Clear();
                        break;
                    default:
                        break;

                }
                
              
            }
            Environment.Exit(0);
            //Console.Clear();
        }
    }
}
