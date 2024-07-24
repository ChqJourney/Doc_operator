using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.Json;

namespace DocFiller
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string templatePath = args[0];
            Console.WriteLine(templatePath);
            var report = JsonSerializer.Deserialize<Report>(args[1]);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(args[0], true))
            {
                foreach (var table in report.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            if (cell.Edits.Count == 0) continue;
                            foreach (var edit in cell.Edits)
                            {
                                
                                switch (edit.InputType+":"+edit.OutputType) 
                                {
                                    case "text:text":
                                        switch (edit.InsertType)
                                        {
                                            case "after":
                                                break;
                                            case "bookmark":
                                                break;
                                        }
                                        break;
                                    case "boolean:checkbox":
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }
        
    }
   
}
