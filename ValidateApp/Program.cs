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

            string templatePath = "D:\\repos\\Doc_operator\\ValidateApp\\iec60598_2_1i_2022-08-26.docx";
            //File.Copy(templatePath, "123.docx", true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                var mainPart = doc.MainDocumentPart;
                var body = mainPart.Document.Body;
                var styles = mainPart.StyleDefinitionsPart.Styles;
                var tables=body.Descendants<Table>().ToList();
                var rows = tables[9].Descendants<TableRow>().ToList();
                var cells = rows[3].Descendants<TableCell>().ToList();
            }

        }

    }
        
    
}
