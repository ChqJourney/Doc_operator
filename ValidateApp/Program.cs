using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Docor;
using System.Collections.Immutable;
using System.Linq;
using ConsoleTables;
using DocumentFormat.OpenXml.Spreadsheet;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace ValidateApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("value1");
            await Task.Delay(1000);
            Console.WriteLine("value2");
            await Task.Delay(1000);
            Console.WriteLine("Value3");
            Console.ReadKey();
            // 输入模板文件路径
            //string templatePath = "iec60598_2_1i_2022-08-26.docx";

            //string outputPath = "output.xlsx";
            //var t = new ConsoleTable("clause no", "title", "blank", "Verdict");
            //var list=new List<List<String>>();
            //using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            //{
            //    var mainPart = doc.MainDocumentPart;
            //    var body=mainPart.Document.Body;

            //    var comments = mainPart.WordprocessingCommentsPart.Comments;

            //    var tables = body.Descendants<Table>().ToList();

            //    //processor.Process();
            //    //var ens = new List<CommentEntity>();

            //    //for (int i = 0; i < tables.Count(); i++)
            //    //{
            //    //    var rows = tables[i].Descendants<TableRow>().ToList();
            //    //    for (int j = 0; j < rows.Count(); j++)
            //    //    {
            //    //        var cells = rows[j].Descendants<TableCell>().ToList();
            //    //        for (int k = 0; k < cells.Count; k++)
            //    //        {
            //    //            var crs = cells[k].Descendants<CommentRangeStart>().FirstOrDefault();
            //    //            if (crs != null && ens.All(a => a.CommentId != crs.Id))
            //    //            {
            //    //                ens.Add(new CommentEntity { TableIdx = i, RowIdx = j, CellIdx = k, CommentId = crs.Id });
            //    //            }
            //    //        }
            //    //    }
            //    //}

            //    //foreach (var en in ens)
            //    //{
            //    //    var dd = comments.ChildElements.Select(s => s as Comment).ToList();
            //    //    en.CommentText = dd.FirstOrDefault(f => f.Id == en.CommentId).InnerText;
            //    //    t.AddRow(en.TableIdx, en.RowIdx, en.CellIdx, en.CommentId, en.CommentText);
            //    //}

            //    for (int i = 6; i < 9; i++)
            //    {
            //            var temp=new List<String>();
            //        for (int j = 0; j < tables[i].Descendants<TableRow>().Count(); j++)
            //        {
            //            var row = tables[i].Elements<TableRow>().ElementAtOrDefault(j);
            //            var cells=row.Descendants<TableCell>().ToList();
            //            if (cells == null) continue;
            //            if (cells.Count == 1)
            //            {
            //                continue;
            //            }
            //            else if(cells.Count ==3)
            //            {

            //            var firstCell = cells[0].InnerText;
            //            var secondCell = cells[1].InnerText;
            //            var thirdCell = cells[2].InnerText;
            //                t.AddRow(firstCell, secondCell, "", thirdCell);
            //                temp.Add(firstCell);
            //                temp.Add(secondCell);
            //                temp.Add("");
            //                temp.Add(thirdCell);
            //            }
            //            else
            //            {
            //                var firstCell = cells[0].InnerText;
            //                var secondCell = cells[1].InnerText;
            //                var thirdCell = cells[2].InnerText;
            //                var fourthCell = cells[3].InnerText;
            //                t.AddRow(firstCell, secondCell, thirdCell,fourthCell);
            //                temp.Add(firstCell);
            //                temp.Add(secondCell);
            //                temp.Add(thirdCell);
            //                temp.Add(fourthCell);
            //            }
            //        } 
            //        list.Add(temp);
            //    }
                //mainPart.Document.Save();

            //}
            //t.Write(ConsoleTables.Format.Minimal);
            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            //{
            //    // 创建 WorkbookPart
            //    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            //    workbookPart.Workbook = new Workbook();

            //    // 创建 WorksheetPart
            //    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            //    worksheetPart.Worksheet = new Worksheet(new SheetData());

            //    // 获取 SheetData
            //    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            //    foreach (var item in list)
            //    {
            //        AddRow(sheetData, item);
            //    }
            //    // 创建 Sheets
            //    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            //    // 创建并附加 Sheet
            //    Sheet sheet = new Sheet()
            //    {
            //        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            //        SheetId = 1,
            //        Name = "Sheet1"
            //    };
            //    sheets.Append(sheet);

            //    // 保存工作簿
            //    workbookPart.Workbook.Save();
            //}
            //Console.ReadKey();
        }
        public static void AddRow(SheetData sheetData, List<string> values)
        {
            Row row = new Row();
            foreach (string value in values)
            {
                Cell cell = new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value)
                };
                row.Append(cell);
            }
            sheetData.Append(row);
        }
    }
        internal class CommentEntity
    {
        public int Id { get; set; }
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public int CellIdx { get; set; }
        public string CommentId { get; set; }
        public string CommentText { get; set; }
    }
    
}
