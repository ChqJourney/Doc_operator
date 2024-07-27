using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection.Metadata.Ecma335;
using System.Text.Json;

namespace TableParse
{
    internal class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            string templatePath = "Templates/iec60598_2_1i_2022-08-26.docx";
            //int tableIdx = 6;
            //string rowType = "complaince";
#else
            string templatePath = args[0];
            //int tableIdx = int.Parse(args[1]);
            //string rowType=args[2];
            
#endif
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                var tables = doc.MainDocumentPart?.Document.Body?.Descendants<Table>().ToList();
                if (tables == null||tables.Count==0) throw new NullReferenceException(nameof(tables));
                var tableIdx = 0;
                foreach (var table in tables)
                {
                    var fTable = new FillTable();
                    fTable.TableIdx = tableIdx;
                    var rows = table.Descendants<TableRow>().ToList();
                    var rowIdx = 0;
                    fTable.Rows = new List<FillRow>();
                    foreach (TableRow row in rows)
                    {
                        var fRow =  new FillRow();
                        fRow.TableIdx = tableIdx;
                        fRow.RowIdx = rowIdx;
                        fRow.Duplicatable = false;
                        var cells = row.Descendants<TableCell>().ToList();
                        var columnIdx = 0;
                        fRow.Cells = new List<FillCell>();
                        foreach (TableCell cell in cells)
                        {
                            var fCell = new FillCell();
                            fCell.ColumnIdx = columnIdx;
                            fCell.RowIdx = rowIdx;
                            var j = cell.Descendants<Justification>().FirstOrDefault();

                            string cellContent = "";
                            if (!string.IsNullOrEmpty(cell.InnerText))
                            {
                                var paras = cell.Descendants<Paragraph>().ToList();
                                var bold = cell.Descendants<Bold>().ToList();
                                if (bold.Count > 0)
                                {
                                    fCell.Bolded = true;
                                }
                                else
                                {
                                    fCell.Bolded = false;
                                }
                                var texts = paras.Select(x => x.InnerText).ToList();
                                cellContent = string.Join('\n', texts);
                            }
                            var cellWidth = cell.TableCellProperties?.TableCellWidth?.Width?.Value;
                            var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val;
                            var vMerge = cell.TableCellProperties?.VerticalMerge?.Val;
                            var hasShading = IsShadingCell(cell);
                            if (j != null)
                            {
                                fCell.IsCentered = j.Val == "center";
                            }
                            fCell.HasShading = hasShading;
                            fCell.OriginalText = cellContent;
                            fCell.CellWidth = int.Parse(cellWidth);
                            fCell.HMerge = hMerge;
                            fCell.VMerge = vMerge;
                            fRow.Cells.Add(fCell);
                            columnIdx++;
                        }
                        rowIdx++;
                        fTable.Rows.Add(fRow);
                    }
                    var result = JsonSerializer.Serialize(fTable);
                    Console.WriteLine(result);
                    tableIdx++;
                }
                
                Environment.Exit(0); 
            }
        }
        private static bool IsShadingCell(TableCell cell)
        {
            var shading = cell.Descendants<Shading>().FirstOrDefault();
            return shading?.Val != null;
        }
    }
}
