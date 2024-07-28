using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Text.Json;

namespace TableParse
{
    internal class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            string templatePath = "Templates/iec60598_2_1i_2022-08-26.docx";
            var tableTypes = new List<int>()
            {
                1,1,1,1,2,2,2,2,2,3,3,3,3,3,3,3,4,3,3,3,3,0,0
            };

#else
            string templatePath = args[0];
            var tableTypes = new List<int>();
            if (args.Length > 1) 
            {
                foreach(char c in args[1])
                {
                    tableTypes.Add(int.Parse(c.ToString()));
                }
            }
            
#endif
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                var tables = doc.MainDocumentPart?.Document.Body?.Descendants<Table>().ToList();
                if (tables == null||tables.Count==0) throw new NullReferenceException(nameof(tables));
                var tableIdx = 0;
                foreach (var table in tables)
                {
                    
                    switch (tableTypes[tableIdx])
                    {
                        case 1:
                        case 3:
                        case 4:
                            var fTable = new FillTable<FillRow>();
                            fTable.TableType=((TableType)tableTypes[tableIdx]).ToString();
                            ParseTable(table, fTable,tableIdx);
                            Console.WriteLine(JsonSerializer.Serialize(fTable));
                            break;
                        case 2:
                            var cTable = new FillTable<ClauseFillRow>();
                            cTable.TableType = ((TableType)tableTypes[tableIdx]).ToString();
                            ParseComplainceTable(table,cTable,tableIdx);
                            Console.Write(JsonSerializer.Serialize(cTable));
                            break;
                        default:
                            break;
                    }
                    
                    tableIdx++;
                }
                
                Environment.Exit(0); 
            }
        }
        private static void ParseTable(Table table,FillTable<FillRow> fTable,int tableIdx)
        {
            fTable.TableIdx = tableIdx;
            var rows = table.Descendants<TableRow>().ToList();
            var rowIdx = 0;
            fTable.Rows = new List<FillRow>();
            foreach (TableRow row in rows)
            {
                var fRow = new FillRow();
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
        }
        
        private static void ParseComplainceTable(Table table,FillTable<ClauseFillRow> fTable,int tableIdx)
        {
            fTable.TableIdx = tableIdx;
            var rows = table.Descendants<TableRow>().ToList();
            fTable.Rows = new List<ClauseFillRow>();
            var currentClauseNo = "";
            var idxInClause = 0;
            foreach(var (row,rowIdx) in rows.Select((row, index) => (row, index)))
            {
                
                var rowType= DectectRowType(row);
                var fRow = new ClauseFillRow();
                fRow.TableIdx = tableIdx;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                fRow.Cells = new List<FillCell>();
                var cells = row.Descendants<TableCell>().ToList();
                if (!string.IsNullOrEmpty(cells[0].InnerText))
                {
                    currentClauseNo = cells[0].InnerText;
                    idxInClause = 0;
                }
                else
                {
                    idxInClause++;
                }
                fRow.ClauseNo = currentClauseNo;
                fRow.IdxUnderClause = idxInClause;
                foreach(var (cell,columnIdx) in cells.Select((cell,index) => (cell, index)))
                {
                    // fetch table title
                    if(rowIdx == 0 && columnIdx == 1 && !string.IsNullOrEmpty(cell.InnerText))
                    {
                        fTable.TableTitle= cell.InnerText;
                    }
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
                }
                
                fTable.Rows.Add(fRow);
            }
         }
        private static bool IsShadingCell(TableCell cell)
        {
            var shading = cell.Descendants<Shading>().FirstOrDefault();
            return shading?.Val != null;
        }
        private static RowType DectectRowType(TableRow row)
        {
            var cells = row.Descendants<TableCell>().ToList();
            if (cells.Count == 1) return RowType.Unknown;
            if (cells.Count == 3)
            {
                if (IsShadingCell(cells[1]))
                {
                    return RowType.SectionHeader;
                }
                return RowType.ClauseHeader;
            }
            else
            {
                if (IsShadingCell(cells[3]))
                {
                    return RowType.InfoItem;
                }
                else
                {
                    return RowType.VerdictItem;
                }
            }

        }
    }
}
