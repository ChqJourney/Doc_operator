using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace DocParser
{
    public class GeneralTable
    {
        public List<GeneralRow> Rows { get; set; } = new List<GeneralRow>();
    }
    public class GeneralRow
    {
        public List<GeneralCell> Cells { get; set; } = new List<GeneralCell>();
    }
    public class GeneralCell
    {
        public int RowIdx { get; set; }
        public int ColumnIdx { get; set; }
        public string CellTxt { get; set; }
        public bool Bolded { get; set; }
        public bool HasShading { get; set; }
        public int CellWidth { get; set; }
        public bool IsCentered { get; set; }

        public string HMerge { get; set; }
        public string VMerge { get; set; }
    }
    public class Rows
    {
        public Task<GeneralTable> ParseGeneralRows(MainDocumentPart main, int tableIdx, Action<String> action = null)
        {
            var table = main.Document.Body?.Descendants<Table>().ElementAt(tableIdx-1);
            if (table == null)throw new NullReferenceException(nameof(table));
            var result = new GeneralTable();
            var rows = table.Descendants<TableRow>().ToList();

            var rowIdx = 0;
            foreach (var row in rows)
            {
                var grow = new GeneralRow();
                
                var cells=row.Descendants<TableCell>().ToList();
                var columnIdx = 0;
                foreach (var cell in cells)
                {
                    var gc=new GeneralCell();
                    gc.RowIdx = rowIdx;
                    gc.ColumnIdx = columnIdx;
                    var j=cell.Descendants<Justification>().FirstOrDefault();

                    string cellContent = "";
                    if (!string.IsNullOrEmpty(cell.InnerText))
                    {
                        var paras = cell.Descendants<Paragraph>().ToList();
                        var bold=cell.Descendants<Bold>().ToList();
                        if (bold.Count > 0) 
                        { 
                            gc.Bolded = true;
                        }
                        else
                        {
                            gc.Bolded = false;
                        }
                        var texts = paras.Select(x => x.InnerText).ToList();
                        cellContent = string.Join('\n', texts);
                    }
                    var cellWidth = cell.TableCellProperties?.TableCellWidth?.Width?.Value;
                    var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val;
                    var vMerge = cell.TableCellProperties?.VerticalMerge?.Val;
                    var hasShading=IsShadingCell(cell);
                    if (j != null)
                    {
                        gc.IsCentered = j.Val == "center";
                    }
                    gc.HasShading = hasShading;
                    gc.CellTxt= cellContent;
                    gc.CellWidth=int.Parse(cellWidth);
                    gc.HMerge = hMerge;
                    gc.VMerge = vMerge;
                    grow.Cells.Add(gc);
                    columnIdx++;
                }
                result.Rows.Add(grow);
                rowIdx++;
            }
            action.Invoke(JsonSerializer.Serialize(result));
            return Task.FromResult(result);
        }
        public Task<List<ReportRow>> ParseRows(MainDocumentPart main,int fromTableIdx, int toTableIdx,Action<String> action=null)
        {
            var tables=main.Document.Body?.Descendants<Table>().Take(new Range(fromTableIdx, toTableIdx)).ToList();
            if(tables==null || tables.Count == 0)
            {
                return Task.FromResult(new List<ReportRow>());
            }
            var resultList=new List<ReportRow>();
            foreach (var table in tables) 
            {
                var rows=table.Descendants<TableRow>().ToList();
                var currentClause = "";
                var i = 0;
                foreach (var row in rows) 
                {
                    var rowType= DectectRowType(row); 
                    var cells=row.Descendants<TableCell>().ToList();
                    if (cells.Count == 1) continue;
                    if (!string.IsNullOrEmpty(cells[0].InnerText))
                    {
                        currentClause = cells[0].InnerText;
                        i = 0;
                    }
                    
                    if (cells.Count == 3)
                    {
                        
                        var tempRow = new ReportRow(rowType,currentClause,i, cells[0].InnerText, cells[1].InnerText, "", cells[2].InnerText);
                        i++;
                        resultList.Add(tempRow);
                    }
                    else
                    {
                            string cellContent="";
                        if (!string.IsNullOrEmpty(cells[2].InnerText))
                        {
                            var paras = cells[2].Descendants<Paragraph>().ToList();
                            var texts=paras.Select(x=>x.InnerText).ToList();
                            cellContent=string.Join('\n',texts);
                        }
                        var tempRow = new ReportRow(rowType, currentClause,i, cells[0].InnerText, cells[1].InnerText, cellContent, cells[3].InnerText);
                        i++;
                        resultList.Add(tempRow);
                    }
                    if (action != null)
                    {
                        action.Invoke(JsonSerializer.Serialize(resultList.Last()));
                    }
                }
                
            }
            return Task.FromResult(resultList);
        }
        public RowType DectectRowType(TableRow row)
        {
            var cells=row.Descendants<TableCell>().ToList();
            if(cells.Count==1)return RowType.Unknown;
            if (cells.Count==3)
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
        private bool IfRemarksNeeded(TableCell cell)
        {
            if (cell.InnerText.Contains("...:")||cell.InnerText.Contains("\u2026\u2026\u2026:"))
            {
                return true;
            }
            return false;
        }
        private bool IsShadingCell(TableCell cell)
        {
            var shading = cell.Descendants<Shading>().FirstOrDefault();
            return shading?.Val != null;
        }
    }
    public class ReportRow
    {
        public ReportRow(RowType rowType,string clauseNo,int idx,string a="",string b="",string c="",string d="")
        {
            CellA = a;
            CellB = b;
            CellC = c;
            CellD = d;
            RowType = rowType;
            ClauseNo = clauseNo;
            IdxUnderClause = idx;
        }
        public RowType RowType { get; set; }
        public string ClauseNo { get; set; }
        public int IdxUnderClause { get; set; }
        public string CellA { get; set; }
        public string CellB { get; set; }
        public string CellC { get; set; }
        public string CellD { get; set; }
        public string Remark { get; set; }
        public string Verdict { get; set; }
    }
   
    public enum RowType
    {
        SectionHeader,
        ClauseHeader,
        VerdictItem,
        InfoItem,
        Unknown
    }
}
