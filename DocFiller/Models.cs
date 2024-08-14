using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace DocFiller
{
    public class Report
    {
        public string ReportNo { get; set; }
        public string TrfPath { get; set; }
        public List<CellEdit> CellEdits { get; set; }
        public List<RowEdit> RowEdits { get; set; }
        public List<TableEdit> TableEdits { get; set; }
        public List<Append> Appends { get; set; }

    }
    public class RowEdit
    {
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public string EditType { get; set; } //delete of duplicate after
    }
    public class TableEdit
    {
        public int TableIdx { get; set; }
        public string EditType { get; set; } //delete or duplicate after
    }
    public class Append
    {
        public string TemplatePath { get; set; }
        public List<CellEdit> CellEdits { get; set; }
        public List<RowEdit> RowEdits { get; set; }
        public List<Append> Appends { get; set; }

    }
    public class GeneralField 
    {
        public int CommentId { get; set; }
        public FieldEdit Edit { get; set; }
    }
    public class FieldEdit
    {
        public string InputValue { get; set; }
        public string TargetType { get; set; }//text,checkbox,textbox,photo
        public string EditType { get; set; }//replace,append,before
        public string AlignType { get; set; }
        public List<string> Hints { get; set; }
        public List<string> RulesBasedOnInput { get; set; }
    }
    public class PhotoBlock
    {

    }

    public class FillTable
    {
        public int TableIdx { get; set; }
        public List<FillRow> Rows { get; set; }
    }
    public class FillRow
    {
        public bool Duplicatable { get; set; }
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public List<FillCell> Cells { get; set; }

    }
    public class ClauseFillRow : FillRow
    {
        public string ClauseNo { get; set; }
        public int IdxUnderClause { get; set; }
    }
    public class FillCell
    {
        public int ColumnIdx { get; set; }
        public string OriginalText { get; set; }
        public List<CellEdit> CellEdits { get; set; }

    }
    public class CellEdit
    {
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public int ColumnIdx { get; set; }
        public string InputValue { get; set; }
        // js type
        public string TargetType { get; set; }
        public string InsertType { get; set; }
        public string AlignType { get; set; }
        public List<string> Hints { get; set; }
        public List<string> RulesBasedOnInput { get; set; }
    }
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
        public int TableIdx { get; set; }
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
            var table = main.Document.Body?.Descendants<Table>().ElementAt(tableIdx - 1);
            if (table == null) throw new NullReferenceException(nameof(table));
            var result = new GeneralTable();
            var rows = table.Descendants<TableRow>().ToList();

            var rowIdx = 0;
            foreach (var row in rows)
            {
                var grow = new GeneralRow();

                var cells = row.Descendants<TableCell>().ToList();
                var columnIdx = 0;
                foreach (var cell in cells)
                {
                    var gc = new GeneralCell();
                    gc.RowIdx = rowIdx;
                    gc.ColumnIdx = columnIdx;
                    var j = cell.Descendants<Justification>().FirstOrDefault();

                    string cellContent = "";
                    if (!string.IsNullOrEmpty(cell.InnerText))
                    {
                        var paras = cell.Descendants<Paragraph>().ToList();
                        var bold = cell.Descendants<Bold>().ToList();
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
                    var hasShading = IsShadingCell(cell);
                    if (j != null)
                    {
                        gc.IsCentered = j.Val == "center";
                    }
                    gc.HasShading = hasShading;
                    gc.CellTxt = cellContent;
                    gc.CellWidth = int.Parse(cellWidth);
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
        public Task<List<ReportRow>> ParseRows(MainDocumentPart main, int fromTableIdx, int toTableIdx, Action<String> action = null)
        {
            var tables = main.Document.Body?.Descendants<Table>().Take(new Range(fromTableIdx, toTableIdx)).ToList();
            if (tables == null || tables.Count == 0)
            {
                return Task.FromResult(new List<ReportRow>());
            }
            var resultList = new List<ReportRow>();
            foreach (var table in tables)
            {
                var rows = table.Descendants<TableRow>().ToList();
                var currentClause = "";
                var i = 0;
                foreach (var row in rows)
                {
                    var rowType = DectectRowType(row);
                    var cells = row.Descendants<TableCell>().ToList();
                    if (cells.Count == 1) continue;
                    if (!string.IsNullOrEmpty(cells[0].InnerText))
                    {
                        currentClause = cells[0].InnerText;
                        i = 0;
                    }

                    if (cells.Count == 3)
                    {

                        var tempRow = new ReportRow(rowType, currentClause, i, cells[0].InnerText, cells[1].InnerText, "", cells[2].InnerText);
                        i++;
                        resultList.Add(tempRow);
                    }
                    else
                    {
                        string cellContent = "";
                        if (!string.IsNullOrEmpty(cells[2].InnerText))
                        {
                            var paras = cells[2].Descendants<Paragraph>().ToList();
                            var texts = paras.Select(x => x.InnerText).ToList();
                            cellContent = string.Join('\n', texts);
                        }
                        var tempRow = new ReportRow(rowType, currentClause, i, cells[0].InnerText, cells[1].InnerText, cellContent, cells[3].InnerText);
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
        private bool IfRemarksNeeded(TableCell cell)
        {
            if (cell.InnerText.Contains("...:") || cell.InnerText.Contains("\u2026\u2026\u2026:"))
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
        public ReportRow(RowType rowType, string clauseNo, int idx, string a = "", string b = "", string c = "", string d = "")
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
