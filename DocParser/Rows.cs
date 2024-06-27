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
    public class Rows
    {
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
                        var tempRow = new ReportRow(rowType, currentClause,i, cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText);
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
        private bool IsShadingCell(TableCell cell,string shadingColor="D9D9D9")
        {
            var shading = cell.Descendants<Shading>().FirstOrDefault();
            if (shading?.Val!=null&&shading.Val==shadingColor)
            {
                
                    return true;
               
            }
            return false;
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
