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
                foreach (var row in rows) 
                {
                    var rowType= DectectRowType(row); 
                    var cells=row.Descendants<TableCell>().ToList();
                    if (cells.Count == 1) continue;
                    if (cells.Count == 3)
                    {
                        var tempRow = new ReportRow(rowType, cells[0].InnerText, cells[1].InnerText, "", cells[2].InnerText);
                       
                        resultList.Add(tempRow);
                    }
                    else
                    {
                        var tempRow = new ReportRow(rowType, cells[0].InnerText, cells[1].InnerText, cells[2].InnerText, cells[3].InnerText);
                        
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
                    if (IfRemarksNeeded(cells[1]))
                    {
                        return RowType.InfoItemWithRemark;
                    }
                    return RowType.InfoItem;
                }
                else
                {
                    if (IfRemarksNeeded(cells[1]))
                    {
                        return RowType.VerdictItemWithRemark;
                    }
                    return RowType.VerdictItem;
                }
            }
            
        }
        private bool IfRemarksNeeded(TableCell cell)
        {
            if (cell.InnerText.Contains("...:"))
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
        public ReportRow(RowType rowType,string a="",string b="",string c="",string d="")
        {
            CellA = a;
            CellB = b;
            CellC = c;
            CellD = d;
            RowType = rowType;
        }
        public RowType RowType { get; set; }
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
        VerdictItemWithRemark,
        InfoItem,
        InfoItemWithRemark,
        Unknown
    }
}
