using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocParser
{
    public class TableRows
    {
        public Task ParseRows(MainDocumentPart main,int fromTableIdx, int toTableIdx)
        {
            var tables=main.Document.Body.Descendants<Table>().Take(new Range(fromTableIdx, toTableIdx)).ToList();
            foreach (var table in tables) 
            {
                var rows=table.Descendants<ReportRow>().ToList();

                var temp=new Table()
            }
        }
    }
    public class ReportRow
    {
        public ReportRow(string a="",string b="",string c="",string d="")
        {
            CellA = a;
            CellB = b;
            CellC = c;
            CellD = d;
        }
        public string CellA { get; set; }
        public string CellB { get; set; }
        public string CellC { get; set; }
        public string CellD { get; set; }
    }
   
    public enum RowType
    {
        SectionHeader,
        ClauseHeader,
        VerdictItem,
        InfoItem
    }
}
