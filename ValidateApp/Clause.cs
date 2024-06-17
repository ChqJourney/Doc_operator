using DocumentFormat.OpenXml.Wordprocessing;

namespace ValidateApp;

public class Section
{
    public string SectionNo { get; set; }
    public string SectionTitle { get; set; }
    public List<Clause> Clauses { get; set; }
}
public class Item
{
    public string ClauseNo {get;set;}
    public string ItemBodyText{get;set;}
}
public class Clause
{
    public String ClauseNo {get;set;}
    public string ClauseTitle{get;set;}
    public List<Item> Items {get;set;}
}
public class WordTableProcessor
{

    public WordTableProcessor(Table table)
    {
        this.Rows = table.Descendants<TableRow>().ToList();
       
    }
    public List<TableRow> Rows { get; set; }
    public List<Section> Sections { get; set; }=new List<Section>();
    private int currentRowIdx;
    public void Process()
    {

        while (currentRowIdx < Rows.Count){
            var row = Rows[currentRowIdx];
            var rowType = DecectRow(row);
            
        }
        
    }

  

    private void FillCell(TableRow row, RowType rowType)
    {
        switch (rowType)
        {
            case RowType.SectionHeader:
                // No fill
                break;
            case RowType.ClauseHeader:
                // Determine based on subsequent rows
                FillClauseHeader(row);
                break;
            case RowType.InfoItem:
                // No fill
                break;
            case RowType.VerdictItem:
                // Fill based on rules
                FillVerdictItem(row);
                break;
        }
    }

    private void FillClauseHeader(TableRow row)
    {
       
    }

    private void FillVerdictItem(TableRow row)
    {
        var fillCell = row.Elements<TableCell>().LastOrDefault();
        if (fillCell != null)
        {
            fillCell.RemoveAllChildren<Paragraph>();
            fillCell.Append(new Paragraph(new Run(new Text("P")))); // or "N/A" based on rules
        }
    }

    public RowType DecectRow(TableRow row)
    {
        if(row.Elements<TableCell>().Count() == 3)
        {
            var cell = row.Elements<TableCell>().ElementAtOrDefault(1);
            if (IsShadingCell(cell))
            {
                return RowType.SectionHeader;
            }
            else
            {
                return RowType.ClauseHeader;
            }
        }
        else
        {

            var cell = row.Elements<TableCell>().ElementAtOrDefault(row.Elements<TableCell>().Count()-1);
            
            if (cell!=null&&IsShadingCell(cell))
            {
                return RowType.InfoItem;
            }
            else if(cell!=null) 
            {
            
                return RowType.VerdictItem;
            }
            else
            {
                return RowType.Unknown;
            }
        }
    }
    

    // 检查单元格是否有底色
    private bool IsShadingCell(TableCell cell)
    {
        var shading = cell.Descendants<Shading>().FirstOrDefault();
        return shading.Fill == "D9D9D9" ;
    }
}


public enum RowType
{
    SectionHeader,
    ClauseHeader,
    InfoItem,
    VerdictItem,
    Unknown
}