using DocumentFormat.OpenXml.Wordprocessing;

namespace ValidateApp;

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
    private readonly string _filePath;
    private readonly List<Func<TableRow, RowType>> _detectors;

    public WordTableProcessor(Table table)
    {
        this.Rows = table.Descendants<TableRow>().ToList();
        _detectors = new List<Func<TableRow, RowType>>
            {
                DetectSectionHeader,
                DetectClauseHeader,
                DetectInfoItem,
                DetectVerdictItem
            };
    }
    public List<TableRow> Rows { get; set; }
    public List<Clause> Clauses { get; set; }
    private int currentRowIdx;
    public void Process()
    {

        while (currentRowIdx < Rows.Count){
            var row = Rows[currentRowIdx];
            var rowType = DetectRowType(row);
            if (rowType == RowType.ClauseHeader)
            {
                var clause = new Clause();
                clause.ClauseNo = row.Elements<TableCell>().ElementAt(0).InnerText;
                clause.ClauseTitle = row.Elements<TableCell>().ElementAt(1).InnerText;
                Clauses.Add(clause);
            }
            while (currentRowIdx < Rows.Count && DetectRowType(Rows[currentRowIdx]) != RowType.ClauseHeader)
            {
                var item = new Item();
                
                item.ClauseNo = row.Elements<TableCell>().ElementAt(0).InnerText;
                item.ItemBodyText = row.Elements<TableCell>().ElementAt(1).InnerText;
                Clauses.Last().Items.Add(item);
                currentRowIdx++;
            }
            // FillCell(row,rowType);
            // currentRowIdx++;
        }
        
    }

    private RowType DetectRowType(TableRow row)
    {
        foreach (var detector in _detectors)
        {
            var result = detector(row);
            if (result != RowType.Unknown)
            {
                return result;
            }
        }
        return RowType.Unknown;
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

    // 根据底色和单元格数检测 Section Header
    private RowType DetectSectionHeader(TableRow row)
    {
        var cell = row.Elements<TableCell>().ElementAtOrDefault(1);
        if (cell != null && !IsWhiteBackground(cell))
        {
            return RowType.SectionHeader;
        }
        return RowType.Unknown;
    }

    // 根据单元格数检测 Clause Header
    private RowType DetectClauseHeader(TableRow row)
    {
        if (row.Elements<TableCell>().Count() == 3)
        {
            return RowType.ClauseHeader;
        }
        return RowType.Unknown;
    }

    // 检测 Info Item
    private RowType DetectInfoItem(TableRow row)
    {
        if (row.Elements<TableCell>().Count() > 3)
        {
            var cell = row.Elements<TableCell>().ElementAtOrDefault(3);
            if (cell != null && !IsWhiteBackground(cell))
            {
                return RowType.InfoItem;
            }
        }
        return RowType.Unknown;
    }

    // 检测 Verdict Item
    private RowType DetectVerdictItem(TableRow row)
    {
        if (row.Elements<TableCell>().Count() > 3)
        {
            var cell = row.Elements<TableCell>().ElementAtOrDefault(3);
            if (cell != null && IsWhiteBackground(cell))
            {
                return RowType.VerdictItem;
            }
        }
        return RowType.Unknown;
    }

    // 检查单元格底色是否为白色
    private bool IsWhiteBackground(TableCell cell)
    {
        var shading = cell.Descendants<Shading>().FirstOrDefault();
        return shading == null || shading.Fill == "FFFFFF" || shading.Fill == "white";
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