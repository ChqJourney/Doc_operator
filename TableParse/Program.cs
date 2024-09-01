using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TableParse
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string action = args[0];
            string templatePath = args[1];
            var tableTypes = new List<int>();
#if DEBUG


            if (args.Length > 2 && args[2] != "")
            {
                foreach (char c in args[2])
                {
                    tableTypes.Add(int.Parse(c.ToString()));
                }
            }

            tableTypes = new List<int>()
            {
                1,1,1,1,2,2,2,2,2,3,3,3,3,3,3,3,4,3,3,3,3,5,5
            };

#else
           
            if (args.Length > 2 && args[2]!="") 
            {
                foreach(char c in args[2])
                {
                    tableTypes.Add(int.Parse(c.ToString()));
                }
            }
            
#endif
            var path=Path.Combine(Environment.CurrentDirectory,templatePath);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                var tables = doc.MainDocumentPart?.Document.Body?.Descendants<Table>().ToList();
                
                if (tables == null || tables.Count == 0) throw new ArgumentNullException(nameof(tables));
#if DEBUG
#else
                if (tableTypes.Count!=0&&tableTypes.Count!=tables.Count)throw new ArgumentNullException(nameof(tableTypes));
#endif
                var tableIdx = 0;
                foreach (var table in tables)
                {

                    switch (action)
                    {
                        case "tables":
                            var fTable = new FillTable<FillRow>();
                            ParseTable(table, fTable, tableIdx);
                            Console.WriteLine(JsonSerializer.Serialize(fTable));
                            break;
                        case "rows":
                            ParseRows(table, tableIdx, tableTypes,doc.MainDocumentPart);
                            break;
                        default:
                            break;
                    }


                    tableIdx++;
                }

                Environment.Exit(0);
            }
        }
        private static void ParseTable(Table table, FillTable<FillRow> fTable, int tableIdx)
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

        private static void ParseRows(Table table, int tableIdx, List<int> tableTypes,MainDocumentPart main)
        {

            var rows = table.Descendants<TableRow>().ToList();
            var currentRowType = tableTypes.Count == 0 ? 1 : tableTypes[tableIdx];

            switch (currentRowType)
            {
                //general
                case 1:
                    parseAndEmitGeneralRows(rows, tableIdx,main);
                    break;
                //compliance
                case 2:
                    parseAndEmitComplianceRows(rows, tableIdx);
                    break;
                //measurement
                case 3:
                    parseAndEmitMeasurementRows(rows, tableIdx);
                    break;
                //component
                case 4:
                    parseAndEmitComponentRows(rows, tableIdx);
                    break;
                //equipment
                case 5:
                    parseAndEmitComponentRows(rows, tableIdx);
                    break;
                default:
                    break;
            }

        }
        private static void parseAndEmitGeneralRows(List<TableRow> rows, int tableIdx,MainDocumentPart main)
        {
            foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
            {
                var fRow = new FillRow();
                fRow.TableIdx = tableIdx;
                fRow.TableType = TableType.General;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                var cells = row.Descendants<TableCell>().ToList();
                fRow.Cells = new List<FillCell>();
                foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
                {
                    var fCell = new FillCell();
                    fCell.ColumnIdx = columnIdx;
                    fCell.RowIdx = rowIdx;
                    

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
                    //var hasShading = IsShadingCell(cell);
                    var j = cell.Descendants<Justification>().FirstOrDefault();
                    if (j != null)
                    {
                        fCell.IsCentered = j.Val == "center";
                    }
                    var pStyle=cell.Descendants<ParagraphStyleId>().FirstOrDefault();
                    if (pStyle != null)
                    {
                        var pStyles = main.StyleDefinitionsPart.Styles.Descendants<Style>().ToList();
                        var st=pStyles.Where(p => p.StyleId == pStyle.Val).SingleOrDefault();
                        var jj = st.Descendants<Justification>().SingleOrDefault();
                        if (jj != null)
                        {
                            fCell.IsCentered = jj.Val == "center";
                        }
                    }
                    //fCell.HasShading = hasShading;
                    fCell.OriginalText = cellContent;
                    fCell.CellWidth = int.Parse(cellWidth);
                    fCell.HMerge = hMerge;
                    fCell.VMerge = vMerge;
                    fRow.Cells.Add(fCell);
                }
                Console.WriteLine(JsonSerializer.Serialize(fRow));
            }
        }
        private static void parseAndEmitComplianceRows(List<TableRow> rows, int tableIdx)
        {
            var currentClauseNo = "";
            var idxInClause = 0;
            foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
            {
                var cells = row.Descendants<TableCell>().ToList();
                var fRow = new ClauseFillRow();
                fRow.Cells = new List<FillCell>();
                fRow.TableIdx = tableIdx;
                fRow.TableType = TableType.Compliance;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                fRow.Deletable = false;
                fRow.RowType = DectectRowType(row);
                if (cells.Count > 1)
                {

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
                if (idxInClause == 0)
                {
                    fRow.ClauseTitle = cells[1].InnerText;
                }
                }

                foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
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
                    var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val?.ToString() ?? null;
                    var vMerge = cell.TableCellProperties?.VerticalMerge?.Val?.ToString() ?? null;
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

                Console.WriteLine(JsonSerializer.Serialize(fRow));
            }
        }
        private static void parseAndEmitMeasurementRows(List<TableRow> rows,int tableIdx)
        {
            var currentClauseNo = "";
            var idxInClause = 0;
            foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
            {
                var cells = row.Descendants<TableCell>().ToList();
                var fRow = new ClauseFillRow();
                fRow.TableType=TableType.Measurement;
                fRow.Cells = new List<FillCell>();
                fRow.TableIdx = tableIdx;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                fRow.Deletable = false;
                if (!string.IsNullOrEmpty(cells[0].InnerText)&&rowIdx==0)
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
                if (idxInClause == 0)
                {
                    fRow.ClauseTitle = cells[1].InnerText;
                }

                foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
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
                    var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val?.ToString() ?? null;
                    var vMerge = cell.TableCellProperties?.VerticalMerge?.Val?.ToString() ?? null;
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

                Console.WriteLine(JsonSerializer.Serialize(fRow));
            }
        }
        private static void parseAndEmitComponentRows(List<TableRow> rows,int tableIdx)
        {
            var currentClauseNo = "";
            var idxInClause = 0;
            foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
            {
                if (!new int[3] { 0, 1, rows.Count - 1 }.Contains(rowIdx)) continue;
                var cells = row.Descendants<TableCell>().ToList();
                var fRow = new ClauseFillRow();
                fRow.Cells = new List<FillCell>();
                fRow.TableType = TableType.Component;
                fRow.TableIdx = tableIdx;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                fRow.Deletable = false;
                if (!string.IsNullOrEmpty(cells[0].InnerText) && rowIdx == 0)
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

                foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
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
                    var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val?.ToString() ?? null;
                    var vMerge = cell.TableCellProperties?.VerticalMerge?.Val?.ToString() ?? null;
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

                Console.WriteLine(JsonSerializer.Serialize(fRow));
            }
        }
        private static void parseAndEmitEquipmentRows(List<TableRow> rows, int tableIdx)
        {
            var currentClauseNo = "";
            var idxInClause = 0;
            foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
            {
                if (!new int[2] { 0, 1}.Contains(rowIdx)) continue;
                var cells = row.Descendants<TableCell>().ToList();
                var fRow = new ClauseFillRow();
                fRow.TableType = TableType.Equipment;
                fRow.Cells = new List<FillCell>();
                fRow.TableIdx = tableIdx;
                fRow.RowIdx = rowIdx;
                fRow.Duplicatable = false;
                fRow.Deletable = false;
                if (!string.IsNullOrEmpty(cells[0].InnerText) && rowIdx == 0)
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

                foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
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
                    var hMerge = cell.TableCellProperties?.HorizontalMerge?.Val?.ToString() ?? null;
                    var vMerge = cell.TableCellProperties?.VerticalMerge?.Val?.ToString() ?? null;
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

                Console.WriteLine(JsonSerializer.Serialize(fRow));
            }
        }
        private static bool IsShadingCell(TableCell cell)
        {
            var shading = cell.Descendants<Shading>().FirstOrDefault();
            return shading != null && shading.Fill?.Value!="auto";
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
            else if(cells.Count==4)
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
            else
            {
                return RowType.Unknown;
            }

        }
    }
    public class IgnoreEmptyStringConverter : JsonConverter<string>
    {
        public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            return reader.GetString();
        }

        public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options)
        {
            if (string.IsNullOrEmpty(value))
            {
                writer.WriteNullValue();
            }
            else
            {
                writer.WriteStringValue(value);
            }
        }
    }
}
