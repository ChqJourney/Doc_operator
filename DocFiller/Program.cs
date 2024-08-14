using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocFiller
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string reportPath = "C:\\Users\\patri\\Desktop\\123\\240400111HZH-005.rpt";
            Console.WriteLine(reportPath);
            string content = File.ReadAllText(reportPath);
            var report = JsonSerializer.Deserialize<Report>(content);
            File.Copy(report.TrfPath, "123.docx", true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open("123.docx", true))
            {
                var tables = doc.MainDocumentPart?.Document.Body?.Descendants<Table>().ToList();
                var paras = doc.MainDocumentPart.HeaderParts.ToList();
                var ids=new List<string>();
                foreach (var item in paras)
                {
                   var id=doc.MainDocumentPart.GetIdOfPart(item);
                    ids.Add(id);
                }
                var refss = doc.MainDocumentPart.Document.Body.Elements();
                // 创建一个新的段落
                Paragraph newSectionParagraph = new Paragraph();

                // 创建段落属性，设置分节符
                ParagraphProperties newSectionProperties = new ParagraphProperties();
                SectionProperties sectionProperties = new SectionProperties();
                newSectionProperties.Append(sectionProperties);
                // 将段落属性添加到段落
                newSectionParagraph.Append(newSectionProperties);
                doc.MainDocumentPart.Document.Body.Append(newSectionParagraph);
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>();
                if (tables == null || tables.Count == 0) throw new ArgumentNullException(nameof(tables));
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<HeaderReference>();

                    // Create the new header reference node.
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = "rId17" });
                }
                foreach (var (table, tableIdx) in tables.Select((table, index) => (table, index)))
                {
                    var rows = table.Descendants<TableRow>().ToList();
                    foreach (var (row, rowIdx) in rows.Select((row, index) => (row, index)))
                    {
                        var cells = row.Descendants<TableCell>().ToList();
                        foreach (var (cell, columnIdx) in cells.Select((cell, index) => (cell, index)))
                        {
                            var edits = report.CellEdits.Where(e => e.TableIdx == tableIdx && e.RowIdx == rowIdx && e.ColumnIdx == columnIdx).ToList();
                            if (edits == null || edits.Count == 0) continue;
                            foreach (var (edit, editIdx) in edits.Select((edit, index) => (edit, index)))
                            {
                                switch (edit.TargetType)
                                {
                                    case "text":
                                        handlePlainText(edit, cell, editIdx);
                                        break;
                                    case "checkbox":
                                        handleCheckBox(edit, cell, editIdx);
                                        break;
                                    case "textbox":
                                        handleTextBox(edit,cell,editIdx);
                                        break;
                                    case "tables":
                                        break;
                                    default:
                                        break;
                                }

                                
                            }
                        }


                    }

                }
                doc.MainDocumentPart?.Document.Save();
            }
        }
        private static void handlePlainText(CellEdit edit, TableCell cell, int paraIdx)
        {

            var para = cell.Descendants<Paragraph>().ElementAt(paraIdx);
            switch (edit.InsertType)
            {
                case "replace":
                    para.RemoveAllChildren();
                    para.AppendChild(new Run(new Text(edit.InputValue)));
                    break;
                case "after":
                    para.AppendChild(new Run(new Text(edit.InputValue)));
                    break;
                case "start":
                    var originRun = para.Descendants<Run>().SingleOrDefault();
                    para.InsertBefore(new Run(new Text(edit.InputValue)), originRun);
                    break;
                default:
                    break;
            }
            if (edit.AlignType != null && edit.AlignType == "center")
            {
                para.ParagraphProperties = new ParagraphProperties();
                para.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Center };
            }

        }
        private static void handleCheckBox(CellEdit edit, TableCell cell, int paraIdx)
        {
            bool[] arr = new bool[edit.InputValue.Length];
            for (int i = 0; i < edit.InputValue.Length; i++)
            {
                arr[i] = edit.InputValue[i] == '1';
            }
            var para=cell.Descendants<Paragraph>().ElementAt(paraIdx);
            if (para == null) return;
            var ffds = para.Descendants<DefaultCheckBoxFormFieldState>().ToList();
            if (ffds == null || ffds.Count == 0) return;
            for (int i = 0; i < ffds.Count; i++)
            {
                ffds[i].Val = OnOffValue.FromBoolean(arr[i]);
            }
        }

        private static void handleTextBox(CellEdit edit, TableCell cell, int paraIdx)
        {
            var para = cell.Descendants<Paragraph>().ElementAt(paraIdx);
            if (para == null) return;
            var txs=para.Descendants<TextInput>().SingleOrDefault();
            if (txs == null) return;
            txs.DefaultTextBoxFormFieldString=new DefaultTextBoxFormFieldString();
            txs.DefaultTextBoxFormFieldString.Val= edit.InputValue;
            
        }

    }

}
