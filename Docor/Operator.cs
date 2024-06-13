using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace Docor
{
    public static class Operator
    {
        public static void ChangeCheckbox(MainDocumentPart mainPart,string bookmark,bool val)
        {
            BookmarkStart bookmarkStart = mainPart.Document.Body.Descendants<BookmarkStart>()
                                                 .FirstOrDefault(b => b.Name == bookmark);

            if (bookmarkStart != null)
            {
                Paragraph parentParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
                if (parentParagraph != null)
                {
                    var sdtElement = parentParagraph.Descendants<FormFieldData>().FirstOrDefault();
                    if (sdtElement != null)
                    {
                        var checkbox = sdtElement.GetFirstChild<CheckBox>();
                        var check = checkbox.Descendants<DefaultCheckBoxFormFieldState>().FirstOrDefault();
                        check.Val = OnOffValue.FromBoolean(val);
                    }
                }
            }
        }
        public static void ReplacePlaceholders(MainDocumentPart mainPart, Dictionary<string, string> replacements)
        {

                // 获取文档正文内容
                var documentBody = mainPart.Document.Body;

                // 获取所有文本节点
                var texts = documentBody.Descendants<Text>().ToList();
                foreach (var text in texts)
                {
                    foreach (var placeholder in replacements.Keys)
                    {
                        // 确保处理被拆分的占位符
                        if (text.Text.Contains(placeholder))
                        {
                            // 替换占位符
                            text.Text = text.Text.Replace(placeholder, replacements[placeholder]);
                        }
                    }
                }
        }


        public static void fillTextField(MainDocumentPart mainPart,int tableIdx,int rowIdx,int cellIdx,string text)
        {
            var documentBody = mainPart.Document.Body;
            var table = documentBody?.Descendants<Table>().Take(new Range(tableIdx,tableIdx+1)).FirstOrDefault();
            var row = table?.Descendants<TableRow>().Take(new Range(rowIdx,rowIdx+1)).FirstOrDefault();
            var cell = row?.Descendants<TableCell>().Take(new Range(cellIdx,cellIdx+1)).FirstOrDefault();
            var x = cell?.Descendants<DefaultTextBoxFormFieldString>().FirstOrDefault();
            x.Val=text;
        }
        public static void VerdictRow(MainDocumentPart mainPart,int rowIdx){
            var documentBody=mainPart.Document.Body;
            var rows=documentBody.Descendants<TableRow>().ToList();
        }
    }
}
