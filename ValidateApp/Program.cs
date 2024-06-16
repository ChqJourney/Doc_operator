using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Docor;
using System.Collections.Immutable;
using System.Linq;
using ConsoleTables;

namespace ValidateApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 输入模板文件路径
            string templatePath = "iec60598_2_1i_2022-08-26.docx";


            var t = new ConsoleTable("table no", "row no", "cell no", "comment id", "comment content");
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
            {
                var mainPart = doc.MainDocumentPart;
                var body=mainPart.Document.Body;

                var comments = mainPart.WordprocessingCommentsPart.Comments;

                var tables = body.Descendants<Table>().ToList();
                var ens = new List<CommentEntity>();

                for (int i = 0; i < tables.Count(); i++)
                {
                    var rows = tables[i].Descendants<TableRow>().ToList();
                    for (int j = 0; j < rows.Count(); j++)
                    {
                        var cells = rows[j].Descendants<TableCell>().ToList();
                        for (int k = 0; k < cells.Count; k++)
                        {
                            var crs = cells[k].Descendants<CommentRangeStart>().FirstOrDefault();
                            if (crs != null && ens.All(a => a.CommentId != crs.Id))
                            {
                                ens.Add(new CommentEntity { TableIdx = i, RowIdx = j, CellIdx = k, CommentId = crs.Id });
                            }
                        }
                    }
                }
               
                foreach (var en in ens)
                {
                    var dd = comments.ChildElements.Select(s => s as Comment).ToList();
                    en.CommentText = dd.FirstOrDefault(f => f.Id == en.CommentId).InnerText;
                    t.AddRow(en.TableIdx,en.RowIdx,en.CellIdx,en.CommentId,en.CommentText);
                }
                
                //mainPart.Document.Save();

            }
            t.Write();
            Console.WriteLine("done");
        }
    }
        internal class CommentEntity
    {
        public int Id { get; set; }
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public int CellIdx { get; set; }
        public string CommentId { get; set; }
        public string CommentText { get; set; }
    }
    
}
