using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Docor;

namespace ValidateApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 输入模板文件路径
            string templatePath = "template.docx";
            // 输出文件路径
            string outputPath = "output.docx";

            // 占位符及其对应的替换值
            var replacements = new Dictionary<string, string>
            {
                { "{name}", "Value1" },
                { "{age}", "Value2" },
                { "{standard}", "Value3" }
            };
            // 复制模板文件作为输出文件
            File.Copy(templatePath, outputPath, true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                var documentBody = mainPart.Document.Body;
                Operator.ChangeCheckbox(mainPart, "Check2", true);
                Operator.ReplacePlaceholders(mainPart, replacements);
                mainPart.Document.Save();
            }
              

            Console.WriteLine("占位符替换完成！");
        }
    }
        
}
