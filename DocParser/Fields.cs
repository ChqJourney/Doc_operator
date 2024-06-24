using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace DocParser
{
    public class Fields
    {
        public Task<List<InputFiled>> CollectFields(MainDocumentPart main,Action<string> action=null)
        { 
            var comments = main.WordprocessingCommentsPart?.Comments.ChildElements.Select(s=>s as Comment).ToList();
            if (comments == null || comments?.Count == 0)
            {
                return Task.FromResult(new List<InputFiled>());
            }
            var result=new List<InputFiled>();
            foreach (var comment in comments)
            {
                var temp = new InputFiled(comment.InnerText,comment.Id);
                var json= JsonSerializer.Serialize(temp);
                if(action != null)
                {
                    action.Invoke(json);
                }

                result.Add(temp);
            }
            return Task.FromResult(result);
        }
        public Task FillFields(MainDocumentPart main,List<InputFiled> fields)
        {
            var crs = main.Document.Body?.Descendants<CommentRangeStart>().ToList();
            foreach(var cr in crs)
            {
                var field = fields?.FirstOrDefault(f => f.CommentId == cr.Id);
                if (field == null) continue;
                
                if (string.IsNullOrEmpty(field.FieldValue?.Trim())) continue;
                switch (field.OutPutType)
                {
                    case OutPutFieldType.Text:
                        cr.Append(new Run(field.FieldValue.Trim()));
                        break;
                    case OutPutFieldType.TextField:
                        FillTextField(cr,field.FieldValue);
                        break;
                    case OutPutFieldType.CheckBox:
                        ChangeCheckbox(cr, Boolean.Parse(field.FieldValue));
                        break;
                    default:
                        break;

                }
            }
            return Task.FromResult(crs);
        }
        private void ChangeCheckbox(CommentRangeStart crs, bool val)
        {
            if(crs==null) return;

             Paragraph parentParagraph = crs.Ancestors<Paragraph>().FirstOrDefault();
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
        private void FillTextField(CommentRangeStart crs, string text)
        {
            var cell = crs.Ancestors<TableCell>().FirstOrDefault();
            var x = cell?.Descendants<DefaultTextBoxFormFieldString>().FirstOrDefault();
            x.Val = text;
        }
    }
    public class InputFiled
    {
        public InputFiled(string commentText,string commentId)
        {
            this.CommentId=commentId;
            this.CommentText=commentText;
            var arr = commentText.Split(':');
            // sample: Text:ReportNo:TextField
            InputType =(InputFieldType) Enum.Parse(typeof(InputFieldType), arr[0]);

            FieldName=arr[1];
            OutPutType = (OutPutFieldType)Enum.Parse(typeof(OutPutFieldType), arr[2]);
        }
        public string FieldName { get; set; }
        public InputFieldType  InputType { get; set; }
        public OutPutFieldType OutPutType { get; set; }
        public string CommentId { get; set; }
        public string CommentText { get; set; }
        public string? FieldValue { get; set; }
    }
    /// <summary>
    /// user input value
    /// </summary>
    public enum InputFieldType
    {
        Text,
        Boolean,
        Picture,
        Unknown
    }
    public enum OutPutFieldType
    {
        Text,
        TextField,
        CheckBox,
    }
}
