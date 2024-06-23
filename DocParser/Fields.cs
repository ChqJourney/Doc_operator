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
                var inputVal=fields?.FirstOrDefault(f=>f.CommentId==cr.Id)?.FieldValue;
                if (!string.IsNullOrEmpty(inputVal))
                {
                    cr.Append(new Run(inputVal));
                }
            }
            return Task.FromResult(crs);
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
            InputFieldType =(InputFieldType) Enum.Parse(typeof(InputFieldType), arr[0]);
            FieldName=arr[1];
        }
        public string FieldName { get; set; }
        public InputFieldType  InputFieldType { get; set; }
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
        TextboxField,
        CheckBox,
    }
}
