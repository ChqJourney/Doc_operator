using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocParser
{
    public class Fields
    {
        public static Task<List<InputFiled>> CollectFields(MainDocumentPart main)
        { 
            var comments = main.WordprocessingCommentsPart?.Comments.ChildElements.Select(s=>s as Comment).ToList();
            if (comments == null || comments.Count == 0)
            {
                return Task.FromResult(new List<InputFiled>());
            }
            var result=new List<InputFiled>();
            foreach (var comment in comments)
            {
                var temp = new InputFiled(comment.InnerText,comment.Id);
                result.Add(temp);
            }
            return Task.FromResult(result);
        }
        public static Task FillFields(MainDocumentPart main,List<InputFiled> fields)
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
            FieldType =(FieldType) Enum.Parse(typeof(FieldType), arr[0]);
            FieldName=arr[1];
        }
        public string FieldName { get; set; }
        public FieldType  FieldType { get; set; }
        public string CommentId { get; set; }
        public string CommentText { get; set; }
        public string? FieldValue { get; set; }
    }
    public enum FieldType
    {
        Text,
        TextArea,
        TextField,
        CheckBox,
        Unknown
    }
}
