using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidateApp
{
    public class Location
    {
        public int Id { get; set; }
        public string LocationType { get; set; }    

    }
    public class CommentLocation:Location
    {
        public string CommentId { get; set; }
        public string CommentText { get; set; }
    }
    public class BookmarkLocation : Location
    {
        public string BookmarkName { get; set; }

    }
    public class TableLocation : Location
    {
        public int TableIdx { get; set;}
        public int RowIdx { get; set; }
        public int ColumnIdx { get; set; }
    }
    public class Operation
    {
        public int Id { get; set; }
        public string OperationType { get; set; }
        public Location OperationLocation { get; set; }
        public List<OperationSource> OperationSources { get; set; }
        
    }
    public class OperationSource
    {
        public int Id { get; set; }
        public string ResultType { get; set; }
        
    }
    public class TextInputSource : OperationSource
    {
        public string Text { get; set; }
    }
    public class EnumInputSource : OperationSource 
    {

        
    }

}
