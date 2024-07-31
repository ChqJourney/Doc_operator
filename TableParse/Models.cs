using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace TableParse
{
    
    public class GeneralField
    {
        public int CommentId { get; set; }
        public FieldEdit Edit { get; set; }
    }
    public class FieldEdit
    {
        public string InputType { get; set; }
        public string InputValue { get; set; }
        public string OutputType { get; set; }
        public string AlignType { get; set; }
        public List<string> Hints { get; set; }
        public List<string> RulesBasedOnInput { get; set; }
    }
    public class PhotoBlock
    {

    }

    public class FillTable<T> where T:FillRow
    {
        public int TableIdx { get; set; }
        public string TableType { get; set; }
        public string TableTitle { get; set; }
        public List<T> Rows { get; set; }
    }
    public class FillComplianceTable
    {
        public int TableIdx { get; set; }
        public string TableType { get; set; }
        public string TableTitle { get; set; }
        public List<ClauseFillRow> Rows { get; set; }
    }
    public class FillRow
    {
        [JsonIgnore(Condition =JsonIgnoreCondition.WhenWritingDefault)]
        public bool Duplicatable { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool Deletable { get; set; }
        public int TableIdx { get; set; }
        public int RowIdx { get; set; }
        public List<FillCell> Cells { get; set; }

    }

    public class ClauseFillRow : FillRow
    {
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string ClauseNo { get; set; }
        public int IdxUnderClause { get; set; }
    }
    public class FillCell
    {
        public int RowIdx { get; set; }
        public int ColumnIdx { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string OriginalText { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool IsCentered { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool Bolded { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public int CellWidth { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string HMerge { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string VMerge { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool HasShading { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public List<CellEdit> Edits { get; set; }

    }
    public class CellEdit
    {
        public string InputValue { get; set; }
        // js type
        public string InputType { get; set; }
        // c# type
        public string OutputType { get; set; }
        public string InsertType { get; set; }
        public string AlignType { get; set; }
        public List<string> Hints { get; set; }
        public List<string> RulesBasedOnInput { get; set; }
    }
    public class GeneralTable
    {
        public List<GeneralRow> Rows { get; set; } = new List<GeneralRow>();
    }
    public class GeneralRow
    {
        public List<GeneralCell> Cells { get; set; } = new List<GeneralCell>();
    }
    public class GeneralCell
    {
        public int RowIdx { get; set; }
        public int TableIdx { get; set; }
        public int ColumnIdx { get; set; }
        public string CellTxt { get; set; }
        public bool Bolded { get; set; }
        public bool HasShading { get; set; }
        public int CellWidth { get; set; }
        public bool IsCentered { get; set; }

        public string HMerge { get; set; }
        public string VMerge { get; set; }
    }
    public enum TableType
    {
        Unknown,
        General,
        Compliance,
        Measurement,
        Component,
        Equipment,
        Photograph
    }
    public enum RowType
    {
        SectionHeader,
        ClauseHeader,
        VerdictItem,
        InfoItem,
        Unknown
    }
}
