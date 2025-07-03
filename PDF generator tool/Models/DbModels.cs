namespace DynamicPdfApp.Models
{
    public class TableModel
    {
        public string TableName { get; set; }
    }

    public class ColumnModel
    {
        public string ColumnName { get; set; }
    }

    public class RecordModel
    {
        public Dictionary<string, object> Fields { get; set; }
    }

    public class TableRequest
    {
        public string ConnectionString { get; set; }
        public string TableName { get; set; }
    }

    public class PdfRequest
    {
        public string ConnectionString { get; set; }
        public string TableName { get; set; }
        public List<string> SelectedFields { get; set; }
    }
}
