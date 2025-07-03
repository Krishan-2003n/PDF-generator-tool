//namespace PDF_generator_tool.Models
//{
//    public class ExportRequest
//    {
//        public string ConnectionString { get; set; }
//        public string TableName { get; set; }
//        public List<string> SelectedFields { get; set; }
//    }
//}


//namespace PDF_generator_tool.Models
//{
//    public class ExportRequest
//    {
//        public string ConnectionString { get; set; }

//        // For single-table export (legacy)
//        public string TableName { get; set; }
//        public List<string> SelectedFields { get; set; }

//        // New: For multi-table export
//        public List<string> TableNames { get; set; }
//        public Dictionary<string, List<string>> SelectedFieldsMap { get; set; }
//    }
//}
namespace PDF_generator_tool.Models
{
    public class ExportRequest
    {
        public string ConnectionString { get; set; }

        // For single-table export (legacy)
        public string TableName { get; set; }
        public List<string> SelectedFields { get; set; }

        // New: For multi-table export
        public List<string> TableNames { get; set; }
        public Dictionary<string, List<string>> SelectedFieldsMap { get; set; }

        // New for ZIP Export
        public bool IncludePdf { get; set; }
        public bool IncludeExcel { get; set; }
        public bool IncludeWord { get; set; }
        public string ZipFileName { get; set; }
    }
}
