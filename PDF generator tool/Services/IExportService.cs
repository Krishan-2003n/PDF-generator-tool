//using PDF_generator_tool.Models;

//namespace PDF_generator_tool.Services
//{
//    public interface IExportService
//    {
//        Task<List<string>> GetTablesAsync(string connectionString);
//        Task<List<string>> GetColumnsAsync(string connectionString, string tableName);
//        Task<List<Dictionary<string, object>>> FetchDataAsync(string connectionString, string tableName, List<string> fields);
//        //Task<List<string>> GetPrimaryKeysAsync(string connectionString, string tableName); // ✅ Corrected async version
//        Task<List<string>> GetPrimaryKeysAsync(string connectionString, string tableName);

//    }
//}
using PDF_generator_tool.Models;

namespace PDF_generator_tool.Services
{
    public interface IExportService
    {
        Task<List<string>> GetTablesAsync(string connectionString);
        Task<List<string>> GetColumnsAsync(string connectionString, string tableName);
        Task<List<Dictionary<string, object>>> FetchDataAsync(string connectionString, string tableName, List<string> fields);
        Task<List<string>> GetPrimaryKeysAsync(string connectionString, string tableName); // <-- FIXED
        Task<List<Dictionary<string, object>>> FetchPreviewDataAsync(string connectionString, string tableName, List<string> fields, int topRows = 5);

    }
}
