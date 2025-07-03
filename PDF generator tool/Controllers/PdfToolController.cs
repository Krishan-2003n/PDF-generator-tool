//using Microsoft.AspNetCore.Mvc;
//using PDF_generator_tool.Models;
//using PDF_generator_tool.Services;
//using System.IO.Compression;

//namespace PDF_generator_tool.Controllers
//{
//    [ApiController]
//    [Route("[controller]")]
//    public class PdfToolController : ControllerBase
//    {
//        private readonly IExportService _exportService;

//        public PdfToolController(IExportService exportService)
//        {
//            _exportService = exportService;
//        }

//        [HttpPost("GenerateAll")]
//        public async Task<IActionResult> GenerateAll([FromBody] ExportRequest request)
//        {
//            var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
//            var primaryKeys = await _exportService.GetPrimaryKeys(request.ConnectionString, request.TableName);

//            var pdfBytes = PdfGenerator.GeneratePdf(data, request.SelectedFields, primaryKeys);
//            var wordBytes = WordGenerator.GenerateWord(data, request.SelectedFields, request.TableName, primaryKeys);
//            var excelBytes = ExcelGenerator.GenerateExcel(data, request.SelectedFields);

//            using var zipStream = new MemoryStream();
//            using (var archive = new ZipArchive(zipStream, ZipArchiveMode.Create, true))
//            {
//                var pdfEntry = archive.CreateEntry($"{request.TableName}.pdf");
//                using (var entryStream = pdfEntry.Open())
//                    entryStream.Write(pdfBytes, 0, pdfBytes.Length);

//                var wordEntry = archive.CreateEntry($"{request.TableName}.docx");
//                using (var entryStream = wordEntry.Open())
//                    entryStream.Write(wordBytes, 0, wordBytes.Length);

//                var excelEntry = archive.CreateEntry($"{request.TableName}.xlsx");
//                using (var entryStream = excelEntry.Open())
//                    entryStream.Write(excelBytes, 0, excelBytes.Length);
//            }

//            zipStream.Position = 0;
//            return File(zipStream.ToArray(), "application/zip", $"{request.TableName}_exports.zip");
//        }
//    }
//}


//using Microsoft.AspNetCore.Mvc;
//using PDF_generator_tool.Models;
//using PDF_generator_tool.Services;

//using Microsoft.AspNetCore.Mvc;
//using PDF_generator_tool.Models;
//using PDF_generator_tool.Services;

//namespace PDF_generator_tool.Controllers
//{
//    public class PdfToolController : Controller
//    {
//        private readonly IExportService _exportService;

//        public PdfToolController(IExportService exportService)
//        {
//            _exportService = exportService;
//        }

//        // 👇 Loads the main UI
//        [HttpGet]
//        public IActionResult Index()
//        {
//            return View();
//        }

//        // ✅ Get Tables
//        [HttpPost]
//        public async Task<IActionResult> GetTables([FromBody] ExportRequest request)
//        {
//            if (string.IsNullOrWhiteSpace(request.ConnectionString))
//                return BadRequest("Connection string is required.");

//            var tables = await _exportService.GetTablesAsync(request.ConnectionString);
//            return Json(tables);
//        }

//        // ✅ Get Columns
//        [HttpPost]
//        public async Task<IActionResult> GetColumns([FromBody] ExportRequest request)
//        {
//            if (string.IsNullOrWhiteSpace(request.ConnectionString) || string.IsNullOrWhiteSpace(request.TableName))
//                return BadRequest("Missing connection string or table name.");

//            var columns = await _exportService.GetColumnsAsync(request.ConnectionString, request.TableName);
//            return Json(columns);
//        }

//        // ✅ Generate PDF
//        [HttpPost]
//        public async Task<IActionResult> GeneratePdf([FromBody] ExportRequest request)
//        {
//            var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
//            var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
//            var pdfBytes = PdfGenerator.GeneratePdf(data, request.SelectedFields, primaryKeys);

//            return File(pdfBytes, "application/pdf", $"{request.TableName}.pdf");
//        }

//        // ✅ Generate Excel
//        [HttpPost]
//        public async Task<IActionResult> GenerateExcel([FromBody] ExportRequest request)
//        {
//            var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
//            var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
//            var bytes = ExcelGenerator.GenerateExcel(data, request.SelectedFields, primaryKeys);
//            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{request.TableName}.xlsx");
//        }

//        // ✅ Generate Word
//        [HttpPost]
//        public async Task<IActionResult> GenerateWord([FromBody] ExportRequest request)
//        {
//            var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
//            var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
//            var bytes = WordGenerator.GenerateWord(data, request.SelectedFields, request.TableName, primaryKeys);
//            return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{request.TableName}.docx");
//        }
//    }
//}
using Microsoft.AspNetCore.Mvc;
using PDF_generator_tool.Models;
using PDF_generator_tool.Services;
using System.IO.Compression;
namespace PDF_generator_tool.Controllers
{
    public class PdfToolController : Controller
    {
        private readonly IExportService _exportService;

        public PdfToolController(IExportService exportService)
        {
            _exportService = exportService;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> GetTables([FromBody] ExportRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ConnectionString))
                return BadRequest("Connection string is required.");

            var tables = await _exportService.GetTablesAsync(request.ConnectionString);
            return Json(tables);
        }

        [HttpPost]
        public async Task<IActionResult> GetColumns([FromBody] ExportRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ConnectionString) || string.IsNullOrWhiteSpace(request.TableName))
                return BadRequest("Missing connection string or table name.");

            var columns = await _exportService.GetColumnsAsync(request.ConnectionString, request.TableName);
            return Json(columns);
        }

        [HttpPost]
        public async Task<IActionResult> GeneratePdf([FromBody] ExportRequest request)
        {
            if (request.TableNames != null && request.TableNames.Any()) // Multi-table case
            {
                var dataMap = new Dictionary<string, List<Dictionary<string, object>>>();
                var primaryKeyMap = new Dictionary<string, List<string>>();
                var selectedMap = request.SelectedFieldsMap;

                foreach (var table in request.TableNames)
                {
                    var fields = selectedMap.ContainsKey(table) ? selectedMap[table] : new List<string>();
                    if (fields.Count == 0) continue;

                    var data = await _exportService.FetchDataAsync(request.ConnectionString, table, fields);
                    var pk = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, table);

                    dataMap[table] = data;
                    primaryKeyMap[table] = pk;
                }

                var bytes = PdfGenerator.GenerateMultiTablePdf(dataMap, selectedMap, primaryKeyMap);
                return File(bytes, "application/pdf", $"MultiTable_Report_{DateTime.Now:yyyyMMddHHmmss}.pdf");
            }
            else // Single-table case
            {
                var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
                var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
                var pdfBytes = PdfGenerator.GeneratePdf(data, request.SelectedFields, primaryKeys);
                return File(pdfBytes, "application/pdf", $"{request.TableName}.pdf");
            }
        }

        //[HttpPost]
        //public async Task<IActionResult> GenerateExcel([FromBody] ExportRequest request)
        //{
        //    var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
        //    var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
        //    var bytes = ExcelGenerator.GenerateExcel(data, request.SelectedFields, primaryKeys);
        //    return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{request.TableName}.xlsx");
        //}

        //[HttpPost]
        //public async Task<IActionResult> GenerateWord([FromBody] ExportRequest request)
        //{
        //    var data = await _exportService.FetchDataAsync(request.ConnectionString, request.TableName, request.SelectedFields);
        //    var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, request.TableName);
        //    var bytes = WordGenerator.GenerateWord(data, request.SelectedFields, request.TableName, primaryKeys);
        //    return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{request.TableName}.docx");
        //}
        [HttpPost]
        public async Task<IActionResult> GenerateExcel([FromBody] ExportRequest request)
        {
            var allData = new Dictionary<string, List<Dictionary<string, object>>>();
            var allPrimaryKeys = new Dictionary<string, List<string>>();

            foreach (var table in request.TableNames)
            {
                var fields = request.SelectedFieldsMap[table];
                var data = await _exportService.FetchDataAsync(request.ConnectionString, table, fields);
                var pk = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, table);

                allData[table] = data;
                allPrimaryKeys[table] = pk;
            }

            var bytes = ExcelGenerator.GenerateExcel(allData, request.SelectedFieldsMap, allPrimaryKeys);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Export_MultiTable_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
        }

        [HttpPost]
        public async Task<IActionResult> GenerateWord([FromBody] ExportRequest request)
        {
            var allData = new Dictionary<string, List<Dictionary<string, object>>>();
            var allPrimaryKeys = new Dictionary<string, List<string>>();

            foreach (var table in request.TableNames)
            {
                var fields = request.SelectedFieldsMap[table];
                var data = await _exportService.FetchDataAsync(request.ConnectionString, table, fields);
                var pk = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, table);

                allData[table] = data;
                allPrimaryKeys[table] = pk;
            }

            var bytes = WordGenerator.GenerateWord(allData, request.SelectedFieldsMap, allPrimaryKeys);
            return File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"Export_MultiTable_{DateTime.Now:yyyyMMddHHmmss}.docx");
        }
        //[HttpPost]
        //public async Task<IActionResult> ExportZip([FromQuery] string format, [FromBody] ExportRequest request)
        //{
        //    if (request.TableNames == null || request.TableNames.Count == 0)
        //        return BadRequest("No tables selected.");

        //    var files = new Dictionary<string, byte[]>();
        //    foreach (var table in request.TableNames)
        //    {
        //        var fields = request.SelectedFieldsMap.ContainsKey(table) ? request.SelectedFieldsMap[table] : new List<string>();
        //        if (fields.Count == 0) continue;

        //        var data = await _exportService.FetchDataAsync(request.ConnectionString, table, fields);
        //        var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, table);
        //        byte[] content = null;
        //        string extension = "", mimeType = "";

        //        switch (format.ToLower())
        //        {
        //            case "pdf":
        //                content = PdfGenerator.GeneratePdf(data, fields, primaryKeys);
        //                extension = ".pdf";
        //                mimeType = "application/pdf";
        //                break;
        //            case "excel":
        //                content = ExcelGenerator.GenerateExcel(new Dictionary<string, List<Dictionary<string, object>>> { [table] = data },
        //                                                       new Dictionary<string, List<string>> { [table] = fields },
        //                                                       new Dictionary<string, List<string>> { [table] = primaryKeys });
        //                extension = ".xlsx";
        //                mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //                break;
        //            case "word":
        //                content = WordGenerator.GenerateWord(new Dictionary<string, List<Dictionary<string, object>>> { [table] = data },
        //                                                     new Dictionary<string, List<string>> { [table] = fields },
        //                                                     new Dictionary<string, List<string>> { [table] = primaryKeys });
        //                extension = ".docx";
        //                mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        //                break;
        //            case "all":
        //                // Recursively call ExportZip for each format (skip here, unless you want multiple ZIPs inside ZIP)
        //                return BadRequest("Cannot export 'all' in a single ZIP. Use buttons individually.");
        //            default:
        //                return BadRequest("Unsupported export format.");
        //        }

        //        files[$"{table}{extension}"] = content;
        //    }

        //    using var ms = new MemoryStream();
        //    using (var archive = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
        //    {
        //        foreach (var file in files)
        //        {
        //            var entry = archive.CreateEntry(file.Key);
        //            using var entryStream = entry.Open();
        //            await entryStream.WriteAsync(file.Value, 0, file.Value.Length);
        //        }
        //    }

        //    return File(ms.ToArray(), "application/zip", $"Export_{format}_ZIP_{DateTime.Now:yyyyMMddHHmmss}.zip");
        //}
        [HttpPost]
        public async Task<IActionResult> ExportZip([FromQuery] string format, [FromQuery] string fileName, [FromBody] ExportRequest request)
        {
            if (request.TableNames == null || request.TableNames.Count == 0)
                return BadRequest("No tables selected.");

            var files = new Dictionary<string, byte[]>();
            foreach (var table in request.TableNames)
            {
                var fields = request.SelectedFieldsMap.ContainsKey(table) ? request.SelectedFieldsMap[table] : new List<string>();
                if (fields.Count == 0) continue;

                var data = await _exportService.FetchDataAsync(request.ConnectionString, table, fields);
                var primaryKeys = await _exportService.GetPrimaryKeysAsync(request.ConnectionString, table);
                byte[] content = null;
                string extension = "", mimeType = "";

                switch (format.ToLower())
                {
                    case "pdf":
                        content = PdfGenerator.GeneratePdf(data, fields, primaryKeys);
                        extension = ".pdf";
                        mimeType = "application/pdf";
                        break;
                    case "excel":
                        content = ExcelGenerator.GenerateExcel(
                            new Dictionary<string, List<Dictionary<string, object>>> { [table] = data },
                            new Dictionary<string, List<string>> { [table] = fields },
                            new Dictionary<string, List<string>> { [table] = primaryKeys }
                        );
                        extension = ".xlsx";
                        mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        break;
                    case "word":
                        content = WordGenerator.GenerateWord(
                            new Dictionary<string, List<Dictionary<string, object>>> { [table] = data },
                            new Dictionary<string, List<string>> { [table] = fields },
                            new Dictionary<string, List<string>> { [table] = primaryKeys }
                        );
                        extension = ".docx";
                        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                        break;
                    case "all":
                        return BadRequest("Cannot export 'all' in a single ZIP. Use buttons individually.");
                    default:
                        return BadRequest("Unsupported export format.");
                }

                files[$"{table}{extension}"] = content;
            }

            using var ms = new MemoryStream();
            using (var archive = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
            {
                foreach (var file in files)
                {
                    var entry = archive.CreateEntry(file.Key);
                    using var entryStream = entry.Open();
                    await entryStream.WriteAsync(file.Value, 0, file.Value.Length);
                }
            }

            var finalZipName = string.IsNullOrWhiteSpace(fileName) ?
                $"Export_{format}_ZIP_{DateTime.Now:yyyyMMddHHmmss}.zip" :
                (fileName.EndsWith(".zip") ? fileName : fileName + ".zip");

            return File(ms.ToArray(), "application/zip", finalZipName);
        }



        [HttpPost]
        public async Task<IActionResult> PreviewData([FromBody] ExportRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ConnectionString) || string.IsNullOrWhiteSpace(request.TableName) || request.SelectedFields == null || !request.SelectedFields.Any())
                return BadRequest("Missing required preview info.");

            var data = await _exportService.FetchPreviewDataAsync(request.ConnectionString, request.TableName, request.SelectedFields, topRows: 10);
            return Json(data);
        }

      

    }
}
