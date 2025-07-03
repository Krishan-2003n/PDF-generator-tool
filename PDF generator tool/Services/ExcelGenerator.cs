//using ClosedXML.Excel;

//public class ExcelGenerator
//{
//    public static byte[] GenerateExcel(List<Dictionary<string, object>> data, List<string> selectedFields)
//    {
//        using var workbook = new XLWorkbook();
//        var worksheet = workbook.Worksheets.Add("Data");

//        // Add header
//        for (int i = 0; i < selectedFields.Count; i++)
//            worksheet.Cell(1, i + 1).Value = selectedFields[i];

//        // Add rows
//        for (int row = 0; row < data.Count; row++)
//        {
//            for (int col = 0; col < selectedFields.Count; col++)
//            {
//                var field = selectedFields[col];
//                var value = data[row].ContainsKey(field) ? data[row][field]?.ToString() : "N/A";
//                worksheet.Cell(row + 2, col + 1).Value = value;
//            }
//        }

//        using var stream = new MemoryStream();
//        workbook.SaveAs(stream);
//        return stream.ToArray();
//    }
//}
//using ClosedXML.Excel;

//namespace PDF_generator_tool.Services
//{
//    public class ExcelGenerator
//    {
//        public static byte[] GenerateExcel(List<Dictionary<string, object>> data, List<string> selectedFields)
//        {
//            using var workbook = new XLWorkbook();
//            var worksheet = workbook.Worksheets.Add("Data");

//            // Headers
//            for (int i = 0; i < selectedFields.Count; i++)
//            {
//                worksheet.Cell(1, i + 1).Value = selectedFields[i];
//                worksheet.Cell(1, i + 1).Style.Font.Bold = true;
//                worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
//                worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
//            }

//            // Data Rows
//            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
//            {
//                var row = data[rowIdx];
//                for (int colIdx = 0; colIdx < selectedFields.Count; colIdx++)
//                {
//                    var field = selectedFields[colIdx];
//                    var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
//                    worksheet.Cell(rowIdx + 2, colIdx + 1).Value = value;
//                }
//            }

//            worksheet.Columns().AdjustToContents();

//            using var ms = new MemoryStream();
//            workbook.SaveAs(ms);
//            return ms.ToArray();
//        }
//    }
//}





//using ClosedXML.Excel;

//namespace PDF_generator_tool.Services
//{
//    public class ExcelGenerator
//    {
//        public static byte[] GenerateExcel(List<Dictionary<string, object>> data, List<string> selectedFields, List<string> primaryKeys)
//        {
//            using var workbook = new XLWorkbook();
//            var worksheet = workbook.Worksheets.Add("Data");

//            // Headers
//            for (int i = 0; i < selectedFields.Count; i++)
//            {
//                var field = selectedFields[i];
//                bool isPK = primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase);
//                worksheet.Cell(1, i + 1).Value = isPK ? $"{field} 🔑" : field;
//                worksheet.Cell(1, i + 1).Style.Font.Bold = true;
//                worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = isPK ? XLColor.LightYellow : XLColor.LightGray;
//                worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
//            }

//            // Data Rows
//            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
//            {
//                var row = data[rowIdx];
//                for (int colIdx = 0; colIdx < selectedFields.Count; colIdx++)
//                {
//                    var field = selectedFields[colIdx];
//                    var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";

//                    var cell = worksheet.Cell(rowIdx + 2, colIdx + 1);
//                    cell.Value = value;

//                    if (primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase))
//                    {
//                        cell.Style.Fill.BackgroundColor = XLColor.LightYellow;
//                    }
//                }
//            }

//            worksheet.Columns().AdjustToContents();

//            using var ms = new MemoryStream();
//            workbook.SaveAs(ms);
//            return ms.ToArray();
//        }
//    }
//}




// ✅ Updated ExcelGenerator.cs for multi-table support
using ClosedXML.Excel;

namespace PDF_generator_tool.Services
{
    public class ExcelGenerator
    {
        public static byte[] GenerateExcel(Dictionary<string, List<Dictionary<string, object>>> dataMap, Dictionary<string, List<string>> selectedFieldsMap, Dictionary<string, List<string>> primaryKeysMap)
        {
            using var workbook = new XLWorkbook();

            foreach (var tableEntry in dataMap)
            {
                string tableName = tableEntry.Key;
                var data = tableEntry.Value;
                var selectedFields = selectedFieldsMap[tableName];
                var primaryKeys = primaryKeysMap.ContainsKey(tableName) ? primaryKeysMap[tableName] : new List<string>();

                var worksheet = workbook.Worksheets.Add(tableName);

                // Headers
                for (int i = 0; i < selectedFields.Count; i++)
                {
                    var field = selectedFields[i];
                    bool isPK = primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase);
                    worksheet.Cell(1, i + 1).Value = isPK ? $"{field} 🔑" : field;
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = isPK ? XLColor.LightYellow : XLColor.LightGray;
                    worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                // Data Rows
                for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
                {
                    var row = data[rowIdx];
                    for (int colIdx = 0; colIdx < selectedFields.Count; colIdx++)
                    {
                        var field = selectedFields[colIdx];
                        var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";

                        var cell = worksheet.Cell(rowIdx + 2, colIdx + 1);
                        cell.Value = value;

                        if (primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase))
                        {
                            cell.Style.Fill.BackgroundColor = XLColor.LightYellow;
                        }
                    }
                }

                worksheet.Columns().AdjustToContents();
            }

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }
    }
}
