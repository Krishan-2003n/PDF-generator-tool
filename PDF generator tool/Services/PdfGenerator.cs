
////TODO:First code
////using iText.Kernel.Colors;
////using iText.Kernel.Geom;
////using iText.Kernel.Pdf;
////using iText.Layout;
////using iText.Layout.Borders;
////using iText.Layout.Element;
////using iText.Layout.Properties;

////namespace PDF_generator_tool.Services
////{

////    public class PdfGenerator
////    {
////        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields)
////        {
////            using var stream = new MemoryStream();
////            var writer = new PdfWriter(stream);
////            var pdf = new PdfDocument(writer);
////            var doc = new Document(pdf, PageSize.A4.Rotate(), false);

////            doc.SetMargins(20, 20, 20, 20);

////            // 1. Header
////            var title = new Paragraph("📊 PDF Report")
////                .SetFontSize(20)

////                .SetTextAlignment(TextAlignment.CENTER)
////                .SetMarginBottom(10);
////            doc.Add(title);

////            // 2. Metadata Info
////            var meta = $"Generated On: {DateTime.Now:dd-MM-yyyy HH:mm}\nTotal Records: {data.Count}";
////            doc.Add(new Paragraph(meta).SetFontSize(10).SetMarginBottom(15));

////            if (data.Count > 0)
////            {
////                // 3. Table
////                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();

////                // Header styling
////                foreach (var col in selectedFields)
////                {
////                    table.AddHeaderCell(new Cell()

////                        .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
////                        .SetBorder(Border.NO_BORDER));
////                }

////                // Data Rows
////                foreach (var row in data)
////                {
////                    foreach (var col in selectedFields)
////                    {
////                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
////                        table.AddCell(new Cell().Add(new Paragraph(value)).SetBorder(Border.NO_BORDER));
////                    }
////                }

////                doc.Add(table);
////                doc.Add(new Paragraph("\n")); // spacing

////                // 4. Optional: Add a Bar Chart
////                var firstField = selectedFields.First(); // Use first column (e.g., "Gender", "Department") for chart
////                var groupedData = data.GroupBy(r => r[firstField]?.ToString() ?? "N/A")
////                                      .ToDictionary(g => g.Key, g => g.Count());

////                doc.Add(new Paragraph($"📈 Distribution by: {firstField}")
////                    .SetFontSize(14)

////                    .SetMarginTop(10)
////                    .SetMarginBottom(10));

////                float chartWidth = 400;
////                float maxCount = groupedData.Values.Max();
////                foreach (var kvp in groupedData)
////                {
////                    float barLength = (kvp.Value / maxCount) * chartWidth;
////                    var chartRow = new Paragraph()
////                        .Add($"{kvp.Key}: ")
////                        .Add(new Text(new string('█', (int)(barLength / 5))).SetFontColor(ColorConstants.BLUE))
////                        .Add($" ({kvp.Value})")
////                        .SetFontSize(10);
////                    doc.Add(chartRow);
////                }
////            }
////            else
////            {
////                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED));
////            }

////            doc.Close();
////            return stream.ToArray();
////        }
////    }
////}

////TODO:second code
////using iText.IO.Font.Constants;
////using iText.Kernel.Colors;
////using iText.Kernel.Font;
////using iText.Kernel.Geom;
////using iText.Kernel.Pdf;
////using iText.Layout;
////using iText.Layout.Borders;
////using iText.Layout.Element;
////using iText.Layout.Properties;

////namespace PDF_generator_tool.Services
////{
////    public class PdfGenerator
////    {
////        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields)
////        {
////            using var stream = new MemoryStream();
////            var writer = new PdfWriter(stream);
////            var pdf = new PdfDocument(writer);
////            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
////            doc.SetMargins(20, 20, 20, 20);

////            PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
////            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

////            // 1. Title
////            var title = new Paragraph("📊 Project Report")
////                .SetFont(boldFont)
////                .SetFontSize(20)
////                .SetTextAlignment(TextAlignment.CENTER)
////                .SetMarginBottom(10);
////            doc.Add(title);

////            // 2. Meta Info
////            var meta = $"Generated On: {DateTime.Now:dd-MM-yyyy HH:mm}\nTotal Records: {data.Count}";
////            doc.Add(new Paragraph(meta).SetFont(regularFont).SetFontSize(10).SetMarginBottom(15));

////            if (data.Count > 0)
////            {
////                // 3. Styled Data Table
////                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();

////                foreach (var col in selectedFields)
////                {
////                    table.AddHeaderCell(new Cell()
////                        .Add(new Paragraph(col).SetFont(boldFont).SetFontSize(10))
////                        .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
////                        .SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f))
////                        .SetPadding(5));
////                }

////                foreach (var row in data)
////                {
////                    foreach (var col in selectedFields)
////                    {
////                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
////                        table.AddCell(new Cell()
////                            .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
////                                .SetFont(regularFont)
////                                .SetFontSize(9))
////                            .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f))
////                            .SetPadding(4));
////                    }
////                }

////                doc.Add(table);
////                doc.Add(new Paragraph("\n"));

////                // 4. Summary Table instead of ID-based bar chart
////                var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
////                var groupedData = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
////                                      .ToDictionary(g => g.Key, g => g.Count());

////                doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
////                    .SetFont(boldFont)
////                    .SetFontSize(14)
////                    .SetMarginTop(10)
////                    .SetMarginBottom(10));

////                var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
////                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));
////                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont)).SetBackgroundColor(ColorConstants.LIGHT_GRAY));

////                foreach (var kvp in groupedData)
////                {
////                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)));
////                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)));
////                }

////                doc.Add(summaryTable);
////            }
////            else
////            {
////                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
////            }

////            doc.Close();
////            return stream.ToArray();
////        }
////    }
////}
//using iText.IO.Font.Constants;
//using iText.Kernel.Colors;
//using iText.Kernel.Font;
//using iText.Kernel.Geom;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Borders;
//using iText.Layout.Element;
//using iText.Layout.Properties;

//namespace PDF_generator_tool.Services
//{
//    public class PdfGenerator
//    {
//        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields)
//        {
//            using var stream = new MemoryStream();
//            var writer = new PdfWriter(stream);
//            var pdf = new PdfDocument(writer);
//            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
//            doc.SetMargins(20, 20, 20, 20);

//            PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
//            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

//            // 1. Title
//            var title = new Paragraph("📊 Project Report")
//                .SetFont(boldFont)
//                .SetFontSize(20)
//                .SetTextAlignment(TextAlignment.CENTER)
//                .SetBackgroundColor(new DeviceRgb(163, 123, 115))
//                .SetFontColor(ColorConstants.WHITE)
//                .SetPadding(8)
//                .SetMarginBottom(15);
//            doc.Add(title);

//            // 2. Meta Info
//            var metaBlock = new Paragraph()
//                .Add(new Text("🗂 Generated On: ").SetFont(boldFont))
//                .Add(new Text($"{DateTime.Now:dd-MM-yyyy HH:mm}\n").SetFont(regularFont))
//                .Add(new Text("📌 Total Records: ").SetFont(boldFont))
//                .Add(new Text(data.Count.ToString()).SetFont(regularFont))
//                .SetFontSize(10)
//                .SetMarginBottom(15)
//                .SetBackgroundColor(new DeviceRgb(238, 238, 238))
//                .SetPadding(8);
//            doc.Add(metaBlock);

//            if (data.Count > 0)
//            {
//                // 3. Styled Data Table
//                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();
//                table.SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f));

//                // Header Row
//                foreach (var col in selectedFields)
//                {
//                    table.AddHeaderCell(new Cell()
//                        .Add(new Paragraph(col).SetFont(boldFont).SetFontSize(10).SetFontColor(ColorConstants.WHITE))
//                        .SetBackgroundColor(new DeviceRgb(125, 79, 80)) // dark red
//                        .SetPadding(6)
//                        .SetBorder(Border.NO_BORDER));
//                }

//                // Data Rows with zebra styling
//                int rowIndex = 0;
//                foreach (var row in data)
//                {
//                    var bgColor = rowIndex % 2 == 0 ? new DeviceRgb(255, 250, 245) : ColorConstants.WHITE;
//                    foreach (var col in selectedFields)
//                    {
//                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
//                        table.AddCell(new Cell()
//                            .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
//                                .SetFont(regularFont)
//                                .SetFontSize(9))
//                            .SetPadding(5)
//                            .SetBackgroundColor(bgColor)
//                            .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f)));
//                    }
//                    rowIndex++;
//                }

//                doc.Add(table);
//                doc.Add(new Paragraph("\n"));

//                // 4. Summary Section
//                var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
//                var groupedData = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
//                                      .ToDictionary(g => g.Key, g => g.Count());

//                doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
//                    .SetFont(boldFont)
//                    .SetFontSize(14)
//                    .SetMarginTop(10)
//                    .SetMarginBottom(10)
//                    .SetTextAlignment(TextAlignment.LEFT));

//                var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));

//                foreach (var kvp in groupedData)
//                {
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)).SetPadding(4));
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)).SetPadding(4));
//                }

//                doc.Add(summaryTable);
//            }
//            else
//            {
//                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
//            }

//            doc.Close();
//            return stream.ToArray();
//        }
//    }
//}
//using iText.IO.Font.Constants;
//using iText.Kernel.Colors;
//using iText.Kernel.Font;
//using iText.Kernel.Geom;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Borders;
//using iText.Layout.Element;
//using iText.Layout.Properties;

//namespace PDF_generator_tool.Services
//{
//    public class PdfGenerator
//    {
//        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields, List<string> primaryKeys)
//        {
//            using var stream = new MemoryStream();
//            var writer = new PdfWriter(stream);
//            var pdf = new PdfDocument(writer);
//            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
//            doc.SetMargins(20, 20, 20, 20);

//            PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
//            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

//            // Title
//            var title = new Paragraph("📊 Project Report")
//                .SetFont(boldFont)
//                .SetFontSize(20)
//                .SetTextAlignment(TextAlignment.CENTER)
//                .SetBackgroundColor(new DeviceRgb(163, 123, 115))
//                .SetFontColor(ColorConstants.WHITE)
//                .SetPadding(8)
//                .SetMarginBottom(15);
//            doc.Add(title);

//            // Meta Info
//            var metaBlock = new Paragraph()
//                .Add(new Text("🗂 Generated On: ").SetFont(boldFont))
//                .Add(new Text($"{DateTime.Now:dd-MM-yyyy HH:mm}\n").SetFont(regularFont))
//                .Add(new Text("📌 Total Records: ").SetFont(boldFont))
//                .Add(new Text(data.Count.ToString()).SetFont(regularFont))
//                .SetFontSize(10)
//                .SetMarginBottom(15)
//                .SetBackgroundColor(new DeviceRgb(238, 238, 238))
//                .SetPadding(8);
//            doc.Add(metaBlock);

//            if (data.Count > 0)
//            {
//                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();
//                table.SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f));

//                // Headers
//                foreach (var col in selectedFields)
//                {
//                    bool isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);
//                    var headerText = isPK ? $"{col} 🔑" : col;

//                    table.AddHeaderCell(new Cell()
//                        .Add(new Paragraph(headerText).SetFont(boldFont).SetFontSize(10).SetFontColor(ColorConstants.WHITE))
//                        .SetBackgroundColor(new DeviceRgb(125, 79, 80))
//                        .SetPadding(6)
//                        .SetBorder(Border.NO_BORDER));
//                }

//                // Rows
//                int rowIndex = 0;
//                foreach (var row in data)
//                {
//                    var bgColor = rowIndex % 2 == 0 ? new DeviceRgb(255, 250, 245) : ColorConstants.WHITE;
//                    foreach (var col in selectedFields)
//                    {
//                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
//                        table.AddCell(new Cell()
//                            .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
//                                .SetFont(regularFont)
//                                .SetFontSize(9))
//                            .SetPadding(5)
//                            .SetBackgroundColor(bgColor)
//                            .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f)));
//                    }
//                    rowIndex++;
//                }

//                doc.Add(table);
//                doc.Add(new Paragraph("\n"));

//                // Summary
//                var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
//                var groupedData = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
//                                      .ToDictionary(g => g.Key, g => g.Count());

//                doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
//                    .SetFont(boldFont)
//                    .SetFontSize(14)
//                    .SetMarginTop(10)
//                    .SetMarginBottom(10)
//                    .SetTextAlignment(TextAlignment.LEFT));

//                var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));

//                foreach (var kvp in groupedData)
//                {
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)).SetPadding(4));
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)).SetPadding(4));
//                }

//                doc.Add(summaryTable);
//            }
//            else
//            {
//                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
//            }

//            doc.Close();
//            return stream.ToArray();
//        }
//    }
//}



//using iText.IO.Font.Constants;
//using iText.Kernel.Colors;
//using iText.Kernel.Font;
//using iText.Kernel.Geom;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Borders;
//using iText.Layout.Element;
//using iText.Layout.Properties;

//namespace PDF_generator_tool.Services
//{
//    public class PdfGenerator
//    {
//        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields, List<string> primaryKeys)
//        {
//            using var stream = new MemoryStream();
//            var writer = new PdfWriter(stream);
//            var pdf = new PdfDocument(writer);
//            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
//            doc.SetMargins(20, 20, 20, 20);

//            PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
//            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

//            // Title
//            var title = new Paragraph("📊 Project Report")
//                .SetFont(boldFont)
//                .SetFontSize(20)
//                .SetTextAlignment(TextAlignment.CENTER)
//                .SetBackgroundColor(new DeviceRgb(163, 123, 115))
//                .SetFontColor(ColorConstants.WHITE)
//                .SetPadding(8)
//                .SetMarginBottom(15);
//            doc.Add(title);

//            // Meta Info
//            var metaBlock = new Paragraph()
//                .Add(new Text("🗂 Generated On: ").SetFont(boldFont))
//                .Add(new Text($"{DateTime.Now:dd-MM-yyyy HH:mm}\n").SetFont(regularFont))
//                .Add(new Text("📌 Total Records: ").SetFont(boldFont))
//                .Add(new Text(data.Count.ToString()).SetFont(regularFont))
//                .SetFontSize(10)
//                .SetMarginBottom(15)
//                .SetBackgroundColor(new DeviceRgb(238, 238, 238))
//                .SetPadding(8);
//            doc.Add(metaBlock);

//            if (data.Count > 0)
//            {
//                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();
//                table.SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f));

//                // Headers
//                foreach (var col in selectedFields)
//                {
//                    bool isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);
//                    var headerText = isPK ? $"{col} 🔑" : col;

//                    table.AddHeaderCell(new Cell()
//                        .Add(new Paragraph(headerText).SetFont(boldFont).SetFontSize(10).SetFontColor(ColorConstants.WHITE))
//                        .SetBackgroundColor(new DeviceRgb(125, 79, 80))
//                        .SetPadding(6)
//                        .SetBorder(Border.NO_BORDER));
//                }

//                // Rows
//                int rowIndex = 0;
//                foreach (var row in data)
//                {
//                    var bgColor = rowIndex % 2 == 0 ? new DeviceRgb(255, 250, 245) : ColorConstants.WHITE;

//                    foreach (var col in selectedFields)
//                    {
//                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
//                        bool isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);

//                        var cell = new Cell()
//                            .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
//                                .SetFont(regularFont)
//                                .SetFontSize(9))
//                            .SetPadding(5)
//                            .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f));

//                        cell.SetBackgroundColor(isPK ? WebColors.GetRGBColor("#fff9c4") : bgColor); // Light yellow if PK
//                        table.AddCell(cell);
//                    }
//                    rowIndex++;
//                }

//                doc.Add(table);
//                doc.Add(new Paragraph("\n"));

//                // Summary
//                var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
//                var groupedData = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
//                                      .ToDictionary(g => g.Key, g => g.Count());

//                doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
//                    .SetFont(boldFont)
//                    .SetFontSize(14)
//                    .SetMarginTop(10)
//                    .SetMarginBottom(10)
//                    .SetTextAlignment(TextAlignment.LEFT));

//                var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));
//                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));

//                foreach (var kvp in groupedData)
//                {
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)).SetPadding(4));
//                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)).SetPadding(4));
//                }

//                doc.Add(summaryTable);
//            }
//            else
//            {
//                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
//            }

//            doc.Close();
//            return stream.ToArray();
//        }
//    }
//}


using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace PDF_generator_tool.Services
{
    public class PdfGenerator
    {
        public static byte[] GeneratePdf(List<Dictionary<string, object>> data, List<string> selectedFields, List<string> primaryKeys)
        {
            using var stream = new MemoryStream();
            var writer = new PdfWriter(stream);
            var pdf = new PdfDocument(writer);
            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
            doc.SetMargins(20, 20, 20, 20);

            PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            // Title
            var title = new Paragraph("📊 Project Report")
                .SetFont(boldFont)
                .SetFontSize(20)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBackgroundColor(new DeviceRgb(163, 123, 115))
                .SetFontColor(ColorConstants.WHITE)
                .SetPadding(8)
                .SetMarginBottom(15);
            doc.Add(title);

            // Meta Info
            var metaBlock = new Paragraph()
                .Add(new Text("🗂 Generated On: ").SetFont(boldFont))
                .Add(new Text($"{DateTime.Now:dd-MM-yyyy HH:mm}\n").SetFont(regularFont))
                .Add(new Text("📌 Total Records: ").SetFont(boldFont))
                .Add(new Text(data.Count.ToString()).SetFont(regularFont))
                .SetFontSize(10)
                .SetMarginBottom(15)
                .SetBackgroundColor(new DeviceRgb(238, 238, 238))
                .SetPadding(8);
            doc.Add(metaBlock);

            if (data.Count > 0)
            {
                var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();
                table.SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f));

                // Headers
                foreach (var col in selectedFields)
                {
                    bool isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);
                    var headerText = isPK ? $"{col} 🔑" : col;

                    table.AddHeaderCell(new Cell()
                        .Add(new Paragraph(headerText).SetFont(boldFont).SetFontSize(10).SetFontColor(ColorConstants.WHITE))
                        .SetBackgroundColor(new DeviceRgb(125, 79, 80))
                        .SetPadding(6)
                        .SetBorder(Border.NO_BORDER));
                }

                // Rows
                int rowIndex = 0;
                foreach (var row in data)
                {
                    var bgColor = rowIndex % 2 == 0 ? new DeviceRgb(255, 250, 245) : ColorConstants.WHITE;

                    foreach (var col in selectedFields)
                    {
                        string value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
                        bool isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);

                        var cell = new Cell()
                            .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
                                .SetFont(regularFont)
                                .SetFontSize(9))
                            .SetPadding(5)
                            .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f));

                        cell.SetBackgroundColor(isPK ? WebColors.GetRGBColor("#fff9c4") : bgColor);
                        table.AddCell(cell);
                    }
                    rowIndex++;
                }

                doc.Add(table);
                doc.Add(new Paragraph("\n"));

                // Summary
                var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
                var groupedData = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
                                      .ToDictionary(g => g.Key, g => g.Count());

                doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
                    .SetFont(boldFont)
                    .SetFontSize(14)
                    .SetMarginTop(10)
                    .SetMarginBottom(10)
                    .SetTextAlignment(TextAlignment.LEFT));

                var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));
                summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));

                foreach (var kvp in groupedData)
                {
                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)).SetPadding(4));
                    summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)).SetPadding(4));
                }

                doc.Add(summaryTable);
            }
            else
            {
                doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
            }

            doc.Close();
            return stream.ToArray();
        }

        public static byte[] GenerateMultiTablePdf(
            Dictionary<string, List<Dictionary<string, object>>> dataMap,
            Dictionary<string, List<string>> selectedFieldsMap,
            Dictionary<string, List<string>> primaryKeysMap)
        {
            using var stream = new MemoryStream();
            var writer = new PdfWriter(stream);
            var pdf = new PdfDocument(writer);
            var doc = new Document(pdf, PageSize.A4.Rotate(), false);
            doc.SetMargins(20, 20, 20, 20);

            var regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            var title = new Paragraph("📊 Multi-Table Project Report")
                .SetFont(boldFont)
                .SetFontSize(20)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBackgroundColor(new DeviceRgb(163, 123, 115))
                .SetFontColor(ColorConstants.WHITE)
                .SetPadding(8)
                .SetMarginBottom(15);
            doc.Add(title);

            foreach (var tableName in dataMap.Keys)
            {
                var data = dataMap[tableName];
                var selectedFields = selectedFieldsMap[tableName];
                var primaryKeys = primaryKeysMap[tableName];

                doc.Add(new Paragraph($"📌 Table: {tableName}")
                    .SetFont(boldFont).SetFontSize(16).SetMarginBottom(8).SetFontColor(ColorConstants.BLACK));

                doc.Add(new Paragraph()
                    .Add(new Text("🗂 Records: ").SetFont(boldFont))
                    .Add(new Text($"{data.Count}\n").SetFont(regularFont))
                    .SetFontSize(10)
                    .SetMarginBottom(10)
                    .SetBackgroundColor(new DeviceRgb(238, 238, 238))
                    .SetPadding(5));

                if (data.Count > 0)
                {
                    var table = new Table(UnitValue.CreatePercentArray(selectedFields.Count)).UseAllAvailableWidth();
                    table.SetBorder(new SolidBorder(ColorConstants.GRAY, 0.5f));

                    foreach (var col in selectedFields)
                    {
                        var isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);
                        var label = isPK ? $"{col} 🔑" : col;
                        table.AddHeaderCell(new Cell()
                            .Add(new Paragraph(label).SetFont(boldFont).SetFontSize(10).SetFontColor(ColorConstants.WHITE))
                            .SetBackgroundColor(new DeviceRgb(125, 79, 80))
                            .SetPadding(6));
                    }

                    int rowIndex = 0;
                    foreach (var row in data)
                    {
                        var bgColor = rowIndex % 2 == 0 ? new DeviceRgb(255, 250, 245) : ColorConstants.WHITE;
                        foreach (var col in selectedFields)
                        {
                            var value = row.ContainsKey(col) ? row[col]?.ToString() ?? "" : "N/A";
                            var isPK = primaryKeys.Contains(col, StringComparer.OrdinalIgnoreCase);
                            var cell = new Cell()
                                .Add(new Paragraph(value.Length > 100 ? value.Substring(0, 97) + "..." : value)
                                .SetFont(regularFont).SetFontSize(9))
                                .SetPadding(5)
                                .SetBorder(new SolidBorder(ColorConstants.LIGHT_GRAY, 0.25f))
                                .SetBackgroundColor(isPK ? WebColors.GetRGBColor("#fff9c4") : bgColor);

                            table.AddCell(cell);
                        }
                        rowIndex++;
                    }

                    doc.Add(table);
                    doc.Add(new Paragraph("\n"));

                    var summaryField = selectedFields.FirstOrDefault(f => !f.ToLower().Contains("id")) ?? selectedFields.First();
                    var grouped = data.GroupBy(r => r[summaryField]?.ToString() ?? "N/A")
                                      .ToDictionary(g => g.Key, g => g.Count());

                    doc.Add(new Paragraph($"📈 Summary by: {summaryField}")
                        .SetFont(boldFont).SetFontSize(12).SetMarginTop(10).SetMarginBottom(10));

                    var summaryTable = new Table(UnitValue.CreatePercentArray(new float[] { 70, 30 })).UseAllAvailableWidth();
                    summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Value").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));
                    summaryTable.AddHeaderCell(new Cell().Add(new Paragraph("Count").SetFont(boldFont).SetFontColor(ColorConstants.WHITE)).SetBackgroundColor(new DeviceRgb(90, 90, 90)).SetPadding(5));

                    foreach (var kvp in grouped)
                    {
                        summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Key).SetFont(regularFont)).SetPadding(4));
                        summaryTable.AddCell(new Cell().Add(new Paragraph(kvp.Value.ToString()).SetFont(regularFont)).SetPadding(4));
                    }

                    doc.Add(summaryTable);
                }
                else
                {
                    doc.Add(new Paragraph("❌ No data found.").SetFontColor(ColorConstants.RED).SetFont(boldFont));
                }

                doc.Add(new AreaBreak());
            }

            doc.Close();
            return stream.ToArray();
        }
    }
}
