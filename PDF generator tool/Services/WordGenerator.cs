//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;
//using System.Text;

//namespace PDF_generator_tool.Services
//{
//    public class WordGenerator
//    {
//        public static byte[] GenerateWord(List<Dictionary<string, object>> data, List<string> selectedFields, string tableName)
//        {
//            using var ms = new MemoryStream();
//            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
//            {
//                var mainPart = wordDoc.AddMainDocumentPart();
//                mainPart.Document = new Document();
//                var body = new Body();

//                // Title
//                body.AppendChild(new Paragraph(new Run(new Text("📋 Project Report")))
//                {
//                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
//                });

//                // Metadata
//                body.AppendChild(new Paragraph(new Run(new Text($"Table: {tableName}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Generated on: {DateTime.Now:dd-MMM-yyyy HH:mm}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Total Records: {data.Count}"))));
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Create table
//                var table = new Table();

//                // Add table style
//                TableProperties tblProps = new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 6 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 6 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 6 },
//                        new RightBorder { Val = BorderValues.Single, Size = 6 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 6 }
//                    )
//                );
//                table.AppendChild(tblProps);

//                // Header row
//                var headerRow = new TableRow();
//                foreach (var field in selectedFields)
//                {
//                    var headerCell = new TableCell(
//                        new Paragraph(new Run(new Text(field)))
//                        {
//                            ParagraphProperties = new ParagraphProperties(
//                                new Justification { Val = JustificationValues.Left }
//                            )
//                        }
//                    );
//                    headerRow.Append(headerCell);
//                }
//                table.Append(headerRow);

//                // Data rows
//                foreach (var row in data)
//                {
//                    var dataRow = new TableRow();
//                    foreach (var field in selectedFields)
//                    {
//                        var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
//                        var cell = new TableCell(
//                            new Paragraph(new Run(new Text(value)))
//                        );
//                        dataRow.Append(cell);
//                    }
//                    table.Append(dataRow);
//                }

//                body.Append(table);
//                mainPart.Document.Append(body);
//                mainPart.Document.Save();
//            }

//            return ms.ToArray();
//        }
//    }
//}

//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;

//namespace PDF_generator_tool.Services
//{
//    public class WordGenerator
//    {
//        public static byte[] GenerateWord(List<Dictionary<string, object>> data, List<string> selectedFields, string tableName, List<string> primaryKeys)
//        {
//            using var ms = new MemoryStream();
//            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
//            {
//                var mainPart = wordDoc.AddMainDocumentPart();
//                mainPart.Document = new Document();
//                var body = new Body();

//                // Title
//                body.AppendChild(new Paragraph(
//                    new Run(new Text("📋 Project Report"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    })
//                {
//                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
//                });

//                // Metadata
//                body.AppendChild(new Paragraph(new Run(new Text($"Table: {tableName}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Generated on: {DateTime.Now:dd-MMM-yyyy HH:mm}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Total Records: {data.Count}"))));
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Table with borders
//                var table = new Table();
//                TableProperties tblProps = new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                );
//                table.AppendChild(tblProps);

//                // Header Row
//                var headerRow = new TableRow();
//                foreach (var field in selectedFields)
//                {
//                    var cell = new TableCell(
//                        new Paragraph(new Run(new Text(field)))
//                        {
//                            ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Left })
//                        }
//                    );
//                    headerRow.Append(cell);
//                }
//                table.Append(headerRow);

//                // Data Rows
//                foreach (var row in data)
//                {
//                    var dataRow = new TableRow();
//                    foreach (var field in selectedFields)
//                    {
//                        var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
//                        var cell = new TableCell(
//                            new Paragraph(new Run(new Text(value)))
//                        );
//                        dataRow.Append(cell);
//                    }
//                    table.Append(dataRow);
//                }

//                body.Append(table);

//                // Summary Table
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                var summaryField = selectedFields.FirstOrDefault(f => !primaryKeys.Contains(f, StringComparer.OrdinalIgnoreCase)) ?? selectedFields.First();
//                var grouped = data.GroupBy(d => d.ContainsKey(summaryField) ? d[summaryField]?.ToString() ?? "N/A" : "N/A")
//                                  .ToDictionary(g => g.Key, g => g.Count());

//                body.AppendChild(new Paragraph(
//                    new Run(new Text($"📈 Summary by: {summaryField}"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    }
//                ));

//                var summaryTable = new Table();
//                summaryTable.AppendChild(new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                ));

//                // Summary Headers
//                var summaryHeader = new TableRow();
//                summaryHeader.Append(
//                    new TableCell(new Paragraph(new Run(new Text("Value")))),
//                    new TableCell(new Paragraph(new Run(new Text("Count"))))
//                );
//                summaryTable.Append(summaryHeader);

//                // Summary Data
//                foreach (var item in grouped)
//                {
//                    var row = new TableRow();
//                    row.Append(
//                        new TableCell(new Paragraph(new Run(new Text(item.Key)))),
//                        new TableCell(new Paragraph(new Run(new Text(item.Value.ToString()))))
//                    );
//                    summaryTable.Append(row);
//                }

//                body.Append(summaryTable);
//                mainPart.Document.Append(body);
//                mainPart.Document.Save();
//            }

//            return ms.ToArray();
//        }
//    }
//}

//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;

//namespace PDF_generator_tool.Services
//{
//    public class WordGenerator
//    {
//        public static byte[] GenerateWord(List<Dictionary<string, object>> data, List<string> selectedFields, string tableName, List<string> primaryKeys)
//        {
//            using var ms = new MemoryStream();
//            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
//            {
//                var mainPart = wordDoc.AddMainDocumentPart();
//                mainPart.Document = new Document();
//                var body = new Body();

//                // Title
//                body.AppendChild(new Paragraph(
//                    new Run(new Text("📋 Project Report"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    })
//                {
//                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
//                });

//                // Metadata
//                body.AppendChild(new Paragraph(new Run(new Text($"Table: {tableName}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Generated on: {DateTime.Now:dd-MMM-yyyy HH:mm}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Total Records: {data.Count}"))));
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Table with borders
//                var table = new Table();
//                TableProperties tblProps = new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                );
//                table.AppendChild(tblProps);

//                // Header Row
//                var headerRow = new TableRow();
//                foreach (var field in selectedFields)
//                {
//                    var headerText = primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase) ? $"{field} 🔑" : field;

//                    var cell = new TableCell(
//                        new Paragraph(new Run(new Text(headerText)))
//                        {
//                            ParagraphProperties = new ParagraphProperties(new Justification { Val = JustificationValues.Left })
//                        }
//                    );
//                    headerRow.Append(cell);
//                }
//                table.Append(headerRow);

//                // Data Rows
//                foreach (var row in data)
//                {
//                    var dataRow = new TableRow();
//                    foreach (var field in selectedFields)
//                    {
//                        var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
//                        var cell = new TableCell(
//                            new Paragraph(new Run(new Text(value)))
//                        );
//                        dataRow.Append(cell);
//                    }
//                    table.Append(dataRow);
//                }

//                body.Append(table);
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Summary Table
//                var summaryField = selectedFields.FirstOrDefault(f => !primaryKeys.Contains(f, StringComparer.OrdinalIgnoreCase)) ?? selectedFields.First();
//                var grouped = data.GroupBy(d => d.ContainsKey(summaryField) ? d[summaryField]?.ToString() ?? "N/A" : "N/A")
//                                  .ToDictionary(g => g.Key, g => g.Count());

//                body.AppendChild(new Paragraph(
//                    new Run(new Text($"📈 Summary by: {summaryField}"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    }
//                ));

//                var summaryTable = new Table();
//                summaryTable.AppendChild(new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                ));

//                // Summary Headers
//                var summaryHeader = new TableRow();
//                summaryHeader.Append(
//                    new TableCell(new Paragraph(new Run(new Text("Value")))),
//                    new TableCell(new Paragraph(new Run(new Text("Count"))))
//                );
//                summaryTable.Append(summaryHeader);

//                // Summary Data
//                foreach (var item in grouped)
//                {
//                    var row = new TableRow();
//                    row.Append(
//                        new TableCell(new Paragraph(new Run(new Text(item.Key)))),
//                        new TableCell(new Paragraph(new Run(new Text(item.Value.ToString()))))
//                    );
//                    summaryTable.Append(row);
//                }

//                body.Append(summaryTable);
//                mainPart.Document.Append(body);
//                mainPart.Document.Save();
//            }

//            return ms.ToArray();
//        }
//    }
//}
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;

//namespace PDF_generator_tool.Services
//{
//    public class WordGenerator
//    {
//        public static byte[] GenerateWord(List<Dictionary<string, object>> data, List<string> selectedFields, string tableName, List<string> primaryKeys)
//        {
//            using var ms = new MemoryStream();
//            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
//            {
//                var mainPart = wordDoc.AddMainDocumentPart();
//                mainPart.Document = new Document();
//                var body = new Body();

//                // Title
//                body.AppendChild(new Paragraph(
//                    new Run(new Text("📋 Project Report"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    })
//                {
//                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
//                });

//                // Metadata
//                body.AppendChild(new Paragraph(new Run(new Text($"Table: {tableName}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Generated on: {DateTime.Now:dd-MMM-yyyy HH:mm}"))));
//                body.AppendChild(new Paragraph(new Run(new Text($"Total Records: {data.Count}"))));
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Table with borders
//                var table = new Table();
//                TableProperties tblProps = new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                );
//                table.AppendChild(tblProps);

//                // Header Row
//                var headerRow = new TableRow();
//                foreach (var field in selectedFields)
//                {
//                    var isPK = primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase);
//                    var headerText = isPK ? $"{field} 🔑" : field;

//                    var headerCell = new TableCell(
//                        new Paragraph(new Run(new Text(headerText)))
//                    );

//                    if (isPK)
//                    {
//                        headerCell.Append(new TableCellProperties(
//                            new Shading() { Fill = "FFF9C4", Val = ShadingPatternValues.Clear, Color = "auto" }
//                        ));
//                    }

//                    headerRow.Append(headerCell);
//                }
//                table.Append(headerRow);

//                // Data Rows
//                foreach (var row in data)
//                {
//                    var dataRow = new TableRow();
//                    foreach (var field in selectedFields)
//                    {
//                        var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
//                        var cell = new TableCell(new Paragraph(new Run(new Text(value))));

//                        if (primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase))
//                        {
//                            cell.Append(new TableCellProperties(
//                                new Shading() { Fill = "FFF9C4", Val = ShadingPatternValues.Clear, Color = "auto" }
//                            ));
//                        }

//                        dataRow.Append(cell);
//                    }
//                    table.Append(dataRow);
//                }

//                body.Append(table);
//                body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

//                // Summary Table
//                var summaryField = selectedFields.FirstOrDefault(f => !primaryKeys.Contains(f, StringComparer.OrdinalIgnoreCase)) ?? selectedFields.First();
//                var grouped = data.GroupBy(d => d.ContainsKey(summaryField) ? d[summaryField]?.ToString() ?? "N/A" : "N/A")
//                                  .ToDictionary(g => g.Key, g => g.Count());

//                body.AppendChild(new Paragraph(
//                    new Run(new Text($"📈 Summary by: {summaryField}"))
//                    {
//                        RunProperties = new RunProperties(new Bold())
//                    }
//                ));

//                var summaryTable = new Table();
//                summaryTable.AppendChild(new TableProperties(
//                    new TableBorders(
//                        new TopBorder { Val = BorderValues.Single, Size = 4 },
//                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
//                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
//                        new RightBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
//                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
//                    )
//                ));

//                // Summary Headers
//                var summaryHeader = new TableRow();
//                summaryHeader.Append(
//                    new TableCell(new Paragraph(new Run(new Text("Value")))),
//                    new TableCell(new Paragraph(new Run(new Text("Count"))))
//                );
//                summaryTable.Append(summaryHeader);

//                // Summary Data
//                foreach (var item in grouped)
//                {
//                    var row = new TableRow();
//                    row.Append(
//                        new TableCell(new Paragraph(new Run(new Text(item.Key)))),
//                        new TableCell(new Paragraph(new Run(new Text(item.Value.ToString()))))
//                    );
//                    summaryTable.Append(row);
//                }

//                body.Append(summaryTable);
//                mainPart.Document.Append(body);
//                mainPart.Document.Save();
//            }

//            return ms.ToArray();
//        }
//    }
//}


// ✅ Updated WordGenerator.cs for multi-table support
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace PDF_generator_tool.Services
{
    public class WordGenerator
    {
        public static byte[] GenerateWord(Dictionary<string, List<Dictionary<string, object>>> dataMap, Dictionary<string, List<string>> selectedFieldsMap, Dictionary<string, List<string>> primaryKeysMap)
        {
            using var ms = new MemoryStream();
            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = new Body();

                foreach (var tableEntry in dataMap)
                {
                    string tableName = tableEntry.Key;
                    var data = tableEntry.Value;
                    var selectedFields = selectedFieldsMap[tableName];
                    var primaryKeys = primaryKeysMap.ContainsKey(tableName) ? primaryKeysMap[tableName] : new List<string>();

                    // Title
                    body.AppendChild(new Paragraph(
                        new Run(new Text($"📋 Project Report - {tableName}"))
                        {
                            RunProperties = new RunProperties(new Bold())
                        })
                    {
                        ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
                    });

                    body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer

                    // Metadata
                    body.AppendChild(new Paragraph(new Run(new Text($"Table: {tableName}"))));
                    body.AppendChild(new Paragraph(new Run(new Text($"Generated on: {DateTime.Now:dd-MMM-yyyy HH:mm}"))));
                    body.AppendChild(new Paragraph(new Run(new Text($"Total Records: {data.Count}"))));
                    body.AppendChild(new Paragraph(new Run(new Text(""))));

                    // Table
                    var table = new Table();
                    table.AppendChild(new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )));

                    // Headers
                    var headerRow = new TableRow();
                    foreach (var field in selectedFields)
                    {
                        var isPK = primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase);
                        var headerText = isPK ? $"{field} 🔑" : field;
                        var cell = new TableCell(new Paragraph(new Run(new Text(headerText))));
                        if (isPK)
                        {
                            cell.Append(new TableCellProperties(new Shading { Fill = "FFF9C4" }));
                        }
                        headerRow.Append(cell);
                    }
                    table.Append(headerRow);

                    // Rows
                    foreach (var row in data)
                    {
                        var tableRow = new TableRow();
                        foreach (var field in selectedFields)
                        {
                            var value = row.ContainsKey(field) ? row[field]?.ToString() ?? "" : "N/A";
                            var cell = new TableCell(new Paragraph(new Run(new Text(value))));
                            if (primaryKeys.Contains(field, StringComparer.OrdinalIgnoreCase))
                            {
                                cell.Append(new TableCellProperties(new Shading { Fill = "FFF9C4" }));
                            }
                            tableRow.Append(cell);
                        }
                        table.Append(tableRow);
                    }

                    body.Append(table);

                    // Summary
                    body.AppendChild(new Paragraph(new Run(new Text(""))));
                    var summaryField = selectedFields.FirstOrDefault(f => !primaryKeys.Contains(f, StringComparer.OrdinalIgnoreCase)) ?? selectedFields.First();
                    var grouped = data.GroupBy(d => d.ContainsKey(summaryField) ? d[summaryField]?.ToString() ?? "N/A" : "N/A")
                                      .ToDictionary(g => g.Key, g => g.Count());

                    body.AppendChild(new Paragraph(
                        new Run(new Text($"📈 Summary by: {summaryField}"))
                        {
                            RunProperties = new RunProperties(new Bold())
                        }));

                    var summaryTable = new Table();
                    summaryTable.AppendChild(new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )));

                    var summaryHeader = new TableRow();
                    summaryHeader.Append(
                        new TableCell(new Paragraph(new Run(new Text("Value")))),
                        new TableCell(new Paragraph(new Run(new Text("Count")))));
                    summaryTable.Append(summaryHeader);

                    foreach (var item in grouped)
                    {
                        var row = new TableRow();
                        row.Append(
                            new TableCell(new Paragraph(new Run(new Text(item.Key)))),
                            new TableCell(new Paragraph(new Run(new Text(item.Value.ToString())))));
                        summaryTable.Append(row);
                    }

                    body.Append(summaryTable);
                    body.AppendChild(new Paragraph(new Run(new Text("")))); // Spacer
                }

                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }

            return ms.ToArray();
        }
    }
}
