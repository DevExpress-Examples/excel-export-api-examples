using System;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;

namespace XLExportExamples
{
    class SparklineActions
    {
        #region Actions
        public static Action<Stream, XlDocumentFormat> AddSparklineGroupAction = AddSparklineGroup;
        public static Action<Stream, XlDocumentFormat> AddSparklineAction = AddSparkline;
        public static Action<Stream, XlDocumentFormat> AdjustScalingAction = AdjustScaling;
        public static Action<Stream, XlDocumentFormat> HighlightValuesAction = HighlightValues;
        public static Action<Stream, XlDocumentFormat> DisplayXAxisAction = DisplayXAxis;
        public static Action<Stream, XlDocumentFormat> SetDateRangeAction = SetDateRange;
        #endregion
        static void AddSparklineGroup(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }
                    for (int i = 0; i < 5; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));

                    string[] columnNames = new string[] { "Product", "Q1", "Q2", "Q3", "Q4" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(columnNames, headerRowFormatting);
                    }

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 4; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }

                    #region #AddSparklineGroup
                    // Create a group of line sparklines.
                    XlSparklineGroup group = new XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 4, 6), XlCellRange.FromLTRB(5, 1, 5, 6));
                    // Set the sparkline weight.
                    group.LineWeight = 1.25;
                    // Display data markers on the sparklines.
                    group.DisplayMarkers = true;
                    sheet.SparklineGroups.Add(group);
                    #endregion #AddSparklineGroup
                }
            }
        }

        static void AddSparkline(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }
                    for (int i = 0; i < 5; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));

                    string[] columnNames = new string[] { "Product", "Q1", "Q2", "Q3", "Q4" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(columnNames, headerRowFormatting);
                    }

                    // Create a group of line sparklines.
                    XlSparklineGroup group = new XlSparklineGroup();
                    // Set the sparkline color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, -0.2);
                    // Set the sparkline weight.
                    group.LineWeight = 1.25;
                    sheet.SparklineGroups.Add(group);

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 4; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }

                            #region #AddSparkline
                            // Add one more sparkline to the existing group.
                            int rowIndex = row.RowIndex;
                            group.Sparklines.Add(new XlSparkline(XlCellRange.FromLTRB(1, rowIndex, 4, rowIndex), XlCellRange.FromLTRB(5, rowIndex, 5, rowIndex)));
                            #endregion #AddSparkline
                        }
                    }
                }
            }
        }

        static void AdjustScaling(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }
                    for (int i = 0; i < 5; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    
                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));

                    string[] columnNames = new string[] { "Product", "Q1", "Q2", "Q3", "Q4" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(columnNames, headerRowFormatting);
                    }

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 4; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 1500);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }

                    #region #AdjustScaling
                    // Create a sparkline group.
                    XlSparklineGroup group = new XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 4, 6), XlCellRange.FromLTRB(5, 1, 5, 6));
                    // Set the sparkline color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0);
                    // Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column;
                    // Set the custom minimum value for the vertical axis.
                    group.MinScaling = XlSparklineAxisScaling.Custom;
                    group.ManualMin = 1000.0;
                    // Set the automatic maximum value for all sparklines in the group. 
                    group.MaxScaling = XlSparklineAxisScaling.Group;
                    sheet.SparklineGroups.Add(group);
                    #endregion #AdjustScaling

                }
            }
        }

        static void HighlightValues(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }
                    for (int i = 0; i < 9; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    
                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = rowFormatting.Clone();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));

                    string[] columnNames = new string[] { "State", "Q1'13", "Q2'13", "Q3'13", "Q4'13", "Q1'14", "Q2'14", "Q3'14", "Q4'14" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(columnNames, headerRowFormatting);
                    }

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 8; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round((random.NextDouble() + 0.5) * 2000 * Math.Sign(random.NextDouble() - 0.4));
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }

                    #region #HighlightValues
                    // Create a sparkline group.                   
                    XlSparklineGroup group = new XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 8, 6), XlCellRange.FromLTRB(9, 1, 9, 6));
                    // Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column;
                    // Set the series color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0);
                    // Set the color for negative points on sparklines. 
                    group.ColorNegative = XlColor.FromTheme(XlThemeColor.Accent2, 0.0);
                    // Set the color for the highest points on sparklines.
                    group.ColorHigh = XlColor.FromTheme(XlThemeColor.Accent6, 0.0);
                    // Highlight the highest and negative points on each sparkline in the group.
                    group.HighlightNegative = true;
                    group.HighlightHighest = true;
                    sheet.SparklineGroups.Add(group);
                    #endregion #HighlightValues
                }
            }
        }

        static void DisplayXAxis(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }
                    for (int i = 0; i < 9; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = rowFormatting.Clone();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));

                    string[] columnNames = new string[] { "State", "Q1'13", "Q2'13", "Q3'13", "Q4'13", "Q1'14", "Q2'14", "Q3'14", "Q4'14" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(columnNames, headerRowFormatting);
                    }

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 8; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round((random.NextDouble() + 0.5) * 2000 * Math.Sign(random.NextDouble() - 0.4));
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }

                    #region #DisplayXAxis
                    // Create a sparkline group.                   
                    XlSparklineGroup group = new XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 8, 6), XlCellRange.FromLTRB(9, 1, 9, 6));
                    // Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column;
                    // Display the horizontal axis.
                    group.DisplayXAxis = true;
                    // Set the series color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0);
                    // Highlight negative points on each sparkline in the group.
                    group.ColorNegative = XlColor.FromTheme(XlThemeColor.Accent2, 0.0);
                    group.HighlightNegative = true;
                    sheet.SparklineGroups.Add(group);
                    #endregion #DisplayXAxis
                }
            }
        }

        static void SetDateRange(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 200;
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                            column.ApplyFormatting((XlNumberFormat)@"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)");
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
                    headerRowFormatting.NumberFormat = XlNumberFormat.ShortDate;

                    object[] headerValues = new object[] { "Product", new DateTime(2015, 3, 15), new DateTime(2015, 4, 1), new DateTime(2015, 6, 1), new DateTime(2015, 10, 1), "Date Axis", "General Axis" };

                    // Generate the header row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(headerValues, headerRowFormatting);
                    }

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" };

                    foreach (string product in products)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                cell.Value = product;
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for (int j = 0; j < 4; j++)
                            {
                                using (IXlCell cell = row.CreateCell())
                                {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                    #region #SetDateRange
                    // Create a group of line sparklines.                    
                    XlSparklineGroup group = new XlSparklineGroup(XlCellRange.Parse("B2:E7"), XlCellRange.Parse("F2:F7"));
                    // Specify the date range for the sparkline group. 
                    group.DateRange = XlCellRange.Parse("B1:E1");
                    // Set the sparkline weight.
                    group.LineWeight = 1.25;
                    // Display data markers on the sparklines.
                    group.DisplayMarkers = true;
                    sheet.SparklineGroups.Add(group);
                    #endregion #SetDateRange
                    // Create another group of line sparklines with defaul general axis type.                    
                    XlSparklineGroup group1 = new XlSparklineGroup(XlCellRange.Parse("B2:E7"), XlCellRange.Parse("G2:G7"));
                    group1.LineWeight = 1.25;
                    group1.DisplayMarkers = true;
                    sheet.SparklineGroups.Add(group1);
                }
            }
        }
    }
}
