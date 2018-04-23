using System;
using DevExpress.Export.Xl;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace XLExportExamples
{
    public static class TableActions
    {
        #region Actions
        public static Action<Stream, XlDocumentFormat> AddTableAction = AddTable;
        public static Action<Stream, XlDocumentFormat> DisableFilteringAction = DisableFiltering;
        public static Action<Stream, XlDocumentFormat> HiddenHeaderRowAction = HiddenHeaderRow;
        public static Action<Stream, XlDocumentFormat> HiddenTotalRowAction = HiddenTotalRow;
        public static Action<Stream, XlDocumentFormat> SideBySideAction = SideBySide;
        public static Action<Stream, XlDocumentFormat> TableStyleAction = TableStyle;
        public static Action<Stream, XlDocumentFormat> TableStyleOptionsAction = TableStyleOptions;
        public static Action<Stream, XlDocumentFormat> CustomFormattingAction = CustomFormatting;
        public static Action<Stream, XlDocumentFormat> CalculatedColumnAction = CalculatedColumn;
        #endregion

        static void AddTable(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];
					
					#region #AddTable
                    IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, true);
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                        // Specify the number format for the "Amount" column and its total cell.
                        XlNumberFormat accounting = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        table.Columns[2].DataFormatting = accounting;
                        table.Columns[2].TotalRowFormatting = accounting;
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);

                    // Create the total row and finish the table.
                    using (IXlRow row = sheet.CreateRow())
                        row.EndTable(table, true);
					#endregion #AddTable
                }
            }
        }

        static void DisableFiltering(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #DisableFiltering
					IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, true);
                        // Disable the filtering functionality for the table. 
                        table.HasAutoFilter = false;
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);

                    // Create the total row and finish the table.
                    using (IXlRow row = sheet.CreateRow())
                        row.EndTable(table, true);
					#endregion #DisableFiltering
                }
            }
        }

        static void HiddenHeaderRow(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #HiddenHeaderRow
					IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the table with the hidden header row.
                        table = row.BeginTable(columnNames, false);
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                        // Populate the first table row with data.
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    }

                    // Generate the remaining table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);

                    // Create the total row and finish the table.
                    using (IXlRow row = sheet.CreateRow())
                        row.EndTable(table, true);
					#endregion #HiddenHeaderRow
                }
            }
        }

        static void HiddenTotalRow(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #HiddenTotalRow
					IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Start generating the table with a header row displayed.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        table = row.BeginTable(columnNames, true);
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);

                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
					#endregion #HiddenTotalRow
                }
            }
        }

        static void SideBySide(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create worksheet columns (A:G) and set their widths.
                    int[] widths = new int[] { 165, 120, 100, 20, 100, 120, 100 };
                    for (int i = 0; i < 7; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #SideBySideTables
                    // Specify two arrays containing column headings for the tables.
                    string[] columnNames1 = new string[] { "Product", "Category", "Amount" };
                    string[] columnNames2 = new string[] { "Region", "Category", "Amount" };

                    // Create the first row in the worksheet from which the tables start.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the first table with a header row displayed.
                        IXlTable table = row.BeginTable(columnNames1, true);
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                        // Specify the number format for the "Amount" column and its total cell.
                        XlNumberFormat accounting = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        table.Columns[2].DataFormatting = accounting;
                        table.Columns[2].TotalRowFormatting = accounting;

                        row.SkipCells(1);

                        // Start generating the second table with a header row displayed.
                        table = row.BeginTable(columnNames2, true);
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                        // Specify the number format for the "Amount" column and its total cell.
                        table.Columns[2].DataFormatting = accounting;
                        table.Columns[2].TotalRowFormatting = accounting;
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                        row.SkipCells(1);
                        row.BulkCells(new object[] { "East", "Dairy Products", 16000 }, null);
                    }
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                        row.SkipCells(1);
                        row.BulkCells(new object[] { "East", "Grains/Cereals", 14500 }, null);
                    }
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15500 }, null);
                        row.SkipCells(1);
                        row.BulkCells(new object[] { "West", "Dairy Products", 16500 }, null);
                    }
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.SkipCells(1);
                        row.BulkCells(new object[] { "West", "Grains/Cereals", 13500 }, null);
                    }

                    // Create the total rows and finish the tables.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        foreach (IXlTable table in sheet.Tables)
                            row.EndTable(table, true);
                    }
                    #endregion #SideBySideTables
                }
            }
        }

        static void TableStyle(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #TableStyle
                    IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, true);

                        // Apply the table style.
                        table.Style.Name = XlBuiltInTableStyleId.Dark7;
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
                    #endregion #TableStyle
                }
            }
        }

        static void TableStyleOptions(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #TableStyleOptions
                    IXlTable table;
                    // Specify an array containing column headings for tables.
                    string[] columnNames = new string[] { "Product", "Category", "Amount" };

                    // Create the row containing the table title.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Disable banded rows" }, XlCellFormatting.Title);
                    sheet.SkipRows(1);

                    // Start generating the table with a header row displayed.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        table = row.BeginTable(columnNames, true);
                        // Disable banded row formatting for the table.
                        table.Style.ShowRowStripes = false;
                    }
                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
                    sheet.SkipRows(1);

                    // Create the row containing the table title.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Enable banded columns" }, XlCellFormatting.Title);
                    sheet.SkipRows(1);

                    // Start generating the table with a header row displayed.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        table = row.BeginTable(columnNames, true);
                        // Apply banded column formatting to the table.
                        table.Style.ShowRowStripes = false;
                        table.Style.ShowColumnStripes = true;
                    }
                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
                    sheet.SkipRows(1);

                    // Create the row containing the table title.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Highlight first column" }, XlCellFormatting.Title);
                    sheet.SkipRows(1);

                    // Start generating the table with a header row displayed.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        table = row.BeginTable(columnNames, true);
                        // Display special formatting for the first column of the table.
                        table.Style.ShowFirstColumn = true;
                    }
                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
                    sheet.SkipRows(1);

                    // Create the row containing the table title.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Highlight last column" }, XlCellFormatting.Title);
                    sheet.SkipRows(1);

                    // Start generating the table with a header row displayed.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        table = row.BeginTable(columnNames, true);
                        // Display special formatting for the last column of the table.
                        table.Style.ShowLastColumn = true;
                    }
                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    // Create the last table row and finish the table.
                    // The total row is not displayed for the table. 
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);
                        row.EndTable(table, false);
                    }
                    #endregion #TableStyleOptions
                }
            }
        }

        static void CustomFormatting(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create columns "A", "B" and "C" and set their widths.
                    int[] widths = new int[] { 165, 120, 100 };
                    for (int i = 0; i < 3; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #CustomFormatting
                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {

                        XlNumberFormat accounting = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";

                        // Create objects containing information about table columns (their names and formatting).
                        List<XlTableColumnInfo> columns = new List<XlTableColumnInfo>();
                        columns.Add(new XlTableColumnInfo("Product"));
                        columns.Add(new XlTableColumnInfo("Category"));
                        columns.Add(new XlTableColumnInfo("Amount"));

                        // Specify formatting settings for the last column of the table.
                        columns[2].HeaderRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent6, -0.3));
                        columns[2].DataFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.9));
                        columns[2].DataFormatting.NumberFormat = accounting;
                        columns[2].TotalRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.8));
                        columns[2].TotalRowFormatting.NumberFormat = accounting;

                        // Specify formatting settings for the header row of the table.
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent6, 0.0));
                        headerRowFormatting.Border = new XlBorder();
                        headerRowFormatting.Border.BottomColor = XlColor.FromArgb(0, 0, 0);
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Dashed;

                        // Start generating the table with a header row displayed.
                        IXlTable table = row.BeginTable(columns, true, headerRowFormatting);
                        // Apply the table style.
                        table.Style.Name = XlBuiltInTableStyleId.Medium16;
                        // Disable banded row formatting for the table.
                        table.Style.ShowRowStripes = false;
                        // Disable the filtering functionality for the table. 
                        table.HasAutoFilter = false;

                        // Specify formatting settings for the total row of the table.
                        table.TotalRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.9));
                        table.TotalRowFormatting.Border = new XlBorder()
                        {
                            BottomColor = XlColor.FromTheme(XlThemeColor.Accent6, 0.0),
                            BottomLineStyle = XlBorderLineStyle.Thick,
                            TopColor = XlColor.FromArgb(0, 0, 0),
                            TopLineStyle = XlBorderLineStyle.Dashed
                        };

                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[2].TotalRowFunction = XlTotalRowFunction.Sum;
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", "Dairy Products", 17000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", "Grains/Cereals", 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", "Dairy Products", 15000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Ravioli Angelo", "Grains/Cereals", 12500 }, null);

                    // Create the total row and finish the table.
                    using (IXlRow row = sheet.CreateRow())
                        row.EndTable(sheet.Tables[0], true);
                    #endregion #CustomFormatting
                }
            }
        }

        static void CalculatedColumn(Stream stream, XlDocumentFormat documentFormat)
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

                    // Create worksheet columns (A:F) and set their widths.
                    int[] widths = new int[] { 165, 100, 100, 100, 100, 110 };
                    for (int i = 0; i < 6; i++)
                        using (IXlColumn column = sheet.CreateColumn())
                            column.WidthInPixels = widths[i];

                    #region #CalculatedColumn
                    IXlTable table;
                    // Specify an array containing column headings for a table.
                    string[] columnNames = new string[] { "Product", "Q1", "Q2", "Q3", "Q4", "Yearly Total" };

                    // Create the first row in the worksheet from which the table starts.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, true);
                        // Specify the total row label.
                        table.Columns[0].TotalRowLabel = "Total";
                        // Specify the function to calculate the total.
                        table.Columns[5].TotalRowFunction = XlTotalRowFunction.Sum;
                        // Specify the number format for numeric values in the table and the total cell of the "Yearly Total" column.
                        XlNumberFormat accounting = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        table.DataFormatting = accounting;
                        table.Columns[5].TotalRowFormatting = accounting;
                        // Set the formula to calculate annual sales of each product
                        // and display results in the "Yearly Total" column.
                        table.Columns[5].SetFormula(XlFunc.Sum(table.GetRowReference("Q1", "Q4")));
                    }

                    // Generate table rows and populate them with data.
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Camembert Pierrot", 17000, 18500, 17500, 18000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Gnocchi di nonna Alice", 15500, 14500, 15000, 14000 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Mascarpone Fabioli", 15000, 15750, 16000, 15500 }, null);
                    using (IXlRow row = sheet.CreateRow())
                        row.BulkCells(new object[] { "Ravioli Angelo", 12500, 11000, 13500, 12000 }, null);

                    // Create the total row and finish the table.
                    using (IXlRow row = sheet.CreateRow())
                        row.EndTable(table, true);
                    #endregion #CalculatedColumn
                }
            }
        }
    }
}
