using System;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;

namespace XLExportExamples
{
    public static class PageViewAndLayoutActions {

        #region Actions
        public static Action<Stream, XlDocumentFormat> FreezeRowAction = FreezeRow;
        public static Action<Stream, XlDocumentFormat> FreezeColumnAction = FreezeColumn;
        public static Action<Stream, XlDocumentFormat> FreezePanesAction = FreezePanes;
        public static Action<Stream, XlDocumentFormat> HeadersFootersAction = HeadersFooters;
        public static Action<Stream, XlDocumentFormat> PageBreaksAction = PageBreaks;
        public static Action<Stream, XlDocumentFormat> PageMarginsAction = PageMargins;
        public static Action<Stream, XlDocumentFormat> PageSetupAction = PageSetup;
        public static Action<Stream, XlDocumentFormat> PrintAreaAction = PrintArea;
        public static Action<Stream, XlDocumentFormat> PrintOptionsAction = PrintOptions;
        public static Action<Stream, XlDocumentFormat> PrintTitlesAction = PrintTitles;
        #endregion

        static void FreezeRow(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #FreezeRow
                    // Freeze the first row in the worksheet.
                    sheet.SplitPosition = new XlCellPosition(0, 1);
                    #endregion #FreezeRow

                    // Generate data for the document.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                }
            }
        }

        static void FreezeColumn(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #FreezeColumn
                    // Freeze the first column in the worksheet.
                    sheet.SplitPosition = new XlCellPosition(1, 0);
                    #endregion #FreezeColumn

                    // Generate data for the document.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                }
            }
        }

        static void FreezePanes(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #FreezePanes
                    // Freeze the first column and the first row.
                    sheet.SplitPosition = new XlCellPosition(1, 1);
                    #endregion #FreezePanes

                    // Generate data for the document.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                }
            }
        }

        static void HeadersFooters(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #HeaderAndFooters
                    // Specify different headers and footers for the odd-numbered and even-numbered pages.
                    sheet.HeaderFooter.DifferentOddEven = true;
                    // Add the bold text to the header left section, 
                    // and insert the workbook name into the header right section.
                    sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.Bold + "Sample report", null, XlHeaderFooter.BookName);
                    // Insert the current page number into the footer right section. 
                    sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR(null, null, XlHeaderFooter.PageNumber);
                    // Insert the workbook file path into the header left section, 
                    // and add the worksheet name to the header right section. 
                    sheet.HeaderFooter.EvenHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.BookPath, null, XlHeaderFooter.SheetName);
                    // Insert the current page number into the footer left section 
                    // and add the current date to the footer right section. 
                    sheet.HeaderFooter.EvenFooter = XlHeaderFooter.FromLCR(XlHeaderFooter.PageNumber, null, XlHeaderFooter.Date);
                    #endregion #HeaderAndFooters

                    // Generate data for the document.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                }
            }
        }

        static void PageBreaks(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PageBreaks
                    // Insert a page break after the column "B".
                    sheet.ColumnPageBreaks.Add(2);
                    // Insert a page break after the tenth row.
                    sheet.RowPageBreaks.Add(10);
                    #endregion #PageBreaks

                    // Generate data for the document.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.Formatting = new XlCellFormatting();
                        column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Sales";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }
                }
            }
        }

        static void PageMargins(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PageMargins
                    sheet.PageMargins = new XlPageMargins();
                    // Set the unit of margin measurement.
                    sheet.PageMargins.PageUnits = XlPageUnits.Centimeters;
                    // Specify page margins.
                    sheet.PageMargins.Left = 2.0;
                    sheet.PageMargins.Right = 1.0;
                    sheet.PageMargins.Top = 1.25;
                    sheet.PageMargins.Bottom = 1.25;
                    // Specify header and footer margins.
                    sheet.PageMargins.Header = 0.7;
                    sheet.PageMargins.Footer = 0.7;
                    #endregion #PageMargins

                    // Generate data for the document.
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Invoke the Page Setup dialog to control margin settings.";
                        }
                    }
                }
            }
        }

        static void PageSetup(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PageSetup
                    // Specify page settings for the worksheet.
                    sheet.PageSetup = new XlPageSetup();
                    // Select the paper size.
                    sheet.PageSetup.PaperKind = System.Drawing.Printing.PaperKind.A4;
                    // Set the page orientation to Landscape.
                    sheet.PageSetup.PageOrientation = XlPageOrientation.Landscape;
                    //  Scale the print area to fit to one page wide.
                    sheet.PageSetup.FitToPage = true;
                    sheet.PageSetup.FitToWidth = 1;
                    sheet.PageSetup.FitToHeight = 0;
                    //  Print in black and white.
                    sheet.PageSetup.BlackAndWhite = true;
                    // Specify the number of copies.
                    sheet.PageSetup.Copies = 2;
                    #endregion #PageSetup

                    // Generate data for the document.
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Invoke the Page Setup dialog to control page settings.";
                        }
                    }
                }
            }
        }

        static void PrintArea(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PrintArea
                    // Set the print area to cells A1:E5.
                    sheet.PrintArea = XlCellRange.FromLTRB(0, 0, 4, 4);
                    #endregion #PrintArea

                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 110;
                        column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom);
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 190;
                    }
                    for(int i = 0; i < 2; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 90;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 130;
                    }
                    sheet.SkipColumns(1);
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 130;
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Employee ID";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Employee name";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Salary";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Bonus";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Department";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Departments";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate data for the document.
                    int[] id = new int[] { 10115, 10709, 10401, 10204 };
                    string[] name = new string[] { "Augusta Delono", "Chris Cadwell", "Frank Diamond", "Simon Newman" };
                    int[] salary = new int[] { 1100, 2000, 1750, 1250 };
                    int[] bonus = new int[] { 50, 180, 100, 80 };
                    int[] deptid = new int[] { 0, 2, 3, 3 };
                    string[] department = new string[] { "Accounting", "IT", "Management", "Manufacturing" };
                    for(int i = 0; i < 4; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = id[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = name[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = salary[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = bonus[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = department[deptid[i]];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            row.SkipCells(1);
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = department[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }
                    // Restrict data entry in the cell range E2:E5 to values in a drop-down list obtained from the cells G2:G5.
                    XlDataValidation validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(4, 1, 4, 4));
                    validation.Type = XlDataValidationType.List;
                    validation.Criteria1 = XlCellRange.FromLTRB(6, 1, 6, 4).AsAbsolute(); 
                    sheet.DataValidations.Add(validation);
                }
            }
        }

        static void PrintOptions(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PrintOptions
                    // Specify print options for the worksheet.
                    sheet.PrintOptions = new XlPrintOptions();
                    // Print row and column headings.
                    sheet.PrintOptions.Headings = true;
                    // Print gridlines.
                    sheet.PrintOptions.GridLines = true;
                    // Center worksheet data on a printed page.
                    sheet.PrintOptions.HorizontalCentered = true;
                    sheet.PrintOptions.VerticalCentered = true;
                    #endregion #PrintOptions

                    // Generate data for the document.
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.Formatting = new XlCellFormatting();
                        column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Sales";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }
                }
            }
        }

        static void PrintTitles(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    #region #PrintTitles
                    // Print the first column and the first row on every page.
                    sheet.PrintTitles.SetColumns(0, 0);
                    sheet.PrintTitles.SetRows(0, 0);
                    #endregion #PrintTitles

                    // Generate data for the document.
                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Generate data rows.
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
