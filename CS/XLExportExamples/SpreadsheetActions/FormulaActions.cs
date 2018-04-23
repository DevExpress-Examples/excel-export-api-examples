using System;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples
{
    public static class FormulaActions {
        
        #region Actions
        public static Action<Stream, XlDocumentFormat> FormulasAction = Formulas;
        public static Action<Stream, XlDocumentFormat> SharedFormulasAction = SharedFormulas;
        public static Action<Stream, XlDocumentFormat> FunctionsAction = Functions;
        #endregion

        static void Formulas(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #Formulas
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 50;
                    }
                    for(int i = 0; i < 2; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 80;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.BodyFont();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
                    // Specify formatting settings for the total row.
                    XlCellFormatting totalRowFormatting = new XlCellFormatting();
                    totalRowFormatting.Font = XlFont.BodyFont();
                    totalRowFormatting.Font.Bold = true;

                    // Generate data for the document.
                    string[] header = new string[] { "Description", "QTY", "Price", "Amount" };
                    string[] product = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                    int[] qty = new int[] { 12, 15, 25, 10 };
                    double[] price = new double[] { 23.25, 15.50, 12.99, 8.95 };

                    // Create the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = header[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Create data rows.
                    for(int i = 0; i < 4; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = product[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = qty[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = price[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                // Set the formula to calculate the amount per product.
                                cell.SetFormula(string.Format("B{0}*C{0}", i + 2));
                            }
                        }
                    }

                    // Create the total row.
                    using(IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(2);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total:";
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the formula to calculate the total amount.
                            cell.SetFormula("SUM(D2:D5)");
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                    }
                }
                #endregion #Formulas
            }
        }

        static void SharedFormulas(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #SharedFormulas
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 50;
                    }
                    for(int i = 0; i < 2; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 80;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.BodyFont();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
                    // Specify formatting settings for the total row.
                    XlCellFormatting totalRowFormatting = new XlCellFormatting();
                    totalRowFormatting.Font = XlFont.BodyFont();
                    totalRowFormatting.Font.Bold = true;

                    // Generate data for the document.
                    string[] header = new string[] { "Description", "QTY", "Price", "Amount" };
                    string[] product = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                    int[] qty = new int[] { 12, 15, 25, 10 };
                    double[] price = new double[] { 23.25, 15.50, 12.99, 8.95 };

                    // Create the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = header[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    // Create data rows.
                    for(int i = 0; i < 4; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = product[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = qty[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = price[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                // Use the shared formula to calculate the amount per product. 
                                if (i == 0)
                                    cell.SetSharedFormula("B2*C2", XlCellRange.FromLTRB(3, 1, 3, 4));
                                else
                                    cell.SetSharedFormula(new XlCellPosition(3, 1));
                            }
                        }
                    }

                    // Create the total row.
                    using(IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(2);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total:";
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the formula to calculate the total amount.
                            cell.SetFormula("SUM(D2:D5)");
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                    }
                }
                #endregion #SharedFormulas
            }
        }

        static void Functions(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #Functions
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create the column "A" and set its width.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    // Create five successive columns and set the specific number format for their cells.
                    for (int i = 0; i < 5; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.BodyFont();
                    rowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0));
                    // Specify formatting settings for the header row.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.BodyFont();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
                    // Specify formatting settings for the total row.
                    XlCellFormatting totalRowFormatting = new XlCellFormatting();
                    totalRowFormatting.Font = XlFont.BodyFont();
                    totalRowFormatting.Font.Bold = true;
                    totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0));

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };

                    // Create the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                            }
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Yearly total";
                            cell.ApplyFormatting(headerRowFormatting);
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                        }
                    }

                    // Create data rows.
                    for(int i = 0; i < 4; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                                cell.ApplyFormatting(rowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8));
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                // Use the SUM function to calculate annual sales for each product.   
                                cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)));
                                cell.ApplyFormatting(rowFormatting);
                                cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                            }
                        }
                    }

                    // Create the total row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total";
                            cell.ApplyFormatting(totalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                        }
                        for(int j = 0; j < 5; j++) {
                            using(IXlCell cell = row.CreateCell()) {
                                // Use the SUBTOTAL function to calculate total sales for each quarter and the entire year.  
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(totalRowFormatting);
                            }
                        }
                    }
                }
                #endregion #Functions
            }
        }

    }
}
