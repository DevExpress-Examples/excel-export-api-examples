using System;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class FormulaActions {

        #region Actions
        public static Action<Stream, XlDocumentFormat> SimpleFormulasAction = SimpleFormula;
        public static Action<Stream, XlDocumentFormat> ComplexFormulasAction = ComplexFormulas;
        public static Action<Stream, XlDocumentFormat> SharedFormulasAction = SharedFormulas;
        public static Action<Stream, XlDocumentFormat> SubtotalsAction = Subtotals;
        #endregion

        static void SimpleFormula(Stream stream, XlDocumentFormat documentFormat) {
            #region #SimpleFormula
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());
            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {
                    // Create worksheet columns and set their widths.
                    for (int i = 0; i < 4; i++) {
                        using (IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 80;
                        }
                    }
                    // Generate data for the document.
                    string[] header = new string[] { "Description", "QTY", "Price", "Amount" };
                    string[] product = new string[] { "Camembert", "Gorgonzola", "Mascarpone", "Mozzarella" };
                    int[] qty = new int[] { 12, 15, 25, 10 };
                    double[] price = new double[] { 23.25, 15.50, 12.99, 8.95 };
                    double discount = 0.2;
                    // Create the header row.
                    using (IXlRow row = sheet.CreateRow()) {
                        for (int i = 0; i < 4; i++) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = header[i];
                            }
                        }
                    }
                    // Create data rows using string formulas.
                    for (int i = 0; i < 4; i++) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = product[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = qty[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = price[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Set the formula to calculate the amount 
                                // applying 20% quantity discount on orders more than 15 items. 
                                cell.SetFormula(String.Format("=IF(B{0}>15,C{0}*B{0}*(1-{1}),C{0}*B{0})", i + 2, discount));
                            }
                        }
                    }

                }
            }
            #endregion #SimpleFormula
        }

        static void ComplexFormulas(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 50;
                    }
                    for (int i = 0; i < 2; i++) {
                        using (IXlColumn column = sheet.CreateColumn()) {
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
                    using (IXlRow row = sheet.CreateRow()) {
                        for (int i = 0; i < 4; i++) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = header[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }

                    #region #Formula_String
                    // Create data rows using string formulas.
                    for (int i = 0; i < 4; i++) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = product[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = qty[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = price[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Set the formula to calculate the amount per product.
                                cell.SetFormula(String.Format("B{0}*C{0}", i + 2));
                            }
                        }
                    }
                    #endregion #Formula_String
                    #region #Formula_IXlFormulaParameter
                    // Create the total row using IXlFormulaParameter.
                    using (IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(2);
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total:";
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                        using (IXlCell cell = row.CreateCell()) {
                            // Set the formula to calculate the total amount plus 10 handling fee.
                            // =SUM($D$2:$D$5)+10
                            IXlFormulaParameter const10 = XlFunc.Param(10);
                            IXlFormulaParameter sumAmountFunction = XlFunc.Sum(XlCellRange.FromLTRB(cell.ColumnIndex, 1, cell.ColumnIndex, row.RowIndex - 1).AsAbsolute());
                            cell.SetFormula(XlOper.Add(sumAmountFunction, const10));
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                    }
                    #endregion #Formula_IXlFormulaParameter
                    #region #Formula_XlExpression
                    // Create a formula using XlExpression.
                    using (IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(2);
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Mean value:";
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                        using (IXlCell cell = row.CreateCell()) {
                            // Set the formula to calculate the mean value.
                            // =$D$6/4
                            XlExpression expression = new XlExpression();
                            expression.Add(new XlPtgRef(new XlCellPosition(cell.ColumnIndex, row.RowIndex - 1, XlPositionType.Absolute, XlPositionType.Absolute)));
                            expression.Add(new XlPtgInt(row.RowIndex - 2));
                            expression.Add(new XlPtgBinaryOperator(XlPtgTypeCode.Div));
                            cell.SetFormula(expression);
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                    }
                    #endregion #Formula_XlExpression
                }
            }
        }

        static void SharedFormulas(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 50;
                    }
                    for (int i = 0; i < 2; i++) {
                        using (IXlColumn column = sheet.CreateColumn()) {
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
                    using (IXlRow row = sheet.CreateRow()) {
                        for (int i = 0; i < 4; i++) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = header[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }
                    #region #SharedFormulas
                    // Create data rows.
                    for (int i = 0; i < 4; i++) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = product[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = qty[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = price[i];
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the shared formula to calculate the amount per product. 
                                if (i == 0)
                                    cell.SetSharedFormula("B2*C2", XlCellRange.FromLTRB(3, 1, 3, 4));
                                else
                                    cell.SetSharedFormula(new XlCellPosition(3, 1));
                            }
                        }
                    }
                    #endregion #SharedFormulas

                    // Create the total row.
                    using (IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(2);
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total:";
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                        using (IXlCell cell = row.CreateCell()) {
                            // Set the formula to calculate the total amount.
                            cell.SetFormula("SUM(D2:D5)");
                            cell.ApplyFormatting(totalRowFormatting);
                        }
                    }
                }

            }
        }

        static void Subtotals(Stream stream, XlDocumentFormat documentFormat) {
            // Declare a variable that indicates the start of the data rows to calculate grand totals.
            int startDataRowForGrandTotal;
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Create the column "A" and set its width.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    // Create five successive columns and set the specific number format for their cells.
                    for (int i = 0; i < 5; i++) {
                        using (IXlColumn column = sheet.CreateColumn()) {
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
                    string[] productsDairy = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                    string[] productsCereals = new string[] { "Gnocchi di nonna Alice", "Gustaf's Knäckebröd", "Ravioli Angelo", "Singaporean Hokkien Fried Mee" };

                    // Create the header row.
                    using (IXlRow row = sheet.CreateRow()) {
                        startDataRowForGrandTotal = row.RowIndex + 1;
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
                        }
                        for (int i = 0; i < 4; i++) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                                cell.ApplyFormatting(headerRowFormatting);
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                            }
                        }
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Yearly total";
                            cell.ApplyFormatting(headerRowFormatting);
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                        }
                    }

                    // Create data rows for Dairy products.
                    for (int i = 0; i < 4; i++) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = productsDairy[i];
                                cell.ApplyFormatting(rowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8));
                            }
                            for (int j = 0; j < 4; j++) {
                                using (IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the SUM function to calculate annual sales for each product.   
                                cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)));
                                cell.ApplyFormatting(rowFormatting);
                                cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                            }
                        }
                    }

                    // Create the total row for Dairies.
                    using (IXlRow row = sheet.CreateRow()) {
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Subtotal";
                            cell.ApplyFormatting(totalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                        }
                        for (int j = 0; j < 5; j++) {
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the SUBTOTAL function to calculate total sales for each quarter and the entire year.  
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(totalRowFormatting);
                            }
                        }
                    }


                    // Create data rows for Cereals.
                    for (int i = 0; i < 4; i++) {
                        using (IXlRow row = sheet.CreateRow()) {
                            using (IXlCell cell = row.CreateCell()) {
                                cell.Value = productsCereals[i];
                                cell.ApplyFormatting(rowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8));
                            }
                            for (int j = 0; j < 4; j++) {
                                using (IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                    cell.ApplyFormatting(rowFormatting);
                                }
                            }
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the SUM function to calculate annual sales for each product.   
                                cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)));
                                cell.ApplyFormatting(rowFormatting);
                                cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                            }
                        }
                    }
                    // Create the total row for Cereals.
                    using (IXlRow row = sheet.CreateRow()) {
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Subtotal";
                            cell.ApplyFormatting(totalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                        }
                        for (int j = 0; j < 5; j++) {
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the SUBTOTAL function to calculate total sales for each quarter and the entire year.  
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(totalRowFormatting);
                            }
                        }
                    }
                    #region #SubtotalFunction
                    // Create the grand total row.
                    using (IXlRow row = sheet.CreateRow()) {
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Grand Total";
                            cell.ApplyFormatting(totalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                        }
                        for (int j = 0; j < 5; j++) {
                            using (IXlCell cell = row.CreateCell()) {
                                // Use the SUBTOTAL function to calculate grand total sales for each quarter and the entire year.  
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, startDataRowForGrandTotal, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(totalRowFormatting);
                            }
                        }
                    }
                    #endregion #SubtotalFunction

                }
            }
        }

    }
}
