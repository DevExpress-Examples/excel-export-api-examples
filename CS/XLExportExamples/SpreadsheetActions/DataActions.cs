using System;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples
{
    public static class DataActions {

        #region Actions
        public static Action<Stream, XlDocumentFormat> AutoFilterAction = AutoFilter;
        public static Action<Stream, XlDocumentFormat> OutlineGroupingAction = OutlineGrouping;
        public static Action<Stream, XlDocumentFormat> DataValidationAction = DataValidation;
        #endregion

        static void AutoFilter(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
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
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));

                    // Generate the header row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Region";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Sales";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate data for the document.
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                    int[] amount = new int[] { 6750, 4500, 3550, 4250, 5500, 6250, 5325, 4235 };
                    for(int i = 0; i < 8; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = (i < 4) ? "East" : "West";
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i % 4];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = amount[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }
                    #region #AutoFilter
                    // Enable filtering for the data range.
                    sheet.AutoFilterRange = sheet.DataRange;
                    #endregion #AutoFilter
                }
            }
        }

        static void OutlineGrouping(Stream stream, XlDocumentFormat documentFormat) {
            #region #Group/Outline
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Specify the summary row and summary column location for the grouped data.
                    sheet.OutlineProperties.SummaryBelow = true;
                    sheet.OutlineProperties.SummaryRight = true;

                    // Create the column "A" and set its width.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }

                    // Begin to group worksheet columns starting from the column "B" to the column "E".
                    sheet.BeginGroup(false);
                    // Create four successive columns ("B", "C", "D" and "E") and set the specific number format for their cells.
                    for (int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }
                    // Finalize the group creation.
                    sheet.EndGroup();

                    // Create the column "F", adjust its width and set the specific number format for its cells.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.Formatting = new XlCellFormatting();
                        column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                    }

                    // Specify formatting settings for cells containing data.
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.BodyFont();
                    rowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0));
                    // Specify formatting settings for the header rows.
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.BodyFont();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
                    // Specify formatting settings for the total rows.
                    XlCellFormatting totalRowFormatting = new XlCellFormatting();
                    totalRowFormatting.Font = XlFont.BodyFont();
                    totalRowFormatting.Font.Bold = true;
                    totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0));
                    // Specify formatting settings for the grand total row.
                    XlCellFormatting grandTotalRowFormatting = new XlCellFormatting();
                    grandTotalRowFormatting.Font = XlFont.BodyFont();
                    grandTotalRowFormatting.Font.Bold = true;
                    grandTotalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.2));

                    // Generate data for the document.
                    Random random = new Random();
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };

                    // Begin to group worksheet rows (create the outer group of rows).
                    sheet.BeginGroup(false);
                    for(int p = 0; p < 2; p++) {
                        // Generate the header row.
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = (p == 0) ? "East" : "West";
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

                        // Create and group data rows (create the inner group of rows containing sales data for the specific region).
                        sheet.BeginGroup(false);
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
                                    cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)));
                                    cell.ApplyFormatting(rowFormatting);
                                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                                }
                            }
                        }
                        // Finalize the group creation.
                        sheet.EndGroup();

                        // Create the total row.
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = "Total";
                                cell.ApplyFormatting(totalRowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                            }
                            for(int j = 0; j < 5; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                    cell.ApplyFormatting(totalRowFormatting);
                                }
                            }
                        }
                    }
                    // Finalize the group creation.
                    sheet.EndGroup();

                    // Create the grand total row.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Grand total";
                            cell.ApplyFormatting(grandTotalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.4));
                        }
                        for(int j = 0; j < 5; j++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, 1, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(grandTotalRowFormatting);
                            }
                        }
                    }
                }
            }

            #endregion #Group/Outline
        }

        static void DataValidation(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create worksheet columns and set their widths.
                    using(IXlColumn column = sheet.CreateColumn()) {
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
                        string[] columnNames = new string[] { "Employee ID", "Employee name", "Salary", "Bonus", "Department" };
                        row.BulkCells(columnNames, headerRowFormatting);
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Departments";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate data for the document.
                    int[] id = new int[] {10115, 10709, 10401, 10204 };
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

                    #region #DataValidation
                    // Apply data validation to cells.
                    // Restrict data entry in the range A2:A5 to a 5-digit number.
                    XlDataValidation validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(0, 1, 0, 4));
                    validation.Type = XlDataValidationType.Custom;
                    validation.Criteria1 = "=AND(ISNUMBER(A2),LEN(A2)=5)";
                    // Add the specified rule to the worksheet collection of data validation rules.
                    sheet.DataValidations.Add(validation);

                    // Restrict data entry in the cell range C2:C5 to a whole number between 600 and 2000.
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(2, 1, 2, 4)); 
                    validation.Type = XlDataValidationType.Whole;
                    validation.Operator = XlDataValidationOperator.Between;
                    validation.Criteria1 = 600;
                    validation.Criteria2 = 2000;
                    // Display the error message.
                    validation.ErrorMessage = "The salary amount must be between 600$ and 2000$.";
                    validation.ErrorTitle = "Warning";
                    validation.ErrorStyle = XlDataValidationErrorStyle.Warning;
                    validation.ShowErrorMessage = true;
                    // Display the input message. 
                    validation.InputPrompt = "Please enter a whole number between 600 and 2000";
                    validation.PromptTitle = "Salary";
                    validation.ShowInputMessage = true;
                    // Add the specified rule to the worksheet collection of data validation rules.
                    sheet.DataValidations.Add(validation);

                    // Restrict data entry in the cell range D2:D5 to a decimal number within the specified limits. 
                    // Bonus cannot be greater than 10% of the salary.
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(3, 1, 3, 4));
                    validation.Type = XlDataValidationType.Decimal;
                    validation.Operator = XlDataValidationOperator.Between;
                    validation.Criteria1 = 0;
                    // Use a formula to specify the validation criterion.
                    validation.Criteria2 = "=C2*0.1";
                    // Display the error message.
                    validation.ErrorMessage = "Bonus cannot be greater than 10% of the salary.";
                    validation.ErrorTitle = "Information";
                    validation.ErrorStyle = XlDataValidationErrorStyle.Information;
                    validation.ShowErrorMessage = true;
                    // Add the specified rule to the worksheet collection of data validation rules.
                    sheet.DataValidations.Add(validation);

                    // Restrict data entry in the cell range E2:E5 to values in a drop-down list obtained from the cells G2:G5.
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(4, 1, 4, 4));
                    validation.Type = XlDataValidationType.List;
                    validation.Criteria1 = XlCellRange.FromLTRB(6, 1, 6, 4).AsAbsolute();
                    // Add the specified rule to the worksheet collection of data validation rules.
                    sheet.DataValidations.Add(validation);
                    #endregion #DataValidation
                }
            }
        }

    }
}
