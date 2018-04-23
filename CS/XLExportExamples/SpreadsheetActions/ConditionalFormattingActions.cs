using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples
{
    public static class ConditionalFormattingActions {
        #region Actions
        public static Action<Stream, XlDocumentFormat> AverageAction = Average;
        public static Action<Stream, XlDocumentFormat> CellIsAction = CellIs;
        public static Action<Stream, XlDocumentFormat> BlanksAction = Blanks;
        public static Action<Stream, XlDocumentFormat> DuplicatesAction = Duplicates;
        public static Action<Stream, XlDocumentFormat> ExpressionAction = Expression;
        public static Action<Stream, XlDocumentFormat> SpecificTextAction = SpecificText;
        public static Action<Stream, XlDocumentFormat> TimePeriodAction = TimePeriod;
        public static Action<Stream, XlDocumentFormat> Top10Action = Top10;
        public static Action<Stream, XlDocumentFormat> DataBarAction = DataBar;
        public static Action<Stream, XlDocumentFormat> IconSetAction = IconSet;
        public static Action<Stream, XlDocumentFormat> ColorScaleAction = ColorScale;

        #endregion

        static void Average(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = i + 1;
                                }
                            }
                        }
                    }

                    #region #AverageRule
                    // Create an instance of the XlConditionalFormatting class.
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10));
                    // Create the rule highlighting values that are above the average in the cell range.
                    XlCondFmtRuleAboveAverage rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.Above;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class.
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10));
                    // Create the rule highlighting values that are above or equal to the average value in the cell range.
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.AboveOrEqual;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class.
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10));
                    // Create the rule highlighting values that are below the average in the cell range.
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.Below;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class.
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10));
                    // Create the rule highlighting values that are below or equal to the average value in the cell range.
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.BelowOrEqual;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #AverageRule
                }
            }
        }

        static void CellIs(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = i + 1;
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = 12 - i;
                            }
                        }
                    }

                    #region #CellIsRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10));
                    // Create the rule to highlight cells whose values are less than 5.
                    XlCondFmtRuleCellIs rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.LessThan;
                    rule.Value = 5;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Create the rule to highlight cells whose values are between 5 and 8.
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.Between;
                    rule.Value = 5;
                    rule.SecondValue = 8;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral;
                    formatting.Rules.Add(rule);
                    // Create the rule to highlight cells whose values are greater than 8.
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.GreaterThan;
                    rule.Value = 8;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10));
                    // Create the rule to highlight cells whose values are greater than a value calculated by a formula. 
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.GreaterThan;
                    rule.Value = "=$A1+3";
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #CellIsRule
                }
            }
        }

        static void Blanks(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {

                    // Generate data for the document.
                    for (int i = 0; i < 10; i++)
                    {
                        using (IXlRow row = sheet.CreateRow())
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                if ((i % 2) == 0)
                                    cell.Value = i + 1;
                            }
                        }
                    }
                    #region #BlanksRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A1:A10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9));
                    // Create the rule to highlight blank cells in the range.
                    XlCondFmtRuleBlanks rule = new XlCondFmtRuleBlanks(true);
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Create the rule to highlight non-blank cells in the range.
                    rule = new XlCondFmtRuleBlanks(false);
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #BlanksRule
                }
            }
        }

        static void Duplicates(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.ColumnIndex * cell.RowIndex + cell.RowIndex + 1;
                                }
                            }
                        }
                    }

                    #region #DuplicatesRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 10));
                    // Create the rule to identify duplicate values in the cell range.
                    formatting.Rules.Add(new XlCondFmtRuleDuplicates() { Formatting = XlCellFormatting.Bad });
                    // Create the rule to identify unique values in the cell range.
                    formatting.Rules.Add(new XlCondFmtRuleUnique() { Formatting = XlCellFormatting.Good });
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #DuplicatesRule
                }
            }
        }

        static void Expression(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    int[] width = new int[] { 80, 150, 90 };
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = width[i];
                            if(i == 2) {
                                column.Formatting = new XlCellFormatting();
                                column.Formatting.NumberFormat = "[$$-409] #,##0.00";
                            }
                        }
                    }
                    string[] columnNames = new string[] { "Account ID", "User Name", "Balance" };
                    using(IXlRow row = sheet.CreateRow()) {
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.Font = XlFont.BodyFont();
                        headerRowFormatting.Font.Bold = true;
                        headerRowFormatting.Border = new XlBorder();
                        headerRowFormatting.Border.BottomColor = Color.Black;
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin;
                        for(int i = 0; i < 3; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = columnNames[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }
                    string[] accountIds = new string[] { "A105", "A114", "B013", "C231", "D101", "D105" };
                    string[] users = new string[] { "Berry Dafoe", "Chris Cadwell", "Esta Mangold", "Liam Bell", "Simon Newman", "Wendy Underwood" };
                    int[] balance = new int[] { 155, 250, 48, 350, -15, 10 };
                    for(int i = 0; i < 6; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = accountIds[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = users[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = balance[i];
                            }
                        }
                    }

                    #region #ExpressionRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A2:C7).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 1, 2, 6));
                    // Create the rule that uses a formula to highlight cells if a value in the column "C" is greater than 0 and less than 50. 
                    XlCondFmtRuleExpression rule = new XlCondFmtRuleExpression("AND($C2>0,$C2<50)");
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlFill.SolidFill(Color.FromArgb(0xff, 0xff, 0xcc));
                    formatting.Rules.Add(rule);
                    // Create the rule that uses a formula to highlight cells if a value in the column "C" is less than or equal to 0. 
                    rule = new XlCondFmtRuleExpression("$C2<=0");
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #ExpressionRule
                }
            }
        }

        static void SpecificText(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    int[] width = new int[] { 250, 180, 100 };
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = width[i];
                            if(i == 2) {
                                column.Formatting = new XlCellFormatting();
                                column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                            }
                        }
                    }
                    string[] columnNames = new string[] { "Product", "Delivery", "Sales" };
                    using(IXlRow row = sheet.CreateRow()) {
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.Font = XlFont.BodyFont();
                        headerRowFormatting.Font.Bold = true;
                        headerRowFormatting.Border = new XlBorder();
                        headerRowFormatting.Border.BottomColor = Color.Black;
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin;
                        for(int i = 0; i < 3; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = columnNames[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Queso Cabrales", "Raclette Courdavault" };
                    string[] deliveries = new string[] { "USA", "Worldwide", "USA", "Ships worldwide", "Worldwide except EU", "EU" };
                    int[] sales = new int[] { 15500, 20250, 12634, 35010, 15234, 10050 };
                    for(int i = 0; i < 6; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = deliveries[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = sales[i];
                            }
                        }
                    }

                    #region #SpecificTextRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (B2:B7).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 1, 1, 6));
                    // Create the rule to highlight cells that contain the given text.
                    XlCondFmtRuleSpecificText rule = new XlCondFmtRuleSpecificText(XlCondFmtSpecificTextType.Contains, "worldwide");
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #SpecificTextRule
                }
            }
        }

        static void TimePeriod(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.ApplyFormatting(XlNumberFormat.ShortDate);
                    }
                    for(int i = 0; i < 10; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = DateTime.Now.AddDays(row.RowIndex - 5);
                            }
                        }
                    }

                    #region #TimePeriodRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A1:A10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9));
                    // Create the rule to highlight yesterday's dates in the cell range.
                    XlCondFmtRuleTimePeriod rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Yesterday;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Create the rule to highlight today's dates in the cell range.
                    rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Today;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Create the rule to highlight tomorrows's dates in the cell range.
                    rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Tomorrow;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #TimePeriodRule
                }
            }
        }

        static void Top10(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 10; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.ColumnIndex * 4 + cell.RowIndex + 1;
                                }
                            }
                        }
                    }
                    #region #TopAndBottomRules
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rules should be applied (A1:D10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 9));
                    // Create the rule to identify bottom 10 values in the cell range.
                    XlCondFmtRuleTop10 rule = new XlCondFmtRuleTop10();
                    rule.Bottom = true;
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Create the rule to identify top 10 values in the cell range.
                    rule = new XlCondFmtRuleTop10();
                    // Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #TopAndBottomRules
                }
            }
        }

        static void DataBar(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 3; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    int rowIndex = cell.RowIndex;
                                    int columnIndex = cell.ColumnIndex;
                                    if(columnIndex == 0)
                                        cell.Value = rowIndex + 1;
                                    else if(columnIndex == 1)
                                        cell.Value = rowIndex - 5;
                                    else
                                        cell.Value = (rowIndex < 5) ? rowIndex + 1 : 11 - rowIndex;
                                }
                            }
                        }
                    }

                    #region #DataBarRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10));
                    // Create the rule to compare values in the cell range using data bars.
                    XlCondFmtRuleDataBar rule = new XlCondFmtRuleDataBar();
                    // Specify the bar color.
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.2);
                    // Specify the solid fill type.
                    rule.GradientFill = false;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10));
                    // Create the rule to compare values in the cell range using data bars.
                    rule = new XlCondFmtRuleDataBar();
                    // Set the positive bar color to green.
                    rule.FillColor = Color.Green;
                    // Set the border color of positive bars to green.
                    rule.BorderColor = Color.Green;
                    // Set the axis color to brown.
                    rule.AxisColor = Color.Brown;
                    // Use the gradient fill type
                    rule.GradientFill = true;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10));
                    // Create the rule to compare values in the cell range using data bars.
                    rule = new XlCondFmtRuleDataBar();
                    // Specify the bar color.
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.2);
                    // Set the minimum length of the data bar.
                    rule.MinLength = 10;
                    // Set the maximum length of the data bar.
                    rule.MaxLength = 90;
                    // Set the value corresponding to the shortest bar.
                    rule.MinValue.ObjectType = XlCondFmtValueObjectType.Number;
                    rule.MinValue.Value = 3;
                    // Set the direction of data bars.
                    rule.Direction = XlDataBarDirection.RightToLeft;
                    // Hide values of cells to which the rule is applied.
                    rule.ShowValues = false;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #DataBarRule
                }
            }
        }

        static void IconSet(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    if(cell.ColumnIndex % 2 == 0)
                                        cell.Value = cell.RowIndex + 1;
                                    else
                                        cell.Value = cell.RowIndex - 5;
                                }
                            }
                        }
                    }

                    #region #IconSetRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10));
                    // Create the rule to apply a specific icon from the "3 Arrows" icon set to each cell in the range based on its value. 
                    XlCondFmtRuleIconSet rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Arrows3;
                    // Set the rule priority.
                    rule.Priority = 1;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10));
                    // Create the rule to apply a specific icon from the "3 Flags" icon set to each cell in the range based on its value. 
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Flags3;
                    // Set the rule priority.
                    rule.Priority = 2;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10));
                    // Create the rule to apply a specific icon from the "5 Ratings" icon set to each cell in the range based on its value. 
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Rating5;
                    // Hide values of cells to which the rule is applied.
                    rule.ShowValues = false;
                    // Set the rule priority.
                    rule.Priority = 3;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify the cell range to which the conditional formatting rule should be applied (D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10));
                    // Create the rule to apply a specific icon from the "4 Traffic Lights" icon set to each cell in the range based on its value. 
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.TrafficLights4;
                    // Reverse the icon order.
                    rule.Reverse = true;
                    // Set the rule priority.
                    rule.Priority = 4;
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #IconSetRule
                }
            }
        }

        static void ColorScale(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Generate data for the document.
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.RowIndex + 1;
                                }
                            }
                        }
                    }

                    #region #ColorScaleRule
                    // Create an instance of the XlConditionalFormatting class. 
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    // Specify cell ranges to which the conditional formatting rule should be applied (A1:A11 and C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); 
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10));
                    // Create the default three-color scale rule to differentiate low, medium and high values in cell ranges.
                    XlCondFmtRuleColorScale rule = new XlCondFmtRuleColorScale();
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);

                    // Create an instance of the XlConditionalFormatting class. 
                    formatting = new XlConditionalFormatting();
                    // Specify cell ranges to which the conditional formatting rule should be applied (B1:B11 and D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10));
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10));
                    // Create the two-color scale rule to differentiate low and high values in cell ranges. 
                    rule = new XlCondFmtRuleColorScale();
                    rule.ColorScaleType = XlCondFmtColorScaleType.ColorScale2;
                    // Set a color corresponding to the minimum value in the cell range.
                    rule.MinColor = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    // Set a color corresponding to the maximum value in the cell range.
                    rule.MaxColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.5);
                    formatting.Rules.Add(rule);
                    // Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting);
                    #endregion #ColorScaleRule
                }
            }
        }

    }
}
