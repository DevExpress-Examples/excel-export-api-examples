using DevExpress.Export.Xl;
using DevExpress.XtraExport.Csv;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;

namespace XLExportExamples {
    public static class CellFormattingActions {

        #region Actions
        public static Action<Stream, XlDocumentFormat> PredefinedFormattingAction = PredefinedFormatting;
        public static Action<Stream, XlDocumentFormat> ThemedFormattingAction = ThemedFormatting;
        public static Action<Stream, XlDocumentFormat> AlignmentAction = Alignment;
        public static Action<Stream, XlDocumentFormat> BordersAction = Borders;
        public static Action<Stream, XlDocumentFormat> FillAction = Fill;
        public static Action<Stream, XlDocumentFormat> FontAction = Font;
        public static Action<Stream, XlDocumentFormat> NumberFormatAction = NumberFormat;
        public static Action<Stream, XlDocumentFormat> RichTextFormattingAction = RichTextFormatting;
        #endregion

        static void RichTextFormatting(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #RichTextFormatting
                // Create a new worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create the first column and set its width.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 180;
                    }
                    // Create the first row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Create the cell A1.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Create an XlRichTextString instance.
                            XlRichTextString richText = new XlRichTextString();
                            // Add three text runs to the collection. 
                            richText.Runs.Add(new XlRichTextRun("Formatted ", XlFont.CustomFont("Arial", 14.0, XlColor.FromArgb(0x53, 0xbb, 0xf4))));
                            richText.Runs.Add(new XlRichTextRun("cell ", XlFont.CustomFont("Century Gothic", 14.0, XlColor.FromArgb(0xf1, 0x77, 0x00))));
                            richText.Runs.Add(new XlRichTextRun("text", XlFont.CustomFont("Consolas", 14.0, XlColor.FromArgb(0xe3, 0x2c, 0x2e))));
                            // Add the rich formatted text to the cell. 
                            cell.SetRichText(richText);
                        }
                    }
                }
                #endregion #RichTextFormatting
            }
        }

            static void PredefinedFormatting(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #PredefinedFormatting
                // Create a new worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create six successive columns and set their widths.
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }
                    
                    // Specify the "Good, Bad and Neutral" formatting category.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Good, Bad and Neutral";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        // Create a cell with the default "Normal" formatting.
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Normal";
                        }
                        // Create a cell and apply the "Bad" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Bad";
                            cell.Formatting = XlCellFormatting.Bad;
                        }
                        // Create a cell and apply the "Good" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Good";
                            cell.Formatting = XlCellFormatting.Good;
                        }
                        // Create a cell and apply the "Neutral" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Neutral";
                            cell.Formatting = XlCellFormatting.Neutral;
                        }
                    }

                    sheet.SkipRows(1);

                    // Specify the "Data and Model" formatting category.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Data and Model";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        // Create a cell and apply the "Calculation" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Calculation";
                            cell.Formatting = XlCellFormatting.Calculation;
                        }
                        // Create a cell and apply the "Check Cell" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Check Cell";
                            cell.Formatting = XlCellFormatting.CheckCell;
                        }
                        // Create a cell and apply the "Explanatory..." predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Explanatory";
                            cell.Formatting = XlCellFormatting.Explanatory;
                        }
                        // Create a cell and apply the "Input" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Input";
                            cell.Formatting = XlCellFormatting.Input;
                        }
                        // Create a cell and apply the "Linked Cell" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Linked Cell";
                            cell.Formatting = XlCellFormatting.LinkedCell;
                        }
                        // Create a cell and apply the "Note" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Note";
                            cell.Formatting = XlCellFormatting.Note;
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        // Create a cell and apply the "Output" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Output";
                            cell.Formatting = XlCellFormatting.Output;
                        }
                        // Create a cell and apply the "Warning Text" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Warning Text";
                            cell.Formatting = XlCellFormatting.WarningText;
                        }
                    }

                    sheet.SkipRows(1);

                    // Specify the "Titles and Headings" formatting category.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Titles and Headings";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        // Create a cell and apply the "Heading 1" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading 1";
                            cell.Formatting = XlCellFormatting.Heading1;
                        }
                        // Create a cell and apply the "Heading 2" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading 2";
                            cell.Formatting = XlCellFormatting.Heading2;
                        }
                        // Create a cell and apply the "Heading 3" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading 3";
                            cell.Formatting = XlCellFormatting.Heading3;
                        }
                        // Create a cell and apply the "Heading 4" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading 4";
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                        // Create a cell and apply the "Title" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Title";
                            cell.Formatting = XlCellFormatting.Title;
                        }
                        // Create a cell and apply the "Total" predefined formatting to it.
                        using (IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total";
                            cell.Formatting = XlCellFormatting.Total;
                        }
                    }
                }
            #endregion #PredefinedFormatting
            }
        }

        static void ThemedFormatting(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #ThemedFormatting
                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create six successive columns and set their widths.
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }
                    
                    // Specify an array that stores six accent colors of the document theme. 
                    XlThemeColor[] themeColors = new XlThemeColor[] { XlThemeColor.Accent1, XlThemeColor.Accent2, XlThemeColor.Accent3, XlThemeColor.Accent4, XlThemeColor.Accent5, XlThemeColor.Accent6 };

                    // Specify the "20% - AccentN" themed cell formatting.
                    // Create a worksheet row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            // Create a new cell in the row.
                            using(IXlCell cell = row.CreateCell()) {
                                // Set the cell value.
                                cell.Value = string.Format("Accent{0} 20%", i + 1);
                                // Apply the themed formatting to the cell using one of the predefined accent colors lightened by 80%.
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.8);
                            }
                        }
                    }

                    // Specify the "40% - AccentN" themed cell formatting.
                    // Create a worksheet row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            // Create a new cell in the row.
                            using(IXlCell cell = row.CreateCell()) {
                                // Set the cell value.
                                cell.Value = string.Format("Accent{0} 40%", i + 1);
                                // Apply the themed formatting to the cell using one of the predefined accent colors lightened by 60%.
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.6);
                            }
                        }
                    }

                    // Specify the "60% - AccentN" themed cell formatting.
                    // Create a worksheet row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            // Create a new cell in the row.
                            using(IXlCell cell = row.CreateCell()) {
                                // Set the cell value.
                                cell.Value = string.Format("Accent{0} 60%", i + 1);
                                // Apply the themed formatting to the cell using one of the predefined accent colors lightened by 40%.
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.4);
                            }
                        }
                    }

                    // Specify the "AccentN" themed cell formatting.
                    // Create a worksheet row.
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            // Create a new cell in the row.
                            using(IXlCell cell = row.CreateCell()) {
                                // Set the cell value.
                                cell.Value = string.Format("Accent{0}", i + 1);
                                // Apply the themed formatting to the cell using one of the predefined accent colors.
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.0);
                            }
                        }
                    }
                }
            #endregion #ThemedFormatting
            }
        }

        static void Alignment(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #Alignment
                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create three successive columns and set their widths.
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 130;
                        }
                    }

                    // Create the first row in the worksheet.
                    using(IXlRow row = sheet.CreateRow()) {
                        // Set the row height.
                        row.HeightInPixels = 40;
                        // Create the first cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Left and Top";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Top));
                        }
                        // Create the second cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Center and Top";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Top));
                        }
                        // Create the third cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Right and Top";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Top));
                        }
                    }

                    // Create the second row in the worksheet.
                    using(IXlRow row = sheet.CreateRow()) {
                        // Set the row height.
                        row.HeightInPixels = 40;
                        // Create the first cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Left and Center";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center));
                        }
                        // Create the second cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Center and Center";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                        }
                        // Create the third cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Right and Center";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center));
                        }
                    }

                    // Create the third row in the worksheet.
                    using(IXlRow row = sheet.CreateRow()) {
                        // Set the row height.
                        row.HeightInPixels = 40;
                        // Create the first cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Left and Bottom";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom));
                        }
                        // Create the second cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Center and Bottom";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Bottom));
                        }
                        // Create the third cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Right and Bottom";
                            // Specify the horizontal and vertical alignment of the cell content.
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                        }
                    }

                    sheet.SkipRows(1);
                    
                    // Create the fifth row in the worksheet.
                    using(IXlRow row = sheet.CreateRow()) {
                        // Create the first cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "The WrapText property is applied to wrap the text within a cell";
                            // Wrap the text within the cell.
                            cell.Formatting = new XlCellAlignment() { WrapText = true };
                        }
                        // Create the second cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Indented text";
                            // Set the indentation of the cell content.
                            cell.Formatting = new XlCellAlignment() { Indent = 2 };
                        }
                        // Create the third cell in the row.
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Rotated text";
                            // Rotate the text within the cell.
                            cell.Formatting = new XlCellAlignment() { TextRotation = 90 };
                        }
                    }
                }
            #endregion #Alignment
            }
        }

        static void Borders(Stream stream, XlDocumentFormat documentFormat) {
            #region #Borders
            // Specify a two-dimensional array that stores possible line styles for a border. 
            XlBorderLineStyle[,] lineStyles = new XlBorderLineStyle[,] {
                        { XlBorderLineStyle.Thin, XlBorderLineStyle.Medium, XlBorderLineStyle.Thick, XlBorderLineStyle.Double },
                        { XlBorderLineStyle.Dotted, XlBorderLineStyle.Dashed, XlBorderLineStyle.DashDot, XlBorderLineStyle.DashDotDot },
                        { XlBorderLineStyle.SlantDashDot, XlBorderLineStyle.MediumDashed, XlBorderLineStyle.MediumDashDot, XlBorderLineStyle.MediumDashDotDot }
                    };

            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {
                    for(int i = 0; i < 3; i++) {
                        sheet.SkipRows(1);
                        // Create a worksheet row.
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                row.SkipCells(1);
                                // Create a new cell in the row.
                                using(IXlCell cell = row.CreateCell()) {
                                    // Set outside borders for the created cell using a particular line style from the lineStyles array.
                                    cell.ApplyFormatting(XlBorder.OutlineBorders(Color.SeaGreen, lineStyles[i, j]));
                                }
                            }
                        }
                    }
                }
            }

            #endregion #Borders
        }

        static void Fill(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #Fill
                // Create a new worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {

                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Fill the cell background using the predefined color.
                            cell.ApplyFormatting(XlFill.SolidFill(Color.Beige));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Fill the cell background using the custom RGB color.
                            cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(0xff, 0x99, 0x66)));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Fill the cell background using the theme color.
                            cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent3, 0.4)));
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Specify the cell background pattern using predefined colors.
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkDown, Color.Red, Color.White));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Specify the cell background pattern using custom RGB colors.
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkTrellis, Color.FromArgb(0xff, 0xff, 0x66), Color.FromArgb(0x66, 0x99, 0xff)));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Specify the cell background pattern using theme colors.
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.LightHorizontal, XlColor.FromTheme(XlThemeColor.Accent1, 0.2), XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                        }
                    }
                }
                #endregion #Fill
            }
        }

        static void Font(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                #region #Font
                // Create a new worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {
                    // Create five successive columns and set their widths.
                    for (int i = 0; i < 5; i++)
                    {
                        using (IXlColumn column = sheet.CreateColumn())
                        {
                            column.WidthInPixels = 100;
                        }
                    }

                    // Create the first row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Create the cell A1.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Body font";
                            // Apply the theme body font to the cell content.
                            cell.ApplyFormatting(XlFont.BodyFont());
                        }

                        // Create the cell B1.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Headings font";
                            // Apply the theme heading font to the cell content.
                            cell.ApplyFormatting(XlFont.HeadingsFont());
                        }

                        // Create the cell C1.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Custom font";
                            // Specify the custom font attributes.
                            XlFont font = new XlFont();
                            font.Name = "Century Gothic";
                            font.SchemeStyle = XlFontSchemeStyles.None;
                            // Apply the custom font to the cell content.
                            cell.ApplyFormatting(font);
                        }
                    }

                    // Create an array that stores different values of font size.
                    int[] fontSizes = new int[] { 11, 14, 18, 24, 36 };
                    // Skip one row in the worksheet.
                    sheet.SkipRows(1);

                    // Create the third row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Create five successive cells (A3:E3) with different font sizes.
                        for (int i = 0; i < 5; i++)
                        {
                            using (IXlCell cell = row.CreateCell())
                            {
                                // Set the cell value that displays the applied font size.
                                cell.Value = string.Format("{0}pt", fontSizes[i]);
                                // Create a font instance of the specified size.
                                XlFont font = new XlFont();
                                font.Size = fontSizes[i];
                                // Apply font settings to the cell content.
                                cell.ApplyFormatting(font);
                            }
                        }
                    }

                    // Skip one row in the worksheet.
                    sheet.SkipRows(1);

                    // Create the fifth row.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        // Create the cell A5.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Red";
                            // Create a font instance and set its color.
                            XlFont font = new XlFont() { Color = Color.Red };
                            // Apply the font color to the cell content.
                            cell.ApplyFormatting(font);
                        }

                        // Create the cell B5. 
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Bold";
                            // Create a font instance and set its style to bold.
                            XlFont font = new XlFont() { Bold = true };
                            // Apply the font style to the cell content.
                            cell.ApplyFormatting(font);
                        }

                        // Create the cell C5. 
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Italic";
                            // Create a font instance and set its style to italic.
                            XlFont font = new XlFont() { Italic = true };
                            // Italicize the cell text.
                            cell.ApplyFormatting(font);
                        }

                        // Create the cell D5. 
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "Underline";
                            // Create a font instance and set the underline type to double.
                            XlFont font = new XlFont() { Underline = XlUnderlineType.Double };
                            // Underline the cell text.
                            cell.ApplyFormatting(font);
                        }

                        // Create the cell E5.
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Set the cell value.
                            cell.Value = "StrikeThrough";
                            // Create a font instance and turn the strikethrough formatting on.
                            XlFont font = new XlFont() { StrikeThrough = true };
                            // Strike the cell text through. 
                            cell.ApplyFormatting(font);
                        }
                    }
                }
                #endregion #Font
            }
        }

        static void NumberFormat(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Specify options for exporting the document to CSV format.
                CsvDataAwareExporterOptions csvOptions = document.Options as CsvDataAwareExporterOptions;
                if(csvOptions != null)
                    csvOptions.WritePreamble = true;

                // Create a worksheet.
                using(IXlSheet sheet = document.CreateSheet()) {
                    // Create six successive columns and set their widths.
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 180;
                        }
                    }

                    #region #ExcelNumberFormat
                    // Create the header row for the "Excel number formats" category.
                    using (IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = "Excel number formats";
                            // Apply the "Heading 4" predefined formatting to the cell.
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                    }
                    // Use the predefined Excel number formats to display data in cells.
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Predefined formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 123.456 as 123.46. 
                            cell.Value = 123.456;
                            cell.Formatting = XlNumberFormat.Number2;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 12345 as 12,345.
                            cell.Value = 12345;
                            cell.Formatting = XlNumberFormat.NumberWithThousandSeparator;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 0.33 as 33%.
                            cell.Value = 0.33;
                            cell.Formatting = XlNumberFormat.Percentage;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display the current date as "mm-dd-yy".  
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlNumberFormat.ShortDate;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display the current time as "h:mm AM/PM".
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlNumberFormat.ShortTime12;
                        }
                    }
                    // Use custom number formats to display data in cells.
                    using (IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Custom formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 4310.45 as $4,310.45.
                            cell.Value = 4310.45;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 3426.75 as €3,426.75.
                            cell.Value = 3426.75;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = @"_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "" - ""??_-;_-@_-";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 0.333 as 33.3%.
                            cell.Value = 0.333;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "0.0%";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Apply the custom number format to the date value.
                            // Display days as Sunday–Saturday, months as January–December, days as 1–31 and years as 1900–9999.
                            cell.Value = DateTime.Now;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "dddd, mmmm d, yyyy";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 0.6234 as 341/547.
                            cell.Value = 0.6234;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "# ???/???";
                        }
                        using (IXlCell cell = row.CreateCell())
                        {
                            // Display text value
                            cell.Value = "test";
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = XlNumberFormat.Text;
                        }

                    }
                    #endregion #ExcelNumberFormat

                    sheet.SkipRows(1);
                    #region #NETNumberFormat
                    // Create the header row for the ".NET number formats" category.
                    using (IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Set the cell value.
                            cell.Value = ".NET number formats";
                            // Apply the "Heading 4" predefined formatting to the cell.
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                    }
                    // Use the standard .NET-style format strings to display data in cells.
                    using (IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Standard formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 123.45 as 123.
                            cell.Value = 123.45;
                            cell.Formatting = XlCellFormatting.FromNetFormat("D", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 12345 as 1.234500E+004.
                            cell.Value = 12345;
                            cell.Formatting = XlCellFormatting.FromNetFormat("E", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 0.33 as 33.00%.
                            cell.Value = 0.33;
                            cell.Formatting = XlCellFormatting.FromNetFormat("P", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display the current date using the short date pattern.
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("d", true);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display the current time using the short time pattern.
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("t", true);
                        }
                    }
                    // Use custom format strings to display data in cells.
                    using (IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Custom formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 123.456 as 123.46. 
                            cell.Value = 123.45;
                            cell.Formatting = XlCellFormatting.FromNetFormat("#0.00", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 12345 as 1.235E+04.
                            cell.Value = 12345;
                            cell.Formatting = XlCellFormatting.FromNetFormat("0.0##e+00", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Display 0.333 as Max=33.3%.
                            cell.Value = 0.333;
                            cell.Formatting = XlCellFormatting.FromNetFormat("Max={0:#.0%}", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Apply the custom format string to the current date. 
                            // Display days as 01–31, months as 01-12 and years as a four-digit number. 
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("dd-MM-yyyy", true);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Apply the custom format string to the current time.
                            // Display hours as 01-12, minutes as 00-59, and add the AM/PM designator. 
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("hh:mm tt", true);
                        }
                    }
                    #endregion #NETNumberFormat
                }
            }
        }
    }
}
