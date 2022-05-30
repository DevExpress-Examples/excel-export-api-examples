using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;

namespace XLExportExamples
{
    public static class PictureActions
    {
        static string imagesPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Images");

        #region Actions
        public static Action<Stream, XlDocumentFormat> InsertPictureAction = InsertPicture;
        public static Action<Stream, XlDocumentFormat> StretchPictureAction = StretchPicture;
        public static Action<Stream, XlDocumentFormat> FitPictureAction = FitPicture;
        public static Action<Stream, XlDocumentFormat> PictureHyperlinkClickAction = PictureHyperlinkClick;
        #endregion

        static void InsertPicture(Stream stream, XlDocumentFormat documentFormat)
        {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #InsertPicture
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {

                    // Insert a picture from a file and anchor it to cells. 
                    using (IXlPicture picture = sheet.CreatePicture())
                    {
                        picture.SetImage(Image.FromFile(Path.Combine(imagesPath, "image1.jpg")), ImageFormat.Jpeg);
                        picture.SetTwoCellAnchor(new XlAnchorPoint(1, 1, 0, 0), new XlAnchorPoint(6, 11, 2, 15), XlAnchorType.TwoCell);
                    }
                }
                #endregion #InsertPicture
            }
        }

        static void StretchPicture(Stream stream, XlDocumentFormat documentFormat)
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
                    sheet.SkipColumns(1);
                    // Create the column "B" and set its width.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 205;
                    }
                    sheet.SkipRows(1);
                    // Create the second row and set its height.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.HeightInPixels = 154;
                    }
                #region #StretchPicture
                    // Insert a picture from a file and stretch it to fill the cell B2.
                    using (IXlPicture picture = sheet.CreatePicture())
                    {
                        picture.SetImage(Image.FromFile(Path.Combine(imagesPath, "image1.jpg")), ImageFormat.Jpeg);
                        picture.StretchToCell(new XlCellPosition(1, 1));
                    }
                }
                #endregion #StretchPicture
            }
        }

        static void FitPicture(Stream stream, XlDocumentFormat documentFormat)
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
                    sheet.SkipColumns(1);
                    // Create the column "B" and set its width.
                    using (IXlColumn column = sheet.CreateColumn())
                    {
                        column.WidthInPixels = 300;
                    }
                    sheet.SkipRows(1);
                    // Create the second row and set its height.
                    using (IXlRow row = sheet.CreateRow())
                    {
                        row.HeightInPixels = 154;
                    }
                #region #FitPicture
                    // Insert a picture from a file to fit in the cell B2.
                    using (IXlPicture picture = sheet.CreatePicture())
                    {
                        picture.SetImage(Image.FromFile(Path.Combine(imagesPath, "image1.jpg")), ImageFormat.Jpeg);
                        picture.FitToCell(new XlCellPosition(1, 1), 300, 154, true);
                    }
                }
                #endregion #FitPicture
            }
        }

        static void PictureHyperlinkClick(Stream stream, XlDocumentFormat documentFormat)
        {
            #region #HyperlinkClick
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using (IXlDocument document = exporter.CreateDocument(stream))
            {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet())
                {

                    // Load a picture from a file and add a hyperlink to it.
                    using (IXlPicture picture = sheet.CreatePicture())
                    {
                        picture.SetImage(Image.FromFile(Path.Combine(imagesPath, "DevExpress.png")), ImageFormat.Png);
                        picture.HyperlinkClick.TargetUri = "http://www.devexpress.com";
                        picture.HyperlinkClick.Tooltip = "Developer Express Inc.";
                        picture.SetTwoCellAnchor(new XlAnchorPoint(1, 1, 0, 0), new XlAnchorPoint(10, 5, 2, 15), XlAnchorType.TwoCell);
                    }
                }
            }
            #endregion #HyperlinkClick

        }
    }
}
