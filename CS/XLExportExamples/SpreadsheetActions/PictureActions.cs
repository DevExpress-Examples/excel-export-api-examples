using System;
using System.Drawing;
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
        #endregion

        static void InsertPicture(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #InsertPicture
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {

                    // Insert a picture from a file and anchor it to cells. 
                    using (IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.SetTwoCellAnchor(new XlAnchorPoint(1, 1, 0, 0), new XlAnchorPoint(6, 11, 2, 15), XlAnchorType.TwoCell);
                    }
                }
                #endregion #InsertPicture
            }
        }

        static void StretchPicture(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #StretchPicture
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {
                    sheet.SkipColumns(1);
                    // Create the column "B" and set its width.
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 205;
                    }
                    sheet.SkipRows(1);
                    // Create the second row and set its height.
                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 154;
                    }

                    // Insert a picture from a file and stretch it to fill the cell B2.
                    using (IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.StretchToCell(new XlCellPosition(1, 1));
                    }
                }
                #endregion #StretchPicture
            }
        }

        static void FitPicture(Stream stream, XlDocumentFormat documentFormat) {
            // Create an exporter instance.
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create a new document.
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                #region #FitPicture
                // Create a worksheet.
                using (IXlSheet sheet = document.CreateSheet()) {
                    sheet.SkipColumns(1);
                    // Create the column "B" and set its width.
                    using (IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 300;
                    }
                    sheet.SkipRows(1);
                    // Create the second row and set its height.
                    using (IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 154;
                    }

                    // Insert a picture form a file to fit in the cell B2.
                    using (IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.FitToCell(new XlCellPosition(1, 1), 300, 154, true);
                    }
                }
                #endregion #FitPicture
            }
        }

    }
}
