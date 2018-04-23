Imports System
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl

Namespace XLExportExamples
    Public NotInheritable Class PictureActions

        Private Sub New()
        End Sub
        Private Shared imagesPath As String = Path.Combine(System.Windows.Forms.Application.StartupPath, "Images")

        #Region "Actions"
        Public Shared InsertPictureAction As Action(Of Stream, XlDocumentFormat) = AddressOf InsertPicture
        Public Shared StretchPictureAction As Action(Of Stream, XlDocumentFormat) = AddressOf StretchPicture
        Public Shared FitPictureAction As Action(Of Stream, XlDocumentFormat) = AddressOf FitPicture
        Public Shared PictureHyperlinkClickAction As Action(Of Stream, XlDocumentFormat) = AddressOf PictureHyperlinkClick
        #End Region

        Private Shared Sub InsertPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

'                #Region "#InsertPicture"
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()

                    ' Insert a picture from a file and anchor it to cells. 
                    Using picture As IXlPicture = sheet.CreatePicture()
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
                        picture.SetTwoCellAnchor(New XlAnchorPoint(1, 1, 0, 0), New XlAnchorPoint(6, 11, 2, 15), XlAnchorType.TwoCell)
                    End Using
                End Using
'                #End Region ' #InsertPicture
            End Using
        End Sub

        Private Shared Sub StretchPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture


                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    sheet.SkipColumns(1)
                    ' Create the column "B" and set its width.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 205
                    End Using
                    sheet.SkipRows(1)
                    ' Create the second row and set its height.
                    Using row As IXlRow = sheet.CreateRow()
                        row.HeightInPixels = 154
                    End Using
'                #Region "#StretchPicture"
                    ' Insert a picture from a file and stretch it to fill the cell B2.
                    Using picture As IXlPicture = sheet.CreatePicture()
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
                        picture.StretchToCell(New XlCellPosition(1, 1))
                    End Using
                End Using
'                #End Region ' #StretchPicture
            End Using
        End Sub

        Private Shared Sub FitPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture


                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    sheet.SkipColumns(1)
                    ' Create the column "B" and set its width.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 300
                    End Using
                    sheet.SkipRows(1)
                    ' Create the second row and set its height.
                    Using row As IXlRow = sheet.CreateRow()
                        row.HeightInPixels = 154
                    End Using
'                #Region "#FitPicture"
                    ' Insert a picture from a file to fit in the cell B2.
                    Using picture As IXlPicture = sheet.CreatePicture()
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
                        picture.FitToCell(New XlCellPosition(1, 1), 300, 154, True)
                    End Using
                End Using
'                #End Region ' #FitPicture
            End Using
        End Sub

        Private Shared Sub PictureHyperlinkClick(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'            #Region "#HyperlinkClick"
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()

                    ' Load a picture from a file and add a hyperlink to it.
                    Using picture As IXlPicture = sheet.CreatePicture()
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "DevExpress.png"))
                        picture.HyperlinkClick.TargetUri = "http://www.devexpress.com"
                        picture.HyperlinkClick.Tooltip = "Developer Express Inc."
                        picture.SetTwoCellAnchor(New XlAnchorPoint(1, 1, 0, 0), New XlAnchorPoint(10, 5, 2, 15), XlAnchorType.TwoCell)
                    End Using
                End Using
            End Using
'            #End Region ' #HyperlinkClick

        End Sub
    End Class
End Namespace
