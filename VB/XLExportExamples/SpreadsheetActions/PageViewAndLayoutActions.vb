Imports System
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl

Namespace XLExportExamples

    Public Module PageViewAndLayoutActions

'#Region "Actions"
        Public FreezeRowAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.FreezeRow

        Public FreezeColumnAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.FreezeColumn

        Public FreezePanesAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.FreezePanes

        Public SheetViewRTLAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.SheetViewRTL

        Public HeadersFootersAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.HeadersFooters

        Public PageBreaksAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PageBreaks

        Public PageMarginsAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PageMargins

        Public PageSetupAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PageSetup

        Public PrintAreaAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PrintArea

        Public PrintOptionsAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PrintOptions

        Public PrintTitlesAction As Action(Of Stream, XlDocumentFormat) = AddressOf XLExportExamples.PageViewAndLayoutActions.PrintTitles

'#End Region
        Private Sub FreezeRow(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#FreezeRow"
                    ' Freeze the first row in the worksheet.
                    sheet.SplitPosition = New XlCellPosition(0, 1)
'#End Region  ' #FreezeRow
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub FreezeColumn(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#FreezeColumn"
                    ' Freeze the first column in the worksheet.
                    sheet.SplitPosition = New XlCellPosition(1, 0)
'#End Region  ' #FreezeColumn
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub FreezePanes(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#FreezePanes"
                    ' Freeze the first column and the first row.
                    sheet.SplitPosition = New XlCellPosition(1, 1)
'#End Region  ' #FreezePanes
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub SheetViewRTL(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#RightToLeftSheetView"
                    ' Display the worksheet from right to left.
                    sheet.ViewOptions.RightToLeft = True
'#End Region  ' #RightToLeftSheetView
                    ' Freeze the first column and the first row.
                    sheet.SplitPosition = New XlCellPosition(1, 1)
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub HeadersFooters(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#HeaderAndFooters"
                    ' Specify different headers and footers for the odd-numbered and even-numbered pages.
                    sheet.HeaderFooter.DifferentOddEven = True
                    ' Add the bold text to the header left section, 
                    ' and insert the workbook name into the header right section.
                    sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.Bold & "Sample report", Nothing, XlHeaderFooter.BookName)
                    ' Insert the current page number into the footer right section. 
                    sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR(Nothing, Nothing, XlHeaderFooter.PageNumber)
                    ' Insert the workbook file path into the header left section, 
                    ' and add the worksheet name to the header right section. 
                    sheet.HeaderFooter.EvenHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.BookPath, Nothing, XlHeaderFooter.SheetName)
                    ' Insert the current page number into the footer left section 
                    ' and add the current date to the footer right section. 
                    sheet.HeaderFooter.EvenFooter = XlHeaderFooter.FromLCR(XlHeaderFooter.PageNumber, Nothing, XlHeaderFooter.Date)
'#End Region  ' #HeaderAndFooters
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub PageBreaks(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PageBreaks"
                    ' Insert a page break after the column "B".
                    sheet.ColumnPageBreaks.Add(2)
                    ' Insert a page break after the tenth row.
                    sheet.RowPageBreaks.Add(10)
'#End Region  ' #PageBreaks
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 100
                        column.Formatting = New XlCellFormatting()
                        column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                    End Using

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Sales"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub PageMargins(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PageMargins"
                    sheet.PageMargins = New XlPageMargins()
                    ' Set the unit of margin measurement.
                    sheet.PageMargins.PageUnits = XlPageUnits.Centimeters
                    ' Specify page margins.
                    sheet.PageMargins.Left = 2.0
                    sheet.PageMargins.Right = 1.0
                    sheet.PageMargins.Top = 1.25
                    sheet.PageMargins.Bottom = 1.25
                    ' Specify header and footer margins.
                    sheet.PageMargins.Header = 0.7
                    sheet.PageMargins.Footer = 0.7
'#End Region  ' #PageMargins
                    ' Generate data for the document.
                    sheet.SkipRows(1)
                    Using row As IXlRow = sheet.CreateRow()
                        row.SkipCells(1)
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Invoke the Page Setup dialog to control margin settings."
                        End Using
                    End Using
                End Using
            End Using
        End Sub

        Private Sub PageSetup(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PageSetup"
                    ' Specify page settings for the worksheet.
                    sheet.PageSetup = New XlPageSetup()
                    ' Select the paper size.
                    sheet.PageSetup.PaperKind = System.Drawing.Printing.PaperKind.A4
                    ' Set the page orientation to Landscape.
                    sheet.PageSetup.PageOrientation = XlPageOrientation.Landscape
                    '  Scale the print area to fit to one page wide.
                    sheet.PageSetup.FitToPage = True
                    sheet.PageSetup.FitToWidth = 1
                    sheet.PageSetup.FitToHeight = 0
                    '  Print in black and white.
                    sheet.PageSetup.BlackAndWhite = True
                    ' Specify the number of copies.
                    sheet.PageSetup.Copies = 2
'#End Region  ' #PageSetup
                    ' Generate data for the document.
                    sheet.SkipRows(1)
                    Using row As IXlRow = sheet.CreateRow()
                        row.SkipCells(1)
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Invoke the Page Setup dialog to control page settings."
                        End Using
                    End Using
                End Using
            End Using
        End Sub

        Private Sub PrintArea(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PrintArea"
                    ' Set the print area to cells A1:E5.
                    sheet.PrintArea = XlCellRange.FromLTRB(0, 0, 4, 4)
'#End Region  ' #PrintArea
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 110
                        column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom)
                    End Using

                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 190
                    End Using

                    For i As Integer = 0 To 2 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 90
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 130
                    End Using

                    sheet.SkipColumns(1)
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 130
                    End Using

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Employee ID"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Employee name"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Salary"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Bonus"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Department"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        row.SkipCells(1)
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Departments"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using
                    End Using

                    ' Generate data for the document.
                    Dim id As Integer() = New Integer() {10115, 10709, 10401, 10204}
                    Dim name As String() = New String() {"Augusta Delono", "Chris Cadwell", "Frank Diamond", "Simon Newman"}
                    Dim salary As Integer() = New Integer() {1100, 2000, 1750, 1250}
                    Dim bonus As Integer() = New Integer() {50, 180, 100, 80}
                    Dim deptid As Integer() = New Integer() {0, 2, 3, 3}
                    Dim department As String() = New String() {"Accounting", "IT", "Management", "Manufacturing"}
                    For i As Integer = 0 To 4 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = id(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = name(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = salary(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = bonus(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = department(deptid(i))
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            row.SkipCells(1)
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = department(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                        End Using
                    Next

                    ' Restrict data entry in the cell range E2:E5 to values in a drop-down list obtained from the cells G2:G5.
                    Dim validation As XlDataValidation = New XlDataValidation()
                    validation.Ranges.Add(XlCellRange.FromLTRB(4, 1, 4, 4))
                    validation.Type = XlDataValidationType.List
                    validation.Criteria1 = XlCellRange.FromLTRB(6, 1, 6, 4).AsAbsolute()
                    sheet.DataValidations.Add(validation)
                End Using
            End Using
        End Sub

        Private Sub PrintOptions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PrintOptions"
                    ' Specify print options for the worksheet.
                    sheet.PrintOptions = New XlPrintOptions()
                    ' Print row and column headings.
                    sheet.PrintOptions.Headings = True
                    ' Print gridlines.
                    sheet.PrintOptions.GridLines = True
                    ' Center worksheet data on a printed page.
                    sheet.PrintOptions.HorizontalCentered = True
                    sheet.PrintOptions.VerticalCentered = True
'#End Region  ' #PrintOptions
                    ' Generate data for the document.
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 100
                        column.Formatting = New XlCellFormatting()
                        column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                    End Using

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Sales"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                        End Using
                    Next
                End Using
            End Using
        End Sub

        Private Sub PrintTitles(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
'#Region "#PrintTitles"
                    ' Print the first column and the first row on every page.
                    sheet.PrintTitles.SetColumns(0, 0)
                    sheet.PrintTitles.SetRows(0, 0)
'#End Region  ' #PrintTitles
                    ' Generate data for the document.
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 250
                    End Using

                    For i As Integer = 0 To 4 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As XlCellFormatting = New XlCellFormatting()
                    rowFormatting.Font = New XlFont()
                    rowFormatting.Font.Name = "Century Gothic"
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                    headerRowFormatting.CopyFrom(rowFormatting)
                    headerRowFormatting.Font.Bold = True
                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                        End Using

                        For i As Integer = 0 To 4 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    ' Generate data rows.
                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"}
                    Dim random As Random = New Random()
                    For i As Integer = 0 To 12 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                            End Using

                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next
                        End Using
                    Next
                End Using
            End Using
        End Sub
    End Module
End Namespace
