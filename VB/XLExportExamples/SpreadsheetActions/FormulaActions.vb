Imports System
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
    Public NotInheritable Class FormulaActions

        Private Sub New()
        End Sub


        #Region "Actions"
        Public Shared FormulasAction As Action(Of Stream, XlDocumentFormat) = AddressOf Formulas
        Public Shared SharedFormulasAction As Action(Of Stream, XlDocumentFormat) = AddressOf SharedFormulas
        Public Shared FunctionsAction As Action(Of Stream, XlDocumentFormat) = AddressOf Functions
        #End Region

        Private Shared Sub Formulas(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

'                #Region "#Formulas"
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()

                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 50
                    End Using
                    For i As Integer = 0 To 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 80
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next i

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.BodyFont()
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))
                    ' Specify formatting settings for the total row.
                    Dim totalRowFormatting As New XlCellFormatting()
                    totalRowFormatting.Font = XlFont.BodyFont()
                    totalRowFormatting.Font.Bold = True

                    ' Generate data for the document.
                    Dim header() As String = { "Description", "QTY", "Price", "Amount" }
                    Dim product() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" }
                    Dim qty() As Integer = { 12, 15, 25, 10 }
                    Dim price() As Double = { 23.25, 15.50, 12.99, 8.95 }

                    ' Create the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        For i As Integer = 0 To 3
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = header(i)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next i
                    End Using

                    ' Create data rows.
                    For i As Integer = 0 To 3
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = qty(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = price(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                ' Set the formula to calculate the amount per product.
                                cell.SetFormula(String.Format("B{0}*C{0}", i + 2))
                            End Using
                        End Using
                    Next i

                    ' Create the total row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.SkipCells(2)
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Total:"
                            cell.ApplyFormatting(totalRowFormatting)
                        End Using
                        Using cell As IXlCell = row.CreateCell()
                            ' Set the formula to calculate the total amount.
                            cell.SetFormula("SUM(D2:D5)")
                            cell.ApplyFormatting(totalRowFormatting)
                        End Using
                    End Using
                End Using
'                #End Region ' #Formulas
            End Using
        End Sub

        Private Shared Sub SharedFormulas(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

'                #Region "#SharedFormulas"
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()

                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 50
                    End Using
                    For i As Integer = 0 To 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 80
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next i

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.BodyFont()
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))
                    ' Specify formatting settings for the total row.
                    Dim totalRowFormatting As New XlCellFormatting()
                    totalRowFormatting.Font = XlFont.BodyFont()
                    totalRowFormatting.Font.Bold = True

                    ' Generate data for the document.
                    Dim header() As String = { "Description", "QTY", "Price", "Amount" }
                    Dim product() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" }
                    Dim qty() As Integer = { 12, 15, 25, 10 }
                    Dim price() As Double = { 23.25, 15.50, 12.99, 8.95 }

                    ' Create the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        For i As Integer = 0 To 3
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = header(i)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next i
                    End Using

                    ' Create data rows.
                    For i As Integer = 0 To 3
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = qty(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = price(i)
                            End Using
                            Using cell As IXlCell = row.CreateCell()
                                ' Use the shared formula to calculate the amount per product. 
                                If i = 0 Then
                                    cell.SetSharedFormula("B2*C2", XlCellRange.FromLTRB(3, 1, 3, 4))
                                Else
                                    cell.SetSharedFormula(New XlCellPosition(3, 1))
                                End If
                            End Using
                        End Using
                    Next i

                    ' Create the total row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.SkipCells(2)
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Total:"
                            cell.ApplyFormatting(totalRowFormatting)
                        End Using
                        Using cell As IXlCell = row.CreateCell()
                            ' Set the formula to calculate the total amount.
                            cell.SetFormula("SUM(D2:D5)")
                            cell.ApplyFormatting(totalRowFormatting)
                        End Using
                    End Using
                End Using
'                #End Region ' #SharedFormulas
            End Using
        End Sub

        Private Shared Sub Functions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

'                #Region "#Functions"
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()

                    ' Create the column "A" and set its width.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    ' Create five successive columns and set the specific number format for their cells.
                    For i As Integer = 0 To 4
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.Formatting = New XlCellFormatting()
                            column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.BodyFont()
                    rowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0))
                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.BodyFont()
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))
                    ' Specify formatting settings for the total row.
                    Dim totalRowFormatting As New XlCellFormatting()
                    totalRowFormatting.Font = XlFont.BodyFont()
                    totalRowFormatting.Font.Bold = True
                    totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0))

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" }

                    ' Create the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Product"
                            cell.ApplyFormatting(headerRowFormatting)
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))
                        End Using
                        For i As Integer = 0 To 3
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = String.Format("Q{0}", i + 1)
                                cell.ApplyFormatting(headerRowFormatting)
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom))
                            End Using
                        Next i
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Yearly total"
                            cell.ApplyFormatting(headerRowFormatting)
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom))
                        End Using
                    End Using

                    ' Create data rows.
                    For i As Integer = 0 To 3
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                                cell.ApplyFormatting(rowFormatting)
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8))
                            End Using
                            For j As Integer = 0 To 3
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                            Using cell As IXlCell = row.CreateCell()
                                ' Use the SUM function to calculate annual sales for each product.   
                                cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)))
                                cell.ApplyFormatting(rowFormatting)
                                cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)))
                            End Using
                        End Using
                    Next i

                    ' Create the total row.
                    Using row As IXlRow = sheet.CreateRow()
                        Using cell As IXlCell = row.CreateCell()
                            cell.Value = "Total"
                            cell.ApplyFormatting(totalRowFormatting)
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6))
                        End Using
                        For j As Integer = 0 To 4
                            Using cell As IXlCell = row.CreateCell()
                                ' Use the SUBTOTAL function to calculate total sales for each quarter and the entire year.  
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, False))
                                cell.ApplyFormatting(totalRowFormatting)
                            End Using
                        Next j
                    End Using
                End Using
'                #End Region ' #Functions
            End Using
        End Sub

    End Class
End Namespace
