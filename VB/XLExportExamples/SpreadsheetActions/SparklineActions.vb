Imports System
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl

Namespace XLExportExamples
    Friend Class SparklineActions
        #Region "Actions"
        Public Shared AddSparklineGroupAction As Action(Of Stream, XlDocumentFormat) = AddressOf AddSparklineGroup
        Public Shared AddSparklineAction As Action(Of Stream, XlDocumentFormat) = AddressOf AddSparkline
        Public Shared AdjustScalingAction As Action(Of Stream, XlDocumentFormat) = AddressOf AdjustScaling
        Public Shared HighlightValuesAction As Action(Of Stream, XlDocumentFormat) = AddressOf HighlightValues
        Public Shared DisplayXAxisAction As Action(Of Stream, XlDocumentFormat) = AddressOf DisplayXAxis
        Public Shared SetDateRangeAction As Action(Of Stream, XlDocumentFormat) = AddressOf SetDateRange
        #End Region
        Private Shared Sub AddSparklineGroup(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    For i As Integer = 0 To 4
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))

                    Dim columnNames() As String = { "Product", "Q1", "Q2", "Q3", "Q4" }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(columnNames, headerRowFormatting)
                    End Using

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 3
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                        End Using
                    Next product

'                    #Region "#AddSparklineGroup"
                    ' Create a group of line sparklines.
                    Dim group As New XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 4, 6), XlCellRange.FromLTRB(5, 1, 5, 6))
                    ' Set the sparkline weight.
                    group.LineWeight = 1.25
                    ' Display data markers on the sparklines.
                    group.DisplayMarkers = True
                    sheet.SparklineGroups.Add(group)
'                    #End Region ' #AddSparklineGroup
                End Using
            End Using
        End Sub

        Private Shared Sub AddSparkline(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    For i As Integer = 0 To 4
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))

                    Dim columnNames() As String = { "Product", "Q1", "Q2", "Q3", "Q4" }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(columnNames, headerRowFormatting)
                    End Using

                    ' Create a group of line sparklines.
                    Dim group As New XlSparklineGroup()
                    ' Set the sparkline color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, -0.2)
                    ' Set the sparkline weight.
                    group.LineWeight = 1.25
                    sheet.SparklineGroups.Add(group)

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 3
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j

'                            #Region "#AddSparkline"
                            ' Add one more sparkline to the existing group.
                            Dim rowIndex As Integer = row.RowIndex
                            group.Sparklines.Add(New XlSparkline(XlCellRange.FromLTRB(1, rowIndex, 4, rowIndex), XlCellRange.FromLTRB(5, rowIndex, 5, rowIndex)))
'                            #End Region ' #AddSparkline
                        End Using
                    Next product
                End Using
            End Using
        End Sub

        Private Shared Sub AdjustScaling(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    For i As Integer = 0 To 4
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))

                    Dim columnNames() As String = { "Product", "Q1", "Q2", "Q3", "Q4" }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(columnNames, headerRowFormatting)
                    End Using

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 3
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 1500)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                        End Using
                    Next product

'                    #Region "#AdjustScaling"
                    ' Create a sparkline group.
                    Dim group As New XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 4, 6), XlCellRange.FromLTRB(5, 1, 5, 6))
                    ' Set the sparkline color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0)
                    ' Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column
                    ' Set the custom minimum value for the vertical axis.
                    group.MinScaling = XlSparklineAxisScaling.Custom
                    group.ManualMin = 1000.0
                    ' Set the automatic maximum value for all sparklines in the group. 
                    group.MaxScaling = XlSparklineAxisScaling.Group
                    sheet.SparklineGroups.Add(group)
'                    #End Region ' #AdjustScaling

                End Using
            End Using
        End Sub

        Private Shared Sub HighlightValues(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    For i As Integer = 0 To 8
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = rowFormatting.Clone()
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))

                    Dim columnNames() As String = { "State", "Q1'13", "Q2'13", "Q3'13", "Q4'13", "Q1'14", "Q2'14", "Q3'14", "Q4'14" }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(columnNames, headerRowFormatting)
                    End Using

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 7
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round((random.NextDouble() + 0.5) * 2000 * Math.Sign(random.NextDouble() - 0.4))
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                        End Using
                    Next product

'                    #Region "#HighlightValues"
                    ' Create a sparkline group.                   
                    Dim group As New XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 8, 6), XlCellRange.FromLTRB(9, 1, 9, 6))
                    ' Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column
                    ' Set the series color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0)
                    ' Set the color for negative points on sparklines. 
                    group.ColorNegative = XlColor.FromTheme(XlThemeColor.Accent2, 0.0)
                    ' Set the color for the highest points on sparklines.
                    group.ColorHigh = XlColor.FromTheme(XlThemeColor.Accent6, 0.0)
                    ' Highlight the highest and negative points on each sparkline in the group.
                    group.HighlightNegative = True
                    group.HighlightHighest = True
                    sheet.SparklineGroups.Add(group)
'                    #End Region ' #HighlightValues
                End Using
            End Using
        End Sub

        Private Shared Sub DisplayXAxis(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using
                    For i As Integer = 0 To 8
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As XlCellFormatting = rowFormatting.Clone()
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))

                    Dim columnNames() As String = { "State", "Q1'13", "Q2'13", "Q3'13", "Q4'13", "Q1'14", "Q2'14", "Q3'14", "Q4'14" }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(columnNames, headerRowFormatting)
                    End Using

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "Alabama", "Arizona", "California", "Colorado", "Connecticut", "Florida" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 7
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round((random.NextDouble() + 0.5) * 2000 * Math.Sign(random.NextDouble() - 0.4))
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                        End Using
                    Next product

'                    #Region "#DisplayXAxis"
                    ' Create a sparkline group.                   
                    Dim group As New XlSparklineGroup(XlCellRange.FromLTRB(1, 1, 8, 6), XlCellRange.FromLTRB(9, 1, 9, 6))
                    ' Change the sparkline group type to "Column".
                    group.SparklineType = XlSparklineType.Column
                    ' Display the horizontal axis.
                    group.DisplayXAxis = True
                    ' Set the series color.
                    group.ColorSeries = XlColor.FromTheme(XlThemeColor.Accent1, 0.0)
                    ' Highlight negative points on each sparkline in the group.
                    group.ColorNegative = XlColor.FromTheme(XlThemeColor.Accent2, 0.0)
                    group.HighlightNegative = True
                    sheet.SparklineGroups.Add(group)
'                    #End Region ' #DisplayXAxis
                End Using
            End Using
        End Sub

        Private Shared Sub SetDateRange(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture

                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns and set their widths.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 200
                    End Using

                    For i As Integer = 0 To 4
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                            column.ApplyFormatting(CType("_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)", XlNumberFormat))
                        End Using
                    Next i

                    ' Specify formatting settings for cells containing data.
                    Dim rowFormatting As New XlCellFormatting()
                    rowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)

                    ' Specify formatting settings for the header row.
                    Dim headerRowFormatting As New XlCellFormatting()
                    headerRowFormatting.Font = XlFont.CustomFont("Century Gothic", 12.0)
                    headerRowFormatting.Font.Bold = True
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0))
                    headerRowFormatting.NumberFormat = XlNumberFormat.ShortDate

                    Dim headerValues() As Object = { "Product", New Date(2015, 10, 1), New Date(2015, 10, 10), New Date(2015, 10, 15), New Date(2015, 10, 25) }

                    ' Generate the header row.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(headerValues, headerRowFormatting)
                    End Using

                    ' Generate data for the document.
                    Dim random As New Random()
                    Dim products() As String = { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD" }

                    For Each product As String In products
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = product
                                cell.ApplyFormatting(rowFormatting)
                            End Using
                            For j As Integer = 0 To 3
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
                                    cell.ApplyFormatting(rowFormatting)
                                End Using
                            Next j
                        End Using
                    Next product
'                    #Region "#SetDateRange"
                    ' Create a group of line sparklines.                    
                    Dim group As New XlSparklineGroup(XlCellRange.Parse("B2:E7"), XlCellRange.Parse("F2:F7"))
                    ' Specify the date range for the sparkline group. 
                    group.DateRange = XlCellRange.Parse("B1:E1")
                    ' Set the sparkline weight.
                    group.LineWeight = 1.25
                    ' Display data markers on the sparklines.
                    group.DisplayMarkers = True
                    sheet.SparklineGroups.Add(group)
'                    #End Region ' #SetDateRange
                End Using
            End Using
        End Sub
    End Class
End Namespace
