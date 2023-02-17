Imports System
Imports DevExpress.Export.Xl
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO

Namespace XLExportExamples

    Public Module TableActions

'#Region "Actions"
        Public AddTableAction As Action(Of Stream, XlDocumentFormat) = AddressOf AddTable

        Public DisableFilteringAction As Action(Of Stream, XlDocumentFormat) = AddressOf DisableFiltering

        Public HiddenHeaderRowAction As Action(Of Stream, XlDocumentFormat) = AddressOf HiddenHeaderRow

        Public HiddenTotalRowAction As Action(Of Stream, XlDocumentFormat) = AddressOf HiddenTotalRow

        Public SideBySideAction As Action(Of Stream, XlDocumentFormat) = AddressOf SideBySide

        Public TableStyleAction As Action(Of Stream, XlDocumentFormat) = AddressOf TableStyle

        Public TableStyleOptionsAction As Action(Of Stream, XlDocumentFormat) = AddressOf TableStyleOptions

        Public CustomFormattingAction As Action(Of Stream, XlDocumentFormat) = AddressOf CustomFormatting

        Public CalculatedColumnAction As Action(Of Stream, XlDocumentFormat) = AddressOf CalculatedColumn

'#End Region
        Private Sub AddTable(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#AddTable"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, True)
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                        ' Specify the number format for the "Amount" column and its total cell.
                        Dim accounting As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        table.Columns(2).DataFormatting = accounting
                        table.Columns(2).TotalRowFormatting = accounting
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                    End Using

                    ' Create the total row and finish the table.
                    Using row As IXlRow = sheet.CreateRow()
                        row.EndTable(table, True)
                    End Using
'#End Region  ' #AddTable
                End Using
            End Using
        End Sub

        Private Sub DisableFiltering(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#DisableFiltering"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, True)
                        ' Disable the filtering functionality for the table. 
                        table.HasAutoFilter = False
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                    End Using

                    ' Create the total row and finish the table.
                    Using row As IXlRow = sheet.CreateRow()
                        row.EndTable(table, True)
                    End Using
'#End Region  ' #DisableFiltering
                End Using
            End Using
        End Sub

        Private Sub HiddenHeaderRow(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#HiddenHeaderRow"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the table with the hidden header row.
                        table = row.BeginTable(columnNames, False)
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                        ' Populate the first table row with data.
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    ' Generate the remaining table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                    End Using

                    ' Create the total row and finish the table.
                    Using row As IXlRow = sheet.CreateRow()
                        row.EndTable(table, True)
                    End Using
'#End Region  ' #HiddenHeaderRow
                End Using
            End Using
        End Sub

        Private Sub HiddenTotalRow(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#HiddenTotalRow"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Start generating the table with a header row displayed.
                    Using row As IXlRow = sheet.CreateRow()
                        table = row.BeginTable(columnNames, True)
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using
'#End Region  ' #HiddenTotalRow
                End Using
            End Using
        End Sub

        Private Sub SideBySide(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns (A:G) and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100, 20, 100, 120, 100}
                    For i As Integer = 0 To 7 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#SideBySideTables"
                    ' Specify two arrays containing column headings for the tables.
                    Dim columnNames1 As String() = New String() {"Product", "Category", "Amount"}
                    Dim columnNames2 As String() = New String() {"Region", "Category", "Amount"}
                    ' Create the first row in the worksheet from which the tables start.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the first table with a header row displayed.
                        Dim table As IXlTable = row.BeginTable(columnNames1, True)
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                        ' Specify the number format for the "Amount" column and its total cell.
                        Dim accounting As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        table.Columns(2).DataFormatting = accounting
                        table.Columns(2).TotalRowFormatting = accounting
                        row.SkipCells(1)
                        ' Start generating the second table with a header row displayed.
                        table = row.BeginTable(columnNames2, True)
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                        ' Specify the number format for the "Amount" column and its total cell.
                        table.Columns(2).DataFormatting = accounting
                        table.Columns(2).TotalRowFormatting = accounting
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                        row.SkipCells(1)
                        row.BulkCells(New Object() {"East", "Dairy Products", 16000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                        row.SkipCells(1)
                        row.BulkCells(New Object() {"East", "Grains/Cereals", 14500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15500}, Nothing)
                        row.SkipCells(1)
                        row.BulkCells(New Object() {"West", "Dairy Products", 16500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.SkipCells(1)
                        row.BulkCells(New Object() {"West", "Grains/Cereals", 13500}, Nothing)
                    End Using

                    ' Create the total rows and finish the tables.
                    Using row As IXlRow = sheet.CreateRow()
                        For Each table As IXlTable In sheet.Tables
                            row.EndTable(table, True)
                        Next
                    End Using
'#End Region  ' #SideBySideTables
                End Using
            End Using
        End Sub

        Private Sub TableStyle(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#TableStyle"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, True)
                        ' Apply the table style.
                        table.Style.Name = XlBuiltInTableStyleId.Dark7
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using
'#End Region  ' #TableStyle
                End Using
            End Using
        End Sub

        Private Sub TableStyleOptions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#TableStyleOptions"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for tables.
                    Dim columnNames As String() = New String() {"Product", "Category", "Amount"}
                    ' Create the row containing the table title.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Disable banded rows"}, XlCellFormatting.Title)
                    End Using

                    sheet.SkipRows(1)
                    ' Start generating the table with a header row displayed.
                    Using row As IXlRow = sheet.CreateRow()
                        table = row.BeginTable(columnNames, True)
                        ' Disable banded row formatting for the table.
                        table.Style.ShowRowStripes = False
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using

                    sheet.SkipRows(1)
                    ' Create the row containing the table title.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Enable banded columns"}, XlCellFormatting.Title)
                    End Using

                    sheet.SkipRows(1)
                    ' Start generating the table with a header row displayed.
                    Using row As IXlRow = sheet.CreateRow()
                        table = row.BeginTable(columnNames, True)
                        ' Apply banded column formatting to the table.
                        table.Style.ShowRowStripes = False
                        table.Style.ShowColumnStripes = True
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using

                    sheet.SkipRows(1)
                    ' Create the row containing the table title.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Highlight first column"}, XlCellFormatting.Title)
                    End Using

                    sheet.SkipRows(1)
                    ' Start generating the table with a header row displayed.
                    Using row As IXlRow = sheet.CreateRow()
                        table = row.BeginTable(columnNames, True)
                        ' Display special formatting for the first column of the table.
                        table.Style.ShowFirstColumn = True
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using

                    sheet.SkipRows(1)
                    ' Create the row containing the table title.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Highlight last column"}, XlCellFormatting.Title)
                    End Using

                    sheet.SkipRows(1)
                    ' Start generating the table with a header row displayed.
                    Using row As IXlRow = sheet.CreateRow()
                        table = row.BeginTable(columnNames, True)
                        ' Display special formatting for the last column of the table.
                        table.Style.ShowLastColumn = True
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    ' Create the last table row and finish the table.
                    ' The total row is not displayed for the table. 
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                        row.EndTable(table, False)
                    End Using
'#End Region  ' #TableStyleOptions
                End Using
            End Using
        End Sub

        Private Sub CustomFormatting(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create columns "A", "B" and "C" and set their widths.
                    Dim widths As Integer() = New Integer() {165, 120, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#CustomFormatting"
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        Dim accounting As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        ' Create objects containing information about table columns (their names and formatting).
                        Dim columns As List(Of XlTableColumnInfo) = New List(Of XlTableColumnInfo)()
                        columns.Add(New XlTableColumnInfo("Product"))
                        columns.Add(New XlTableColumnInfo("Category"))
                        columns.Add(New XlTableColumnInfo("Amount"))
                        ' Specify formatting settings for the last column of the table.
                        columns(2).HeaderRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent6, -0.3))
                        columns(2).DataFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.9))
                        columns(2).DataFormatting.NumberFormat = accounting
                        columns(2).TotalRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.8))
                        columns(2).TotalRowFormatting.NumberFormat = accounting
                        ' Specify formatting settings for the header row of the table.
                        Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                        headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent6, 0.0))
                        headerRowFormatting.Border = New XlBorder()
                        headerRowFormatting.Border.BottomColor = XlColor.FromArgb(0, 0, 0)
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Dashed
                        ' Start generating the table with a header row displayed.
                        Dim table As IXlTable = row.BeginTable(columns, True, headerRowFormatting)
                        ' Apply the table style.
                        table.Style.Name = XlBuiltInTableStyleId.Medium16
                        ' Disable banded row formatting for the table.
                        table.Style.ShowRowStripes = False
                        ' Disable the filtering functionality for the table. 
                        table.HasAutoFilter = False
                        ' Specify formatting settings for the total row of the table.
                        table.TotalRowFormatting = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Dark1, 0.9))
                        table.TotalRowFormatting.Border = New XlBorder() With {.BottomColor = XlColor.FromTheme(XlThemeColor.Accent6, 0.0), .BottomLineStyle = XlBorderLineStyle.Thick, .TopColor = XlColor.FromArgb(0, 0, 0), .TopLineStyle = XlBorderLineStyle.Dashed}
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(2).TotalRowFunction = XlTotalRowFunction.Sum
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", "Dairy Products", 17000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", "Grains/Cereals", 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", "Dairy Products", 15000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", "Grains/Cereals", 12500}, Nothing)
                    End Using

                    ' Create the total row and finish the table.
                    Using row As IXlRow = sheet.CreateRow()
                        row.EndTable(sheet.Tables(0), True)
                    End Using
'#End Region  ' #CustomFormatting
                End Using
            End Using
        End Sub

        Private Sub CalculatedColumn(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Create worksheet columns (A:F) and set their widths.
                    Dim widths As Integer() = New Integer() {165, 100, 100, 100, 100, 110}
                    For i As Integer = 0 To 6 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = widths(i)
                        End Using
                    Next

'#Region "#CalculatedColumn"
                    Dim table As IXlTable
                    ' Specify an array containing column headings for a table.
                    Dim columnNames As String() = New String() {"Product", "Q1", "Q2", "Q3", "Q4", "Yearly Total"}
                    ' Create the first row in the worksheet from which the table starts.
                    Using row As IXlRow = sheet.CreateRow()
                        ' Start generating the table with a header row displayed.
                        table = row.BeginTable(columnNames, True)
                        ' Specify the total row label.
                        table.Columns(0).TotalRowLabel = "Total"
                        ' Specify the function to calculate the total.
                        table.Columns(5).TotalRowFunction = XlTotalRowFunction.Sum
                        ' Specify the number format for numeric values in the table and the total cell of the "Yearly Total" column.
                        Dim accounting As XlNumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                        table.DataFormatting = accounting
                        table.Columns(5).TotalRowFormatting = accounting
                        ' Set the formula to calculate annual sales of each product
                        ' and display results in the "Yearly Total" column.
                        table.Columns(5).SetFormula(XlFunc.Sum(table.GetRowReference("Q1", "Q4")))
                    End Using

                    ' Generate table rows and populate them with data.
                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Camembert Pierrot", 17000, 18500, 17500, 18000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Gnocchi di nonna Alice", 15500, 14500, 15000, 14000}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Mascarpone Fabioli", 15000, 15750, 16000, 15500}, Nothing)
                    End Using

                    Using row As IXlRow = sheet.CreateRow()
                        row.BulkCells(New Object() {"Ravioli Angelo", 12500, 11000, 13500, 12000}, Nothing)
                    End Using

                    ' Create the total row and finish the table.
                    Using row As IXlRow = sheet.CreateRow()
                        row.EndTable(table, True)
                    End Using
'#End Region  ' #CalculatedColumn
                End Using
            End Using
        End Sub
    End Module
End Namespace
