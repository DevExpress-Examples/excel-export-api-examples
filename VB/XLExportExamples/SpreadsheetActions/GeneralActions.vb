Imports System
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl

Namespace XLExportExamples
	Public NotInheritable Class GeneralActions

		Private Sub New()
		End Sub


		#Region "Actions"
		Public Shared CreateDocumentAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateDocument
		Public Shared CreateSheetAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateSheet
		Public Shared CreateHiddenSheetAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateHiddenSheet
		Public Shared HideGridlinesAction As Action(Of Stream, XlDocumentFormat) = AddressOf HideGridlines
		Public Shared HideHeadersAction As Action(Of Stream, XlDocumentFormat) = AddressOf HideHeaders
		Public Shared CreateColumnsAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateColumns
		Public Shared CreateRowsAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateRows
		Public Shared CreateCellsAction As Action(Of Stream, XlDocumentFormat) = AddressOf CreateCells
		Public Shared MergeCellsAction As Action(Of Stream, XlDocumentFormat) = AddressOf MergeCells
		#End Region

		Private Shared Sub CreateDocument(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateDocument"
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document and write it to the specified stream.
			Using document As IXlDocument = exporter.CreateDocument(stream)
				' Specify the document culture. 
				document.Options.Culture = CultureInfo.CurrentCulture
			End Using
'			#End Region ' #CreateDocument
		End Sub

		Private Shared Sub CreateSheet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
'			#Region "#CreateSheet"
			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create a new worksheet under the specified name. 
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Sales report"
				End Using
			End Using
'			#End Region ' #CreateSheet
		End Sub

		Private Shared Sub CreateHiddenSheet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture

'				#Region "#CreateHiddenSheet"
				' Create the first worksheet. 
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Sales report"
				End Using

				' Create the second worksheet and specify its visibility.
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Sales data"
					sheet.VisibleState = XlSheetVisibleState.Hidden
				End Using
'				#End Region ' #CreateHiddenSheet
			End Using
		End Sub

		Private Shared Sub HideHeaders(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture

'				#Region "#HideHeaders"
				' Create a worksheet. 
				Using sheet As IXlSheet = document.CreateSheet()
					' Hide row and column headers in the worksheet.
					sheet.ViewOptions.ShowRowColumnHeaders = False
				End Using
'				#End Region ' #HideHeaders
			End Using
		End Sub

		Private Shared Sub HideGridlines(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture

'				#Region "#HideGridlines"
				' Create a worksheet. 
				Using sheet As IXlSheet = document.CreateSheet()
					' Hide gridlines on the worksheet.
					sheet.ViewOptions.ShowGridLines = False
				End Using
'				#End Region ' #HideGridlines
			End Using
		End Sub

		Private Shared Sub CreateColumns(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture
'				#Region "#CreateColumns"
				' Create a worksheet.
				Using sheet As IXlSheet = document.CreateSheet()

					' Create the column A and set its width to 100 pixels.
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 100
					End Using

					' Hide the column B in the worksheet.
					Using column As IXlColumn = sheet.CreateColumn()
						column.IsHidden = True
					End Using

					' Create the column D and set its width to 24.5 characters.
					Using column As IXlColumn = sheet.CreateColumn(3)
						column.WidthInCharacters = 24.5F
					End Using
				End Using
'				#End Region ' #CreateColumns
			End Using
		End Sub

		Private Shared Sub CreateRows(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture
'				#Region "#CreateRows"
				' Create a worksheet.
				Using sheet As IXlSheet = document.CreateSheet()

					' Create the first row and set its height to 40 pixels.
					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 40
					End Using

					' Hide the third row in the worksheet.
					Using row As IXlRow = sheet.CreateRow(2)
						row.IsHidden = True
					End Using
				End Using
'				#End Region ' #CreateRows
			End Using
		End Sub

		Private Shared Sub CreateCells(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Specify the document culture.
				document.Options.Culture = CultureInfo.CurrentCulture
'				#Region "#CreateCells"
				' Create a worksheet.
				Using sheet As IXlSheet = document.CreateSheet()
					' Create the column A and set its width. 
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 150
					End Using

					' Create the first row.
					Using row As IXlRow = sheet.CreateRow()

						' Create the cell A1 and set its value.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Numeric value:"
						End Using

						' Create the cell B1 and assign the numeric value to it.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 123.45
						End Using
					End Using

					' Create the second row.
					Using row As IXlRow = sheet.CreateRow()

						' Create the cell A2 and set its value.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Text value:"
						End Using

						' Create the cell B2 and assign the text value to it.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "abc"
						End Using
					End Using

					' Create the third row.
					Using row As IXlRow = sheet.CreateRow()

						' Create the cell A3 and set its value.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Boolean value:"
						End Using

						' Create the cell B3 and assign the boolean value to it.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = True
						End Using
					End Using

					' Create the fourth row.
					Using row As IXlRow = sheet.CreateRow()

						' Create the cell A4 and set its value.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Error value:"
						End Using

						' Create the cell B4 and assign an error value to it.
						Using cell As IXlCell = row.CreateCell()
							cell.Value = XlVariantValue.ErrorName
						End Using
					End Using
				End Using
'				#End Region ' #CreateCells
			End Using
		End Sub

		Private Shared Sub MergeCells(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
			' Create an exporter instance.
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create a new document.
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create a worksheet.
				Using sheet As IXlSheet = document.CreateSheet()
					' Create the first row in the worksheet.
					Using row As IXlRow = sheet.CreateRow()
						' Create a cell.
						Using cell As IXlCell = row.CreateCell()
							' Set the cell value.
							cell.Value = "Merged cells A1 to E1"
							' Align the cell content.
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
						End Using
					End Using

					' Create the second row in the worksheet.
					Using row As IXlRow = sheet.CreateRow()
						' Create a cell.
						Using cell As IXlCell = row.CreateCell()
							' Set the cell value.
							cell.Value = "Merged cells A2 to A5"
							' Align the cell content.
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
							' Wrap the text within the cell.
							cell.Formatting.Alignment.WrapText = True
						End Using
						' Create a cell.
						Using cell As IXlCell = row.CreateCell()
							' Set the cell value.
							cell.Value = "Merged cells B2 to E5"
							' Align the cell content.
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
						End Using
					End Using

'					#Region "#MergeCells"
					' Merge cells contained in the range A1:E1.
					sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 0, 4, 0))

					' Merge cells contained in the range A2:A5.
					sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 1, 0, 4))

					' Merge cells contained in the range B2:E5.
					sheet.MergedCells.Add(XlCellRange.FromLTRB(1, 1, 4, 4))
'					#End Region ' #MergeCells
				End Using
			End Using
		End Sub

	End Class
End Namespace
