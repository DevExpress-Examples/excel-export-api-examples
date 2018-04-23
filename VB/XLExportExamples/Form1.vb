Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics
Imports DevExpress.Export.Xl

Namespace XLExportExamples
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            InitTreeListControl()
        End Sub

        Private Sub InitTreeListControl()
            Dim examples As New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub
        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'            #Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Basic Actions"))
            examples.Add(New SpreadsheetNode("Cell Formatting"))
            examples.Add(New SpreadsheetNode("Conditional Formatting"))
            examples.Add(New SpreadsheetNode("Data Actions"))
            examples.Add(New SpreadsheetNode("Formula Actions"))
            examples.Add(New SpreadsheetNode("Page View and Layout"))
            examples.Add(New SpreadsheetNode("Pictures"))
            examples.Add(New SpreadsheetNode("Sparklines"))
            examples.Add(New SpreadsheetNode("Miscellaneous"))
'            #End Region

'            #Region "ExampleNodes"
            ' Add nodes to the "Basic Actions" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Create Document", GeneralActions.CreateDocumentAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Worksheet", GeneralActions.CreateSheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Worksheet", GeneralActions.CreateHiddenSheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Gridlines", GeneralActions.HideGridlinesAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Row and Column Headers", GeneralActions.HideHeadersAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Columns", GeneralActions.CreateColumnsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Rows", GeneralActions.CreateRowsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Cells", GeneralActions.CreateCellsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Merge Cells", GeneralActions.MergeCellsAction))

            ' Add nodes to the "Cell Formatting" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Predefined Style-Like Formatting", CellFormattingActions.PredefinedFormattingAction))
            examples(1).Groups.Add(New SpreadsheetExample("Themed Formatting", CellFormattingActions.ThemedFormattingAction))
            examples(1).Groups.Add(New SpreadsheetExample("Alignment", CellFormattingActions.AlignmentAction))
            examples(1).Groups.Add(New SpreadsheetExample("Borders", CellFormattingActions.BordersAction))
            examples(1).Groups.Add(New SpreadsheetExample("Fill", CellFormattingActions.FillAction))
            examples(1).Groups.Add(New SpreadsheetExample("Font", CellFormattingActions.FontAction))
            examples(1).Groups.Add(New SpreadsheetExample("Number Format", CellFormattingActions.NumberFormatAction))
            examples(1).Groups.Add(New SpreadsheetExample("Rich Text Formatting", CellFormattingActions.RichTextFormattingAction))

            ' Add nodes to the "Conditional Formatting" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Less Than/Greater Than/Between Rules", ConditionalFormattingActions.CellIsAction))
            examples(2).Groups.Add(New SpreadsheetExample("Text that Contains... Rule", ConditionalFormattingActions.SpecificTextAction))
            examples(2).Groups.Add(New SpreadsheetExample("A Date Occurring... Rule", ConditionalFormattingActions.TimePeriodAction))
            examples(2).Groups.Add(New SpreadsheetExample("Duplicate Values", ConditionalFormattingActions.DuplicatesAction))
            examples(2).Groups.Add(New SpreadsheetExample("Blank/Non-Blank Cells", ConditionalFormattingActions.BlanksAction))
            examples(2).Groups.Add(New SpreadsheetExample("Formula-Based Rules", ConditionalFormattingActions.ExpressionAction))
            examples(2).Groups.Add(New SpreadsheetExample("Top/Bottom Rules", ConditionalFormattingActions.Top10Action))
            examples(2).Groups.Add(New SpreadsheetExample("Above/Below Average Rules", ConditionalFormattingActions.AverageAction))
            examples(2).Groups.Add(New SpreadsheetExample("Color Scales", ConditionalFormattingActions.ColorScaleAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Bars", ConditionalFormattingActions.DataBarAction))
            examples(2).Groups.Add(New SpreadsheetExample("Icon Sets", ConditionalFormattingActions.IconSetAction))

            ' Add nodes to the "Data Actions" group of examples.
            examples(3).Groups.Add(New SpreadsheetExample("Enable Filtering", DataActions.AutoFilterAction))
            examples(3).Groups.Add(New SpreadsheetExample("Outline Data", DataActions.OutlineGroupingAction))
            examples(3).Groups.Add(New SpreadsheetExample("Data Validation", DataActions.DataValidationAction))

            ' Add nodes to the "Formula Actions" group of examples.
            examples(4).Groups.Add(New SpreadsheetExample("Formulas", FormulaActions.FormulasAction))
            examples(4).Groups.Add(New SpreadsheetExample("Shared Formulas", FormulaActions.SharedFormulasAction))
            examples(4).Groups.Add(New SpreadsheetExample("Functions", FormulaActions.FunctionsAction))

            ' Add nodes to the "Page View and Layout" group of examples.
            examples(5).Groups.Add(New SpreadsheetExample("Freeze Row", PageViewAndLayoutActions.FreezeRowAction))
            examples(5).Groups.Add(New SpreadsheetExample("Freeze Column", PageViewAndLayoutActions.FreezeColumnAction))
            examples(5).Groups.Add(New SpreadsheetExample("Freeze Panes", PageViewAndLayoutActions.FreezePanesAction))
            examples(5).Groups.Add(New SpreadsheetExample("Right-To-Left View", PageViewAndLayoutActions.SheetViewRTLAction))
            examples(5).Groups.Add(New SpreadsheetExample("Headers and Footers", PageViewAndLayoutActions.HeadersFootersAction))
            examples(5).Groups.Add(New SpreadsheetExample("Page Breaks", PageViewAndLayoutActions.PageBreaksAction))
            examples(5).Groups.Add(New SpreadsheetExample("Page Margins", PageViewAndLayoutActions.PageMarginsAction))
            examples(5).Groups.Add(New SpreadsheetExample("Page Setup", PageViewAndLayoutActions.PageSetupAction))
            examples(5).Groups.Add(New SpreadsheetExample("Print Area", PageViewAndLayoutActions.PrintAreaAction))
            examples(5).Groups.Add(New SpreadsheetExample("Print Options", PageViewAndLayoutActions.PrintOptionsAction))
            examples(5).Groups.Add(New SpreadsheetExample("Print Titles", PageViewAndLayoutActions.PrintTitlesAction))

            ' Add nodes to the "Pictures" group of examples.
            examples(6).Groups.Add(New SpreadsheetExample("Insert Picture", PictureActions.InsertPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Stretch Picture", PictureActions.StretchPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Fit Picture in Cell", PictureActions.FitPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Picture Hyperlink", PictureActions.PictureHyperlinkClickAction))

            ' Add nodes to the "Sparklines" group of examples.
            examples(7).Groups.Add(New SpreadsheetExample("Add Sparkline Group", SparklineActions.AddSparklineGroupAction))
            examples(7).Groups.Add(New SpreadsheetExample("Add Sparkline to Group", SparklineActions.AddSparklineAction))
            examples(7).Groups.Add(New SpreadsheetExample("Adjust Scaling", SparklineActions.AdjustScalingAction))
            examples(7).Groups.Add(New SpreadsheetExample("Highlight Values", SparklineActions.HighlightValuesAction))
            examples(7).Groups.Add(New SpreadsheetExample("Display X-axis", SparklineActions.DisplayXAxisAction))
            examples(7).Groups.Add(New SpreadsheetExample("Set Date Range", SparklineActions.SetDateRangeAction))


            ' Add nodes to the "Miscellaneous" group of examples.
            examples(8).Groups.Add(New SpreadsheetExample("Insert Hyperlinks", MiscellaneousActions.HyperlinksAction))
            examples(8).Groups.Add(New SpreadsheetExample("Document Properties", MiscellaneousActions.DocumentPropertiesAction))
            examples(8).Groups.Add(New SpreadsheetExample("Document Options and Restrictions", MiscellaneousActions.DocumentOptionsAction))
            examples(8).Groups.Add(New SpreadsheetExample("CSV Export Options", MiscellaneousActions.CsvExportOptionsAction))

'            #End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub


         Private Sub RunExample(ByVal filePath As String, ByVal documentFormat As XlDocumentFormat)
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then
                Return
            End If
            Using stream As New FileStream(filePath, FileMode.Create)
                Dim action As Action(Of Stream, XlDocumentFormat) = example.Action
                action(stream, documentFormat)
            End Using
            Process.Start(filePath)
         End Sub

        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLSX.Click
            If treeList1.FocusedNode Is Nothing Then
                Return
            End If
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            RunExample(fileName, XlDocumentFormat.Xlsx)
        End Sub

        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToXLS.Click
            If treeList1.FocusedNode Is Nothing Then
                Return
            End If
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            RunExample(fileName, XlDocumentFormat.Xls)
        End Sub

        Private Sub btnExportToCSV_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExportToCSV.Click
            If treeList1.FocusedNode Is Nothing Then
                Return
            End If
            Dim fileName As String = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv")
            If String.IsNullOrEmpty(fileName) Then
                Return
            End If
            RunExample(fileName, XlDocumentFormat.Csv)
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaulName As String) As String
            Dim sfDialog As New SaveFileDialog()
            sfDialog.Filter = filter
            sfDialog.FileName = defaulName
            If sfDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                Return Nothing
            End If
            Return sfDialog.FileName
        End Function

        Private Sub EnableButtons(ByVal exampleName As String)
            btnExportToCSV.Enabled = Not XLExportDisabledCSVExamples.Examples.Contains(exampleName)
            btnExportToXLS.Enabled = Not XLExportForbiddenXLSExamples.Examples.Contains(exampleName)
        End Sub

        Private Sub treeList1_FocusedNodeChanged(ByVal sender As Object, ByVal e As DevExpress.XtraTreeList.FocusedNodeChangedEventArgs) Handles treeList1.FocusedNodeChanged
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then
                Return
            End If
            EnableButtons(example.Name)
        End Sub
    End Class
End Namespace
