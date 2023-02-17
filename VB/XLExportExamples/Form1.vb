Imports System
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics
Imports DevExpress.Export.Xl

Namespace XLExportExamples

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            InitTreeListControl()
        End Sub

        Private Sub InitTreeListControl()
            Dim examples As GroupsOfSpreadsheetExamples = New GroupsOfSpreadsheetExamples()
            InitData(examples)
            DataBinding(examples)
        End Sub

        Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'#Region "GroupNodes"
            examples.Add(New SpreadsheetNode("Basic Actions"))
            examples.Add(New SpreadsheetNode("Cell Formatting"))
            examples.Add(New SpreadsheetNode("Conditional Formatting"))
            examples.Add(New SpreadsheetNode("Data Actions"))
            examples.Add(New SpreadsheetNode("Formula Actions"))
            examples.Add(New SpreadsheetNode("Page View and Layout"))
            examples.Add(New SpreadsheetNode("Pictures"))
            examples.Add(New SpreadsheetNode("Sparklines"))
            examples.Add(New SpreadsheetNode("Miscellaneous"))
            examples.Add(New SpreadsheetNode("Tables"))
'#End Region
'#Region "ExampleNodes"
            ' Add nodes to the "Basic Actions" group of examples.
            examples(0).Groups.Add(New SpreadsheetExample("Create Document", CreateDocumentAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Worksheet", CreateSheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Worksheet", CreateHiddenSheetAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Gridlines", HideGridlinesAction))
            examples(0).Groups.Add(New SpreadsheetExample("Hide Row and Column Headers", HideHeadersAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Columns", CreateColumnsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Rows", CreateRowsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Create Cells", CreateCellsAction))
            examples(0).Groups.Add(New SpreadsheetExample("Merge Cells", MergeCellsAction))
            ' Add nodes to the "Cell Formatting" group of examples.
            examples(1).Groups.Add(New SpreadsheetExample("Predefined Style-Like Formatting", PredefinedFormattingAction))
            examples(1).Groups.Add(New SpreadsheetExample("Themed Formatting", ThemedFormattingAction))
            examples(1).Groups.Add(New SpreadsheetExample("Alignment", AlignmentAction))
            examples(1).Groups.Add(New SpreadsheetExample("Borders", BordersAction))
            examples(1).Groups.Add(New SpreadsheetExample("Fill", FillAction))
            examples(1).Groups.Add(New SpreadsheetExample("Font", FontAction))
            examples(1).Groups.Add(New SpreadsheetExample("Number Format", NumberFormatAction))
            examples(1).Groups.Add(New SpreadsheetExample("Rich Text Formatting", RichTextFormattingAction))
            ' Add nodes to the "Conditional Formatting" group of examples.
            examples(2).Groups.Add(New SpreadsheetExample("Less Than/Greater Than/Between Rules", CellIsAction))
            examples(2).Groups.Add(New SpreadsheetExample("Text that Contains... Rule", SpecificTextAction))
            examples(2).Groups.Add(New SpreadsheetExample("A Date Occurring... Rule", TimePeriodAction))
            examples(2).Groups.Add(New SpreadsheetExample("Duplicate Values", DuplicatesAction))
            examples(2).Groups.Add(New SpreadsheetExample("Blank/Non-Blank Cells", BlanksAction))
            examples(2).Groups.Add(New SpreadsheetExample("Formula-Based Rules", ExpressionAction))
            examples(2).Groups.Add(New SpreadsheetExample("Top/Bottom Rules", Top10Action))
            examples(2).Groups.Add(New SpreadsheetExample("Above/Below Average Rules", AverageAction))
            examples(2).Groups.Add(New SpreadsheetExample("Color Scales", ColorScaleAction))
            examples(2).Groups.Add(New SpreadsheetExample("Data Bars", DataBarAction))
            examples(2).Groups.Add(New SpreadsheetExample("Icon Sets", IconSetAction))
            ' Add nodes to the "Data Actions" group of examples.
            examples(3).Groups.Add(New SpreadsheetExample("Enable Filtering", AutoFilterAction))
            examples(3).Groups.Add(New SpreadsheetExample("Outline Data", OutlineGroupingAction))
            examples(3).Groups.Add(New SpreadsheetExample("Data Validation", DataValidationAction))
            ' Add nodes to the "Formula Actions" group of examples.
            examples(4).Groups.Add(New SpreadsheetExample("Simple Formulas", SimpleFormulasAction))
            examples(4).Groups.Add(New SpreadsheetExample("Complex Formulas", ComplexFormulasAction))
            examples(4).Groups.Add(New SpreadsheetExample("Shared Formulas", SharedFormulasAction))
            examples(4).Groups.Add(New SpreadsheetExample("Subtotals", SubtotalsAction))
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
            examples(6).Groups.Add(New SpreadsheetExample("Insert Picture", InsertPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Stretch Picture", StretchPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Fit Picture in Cell", FitPictureAction))
            examples(6).Groups.Add(New SpreadsheetExample("Picture Hyperlink", PictureHyperlinkClickAction))
            ' Add nodes to the "Sparklines" group of examples.
            examples(7).Groups.Add(New SpreadsheetExample("Add Sparkline Group", SparklineActions.AddSparklineGroupAction))
            examples(7).Groups.Add(New SpreadsheetExample("Add Sparkline to Group", SparklineActions.AddSparklineAction))
            examples(7).Groups.Add(New SpreadsheetExample("Adjust Scaling", SparklineActions.AdjustScalingAction))
            examples(7).Groups.Add(New SpreadsheetExample("Highlight Values", SparklineActions.HighlightValuesAction))
            examples(7).Groups.Add(New SpreadsheetExample("Display X-axis", SparklineActions.DisplayXAxisAction))
            examples(7).Groups.Add(New SpreadsheetExample("Set Date Range", SparklineActions.SetDateRangeAction))
            ' Add nodes to the "Miscellaneous" group of examples.
            examples(8).Groups.Add(New SpreadsheetExample("Insert Hyperlinks", HyperlinksAction))
            examples(8).Groups.Add(New SpreadsheetExample("Document Properties", DocumentPropertiesAction))
            examples(8).Groups.Add(New SpreadsheetExample("Document Options and Restrictions", DocumentOptionsAction))
            examples(8).Groups.Add(New SpreadsheetExample("CSV Export Options", CsvExportOptionsAction))
            ' Add nodes to the "Tables" group of examples.
            examples(9).Groups.Add(New SpreadsheetExample("Create Table", AddTableAction))
            examples(9).Groups.Add(New SpreadsheetExample("Disable Filtering", DisableFilteringAction))
            examples(9).Groups.Add(New SpreadsheetExample("Hidden Header Row", HiddenHeaderRowAction))
            examples(9).Groups.Add(New SpreadsheetExample("Hidden Total Row", HiddenTotalRowAction))
            examples(9).Groups.Add(New SpreadsheetExample("Side-By-Side Tables", SideBySideAction))
            examples(9).Groups.Add(New SpreadsheetExample("Table Style", TableStyleAction))
            examples(9).Groups.Add(New SpreadsheetExample("Table Style Options", TableStyleOptionsAction))
            examples(9).Groups.Add(New SpreadsheetExample("Custom Formatting", CustomFormattingAction))
            examples(9).Groups.Add(New SpreadsheetExample("Calculated Column", CalculatedColumnAction))
'#End Region
        End Sub

        Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
            treeList1.DataSource = examples
            treeList1.ExpandAll()
            treeList1.BestFitColumns()
        End Sub

        Private Sub RunExample(ByVal filePath As String, ByVal documentFormat As XlDocumentFormat)
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then Return
            Using stream As FileStream = New FileStream(filePath, FileMode.Create)
                Dim action As Action(Of Stream, XlDocumentFormat) = example.Action
                action(stream, documentFormat)
            End Using

            Call Process.Start(filePath)
        End Sub

        Private Sub btnExportToXLSX_Click(ByVal sender As Object, ByVal e As EventArgs)
            If treeList1.FocusedNode Is Nothing Then Return
            Dim fileName As String = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx")
            If String.IsNullOrEmpty(fileName) Then Return
            RunExample(fileName, XlDocumentFormat.Xlsx)
        End Sub

        Private Sub btnExportToXLS_Click(ByVal sender As Object, ByVal e As EventArgs)
            If treeList1.FocusedNode Is Nothing Then Return
            Dim fileName As String = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls")
            If String.IsNullOrEmpty(fileName) Then Return
            RunExample(fileName, XlDocumentFormat.Xls)
        End Sub

        Private Sub btnExportToCSV_Click(ByVal sender As Object, ByVal e As EventArgs)
            If treeList1.FocusedNode Is Nothing Then Return
            Dim fileName As String = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv")
            If String.IsNullOrEmpty(fileName) Then Return
            RunExample(fileName, XlDocumentFormat.Csv)
        End Sub

        Private Function GetSaveFileName(ByVal filter As String, ByVal defaulName As String) As String
            Dim sfDialog As SaveFileDialog = New SaveFileDialog()
            sfDialog.Filter = filter
            sfDialog.FileName = defaulName
            If sfDialog.ShowDialog() <> DialogResult.OK Then Return Nothing
            Return sfDialog.FileName
        End Function

        Private Sub EnableButtons(ByVal exampleName As String)
            btnExportToCSV.Enabled = Not XLExportDisabledCSVExamples.Examples.Contains(exampleName)
            btnExportToXLS.Enabled = Not XLExportForbiddenXLSExamples.Examples.Contains(exampleName)
        End Sub

        Private Sub treeList1_FocusedNodeChanged(ByVal sender As Object, ByVal e As DevExpress.XtraTreeList.FocusedNodeChangedEventArgs)
            Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
            If example Is Nothing Then Return
            EnableButtons(example.Name)
        End Sub
    End Class
End Namespace
