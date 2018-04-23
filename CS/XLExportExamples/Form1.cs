using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using DevExpress.Export.Xl;

namespace XLExportExamples
{
    public partial class Form1 : Form {

        public Form1() {
            InitializeComponent();
            InitTreeListControl();
        }

        void InitTreeListControl() {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }
        void InitData(GroupsOfSpreadsheetExamples examples) {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Basic Actions"));
            examples.Add(new SpreadsheetNode("Cell Formatting"));
            examples.Add(new SpreadsheetNode("Conditional Formatting"));
            examples.Add(new SpreadsheetNode("Data Actions"));
            examples.Add(new SpreadsheetNode("Formula Actions"));
            examples.Add(new SpreadsheetNode("Page View and Layout"));
            examples.Add(new SpreadsheetNode("Pictures"));
            examples.Add(new SpreadsheetNode("Sparklines"));
            examples.Add(new SpreadsheetNode("Miscellaneous"));
            examples.Add(new SpreadsheetNode("Tables"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Basic Actions" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Create Document", GeneralActions.CreateDocumentAction));
            examples[0].Groups.Add(new SpreadsheetExample("Create Worksheet", GeneralActions.CreateSheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Hide Worksheet", GeneralActions.CreateHiddenSheetAction));
            examples[0].Groups.Add(new SpreadsheetExample("Hide Gridlines", GeneralActions.HideGridlinesAction));
            examples[0].Groups.Add(new SpreadsheetExample("Hide Row and Column Headers", GeneralActions.HideHeadersAction));
            examples[0].Groups.Add(new SpreadsheetExample("Create Columns", GeneralActions.CreateColumnsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Create Rows", GeneralActions.CreateRowsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Create Cells", GeneralActions.CreateCellsAction));
            examples[0].Groups.Add(new SpreadsheetExample("Merge Cells", GeneralActions.MergeCellsAction));

            // Add nodes to the "Cell Formatting" group of examples.
            examples[1].Groups.Add(new SpreadsheetExample("Predefined Style-Like Formatting", CellFormattingActions.PredefinedFormattingAction));
            examples[1].Groups.Add(new SpreadsheetExample("Themed Formatting", CellFormattingActions.ThemedFormattingAction));
            examples[1].Groups.Add(new SpreadsheetExample("Alignment", CellFormattingActions.AlignmentAction));
            examples[1].Groups.Add(new SpreadsheetExample("Borders", CellFormattingActions.BordersAction));
            examples[1].Groups.Add(new SpreadsheetExample("Fill", CellFormattingActions.FillAction));
            examples[1].Groups.Add(new SpreadsheetExample("Font", CellFormattingActions.FontAction));
            examples[1].Groups.Add(new SpreadsheetExample("Number Format", CellFormattingActions.NumberFormatAction));
            examples[1].Groups.Add(new SpreadsheetExample("Rich Text Formatting", CellFormattingActions.RichTextFormattingAction));

            // Add nodes to the "Conditional Formatting" group of examples.
            examples[2].Groups.Add(new SpreadsheetExample("Less Than/Greater Than/Between Rules", ConditionalFormattingActions.CellIsAction));
            examples[2].Groups.Add(new SpreadsheetExample("Text that Contains... Rule", ConditionalFormattingActions.SpecificTextAction));
            examples[2].Groups.Add(new SpreadsheetExample("A Date Occurring... Rule", ConditionalFormattingActions.TimePeriodAction));
            examples[2].Groups.Add(new SpreadsheetExample("Duplicate Values", ConditionalFormattingActions.DuplicatesAction));
            examples[2].Groups.Add(new SpreadsheetExample("Blank/Non-Blank Cells", ConditionalFormattingActions.BlanksAction));
            examples[2].Groups.Add(new SpreadsheetExample("Formula-Based Rules", ConditionalFormattingActions.ExpressionAction));
            examples[2].Groups.Add(new SpreadsheetExample("Top/Bottom Rules", ConditionalFormattingActions.Top10Action));
            examples[2].Groups.Add(new SpreadsheetExample("Above/Below Average Rules", ConditionalFormattingActions.AverageAction));
            examples[2].Groups.Add(new SpreadsheetExample("Color Scales", ConditionalFormattingActions.ColorScaleAction));
            examples[2].Groups.Add(new SpreadsheetExample("Data Bars", ConditionalFormattingActions.DataBarAction));
            examples[2].Groups.Add(new SpreadsheetExample("Icon Sets", ConditionalFormattingActions.IconSetAction));

            // Add nodes to the "Data Actions" group of examples.
            examples[3].Groups.Add(new SpreadsheetExample("Enable Filtering", DataActions.AutoFilterAction));
            examples[3].Groups.Add(new SpreadsheetExample("Outline Data", DataActions.OutlineGroupingAction));
            examples[3].Groups.Add(new SpreadsheetExample("Data Validation", DataActions.DataValidationAction));

            // Add nodes to the "Formula Actions" group of examples.
            examples[4].Groups.Add(new SpreadsheetExample("Simple Formulas", FormulaActions.SimpleFormulasAction));
            examples[4].Groups.Add(new SpreadsheetExample("Complex Formulas", FormulaActions.ComplexFormulasAction));
            examples[4].Groups.Add(new SpreadsheetExample("Shared Formulas", FormulaActions.SharedFormulasAction));
            examples[4].Groups.Add(new SpreadsheetExample("Subtotals", FormulaActions.SubtotalsAction));

            // Add nodes to the "Page View and Layout" group of examples.
            examples[5].Groups.Add(new SpreadsheetExample("Freeze Row", PageViewAndLayoutActions.FreezeRowAction));
            examples[5].Groups.Add(new SpreadsheetExample("Freeze Column", PageViewAndLayoutActions.FreezeColumnAction));
            examples[5].Groups.Add(new SpreadsheetExample("Freeze Panes", PageViewAndLayoutActions.FreezePanesAction));
            examples[5].Groups.Add(new SpreadsheetExample("Right-To-Left View", PageViewAndLayoutActions.SheetViewRTLAction));
            examples[5].Groups.Add(new SpreadsheetExample("Headers and Footers", PageViewAndLayoutActions.HeadersFootersAction));
            examples[5].Groups.Add(new SpreadsheetExample("Page Breaks", PageViewAndLayoutActions.PageBreaksAction));
            examples[5].Groups.Add(new SpreadsheetExample("Page Margins", PageViewAndLayoutActions.PageMarginsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Page Setup", PageViewAndLayoutActions.PageSetupAction));
            examples[5].Groups.Add(new SpreadsheetExample("Print Area", PageViewAndLayoutActions.PrintAreaAction));
            examples[5].Groups.Add(new SpreadsheetExample("Print Options", PageViewAndLayoutActions.PrintOptionsAction));
            examples[5].Groups.Add(new SpreadsheetExample("Print Titles", PageViewAndLayoutActions.PrintTitlesAction));

            // Add nodes to the "Pictures" group of examples.
            examples[6].Groups.Add(new SpreadsheetExample("Insert Picture", PictureActions.InsertPictureAction));
            examples[6].Groups.Add(new SpreadsheetExample("Stretch Picture", PictureActions.StretchPictureAction));
            examples[6].Groups.Add(new SpreadsheetExample("Fit Picture in Cell", PictureActions.FitPictureAction));
            examples[6].Groups.Add(new SpreadsheetExample("Picture Hyperlink", PictureActions.PictureHyperlinkClickAction));

            // Add nodes to the "Sparklines" group of examples.
            examples[7].Groups.Add(new SpreadsheetExample("Add Sparkline Group", SparklineActions.AddSparklineGroupAction));
            examples[7].Groups.Add(new SpreadsheetExample("Add Sparkline to Group", SparklineActions.AddSparklineAction));
            examples[7].Groups.Add(new SpreadsheetExample("Adjust Scaling", SparklineActions.AdjustScalingAction));
            examples[7].Groups.Add(new SpreadsheetExample("Highlight Values", SparklineActions.HighlightValuesAction));
            examples[7].Groups.Add(new SpreadsheetExample("Display X-axis", SparklineActions.DisplayXAxisAction));
            examples[7].Groups.Add(new SpreadsheetExample("Set Date Range", SparklineActions.SetDateRangeAction));


            // Add nodes to the "Miscellaneous" group of examples.
            examples[8].Groups.Add(new SpreadsheetExample("Insert Hyperlinks", MiscellaneousActions.HyperlinksAction));
            examples[8].Groups.Add(new SpreadsheetExample("Document Properties", MiscellaneousActions.DocumentPropertiesAction));
            examples[8].Groups.Add(new SpreadsheetExample("Document Options and Restrictions", MiscellaneousActions.DocumentOptionsAction));
            examples[8].Groups.Add(new SpreadsheetExample("CSV Export Options", MiscellaneousActions.CsvExportOptionsAction));

            // Add nodes to the "Tables" group of examples.
            examples[9].Groups.Add(new SpreadsheetExample("Create Table", TableActions.AddTableAction));
            examples[9].Groups.Add(new SpreadsheetExample("Disable Filtering", TableActions.DisableFilteringAction));
            examples[9].Groups.Add(new SpreadsheetExample("Hidden Header Row", TableActions.HiddenHeaderRowAction));
            examples[9].Groups.Add(new SpreadsheetExample("Hidden Total Row", TableActions.HiddenTotalRowAction));
            examples[9].Groups.Add(new SpreadsheetExample("Side-By-Side Tables", TableActions.SideBySideAction));
            examples[9].Groups.Add(new SpreadsheetExample("Table Style", TableActions.TableStyleAction));
            examples[9].Groups.Add(new SpreadsheetExample("Table Style Options", TableActions.TableStyleOptionsAction));
            examples[9].Groups.Add(new SpreadsheetExample("Custom Formatting", TableActions.CustomFormattingAction));
            examples[9].Groups.Add(new SpreadsheetExample("Calculated Column", TableActions.CalculatedColumnAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples) {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }
        

         void RunExample(string filePath, XlDocumentFormat documentFormat)
        {
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                Action<Stream, XlDocumentFormat> action = example.Action;
                action(stream, documentFormat);
            }
            Process.Start(filePath);
        }

        void btnExportToXLSX_Click(object sender, EventArgs e)
        {
            if (treeList1.FocusedNode == null) return;
            string fileName = GetSaveFileName("Excel Workbook files(*.xlsx)|*.xlsx", "Document.xlsx");
            if (string.IsNullOrEmpty(fileName))
                return;
            RunExample(fileName, XlDocumentFormat.Xlsx);
        }

        void btnExportToXLS_Click(object sender, EventArgs e)
        {
            if (treeList1.FocusedNode == null) return;
            string fileName = GetSaveFileName("Excel 97-2003 Workbook files(*.xls)|*.xls", "Document.xls");
            if (string.IsNullOrEmpty(fileName))
                return;
            RunExample(fileName, XlDocumentFormat.Xls);
        }

        void btnExportToCSV_Click(object sender, EventArgs e)
        {
            if (treeList1.FocusedNode == null) return;
            string fileName = GetSaveFileName("CSV (Comma delimited files)(*.csv)|*.csv", "Document.csv");
            if (string.IsNullOrEmpty(fileName))
                return;
            RunExample(fileName, XlDocumentFormat.Csv);
        }

        string GetSaveFileName(string filter, string defaulName)
        {
            SaveFileDialog sfDialog = new SaveFileDialog();
            sfDialog.Filter = filter;
            sfDialog.FileName = defaulName;
            if (sfDialog.ShowDialog() != DialogResult.OK)
                return null;
            return sfDialog.FileName;
        }

        void EnableButtons(string exampleName)
        {
            btnExportToCSV.Enabled = !XLExportDisabledCSVExamples.Examples.Contains(exampleName);
            btnExportToXLS.Enabled = !XLExportForbiddenXLSExamples.Examples.Contains(exampleName);
        }

        private void treeList1_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
        {
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null) return;
            EnableButtons(example.Name);
        }
    }
}
