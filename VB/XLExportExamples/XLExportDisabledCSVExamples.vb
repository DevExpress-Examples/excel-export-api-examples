Imports System.Collections.Generic

Namespace XLExportExamples

'#Region "XLExportDisabledCSVExamples"
    Public Module XLExportDisabledCSVExamples

        Public Examples As HashSet(Of String) = CreateNonCSVExamples()

        Private Function CreateNonCSVExamples() As HashSet(Of String)
            Dim result As HashSet(Of String) = New HashSet(Of String)()
            result.Add("Hide Worksheet")
            result.Add("Hide Gridlines")
            result.Add("Hide Row and Column Headers")
            result.Add("Merge Cells")
            result.Add("Predefined Style-Like Formatting")
            result.Add("Themed Formatting")
            result.Add("Alignment")
            result.Add("Borders")
            result.Add("Fill")
            result.Add("Font")
            result.Add("Rich Text Formatting")
            result.Add("Above/Below Average Rules")
            result.Add("Less Than/Greater Than/Between Rules")
            result.Add("Blank/Non-Blank Cells")
            result.Add("Duplicate Values")
            result.Add("Formula-Based Rules")
            result.Add("Text that Contains... Rule")
            result.Add("A Date Occurring... Rule")
            result.Add("Top/Bottom Rules")
            result.Add("Data Bars")
            result.Add("Icon Sets")
            result.Add("Color Scales")
            result.Add("Enable Filtering")
            result.Add("Outline Data")
            result.Add("Data Validation")
            result.Add("Right-To-Left View")
            result.Add("Freeze Row")
            result.Add("Freeze Column")
            result.Add("Freeze Panes")
            result.Add("Headers and Footers")
            result.Add("Page Breaks")
            result.Add("Page Margins")
            result.Add("Page Setup")
            result.Add("Print Area")
            result.Add("Print Options")
            result.Add("Print Titles")
            result.Add("Insert Hyperlinks")
            result.Add("Formulas")
            result.Add("Shared Formulas")
            result.Add("Functions")
            result.Add("Document Properties")
            result.Add("Insert Picture")
            result.Add("Stretch Picture")
            result.Add("Fit Picture in Cell")
            result.Add("Picture Hyperlink")
            result.Add("Add Sparkline Group")
            result.Add("Add Sparkline to Group")
            result.Add("Adjust Scaling")
            result.Add("Highlight Values")
            result.Add("Display X-axis")
            result.Add("Set Date Range")
            result.Add("Create Table")
            result.Add("Disable Filtering")
            result.Add("Hidden Header Row")
            result.Add("Hidden Total Row")
            result.Add("Side-By-Side Tables")
            result.Add("Table Style")
            result.Add("Table Style Options")
            result.Add("Custom Formatting")
            result.Add("Calculated Column")
            Return result
        End Function
    End Module

'#End Region
'#Region "XLExportForbiddenXLSModules"
    Public Module XLExportForbiddenXLSExamples

        Public Examples As HashSet(Of String) = CreateNonXLSExamples()

        Private Function CreateNonXLSExamples() As HashSet(Of String)
            Dim result As HashSet(Of String) = New HashSet(Of String)()
            result.Add("Add Sparkline Group")
            result.Add("Add Sparkline to Group")
            result.Add("Adjust Scaling")
            result.Add("Highlight Values")
            result.Add("Display X-axis")
            result.Add("Set Date Range")
            Return result
        End Function
    End Module
'#End Region
End Namespace
