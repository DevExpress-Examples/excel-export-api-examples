using System;
using System.Collections.Generic;

namespace XLExportExamples
{
    #region XLExportDisabledCSVExamples
    public static class XLExportDisabledCSVExamples {
        public static HashSet<string> Examples = CreateNonCSVExamples();
        static HashSet<string> CreateNonCSVExamples() {
            HashSet<string> result = new HashSet<string>();
            result.Add("Hide Worksheet");
            result.Add("Merge Cells");
            result.Add("Predefined Style-Like Formatting");
            result.Add("Themed Formatting");
            result.Add("Alignment");
            result.Add("Borders");
            result.Add("Fill");
            result.Add("Font");
            result.Add("Above/Below Average Rules");
            result.Add("Less Than/Greater Than/Between Rules");
            result.Add("Blank/Non-Blank Cells");
            result.Add("Duplicate Values");
            result.Add("Formula-Based Rules");
            result.Add("Text that Contains... Rule");
            result.Add("A Date Occurring... Rule");
            result.Add("Top/Bottom Rules");
            result.Add("Data Bars");
            result.Add("Icon Sets");
            result.Add("Color Scales");
            result.Add("Enable Filtering");
            result.Add("Outline Data");
            result.Add("Data Validation");
            result.Add("Freeze Row");
            result.Add("Freeze Column");
            result.Add("Freeze Panes");
            result.Add("Headers and Footers");
            result.Add("Page Breaks");
            result.Add("Page Margins");
            result.Add("Page Setup");
            result.Add("Print Area");
            result.Add("Print Options");
            result.Add("Print Titles");
            result.Add("Insert Hyperlinks");
            result.Add("Formulas");
            result.Add("Shared Formulas");
            result.Add("Functions");
            result.Add("Document Properties");
            result.Add("Insert Picture");
            result.Add("Stretch Picture");
            result.Add("Fit Picture In Cell");
            return result;
        }
    }
    #endregion
}
