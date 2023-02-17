Imports System
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet

Namespace XLExportExamples

    Public Module ConditionalFormattingActions

'#Region "Actions"
        Public AverageAction As Action(Of Stream, XlDocumentFormat) = AddressOf Average

        Public CellIsAction As Action(Of Stream, XlDocumentFormat) = AddressOf CellIs

        Public BlanksAction As Action(Of Stream, XlDocumentFormat) = AddressOf Blanks

        Public DuplicatesAction As Action(Of Stream, XlDocumentFormat) = AddressOf Duplicates

        Public ExpressionAction As Action(Of Stream, XlDocumentFormat) = AddressOf Expression

        Public SpecificTextAction As Action(Of Stream, XlDocumentFormat) = AddressOf SpecificText

        Public TimePeriodAction As Action(Of Stream, XlDocumentFormat) = AddressOf TimePeriod

        Public Top10Action As Action(Of Stream, XlDocumentFormat) = AddressOf Top10

        Public DataBarAction As Action(Of Stream, XlDocumentFormat) = AddressOf DataBar

        Public IconSetAction As Action(Of Stream, XlDocumentFormat) = AddressOf IconSet

        Public ColorScaleAction As Action(Of Stream, XlDocumentFormat) = AddressOf ColorScale

'#End Region
        Private Sub Average(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = i + 1
                                End Using
                            Next
                        End Using
                    Next

'#Region "#AverageRule"
                    ' Create an instance of the XlConditionalFormatting class.
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10))
                    ' Create the rule highlighting values that are above the average in the cell range.
                    Dim rule As XlCondFmtRuleAboveAverage = New XlCondFmtRuleAboveAverage()
                    rule.Condition = XlCondFmtAverageCondition.Above
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class.
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10))
                    ' Create the rule highlighting values that are above or equal to the average value in the cell range.
                    rule = New XlCondFmtRuleAboveAverage()
                    rule.Condition = XlCondFmtAverageCondition.AboveOrEqual
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class.
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10))
                    ' Create the rule highlighting values that are below the average in the cell range.
                    rule = New XlCondFmtRuleAboveAverage()
                    rule.Condition = XlCondFmtAverageCondition.Below
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class.
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10))
                    ' Create the rule highlighting values that are below or equal to the average value in the cell range.
                    rule = New XlCondFmtRuleAboveAverage()
                    rule.Condition = XlCondFmtAverageCondition.BelowOrEqual
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #AverageRule
                End Using
            End Using
        End Sub

        Private Sub CellIs(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = i + 1
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = 12 - i
                            End Using
                        End Using
                    Next

'#Region "#CellIsRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10))
                    ' Create the rule to highlight cells whose values are less than 5.
                    Dim rule As XlCondFmtRuleCellIs = New XlCondFmtRuleCellIs()
                    rule.Operator = XlCondFmtOperator.LessThan
                    rule.Value = 5
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Create the rule to highlight cells whose values are between 5 and 8.
                    rule = New XlCondFmtRuleCellIs()
                    rule.Operator = XlCondFmtOperator.Between
                    rule.Value = 5
                    rule.SecondValue = 8
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral
                    formatting.Rules.Add(rule)
                    ' Create the rule to highlight cells whose values are greater than 8.
                    rule = New XlCondFmtRuleCellIs()
                    rule.Operator = XlCondFmtOperator.GreaterThan
                    rule.Value = 8
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10))
                    ' Create the rule to highlight cells whose values are greater than a value calculated by a formula. 
                    rule = New XlCondFmtRuleCellIs()
                    rule.Operator = XlCondFmtOperator.GreaterThan
                    rule.Value = "=$A1+3"
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #CellIsRule
                End Using
            End Using
        End Sub

        Private Sub Blanks(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 10 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                If i Mod 2 = 0 Then cell.Value = i + 1
                            End Using
                        End Using
                    Next

'#Region "#BlanksRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A1:A10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9))
                    ' Create the rule to highlight blank cells in the range.
                    Dim rule As XlCondFmtRuleBlanks = New XlCondFmtRuleBlanks(True)
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Create the rule to highlight non-blank cells in the range.
                    rule = New XlCondFmtRuleBlanks(False)
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #BlanksRule
                End Using
            End Using
        End Sub

        Private Sub Duplicates(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = cell.ColumnIndex * cell.RowIndex + cell.RowIndex + 1
                                End Using
                            Next
                        End Using
                    Next

'#Region "#DuplicatesRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 10))
                    ' Create the rule to identify duplicate values in the cell range.
                    formatting.Rules.Add(New XlCondFmtRuleDuplicates() With {.Formatting = XlCellFormatting.Bad})
                    ' Create the rule to identify unique values in the cell range.
                    formatting.Rules.Add(New XlCondFmtRuleUnique() With {.Formatting = XlCellFormatting.Good})
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #DuplicatesRule
                End Using
            End Using
        End Sub

        Private Sub Expression(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    Dim width As Integer() = New Integer() {80, 150, 90}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = width(i)
                            If i = 2 Then
                                column.Formatting = New XlCellFormatting()
                                column.Formatting.NumberFormat = "[$$-409] #,##0.00"
                            End If
                        End Using
                    Next

                    Dim columnNames As String() = New String() {"Account ID", "User Name", "Balance"}
                    Using row As IXlRow = sheet.CreateRow()
                        Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                        headerRowFormatting.Font = XlFont.BodyFont()
                        headerRowFormatting.Font.Bold = True
                        headerRowFormatting.Border = New XlBorder()
                        headerRowFormatting.Border.BottomColor = Color.Black
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin
                        For i As Integer = 0 To 3 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = columnNames(i)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    Dim accountIds As String() = New String() {"A105", "A114", "B013", "C231", "D101", "D105"}
                    Dim users As String() = New String() {"Berry Dafoe", "Chris Cadwell", "Esta Mangold", "Liam Bell", "Simon Newman", "Wendy Underwood"}
                    Dim balance As Integer() = New Integer() {155, 250, 48, 350, -15, 10}
                    For i As Integer = 0 To 6 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = accountIds(i)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = users(i)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = balance(i)
                            End Using
                        End Using
                    Next

'#Region "#ExpressionRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A2:C7).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 1, 2, 6))
                    ' Create the rule that uses a formula to highlight cells if a value in the column "C" is greater than 0 and less than 50. 
                    Dim rule As XlCondFmtRuleExpression = New XlCondFmtRuleExpression("AND($C2>0,$C2<50)")
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlFill.SolidFill(Color.FromArgb(&HfF, &HfF, &Hcc))
                    formatting.Rules.Add(rule)
                    ' Create the rule that uses a formula to highlight cells if a value in the column "C" is less than or equal to 0. 
                    rule = New XlCondFmtRuleExpression("$C2<=0")
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #ExpressionRule
                End Using
            End Using
        End Sub

        Private Sub SpecificText(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    Dim width As Integer() = New Integer() {250, 180, 100}
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = width(i)
                            If i = 2 Then
                                column.Formatting = New XlCellFormatting()
                                column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
                            End If
                        End Using
                    Next

                    Dim columnNames As String() = New String() {"Product", "Delivery", "Sales"}
                    Using row As IXlRow = sheet.CreateRow()
                        Dim headerRowFormatting As XlCellFormatting = New XlCellFormatting()
                        headerRowFormatting.Font = XlFont.BodyFont()
                        headerRowFormatting.Font.Bold = True
                        headerRowFormatting.Border = New XlBorder()
                        headerRowFormatting.Border.BottomColor = Color.Black
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin
                        For i As Integer = 0 To 3 - 1
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = columnNames(i)
                                cell.ApplyFormatting(headerRowFormatting)
                            End Using
                        Next
                    End Using

                    Dim products As String() = New String() {"Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Queso Cabrales", "Raclette Courdavault"}
                    Dim deliveries As String() = New String() {"USA", "Worldwide", "USA", "Ships worldwide", "Worldwide except EU", "EU"}
                    Dim sales As Integer() = New Integer() {15500, 20250, 12634, 35010, 15234, 10050}
                    For i As Integer = 0 To 6 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = products(i)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = deliveries(i)
                            End Using

                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = sales(i)
                            End Using
                        End Using
                    Next

'#Region "#SpecificTextRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (B2:B7).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 1, 1, 6))
                    ' Create the rule to highlight cells that contain the given text.
                    Dim rule As XlCondFmtRuleSpecificText = New XlCondFmtRuleSpecificText(XlCondFmtSpecificTextType.Contains, "worldwide")
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #SpecificTextRule
                End Using
            End Using
        End Sub

        Private Sub TimePeriod(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    Using column As IXlColumn = sheet.CreateColumn()
                        column.WidthInPixels = 100
                        column.ApplyFormatting(XlNumberFormat.ShortDate)
                    End Using

                    For i As Integer = 0 To 10 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            Using cell As IXlCell = row.CreateCell()
                                cell.Value = Date.Now.AddDays(row.RowIndex - 5)
                            End Using
                        End Using
                    Next

'#Region "#TimePeriodRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A1:A10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9))
                    ' Create the rule to highlight yesterday's dates in the cell range.
                    Dim rule As XlCondFmtRuleTimePeriod = New XlCondFmtRuleTimePeriod()
                    rule.TimePeriod = XlCondFmtTimePeriod.Yesterday
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Create the rule to highlight today's dates in the cell range.
                    rule = New XlCondFmtRuleTimePeriod()
                    rule.TimePeriod = XlCondFmtTimePeriod.Today
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Create the rule to highlight tomorrows's dates in the cell range.
                    rule = New XlCondFmtRuleTimePeriod()
                    rule.TimePeriod = XlCondFmtTimePeriod.Tomorrow
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Neutral
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #TimePeriodRule
                End Using
            End Using
        End Sub

        Private Sub Top10(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 10 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = cell.ColumnIndex * 4 + cell.RowIndex + 1
                                End Using
                            Next
                        End Using
                    Next

'#Region "#TopAndBottomRules"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rules should be applied (A1:D10).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 9))
                    ' Create the rule to identify bottom 10 values in the cell range.
                    Dim rule As XlCondFmtRuleTop10 = New XlCondFmtRuleTop10()
                    rule.Bottom = True
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Bad
                    formatting.Rules.Add(rule)
                    ' Create the rule to identify top 10 values in the cell range.
                    rule = New XlCondFmtRuleTop10()
                    ' Specify formatting settings to be applied to cells if the condition is true.
                    rule.Formatting = XlCellFormatting.Good
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #TopAndBottomRules
                End Using
            End Using
        End Sub

        Private Sub DataBar(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 3 - 1
                        Using column As IXlColumn = sheet.CreateColumn()
                            column.WidthInPixels = 100
                        End Using
                    Next

                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 3 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    Dim rowIndex As Integer = cell.RowIndex
                                    Dim columnIndex As Integer = cell.ColumnIndex
                                    If columnIndex = 0 Then
                                        cell.Value = rowIndex + 1
                                    ElseIf columnIndex = 1 Then
                                        cell.Value = rowIndex - 5
                                    Else
                                        cell.Value = If(rowIndex < 5, rowIndex + 1, 11 - rowIndex)
                                    End If
                                End Using
                            Next
                        End Using
                    Next

'#Region "#DataBarRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10))
                    ' Create the rule to compare values in the cell range using data bars.
                    Dim rule As XlCondFmtRuleDataBar = New XlCondFmtRuleDataBar()
                    ' Specify the bar color.
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.2)
                    ' Specify the solid fill type.
                    rule.GradientFill = False
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10))
                    ' Create the rule to compare values in the cell range using data bars.
                    rule = New XlCondFmtRuleDataBar()
                    ' Set the positive bar color to green.
                    rule.FillColor = Color.Green
                    ' Set the border color of positive bars to green.
                    rule.BorderColor = Color.Green
                    ' Set the axis color to brown.
                    rule.AxisColor = Color.Brown
                    ' Use the gradient fill type
                    rule.GradientFill = True
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10))
                    ' Create the rule to compare values in the cell range using data bars.
                    rule = New XlCondFmtRuleDataBar()
                    ' Specify the bar color.
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.2)
                    ' Set the minimum length of the data bar.
                    rule.MinLength = 10
                    ' Set the maximum length of the data bar.
                    rule.MaxLength = 90
                    ' Set the value corresponding to the shortest bar.
                    rule.MinValue.ObjectType = XlCondFmtValueObjectType.Number
                    rule.MinValue.Value = 3
                    ' Set the direction of data bars.
                    rule.Direction = XlDataBarDirection.RightToLeft
                    ' Hide values of cells to which the rule is applied.
                    rule.ShowValues = False
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #DataBarRule
                End Using
            End Using
        End Sub

        Private Sub IconSet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    If cell.ColumnIndex Mod 2 = 0 Then
                                        cell.Value = cell.RowIndex + 1
                                    Else
                                        cell.Value = cell.RowIndex - 5
                                    End If
                                End Using
                            Next
                        End Using
                    Next

'#Region "#IconSetRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (A1:A11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10))
                    ' Create the rule to apply a specific icon from the "3 Arrows" icon set to each cell in the range based on its value. 
                    Dim rule As XlCondFmtRuleIconSet = New XlCondFmtRuleIconSet()
                    rule.IconSetType = XlCondFmtIconSetType.Arrows3
                    ' Set the rule priority.
                    rule.Priority = 1
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (B1:B11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10))
                    ' Create the rule to apply a specific icon from the "3 Flags" icon set to each cell in the range based on its value. 
                    rule = New XlCondFmtRuleIconSet()
                    rule.IconSetType = XlCondFmtIconSetType.Flags3
                    ' Set the rule priority.
                    rule.Priority = 2
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10))
                    ' Create the rule to apply a specific icon from the "5 Ratings" icon set to each cell in the range based on its value. 
                    rule = New XlCondFmtRuleIconSet()
                    rule.IconSetType = XlCondFmtIconSetType.Rating5
                    ' Hide values of cells to which the rule is applied.
                    rule.ShowValues = False
                    ' Set the rule priority.
                    rule.Priority = 3
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify the cell range to which the conditional formatting rule should be applied (D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10))
                    ' Create the rule to apply a specific icon from the "4 Traffic Lights" icon set to each cell in the range based on its value. 
                    rule = New XlCondFmtRuleIconSet()
                    rule.IconSetType = XlCondFmtIconSetType.TrafficLights4
                    ' Reverse the icon order.
                    rule.Reverse = True
                    ' Set the rule priority.
                    rule.Priority = 4
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #IconSetRule
                End Using
            End Using
        End Sub

        Private Sub ColorScale(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
            ' Create an exporter instance.
            Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)
            ' Create a new document.
            Using document As IXlDocument = exporter.CreateDocument(stream)
                document.Options.Culture = CultureInfo.CurrentCulture
                ' Create a worksheet.
                Using sheet As IXlSheet = document.CreateSheet()
                    ' Generate data for the document.
                    For i As Integer = 0 To 11 - 1
                        Using row As IXlRow = sheet.CreateRow()
                            For j As Integer = 0 To 4 - 1
                                Using cell As IXlCell = row.CreateCell()
                                    cell.Value = cell.RowIndex + 1
                                End Using
                            Next
                        End Using
                    Next

'#Region "#ColorScaleRule"
                    ' Create an instance of the XlConditionalFormatting class. 
                    Dim formatting As XlConditionalFormatting = New XlConditionalFormatting()
                    ' Specify cell ranges to which the conditional formatting rule should be applied (A1:A11 and C1:C11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10))
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10))
                    ' Create the default three-color scale rule to differentiate low, medium and high values in cell ranges.
                    Dim rule As XlCondFmtRuleColorScale = New XlCondFmtRuleColorScale()
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
                    ' Create an instance of the XlConditionalFormatting class. 
                    formatting = New XlConditionalFormatting()
                    ' Specify cell ranges to which the conditional formatting rule should be applied (B1:B11 and D1:D11).
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10))
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10))
                    ' Create the two-color scale rule to differentiate low and high values in cell ranges. 
                    rule = New XlCondFmtRuleColorScale()
                    rule.ColorScaleType = XlCondFmtColorScaleType.ColorScale2
                    ' Set a color corresponding to the minimum value in the cell range.
                    rule.MinColor = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
                    ' Set a color corresponding to the maximum value in the cell range.
                    rule.MaxColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.5)
                    formatting.Rules.Add(rule)
                    ' Add the specified format options to the worksheet collection of conditional formats.
                    sheet.ConditionalFormattings.Add(formatting)
'#End Region  ' #ColorScaleRule
                End Using
            End Using
        End Sub
    End Module
End Namespace
