Imports System
Imports System.ComponentModel
Imports DevExpress.XtraTreeList
Imports DevExpress.Export.Xl
Imports System.IO

Namespace XLExportExamples

    Public Class SpreadsheetNode

        Private groupsField As GroupsOfSpreadsheetExamples = New GroupsOfSpreadsheetExamples()

        Private ownerField As GroupsOfSpreadsheetExamples

        Public Sub New(ByVal name As String)
            Me.Name = name
        End Sub

        <Browsable(False)>
        Public ReadOnly Property Groups As GroupsOfSpreadsheetExamples
            Get
                Return groupsField
            End Get
        End Property

        Public Property Name As String

        <Browsable(False)>
        Public Property Owner As GroupsOfSpreadsheetExamples
            Get
                Return ownerField
            End Get

            Set(ByVal value As GroupsOfSpreadsheetExamples)
                ownerField = value
            End Set
        End Property
    End Class

    Public Class SpreadsheetExample
        Inherits SpreadsheetNode

        Private _Action As Action(Of System.IO.Stream, DevExpress.Export.Xl.XlDocumentFormat)

        Public Sub New(ByVal name As String, ByVal action As Action(Of Stream, XlDocumentFormat))
            MyBase.New(name)
            Me.Action = action
        End Sub

        Public Property Action As Action(Of Stream, XlDocumentFormat)
            Get
                Return _Action
            End Get

            Private Set(ByVal value As Action(Of Stream, XlDocumentFormat))
                _Action = value
            End Set
        End Property
    End Class

    Public Class GroupsOfSpreadsheetExamples
        Inherits BindingList(Of SpreadsheetNode)
        Implements TreeList.IVirtualTreeListData

        Private Sub VirtualTreeGetChildNodes(ByVal info As VirtualTreeGetChildNodesInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetChildNodes
            Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
            info.Children = obj.Groups
        End Sub

        Protected Overrides Sub InsertItem(ByVal index As Integer, ByVal item As SpreadsheetNode)
            item.Owner = Me
            MyBase.InsertItem(index, item)
        End Sub

        Private Sub VirtualTreeGetCellValue(ByVal info As VirtualTreeGetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeGetCellValue
            Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
            Select Case info.Column.Caption
                Case "Name"
                    info.CellData = obj.Name
            End Select
        End Sub

        Private Sub VirtualTreeSetCellValue(ByVal info As VirtualTreeSetCellValueInfo) Implements TreeList.IVirtualTreeListData.VirtualTreeSetCellValue
            Dim obj As SpreadsheetNode = TryCast(info.Node, SpreadsheetNode)
            Select Case info.Column.Caption
                Case "Name"
                    obj.Name = CStr(info.NewCellData)
            End Select
        End Sub
    End Class
End Namespace
