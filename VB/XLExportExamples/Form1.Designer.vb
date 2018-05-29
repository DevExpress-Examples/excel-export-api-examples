Imports System.Drawing
Namespace XLExportExamples
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.treeListColumn1 = New DevExpress.XtraTreeList.Columns.TreeListColumn()
            Me.splitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.btnExportToCSV = New DevExpress.XtraEditors.SimpleButton()
            Me.btnExportToXLS = New DevExpress.XtraEditors.SimpleButton()
            Me.btnExportToXLSX = New DevExpress.XtraEditors.SimpleButton()
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' treeList1
            ' 
            Me.treeList1.Appearance.FocusedCell.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold)
            Me.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Blue
            Me.treeList1.Appearance.FocusedCell.Options.UseFont = True
            Me.treeList1.Appearance.FocusedCell.Options.UseForeColor = True
            Me.treeList1.Columns.AddRange(New DevExpress.XtraTreeList.Columns.TreeListColumn() { Me.treeListColumn1})
            Me.treeList1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.treeList1.Location = New System.Drawing.Point(0, 0)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.OptionsBehavior.Editable = False
            Me.treeList1.OptionsView.ShowColumns = False
            Me.treeList1.OptionsView.ShowIndicator = False
            Me.treeList1.Size = New System.Drawing.Size(497, 614)
            Me.treeList1.TabIndex = 0
            ' 
            ' treeListColumn1
            ' 
            Me.treeListColumn1.Caption = "Name"
            Me.treeListColumn1.FieldName = "Name"
            Me.treeListColumn1.Name = "treeListColumn1"
            Me.treeListColumn1.Visible = True
            Me.treeListColumn1.VisibleIndex = 0
            Me.treeListColumn1.Width = 92
            ' 
            ' splitContainerControl1
            ' 
            Me.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.splitContainerControl1.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.splitContainerControl1.Horizontal = False
            Me.splitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl1.Name = "splitContainerControl1"
            Me.splitContainerControl1.Panel1.Controls.Add(Me.treeList1)
            Me.splitContainerControl1.Panel1.Text = "Panel1"
            Me.splitContainerControl1.Panel2.Controls.Add(Me.btnExportToCSV)
            Me.splitContainerControl1.Panel2.Controls.Add(Me.btnExportToXLS)
            Me.splitContainerControl1.Panel2.Controls.Add(Me.btnExportToXLSX)
            Me.splitContainerControl1.Panel2.Text = "Panel2"
            Me.splitContainerControl1.Size = New System.Drawing.Size(497, 700)
            Me.splitContainerControl1.SplitterPosition = 81
            Me.splitContainerControl1.TabIndex = 2
            Me.splitContainerControl1.Text = "splitContainerControl1"
            ' 
            ' btnExportToCSV
            ' 
            Me.btnExportToCSV.Image = (CType(resources.GetObject("btnExportToCSV.Image"), System.Drawing.Image))
            Me.btnExportToCSV.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter
            Me.btnExportToCSV.Location = New System.Drawing.Point(299, 10)
            Me.btnExportToCSV.Name = "btnExportToCSV"
            Me.btnExportToCSV.Size = New System.Drawing.Size(90, 59)
            Me.btnExportToCSV.TabIndex = 5
            Me.btnExportToCSV.Text = "Export to CSV"
            ' 
            ' btnExportToXLS
            ' 
            Me.btnExportToXLS.Image = (CType(resources.GetObject("btnExportToXLS.Image"), System.Drawing.Image))
            Me.btnExportToXLS.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter
            Me.btnExportToXLS.Location = New System.Drawing.Point(203, 10)
            Me.btnExportToXLS.Name = "btnExportToXLS"
            Me.btnExportToXLS.Size = New System.Drawing.Size(90, 59)
            Me.btnExportToXLS.TabIndex = 4
            Me.btnExportToXLS.Text = "Export to XLS"
            ' 
            ' btnExportToXLSX
            ' 
            Me.btnExportToXLSX.Image = (CType(resources.GetObject("btnExportToXLSX.Image"), System.Drawing.Image))
            Me.btnExportToXLSX.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter
            Me.btnExportToXLSX.Location = New System.Drawing.Point(107, 10)
            Me.btnExportToXLSX.Name = "btnExportToXLSX"
            Me.btnExportToXLSX.Size = New System.Drawing.Size(90, 59)
            Me.btnExportToXLSX.TabIndex = 3
            Me.btnExportToXLSX.Text = "Export to XLSX"
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(497, 700)
            Me.Controls.Add(Me.splitContainerControl1)
            Me.Name = "Form1"
            Me.Text = "XL Export Library Examples"
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.splitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl1.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

        #End Region

        Private WithEvents treeList1 As DevExpress.XtraTreeList.TreeList
        Private treeListColumn1 As DevExpress.XtraTreeList.Columns.TreeListColumn
        Private splitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl
        Private WithEvents btnExportToCSV As DevExpress.XtraEditors.SimpleButton
        Private WithEvents btnExportToXLS As DevExpress.XtraEditors.SimpleButton
        Private WithEvents btnExportToXLSX As DevExpress.XtraEditors.SimpleButton
    End Class
End Namespace

