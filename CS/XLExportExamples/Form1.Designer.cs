using System.Drawing;
namespace XLExportExamples {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.treeList1 = new DevExpress.XtraTreeList.TreeList();
            this.treeListColumn1 = new DevExpress.XtraTreeList.Columns.TreeListColumn();
            this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
            this.btnExportToCSV = new DevExpress.XtraEditors.SimpleButton();
            this.btnExportToXLS = new DevExpress.XtraEditors.SimpleButton();
            this.btnExportToXLSX = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
            this.splitContainerControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeList1
            // 
            this.treeList1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Blue;
            this.treeList1.Appearance.FocusedCell.Options.UseFont = true;
            this.treeList1.Appearance.FocusedCell.Options.UseForeColor = true;
            this.treeList1.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
            this.treeListColumn1});
            this.treeList1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeList1.Location = new System.Drawing.Point(0, 0);
            this.treeList1.Name = "treeList1";
            this.treeList1.OptionsBehavior.Editable = false;
            this.treeList1.OptionsView.ShowColumns = false;
            this.treeList1.OptionsView.ShowIndicator = false;
            this.treeList1.Size = new System.Drawing.Size(497, 614);
            this.treeList1.TabIndex = 0;
            this.treeList1.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.treeList1_FocusedNodeChanged);
            // 
            // treeListColumn1
            // 
            this.treeListColumn1.Caption = "Name";
            this.treeListColumn1.FieldName = "Name";
            this.treeListColumn1.Name = "treeListColumn1";
            this.treeListColumn1.Visible = true;
            this.treeListColumn1.VisibleIndex = 0;
            this.treeListColumn1.Width = 92;
            // 
            // splitContainerControl1
            // 
            this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerControl1.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2;
            this.splitContainerControl1.Horizontal = false;
            this.splitContainerControl1.Location = new System.Drawing.Point(0, 0);
            this.splitContainerControl1.Name = "splitContainerControl1";
            this.splitContainerControl1.Panel1.Controls.Add(this.treeList1);
            this.splitContainerControl1.Panel1.Text = "Panel1";
            this.splitContainerControl1.Panel2.Controls.Add(this.btnExportToCSV);
            this.splitContainerControl1.Panel2.Controls.Add(this.btnExportToXLS);
            this.splitContainerControl1.Panel2.Controls.Add(this.btnExportToXLSX);
            this.splitContainerControl1.Panel2.Text = "Panel2";
            this.splitContainerControl1.Size = new System.Drawing.Size(497, 700);
            this.splitContainerControl1.SplitterPosition = 81;
            this.splitContainerControl1.TabIndex = 2;
            this.splitContainerControl1.Text = "splitContainerControl1";
            // 
            // btnExportToCSV
            // 
            this.btnExportToCSV.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToCSV.Image")));
            this.btnExportToCSV.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter;
            this.btnExportToCSV.Location = new System.Drawing.Point(299, 10);
            this.btnExportToCSV.Name = "btnExportToCSV";
            this.btnExportToCSV.Size = new System.Drawing.Size(90, 59);
            this.btnExportToCSV.TabIndex = 5;
            this.btnExportToCSV.Text = "Export to CSV";
            this.btnExportToCSV.Click += new System.EventHandler(this.btnExportToCSV_Click);
            // 
            // btnExportToXLS
            // 
            this.btnExportToXLS.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToXLS.Image")));
            this.btnExportToXLS.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter;
            this.btnExportToXLS.Location = new System.Drawing.Point(203, 10);
            this.btnExportToXLS.Name = "btnExportToXLS";
            this.btnExportToXLS.Size = new System.Drawing.Size(90, 59);
            this.btnExportToXLS.TabIndex = 4;
            this.btnExportToXLS.Text = "Export to XLS";
            this.btnExportToXLS.Click += new System.EventHandler(this.btnExportToXLS_Click);
            // 
            // btnExportToXLSX
            // 
            this.btnExportToXLSX.Image = ((System.Drawing.Image)(resources.GetObject("btnExportToXLSX.Image")));
            this.btnExportToXLSX.ImageLocation = DevExpress.XtraEditors.ImageLocation.TopCenter;
            this.btnExportToXLSX.Location = new System.Drawing.Point(107, 10);
            this.btnExportToXLSX.Name = "btnExportToXLSX";
            this.btnExportToXLSX.Size = new System.Drawing.Size(90, 59);
            this.btnExportToXLSX.TabIndex = 3;
            this.btnExportToXLSX.Text = "Export to XLSX";
            this.btnExportToXLSX.Click += new System.EventHandler(this.btnExportToXLSX_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 700);
            this.Controls.Add(this.splitContainerControl1);
            this.Name = "Form1";
            this.Text = "XL Export Library Examples";
            ((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
            this.splitContainerControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraTreeList.TreeList treeList1;
        private DevExpress.XtraTreeList.Columns.TreeListColumn treeListColumn1;
        private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
        private DevExpress.XtraEditors.SimpleButton btnExportToCSV;
        private DevExpress.XtraEditors.SimpleButton btnExportToXLS;
        private DevExpress.XtraEditors.SimpleButton btnExportToXLSX;
    }
}

