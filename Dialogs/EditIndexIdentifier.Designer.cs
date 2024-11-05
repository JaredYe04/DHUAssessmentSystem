namespace 考核系统.Dialogs
{
    partial class EditIndexIdentifier
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.IndexIdentifierSplit = new System.Windows.Forms.TableLayoutPanel();
            this.indexIdentifierDataGrid = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.identifier_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox10.SuspendLayout();
            this.IndexIdentifierSplit.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.indexIdentifierDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.IndexIdentifierSplit);
            this.groupBox10.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox10.Location = new System.Drawing.Point(0, 0);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(338, 623);
            this.groupBox10.TabIndex = 8;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "指标类别";
            // 
            // IndexIdentifierSplit
            // 
            this.IndexIdentifierSplit.ColumnCount = 1;
            this.IndexIdentifierSplit.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.IndexIdentifierSplit.Controls.Add(this.indexIdentifierDataGrid, 0, 0);
            this.IndexIdentifierSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.IndexIdentifierSplit.Location = new System.Drawing.Point(3, 24);
            this.IndexIdentifierSplit.Name = "IndexIdentifierSplit";
            this.IndexIdentifierSplit.RowCount = 1;
            this.IndexIdentifierSplit.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.IndexIdentifierSplit.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 648F));
            this.IndexIdentifierSplit.Size = new System.Drawing.Size(332, 596);
            this.IndexIdentifierSplit.TabIndex = 0;
            // 
            // indexIdentifierDataGrid
            // 
            this.indexIdentifierDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.indexIdentifierDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.identifier_name});
            this.indexIdentifierDataGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.indexIdentifierDataGrid.Location = new System.Drawing.Point(3, 4);
            this.indexIdentifierDataGrid.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.indexIdentifierDataGrid.Name = "indexIdentifierDataGrid";
            this.indexIdentifierDataGrid.RowHeadersWidth = 51;
            this.indexIdentifierDataGrid.RowTemplate.Height = 27;
            this.indexIdentifierDataGrid.Size = new System.Drawing.Size(326, 588);
            this.indexIdentifierDataGrid.TabIndex = 7;
            this.indexIdentifierDataGrid.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.indexIdentifierDataGrid_CellEndEdit);
            this.indexIdentifierDataGrid.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.indexIdentifierDataGrid_RowsRemoved);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "分类标号";
            this.dataGridViewTextBoxColumn1.MinimumWidth = 6;
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Width = 120;
            // 
            // identifier_name
            // 
            this.identifier_name.HeaderText = "分类名称";
            this.identifier_name.MinimumWidth = 8;
            this.identifier_name.Name = "identifier_name";
            this.identifier_name.Width = 120;
            // 
            // EditIndexIdentifier
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(338, 623);
            this.Controls.Add(this.groupBox10);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "EditIndexIdentifier";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "编辑指标类别";
            this.Load += new System.EventHandler(this.EditIndexIdentifier_Load);
            this.groupBox10.ResumeLayout(false);
            this.IndexIdentifierSplit.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.indexIdentifierDataGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.TableLayoutPanel IndexIdentifierSplit;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn identifier_name;
        public System.Windows.Forms.DataGridView indexIdentifierDataGrid;
    }
}