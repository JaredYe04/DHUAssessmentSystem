namespace 考核系统.Dialogs
{
    partial class ErrorInfoDialog
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
            this.textErrorInfo = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // textErrorInfo
            // 
            this.textErrorInfo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textErrorInfo.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textErrorInfo.Location = new System.Drawing.Point(0, 0);
            this.textErrorInfo.Multiline = true;
            this.textErrorInfo.Name = "textErrorInfo";
            this.textErrorInfo.ReadOnly = true;
            this.textErrorInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textErrorInfo.Size = new System.Drawing.Size(671, 462);
            this.textErrorInfo.TabIndex = 0;
            // 
            // ErrorInfoDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(671, 462);
            this.Controls.Add(this.textErrorInfo);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorInfoDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "错误信息";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox textErrorInfo;
    }
}