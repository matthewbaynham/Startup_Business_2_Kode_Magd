namespace KodeMagd.Rename
{
    partial class FrmRenameModuleOrForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRenameModuleOrForm));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.lstModuleType = new System.Windows.Forms.CheckedListBox();
            this.dgModule = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgModule)).BeginInit();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 3;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 378);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 28);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lstModuleType
            // 
            this.lstModuleType.CheckOnClick = true;
            this.lstModuleType.FormattingEnabled = true;
            this.lstModuleType.Location = new System.Drawing.Point(12, 12);
            this.lstModuleType.Name = "lstModuleType";
            this.lstModuleType.Size = new System.Drawing.Size(218, 109);
            this.lstModuleType.TabIndex = 0;
            this.lstModuleType.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstModuleType_KeyUp);
            this.lstModuleType.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstModuleType_MouseUp);
            // 
            // dgModule
            // 
            this.dgModule.AllowUserToAddRows = false;
            this.dgModule.AllowUserToDeleteRows = false;
            this.dgModule.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgModule.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName});
            this.dgModule.Location = new System.Drawing.Point(245, 13);
            this.dgModule.Name = "dgModule";
            this.dgModule.RowHeadersVisible = false;
            this.dgModule.Size = new System.Drawing.Size(375, 344);
            this.dgModule.TabIndex = 1;
            this.dgModule.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgModule_CellValidating);
            // 
            // ColName
            // 
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            this.ColName.Width = 300;
            // 
            // FrmRenameModuleOrForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.dgModule);
            this.Controls.Add(this.lstModuleType);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmRenameModuleOrForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmRenameModuleOrForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmRenameModuleOrForm_KeyDown);
            this.Resize += new System.EventHandler(this.FrmRenameModuleOrForm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgModule)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckedListBox lstModuleType;
        private System.Windows.Forms.DataGridView dgModule;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
    }
}