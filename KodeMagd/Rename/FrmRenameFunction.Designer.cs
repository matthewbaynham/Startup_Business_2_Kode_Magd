namespace KodeMagd.Rename
{
    partial class FrmRenameFunction
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRenameFunction));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.lstModuleType = new System.Windows.Forms.CheckedListBox();
            this.lstModule = new System.Windows.Forms.ListBox();
            this.dgFunctions = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColScope = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgFunctions)).BeginInit();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 4;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 393);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 28);
            this.btnClose.TabIndex = 3;
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
            // lstModule
            // 
            this.lstModule.FormattingEnabled = true;
            this.lstModule.Location = new System.Drawing.Point(12, 157);
            this.lstModule.Name = "lstModule";
            this.lstModule.Size = new System.Drawing.Size(218, 264);
            this.lstModule.TabIndex = 1;
            this.lstModule.SelectedIndexChanged += new System.EventHandler(this.lstModule_SelectedIndexChanged);
            // 
            // dgFunctions
            // 
            this.dgFunctions.AllowUserToAddRows = false;
            this.dgFunctions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFunctions.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColScope,
            this.ColType});
            this.dgFunctions.Location = new System.Drawing.Point(241, 16);
            this.dgFunctions.Name = "dgFunctions";
            this.dgFunctions.RowHeadersVisible = false;
            this.dgFunctions.Size = new System.Drawing.Size(378, 362);
            this.dgFunctions.TabIndex = 2;
            this.dgFunctions.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgFunctions_CellValidating);
            // 
            // ColName
            // 
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            this.ColName.Width = 200;
            // 
            // ColScope
            // 
            this.ColScope.HeaderText = "Scope";
            this.ColScope.Name = "ColScope";
            this.ColScope.ReadOnly = true;
            this.ColScope.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColScope.Width = 50;
            // 
            // ColType
            // 
            this.ColType.HeaderText = "Type";
            this.ColType.MinimumWidth = 50;
            this.ColType.Name = "ColType";
            this.ColType.ReadOnly = true;
            this.ColType.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColType.Width = 150;
            // 
            // FrmRenameFunction
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.dgFunctions);
            this.Controls.Add(this.lstModule);
            this.Controls.Add(this.lstModuleType);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmRenameFunction";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmRenameFunction_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmRenameFunction_KeyDown);
            this.Resize += new System.EventHandler(this.FrmRenameFunction_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgFunctions)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckedListBox lstModuleType;
        private System.Windows.Forms.ListBox lstModule;
        private System.Windows.Forms.DataGridView dgFunctions;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColScope;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColType;
    }
}