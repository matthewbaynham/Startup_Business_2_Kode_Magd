namespace KodeMagd.Rename
{
    partial class FrmRenameVariable
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRenameVariable));
            this.lblModule = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblVariables = new System.Windows.Forms.Label();
            this.dgVariables = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColFunction = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColIndex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblModuleType = new System.Windows.Forms.Label();
            this.cmbModuleType = new System.Windows.Forms.ComboBox();
            this.lstModule = new System.Windows.Forms.ListBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgVariables)).BeginInit();
            this.SuspendLayout();
            // 
            // lblModule
            // 
            this.lblModule.AutoSize = true;
            this.lblModule.Location = new System.Drawing.Point(12, 44);
            this.lblModule.Name = "lblModule";
            this.lblModule.Size = new System.Drawing.Size(42, 13);
            this.lblModule.TabIndex = 2;
            this.lblModule.Text = "Module";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 7;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(535, 381);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblVariables
            // 
            this.lblVariables.AutoSize = true;
            this.lblVariables.Location = new System.Drawing.Point(264, 44);
            this.lblVariables.Name = "lblVariables";
            this.lblVariables.Size = new System.Drawing.Size(50, 13);
            this.lblVariables.TabIndex = 4;
            this.lblVariables.Text = "Variables";
            // 
            // dgVariables
            // 
            this.dgVariables.AllowUserToAddRows = false;
            this.dgVariables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgVariables.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColType,
            this.ColFunction,
            this.ColIndex});
            this.dgVariables.Location = new System.Drawing.Point(267, 69);
            this.dgVariables.Name = "dgVariables";
            this.dgVariables.RowHeadersVisible = false;
            this.dgVariables.Size = new System.Drawing.Size(342, 299);
            this.dgVariables.TabIndex = 5;
            this.dgVariables.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgVariables_CellValidating);
            // 
            // ColName
            // 
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            // 
            // ColType
            // 
            this.ColType.HeaderText = "Type";
            this.ColType.Name = "ColType";
            this.ColType.ReadOnly = true;
            this.ColType.Width = 60;
            // 
            // ColFunction
            // 
            this.ColFunction.HeaderText = "Function / Sub / Property";
            this.ColFunction.Name = "ColFunction";
            this.ColFunction.ReadOnly = true;
            // 
            // ColIndex
            // 
            this.ColIndex.HeaderText = "Index";
            this.ColIndex.Name = "ColIndex";
            this.ColIndex.ReadOnly = true;
            this.ColIndex.Visible = false;
            // 
            // lblModuleType
            // 
            this.lblModuleType.AutoSize = true;
            this.lblModuleType.Location = new System.Drawing.Point(13, 15);
            this.lblModuleType.Name = "lblModuleType";
            this.lblModuleType.Size = new System.Drawing.Size(94, 13);
            this.lblModuleType.TabIndex = 0;
            this.lblModuleType.Text = "Filter Module Type";
            // 
            // cmbModuleType
            // 
            this.cmbModuleType.FormattingEnabled = true;
            this.cmbModuleType.Location = new System.Drawing.Point(122, 12);
            this.cmbModuleType.Name = "cmbModuleType";
            this.cmbModuleType.Size = new System.Drawing.Size(136, 21);
            this.cmbModuleType.TabIndex = 1;
            this.cmbModuleType.SelectedIndexChanged += new System.EventHandler(this.cmbModuleType_SelectedIndexChanged);
            // 
            // lstModule
            // 
            this.lstModule.FormattingEnabled = true;
            this.lstModule.Location = new System.Drawing.Point(12, 69);
            this.lstModule.Name = "lstModule";
            this.lstModule.Size = new System.Drawing.Size(246, 342);
            this.lstModule.TabIndex = 3;
            this.lstModule.SelectedIndexChanged += new System.EventHandler(this.lstModule_SelectedIndexChanged);
            // 
            // FrmRenameVariable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.lstModule);
            this.Controls.Add(this.cmbModuleType);
            this.Controls.Add(this.lblModuleType);
            this.Controls.Add(this.dgVariables);
            this.Controls.Add(this.lblVariables);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblModule);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmRenameVariable";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmRenameVariable";
            this.Load += new System.EventHandler(this.FrmRenameVariable_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmRenameVariable_KeyDown);
            this.Resize += new System.EventHandler(this.FrmRenameVariable_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgVariables)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblModule;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblVariables;
        private System.Windows.Forms.DataGridView dgVariables;
        private System.Windows.Forms.Label lblModuleType;
        private System.Windows.Forms.ComboBox cmbModuleType;
        private System.Windows.Forms.ListBox lstModule;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColType;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColFunction;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColIndex;
    }
}