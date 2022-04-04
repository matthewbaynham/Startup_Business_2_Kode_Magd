namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_ExcuteStoredProcedure
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_ExcuteStoredProcedure));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtSPName = new System.Windows.Forms.TextBox();
            this.chkExecuteAsynchronously = new System.Windows.Forms.CheckBox();
            this.lblSPName = new System.Windows.Forms.Label();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.txtConnectionString = new System.Windows.Forms.TextBox();
            this.btnRecent = new System.Windows.Forms.Button();
            this.dgParameters = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColDirection = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColVariable = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnAdd = new System.Windows.Forms.Button();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.chkAddReferences = new System.Windows.Forms.CheckBox();
            this.btnBuildConnectionString = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 395);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(464, 395);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 10;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtSPName
            // 
            this.txtSPName.Location = new System.Drawing.Point(133, 9);
            this.txtSPName.Name = "txtSPName";
            this.txtSPName.Size = new System.Drawing.Size(339, 20);
            this.txtSPName.TabIndex = 1;
            // 
            // chkExecuteAsynchronously
            // 
            this.chkExecuteAsynchronously.AutoSize = true;
            this.chkExecuteAsynchronously.Location = new System.Drawing.Point(478, 12);
            this.chkExecuteAsynchronously.Name = "chkExecuteAsynchronously";
            this.chkExecuteAsynchronously.Size = new System.Drawing.Size(142, 17);
            this.chkExecuteAsynchronously.TabIndex = 2;
            this.chkExecuteAsynchronously.Text = "Execute Asynchronously";
            this.chkExecuteAsynchronously.UseVisualStyleBackColor = true;
            // 
            // lblSPName
            // 
            this.lblSPName.AutoSize = true;
            this.lblSPName.Location = new System.Drawing.Point(12, 9);
            this.lblSPName.Name = "lblSPName";
            this.lblSPName.Size = new System.Drawing.Size(115, 13);
            this.lblSPName.TabIndex = 0;
            this.lblSPName.Text = "Store Procedure Name";
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(12, 52);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 3;
            this.lblConnectionString.Text = "Connection String";
            // 
            // txtConnectionString
            // 
            this.txtConnectionString.Location = new System.Drawing.Point(12, 80);
            this.txtConnectionString.Multiline = true;
            this.txtConnectionString.Name = "txtConnectionString";
            this.txtConnectionString.Size = new System.Drawing.Size(607, 62);
            this.txtConnectionString.TabIndex = 6;
            // 
            // btnRecent
            // 
            this.btnRecent.Location = new System.Drawing.Point(121, 47);
            this.btnRecent.Name = "btnRecent";
            this.btnRecent.Size = new System.Drawing.Size(75, 23);
            this.btnRecent.TabIndex = 4;
            this.btnRecent.Text = "&Recent";
            this.btnRecent.UseVisualStyleBackColor = true;
            this.btnRecent.Click += new System.EventHandler(this.btnRecent_Click);
            // 
            // dgParameters
            // 
            this.dgParameters.AllowUserToAddRows = false;
            this.dgParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColType,
            this.ColSize,
            this.ColDirection,
            this.ColVariable,
            this.ColValue});
            this.dgParameters.Location = new System.Drawing.Point(14, 155);
            this.dgParameters.Name = "dgParameters";
            this.dgParameters.Size = new System.Drawing.Size(605, 223);
            this.dgParameters.TabIndex = 7;
            this.dgParameters.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgParameters_CellValidating);
            this.dgParameters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgParameters_RowsAdded);
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
            // 
            // ColSize
            // 
            this.ColSize.HeaderText = "Size";
            this.ColSize.Name = "ColSize";
            this.ColSize.Width = 50;
            // 
            // ColDirection
            // 
            this.ColDirection.HeaderText = "Direction";
            this.ColDirection.Name = "ColDirection";
            // 
            // ColVariable
            // 
            this.ColVariable.HeaderText = "Variable";
            this.ColVariable.Name = "ColVariable";
            this.ColVariable.Width = 50;
            // 
            // ColValue
            // 
            this.ColValue.HeaderText = "Value";
            this.ColValue.Name = "ColValue";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(12, 395);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(60, 23);
            this.btnAdd.TabIndex = 8;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 12;
            this.ssStatus.Text = "status";
            // 
            // chkAddReferences
            // 
            this.chkAddReferences.AutoSize = true;
            this.chkAddReferences.Location = new System.Drawing.Point(355, 399);
            this.chkAddReferences.Name = "chkAddReferences";
            this.chkAddReferences.Size = new System.Drawing.Size(103, 17);
            this.chkAddReferences.TabIndex = 9;
            this.chkAddReferences.Text = "Add References";
            this.chkAddReferences.UseVisualStyleBackColor = true;
            // 
            // btnBuildConnectionString
            // 
            this.btnBuildConnectionString.Location = new System.Drawing.Point(208, 49);
            this.btnBuildConnectionString.Name = "btnBuildConnectionString";
            this.btnBuildConnectionString.Size = new System.Drawing.Size(75, 23);
            this.btnBuildConnectionString.TabIndex = 5;
            this.btnBuildConnectionString.Text = "&Build";
            this.btnBuildConnectionString.UseVisualStyleBackColor = true;
            this.btnBuildConnectionString.Click += new System.EventHandler(this.btnBuildConnectionString_Click);
            // 
            // FrmInsertCode_ExcuteStoredProcedure
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.btnBuildConnectionString);
            this.Controls.Add(this.chkAddReferences);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.dgParameters);
            this.Controls.Add(this.btnRecent);
            this.Controls.Add(this.txtConnectionString);
            this.Controls.Add(this.lblConnectionString);
            this.Controls.Add(this.lblSPName);
            this.Controls.Add(this.chkExecuteAsynchronously);
            this.Controls.Add(this.txtSPName);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_ExcuteStoredProcedure";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_ExcuteStoredProcedure_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_ExcuteStoredProcedure_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_ExcuteStoredProcedure_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtSPName;
        private System.Windows.Forms.CheckBox chkExecuteAsynchronously;
        private System.Windows.Forms.Label lblSPName;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TextBox txtConnectionString;
        private System.Windows.Forms.Button btnRecent;
        private System.Windows.Forms.DataGridView dgParameters;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColType;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSize;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDirection;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColVariable;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValue;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.CheckBox chkAddReferences;
        private System.Windows.Forms.Button btnBuildConnectionString;
    }
}