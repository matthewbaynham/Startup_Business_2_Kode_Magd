namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_PivotTable
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_PivotTable));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.grpSourceType = new System.Windows.Forms.GroupBox();
            this.optNamedRange = new System.Windows.Forms.RadioButton();
            this.optDatabase = new System.Windows.Forms.RadioButton();
            this.optSelectedRange = new System.Windows.Forms.RadioButton();
            this.lblSource = new System.Windows.Forms.Label();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.txtConnectionString = new System.Windows.Forms.TextBox();
            this.dgFields = new System.Windows.Forms.DataGridView();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOrientation = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.lblFields = new System.Windows.Forms.Label();
            this.cmbSource = new System.Windows.Forms.ComboBox();
            this.lblCommandType = new System.Windows.Forms.Label();
            this.cmbCommandType = new System.Windows.Forms.ComboBox();
            this.btnConnectionStringExpend = new System.Windows.Forms.Button();
            this.btnSourceExpand = new System.Windows.Forms.Button();
            this.lblWarning = new System.Windows.Forms.Label();
            this.btnConnectionStringRecent = new System.Windows.Forms.Button();
            this.chkNewSheet = new System.Windows.Forms.CheckBox();
            this.lblDestination = new System.Windows.Forms.Label();
            this.cmbSheetName = new System.Windows.Forms.ComboBox();
            this.txtSheetName = new System.Windows.Forms.TextBox();
            this.lblSheetName = new System.Windows.Forms.Label();
            this.lblAddress = new System.Windows.Forms.Label();
            this.txtAddress = new System.Windows.Forms.TextBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnConnectionStringBuild = new System.Windows.Forms.Button();
            this.grpSourceType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 402);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 22;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(454, 402);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 21;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // grpSourceType
            // 
            this.grpSourceType.Controls.Add(this.optNamedRange);
            this.grpSourceType.Controls.Add(this.optDatabase);
            this.grpSourceType.Controls.Add(this.optSelectedRange);
            this.grpSourceType.Location = new System.Drawing.Point(12, 12);
            this.grpSourceType.Name = "grpSourceType";
            this.grpSourceType.Size = new System.Drawing.Size(128, 91);
            this.grpSourceType.TabIndex = 0;
            this.grpSourceType.TabStop = false;
            this.grpSourceType.Text = "Source Type";
            // 
            // optNamedRange
            // 
            this.optNamedRange.AutoSize = true;
            this.optNamedRange.Location = new System.Drawing.Point(17, 42);
            this.optNamedRange.Name = "optNamedRange";
            this.optNamedRange.Size = new System.Drawing.Size(94, 17);
            this.optNamedRange.TabIndex = 1;
            this.optNamedRange.TabStop = true;
            this.optNamedRange.Text = "Named Range";
            this.optNamedRange.UseVisualStyleBackColor = true;
            this.optNamedRange.CheckedChanged += new System.EventHandler(this.optNamedRange_CheckedChanged);
            // 
            // optDatabase
            // 
            this.optDatabase.AutoSize = true;
            this.optDatabase.Location = new System.Drawing.Point(17, 65);
            this.optDatabase.Name = "optDatabase";
            this.optDatabase.Size = new System.Drawing.Size(71, 17);
            this.optDatabase.TabIndex = 2;
            this.optDatabase.TabStop = true;
            this.optDatabase.Text = "Database";
            this.optDatabase.UseVisualStyleBackColor = true;
            this.optDatabase.CheckedChanged += new System.EventHandler(this.optDatabase_CheckedChanged);
            // 
            // optSelectedRange
            // 
            this.optSelectedRange.AutoSize = true;
            this.optSelectedRange.Location = new System.Drawing.Point(17, 19);
            this.optSelectedRange.Name = "optSelectedRange";
            this.optSelectedRange.Size = new System.Drawing.Size(102, 17);
            this.optSelectedRange.TabIndex = 0;
            this.optSelectedRange.TabStop = true;
            this.optSelectedRange.Text = "Selected Range";
            this.optSelectedRange.UseVisualStyleBackColor = true;
            this.optSelectedRange.CheckedChanged += new System.EventHandler(this.optSelectedRange_CheckedChanged);
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(10, 117);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 7;
            this.lblSource.Text = "Source";
            // 
            // txtSource
            // 
            this.txtSource.AcceptsTab = true;
            this.txtSource.AllowDrop = true;
            this.txtSource.Location = new System.Drawing.Point(13, 133);
            this.txtSource.Multiline = true;
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(282, 94);
            this.txtSource.TabIndex = 9;
            this.txtSource.TextChanged += new System.EventHandler(this.txtSource_TextChanged);
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(9, 271);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 13;
            this.lblConnectionString.Text = "Connection String";
            // 
            // txtConnectionString
            // 
            this.txtConnectionString.AcceptsTab = true;
            this.txtConnectionString.AllowDrop = true;
            this.txtConnectionString.Location = new System.Drawing.Point(12, 287);
            this.txtConnectionString.Multiline = true;
            this.txtConnectionString.Name = "txtConnectionString";
            this.txtConnectionString.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtConnectionString.Size = new System.Drawing.Size(283, 68);
            this.txtConnectionString.TabIndex = 14;
            this.txtConnectionString.TextChanged += new System.EventHandler(this.txtConnectionString_TextChanged);
            // 
            // dgFields
            // 
            this.dgFields.AllowUserToAddRows = false;
            this.dgFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colName,
            this.colOrientation});
            this.dgFields.Location = new System.Drawing.Point(363, 31);
            this.dgFields.Name = "dgFields";
            this.dgFields.Size = new System.Drawing.Size(256, 357);
            this.dgFields.TabIndex = 20;
            this.dgFields.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgFields_RowsAdded);
            // 
            // colName
            // 
            this.colName.HeaderText = "Name";
            this.colName.Name = "colName";
            // 
            // colOrientation
            // 
            this.colOrientation.HeaderText = "Orientation";
            this.colOrientation.Name = "colOrientation";
            // 
            // lblFields
            // 
            this.lblFields.AutoSize = true;
            this.lblFields.Location = new System.Drawing.Point(360, 9);
            this.lblFields.Name = "lblFields";
            this.lblFields.Size = new System.Drawing.Size(34, 13);
            this.lblFields.TabIndex = 19;
            this.lblFields.Text = "Fields";
            // 
            // cmbSource
            // 
            this.cmbSource.FormattingEnabled = true;
            this.cmbSource.Location = new System.Drawing.Point(12, 133);
            this.cmbSource.Name = "cmbSource";
            this.cmbSource.Size = new System.Drawing.Size(283, 21);
            this.cmbSource.TabIndex = 8;
            this.cmbSource.SelectedIndexChanged += new System.EventHandler(this.cmbSource_SelectedIndexChanged);
            // 
            // lblCommandType
            // 
            this.lblCommandType.AutoSize = true;
            this.lblCommandType.Location = new System.Drawing.Point(10, 231);
            this.lblCommandType.Name = "lblCommandType";
            this.lblCommandType.Size = new System.Drawing.Size(81, 13);
            this.lblCommandType.TabIndex = 11;
            this.lblCommandType.Text = "Command Type";
            // 
            // cmbCommandType
            // 
            this.cmbCommandType.FormattingEnabled = true;
            this.cmbCommandType.Location = new System.Drawing.Point(13, 247);
            this.cmbCommandType.Name = "cmbCommandType";
            this.cmbCommandType.Size = new System.Drawing.Size(282, 21);
            this.cmbCommandType.TabIndex = 12;
            this.cmbCommandType.TextChanged += new System.EventHandler(this.cmbCommandType_TextChanged);
            // 
            // btnConnectionStringExpend
            // 
            this.btnConnectionStringExpend.Location = new System.Drawing.Point(301, 287);
            this.btnConnectionStringExpend.Name = "btnConnectionStringExpend";
            this.btnConnectionStringExpend.Size = new System.Drawing.Size(55, 23);
            this.btnConnectionStringExpend.TabIndex = 15;
            this.btnConnectionStringExpend.Text = ". . .";
            this.btnConnectionStringExpend.UseVisualStyleBackColor = true;
            this.btnConnectionStringExpend.Click += new System.EventHandler(this.btnConnectionStringExpend_Click);
            // 
            // btnSourceExpand
            // 
            this.btnSourceExpand.Location = new System.Drawing.Point(301, 131);
            this.btnSourceExpand.Name = "btnSourceExpand";
            this.btnSourceExpand.Size = new System.Drawing.Size(56, 23);
            this.btnSourceExpand.TabIndex = 10;
            this.btnSourceExpand.Text = ". . .";
            this.btnSourceExpand.UseVisualStyleBackColor = true;
            this.btnSourceExpand.Click += new System.EventHandler(this.btnSourceExpand_Click);
            // 
            // lblWarning
            // 
            this.lblWarning.Location = new System.Drawing.Point(10, 367);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(347, 58);
            this.lblWarning.TabIndex = 18;
            this.lblWarning.Text = "Warning";
            // 
            // btnConnectionStringRecent
            // 
            this.btnConnectionStringRecent.Location = new System.Drawing.Point(301, 316);
            this.btnConnectionStringRecent.Name = "btnConnectionStringRecent";
            this.btnConnectionStringRecent.Size = new System.Drawing.Size(55, 23);
            this.btnConnectionStringRecent.TabIndex = 16;
            this.btnConnectionStringRecent.Text = "&Recent";
            this.btnConnectionStringRecent.UseVisualStyleBackColor = true;
            this.btnConnectionStringRecent.Click += new System.EventHandler(this.btnConnectionStringRecent_Click);
            // 
            // chkNewSheet
            // 
            this.chkNewSheet.AutoSize = true;
            this.chkNewSheet.Location = new System.Drawing.Point(275, 12);
            this.chkNewSheet.Name = "chkNewSheet";
            this.chkNewSheet.Size = new System.Drawing.Size(79, 17);
            this.chkNewSheet.TabIndex = 3;
            this.chkNewSheet.Text = "New Sheet";
            this.chkNewSheet.UseVisualStyleBackColor = true;
            this.chkNewSheet.CheckedChanged += new System.EventHandler(this.chkNewSheet_CheckedChanged);
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(148, 8);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(60, 13);
            this.lblDestination.TabIndex = 1;
            this.lblDestination.Text = "Destination";
            // 
            // cmbSheetName
            // 
            this.cmbSheetName.FormattingEnabled = true;
            this.cmbSheetName.Location = new System.Drawing.Point(146, 44);
            this.cmbSheetName.Name = "cmbSheetName";
            this.cmbSheetName.Size = new System.Drawing.Size(208, 21);
            this.cmbSheetName.TabIndex = 5;
            // 
            // txtSheetName
            // 
            this.txtSheetName.Location = new System.Drawing.Point(146, 44);
            this.txtSheetName.Name = "txtSheetName";
            this.txtSheetName.Size = new System.Drawing.Size(208, 20);
            this.txtSheetName.TabIndex = 4;
            // 
            // lblSheetName
            // 
            this.lblSheetName.AutoSize = true;
            this.lblSheetName.Location = new System.Drawing.Point(147, 28);
            this.lblSheetName.Name = "lblSheetName";
            this.lblSheetName.Size = new System.Drawing.Size(66, 13);
            this.lblSheetName.TabIndex = 2;
            this.lblSheetName.Text = "Sheet Name";
            // 
            // lblAddress
            // 
            this.lblAddress.AutoSize = true;
            this.lblAddress.Location = new System.Drawing.Point(149, 72);
            this.lblAddress.Name = "lblAddress";
            this.lblAddress.Size = new System.Drawing.Size(45, 13);
            this.lblAddress.TabIndex = 4;
            this.lblAddress.Text = "Address";
            // 
            // txtAddress
            // 
            this.txtAddress.Location = new System.Drawing.Point(146, 89);
            this.txtAddress.Name = "txtAddress";
            this.txtAddress.Size = new System.Drawing.Size(211, 20);
            this.txtAddress.TabIndex = 5;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 23;
            this.ssStatus.Text = "status";
            // 
            // btnConnectionStringBuild
            // 
            this.btnConnectionStringBuild.Location = new System.Drawing.Point(301, 345);
            this.btnConnectionStringBuild.Name = "btnConnectionStringBuild";
            this.btnConnectionStringBuild.Size = new System.Drawing.Size(55, 23);
            this.btnConnectionStringBuild.TabIndex = 17;
            this.btnConnectionStringBuild.Text = "&Build";
            this.btnConnectionStringBuild.UseVisualStyleBackColor = true;
            this.btnConnectionStringBuild.Click += new System.EventHandler(this.btnBuildConnectionString_Click);
            // 
            // FrmInsertCode_PivotTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.btnConnectionStringBuild);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.txtAddress);
            this.Controls.Add(this.lblAddress);
            this.Controls.Add(this.lblSheetName);
            this.Controls.Add(this.txtSheetName);
            this.Controls.Add(this.cmbSheetName);
            this.Controls.Add(this.lblDestination);
            this.Controls.Add(this.chkNewSheet);
            this.Controls.Add(this.btnConnectionStringRecent);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.btnSourceExpand);
            this.Controls.Add(this.btnConnectionStringExpend);
            this.Controls.Add(this.cmbCommandType);
            this.Controls.Add(this.lblCommandType);
            this.Controls.Add(this.cmbSource);
            this.Controls.Add(this.lblFields);
            this.Controls.Add(this.dgFields);
            this.Controls.Add(this.txtConnectionString);
            this.Controls.Add(this.lblConnectionString);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.grpSourceType);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_PivotTable";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_PivotTable_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_PivotTable_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_PivotTable_Resize);
            this.grpSourceType.ResumeLayout(false);
            this.grpSourceType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox grpSourceType;
        private System.Windows.Forms.RadioButton optNamedRange;
        private System.Windows.Forms.RadioButton optDatabase;
        private System.Windows.Forms.RadioButton optSelectedRange;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TextBox txtConnectionString;
        private System.Windows.Forms.DataGridView dgFields;
        private System.Windows.Forms.Label lblFields;
        private System.Windows.Forms.ComboBox cmbSource;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewComboBoxColumn colOrientation;
        private System.Windows.Forms.Label lblCommandType;
        private System.Windows.Forms.ComboBox cmbCommandType;
        private System.Windows.Forms.Button btnConnectionStringExpend;
        private System.Windows.Forms.Button btnSourceExpand;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Button btnConnectionStringRecent;
        private System.Windows.Forms.CheckBox chkNewSheet;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.ComboBox cmbSheetName;
        private System.Windows.Forms.TextBox txtSheetName;
        private System.Windows.Forms.Label lblSheetName;
        private System.Windows.Forms.Label lblAddress;
        private System.Windows.Forms.TextBox txtAddress;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnConnectionStringBuild;
    }
}