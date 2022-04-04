namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_TextFile
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_TextFile));
            this.btnCancel = new System.Windows.Forms.Button();
            this.grpDirection = new System.Windows.Forms.GroupBox();
            this.optDirectionWrite = new System.Windows.Forms.RadioButton();
            this.optDirectionRead = new System.Windows.Forms.RadioButton();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.txtVariableName = new System.Windows.Forms.TextBox();
            this.lblVariableName = new System.Windows.Forms.Label();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.ofdBrowseOpen = new System.Windows.Forms.OpenFileDialog();
            this.ofdBrowseSave = new System.Windows.Forms.SaveFileDialog();
            this.optFixedFieldLength = new System.Windows.Forms.RadioButton();
            this.optDelimited = new System.Windows.Forms.RadioButton();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.dgFixedColumns = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColStartChar = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColDataType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.txtDelimiterOther = new System.Windows.Forms.TextBox();
            this.optSemiColon = new System.Windows.Forms.RadioButton();
            this.optColon = new System.Windows.Forms.RadioButton();
            this.optComma = new System.Windows.Forms.RadioButton();
            this.optOther = new System.Windows.Forms.RadioButton();
            this.optTab = new System.Windows.Forms.RadioButton();
            this.grpDelimiter = new System.Windows.Forms.GroupBox();
            this.btnAddColumn = new System.Windows.Forms.Button();
            this.chkAutoupdatePositions = new System.Windows.Forms.CheckBox();
            this.chkAddReferences = new System.Windows.Forms.CheckBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.bgwStatusUpdater = new System.ComponentModel.BackgroundWorker();
            this.grpDirection.SuspendLayout();
            this.grpType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFixedColumns)).BeginInit();
            this.grpDelimiter.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(523, 410);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(89, 22);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grpDirection
            // 
            this.grpDirection.Controls.Add(this.optDirectionWrite);
            this.grpDirection.Controls.Add(this.optDirectionRead);
            this.grpDirection.Location = new System.Drawing.Point(12, 12);
            this.grpDirection.Name = "grpDirection";
            this.grpDirection.Size = new System.Drawing.Size(104, 74);
            this.grpDirection.TabIndex = 0;
            this.grpDirection.TabStop = false;
            this.grpDirection.Text = "Direction";
            // 
            // optDirectionWrite
            // 
            this.optDirectionWrite.AutoSize = true;
            this.optDirectionWrite.Location = new System.Drawing.Point(6, 44);
            this.optDirectionWrite.Name = "optDirectionWrite";
            this.optDirectionWrite.Size = new System.Drawing.Size(50, 17);
            this.optDirectionWrite.TabIndex = 1;
            this.optDirectionWrite.TabStop = true;
            this.optDirectionWrite.Text = "Write";
            this.optDirectionWrite.UseVisualStyleBackColor = true;
            this.optDirectionWrite.CheckedChanged += new System.EventHandler(this.optDirectionWrite_CheckedChanged);
            // 
            // optDirectionRead
            // 
            this.optDirectionRead.AutoSize = true;
            this.optDirectionRead.Location = new System.Drawing.Point(6, 19);
            this.optDirectionRead.Name = "optDirectionRead";
            this.optDirectionRead.Size = new System.Drawing.Size(51, 17);
            this.optDirectionRead.TabIndex = 0;
            this.optDirectionRead.TabStop = true;
            this.optDirectionRead.Text = "Read";
            this.optDirectionRead.UseVisualStyleBackColor = true;
            this.optDirectionRead.CheckedChanged += new System.EventHandler(this.optDirectionRead_CheckedChanged);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(442, 409);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 12;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(12, 115);
            this.txtPath.Multiline = true;
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(610, 42);
            this.txtPath.TabIndex = 6;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(9, 99);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(29, 13);
            this.lblPath.TabIndex = 5;
            this.lblPath.Text = "Path";
            // 
            // txtVariableName
            // 
            this.txtVariableName.Location = new System.Drawing.Point(256, 33);
            this.txtVariableName.Name = "txtVariableName";
            this.txtVariableName.Size = new System.Drawing.Size(366, 20);
            this.txtVariableName.TabIndex = 3;
            // 
            // lblVariableName
            // 
            this.lblVariableName.AutoSize = true;
            this.lblVariableName.Location = new System.Drawing.Point(253, 12);
            this.lblVariableName.Name = "lblVariableName";
            this.lblVariableName.Size = new System.Drawing.Size(70, 13);
            this.lblVariableName.TabIndex = 2;
            this.lblVariableName.Text = "Name (Suffix)";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(547, 86);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.Text = "&Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // optFixedFieldLength
            // 
            this.optFixedFieldLength.AutoSize = true;
            this.optFixedFieldLength.Location = new System.Drawing.Point(6, 42);
            this.optFixedFieldLength.Name = "optFixedFieldLength";
            this.optFixedFieldLength.Size = new System.Drawing.Size(111, 17);
            this.optFixedFieldLength.TabIndex = 1;
            this.optFixedFieldLength.TabStop = true;
            this.optFixedFieldLength.Text = "Fixed Field Length";
            this.optFixedFieldLength.UseVisualStyleBackColor = true;
            this.optFixedFieldLength.CheckedChanged += new System.EventHandler(this.optFixedFieldLength_CheckedChanged);
            // 
            // optDelimited
            // 
            this.optDelimited.AutoSize = true;
            this.optDelimited.Location = new System.Drawing.Point(6, 19);
            this.optDelimited.Name = "optDelimited";
            this.optDelimited.Size = new System.Drawing.Size(68, 17);
            this.optDelimited.TabIndex = 0;
            this.optDelimited.TabStop = true;
            this.optDelimited.Text = "Delimited";
            this.optDelimited.UseVisualStyleBackColor = true;
            this.optDelimited.CheckedChanged += new System.EventHandler(this.optDelimited_CheckedChanged);
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optFixedFieldLength);
            this.grpType.Controls.Add(this.optDelimited);
            this.grpType.Location = new System.Drawing.Point(122, 12);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(125, 74);
            this.grpType.TabIndex = 1;
            this.grpType.TabStop = false;
            this.grpType.Text = "Type";
            // 
            // dgFixedColumns
            // 
            this.dgFixedColumns.AllowUserToAddRows = false;
            this.dgFixedColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFixedColumns.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColStartChar,
            this.ColSize,
            this.ColDataType});
            this.dgFixedColumns.Location = new System.Drawing.Point(12, 173);
            this.dgFixedColumns.Name = "dgFixedColumns";
            this.dgFixedColumns.Size = new System.Drawing.Size(610, 150);
            this.dgFixedColumns.TabIndex = 3;
            this.dgFixedColumns.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgFixedColumns_CellEndEdit);
            this.dgFixedColumns.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgFixedColumns_RowsAdded);
            // 
            // ColName
            // 
            this.ColName.FillWeight = 150F;
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            // 
            // ColStartChar
            // 
            this.ColStartChar.HeaderText = "Start Char";
            this.ColStartChar.Name = "ColStartChar";
            this.ColStartChar.Width = 50;
            // 
            // ColSize
            // 
            this.ColSize.HeaderText = "Char Count";
            this.ColSize.Name = "ColSize";
            this.ColSize.Width = 50;
            // 
            // ColDataType
            // 
            this.ColDataType.HeaderText = "Data Type";
            this.ColDataType.Name = "ColDataType";
            this.ColDataType.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColDataType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // txtDelimiterOther
            // 
            this.txtDelimiterOther.Location = new System.Drawing.Point(209, 19);
            this.txtDelimiterOther.MaxLength = 1;
            this.txtDelimiterOther.Name = "txtDelimiterOther";
            this.txtDelimiterOther.Size = new System.Drawing.Size(38, 20);
            this.txtDelimiterOther.TabIndex = 5;
            // 
            // optSemiColon
            // 
            this.optSemiColon.AutoSize = true;
            this.optSemiColon.Location = new System.Drawing.Point(9, 88);
            this.optSemiColon.Name = "optSemiColon";
            this.optSemiColon.Size = new System.Drawing.Size(78, 17);
            this.optSemiColon.TabIndex = 3;
            this.optSemiColon.TabStop = true;
            this.optSemiColon.Text = "Semi Colon";
            this.optSemiColon.UseVisualStyleBackColor = true;
            this.optSemiColon.CheckedChanged += new System.EventHandler(this.optSemiColon_CheckedChanged);
            // 
            // optColon
            // 
            this.optColon.AutoSize = true;
            this.optColon.Location = new System.Drawing.Point(9, 65);
            this.optColon.Name = "optColon";
            this.optColon.Size = new System.Drawing.Size(52, 17);
            this.optColon.TabIndex = 2;
            this.optColon.TabStop = true;
            this.optColon.Text = "Colon";
            this.optColon.UseVisualStyleBackColor = true;
            this.optColon.CheckedChanged += new System.EventHandler(this.optColon_CheckedChanged);
            // 
            // optComma
            // 
            this.optComma.AutoSize = true;
            this.optComma.Location = new System.Drawing.Point(9, 42);
            this.optComma.Name = "optComma";
            this.optComma.Size = new System.Drawing.Size(60, 17);
            this.optComma.TabIndex = 1;
            this.optComma.TabStop = true;
            this.optComma.Text = "Comma";
            this.optComma.UseVisualStyleBackColor = true;
            this.optComma.CheckedChanged += new System.EventHandler(this.optComma_CheckedChanged);
            // 
            // optOther
            // 
            this.optOther.AutoSize = true;
            this.optOther.Location = new System.Drawing.Point(152, 19);
            this.optOther.Name = "optOther";
            this.optOther.Size = new System.Drawing.Size(51, 17);
            this.optOther.TabIndex = 4;
            this.optOther.TabStop = true;
            this.optOther.Text = "Other";
            this.optOther.UseVisualStyleBackColor = true;
            this.optOther.CheckedChanged += new System.EventHandler(this.optOther_CheckedChanged);
            // 
            // optTab
            // 
            this.optTab.AutoSize = true;
            this.optTab.Location = new System.Drawing.Point(9, 19);
            this.optTab.Name = "optTab";
            this.optTab.Size = new System.Drawing.Size(44, 17);
            this.optTab.TabIndex = 0;
            this.optTab.TabStop = true;
            this.optTab.Text = "Tab";
            this.optTab.UseVisualStyleBackColor = true;
            this.optTab.CheckedChanged += new System.EventHandler(this.optTab_CheckedChanged);
            // 
            // grpDelimiter
            // 
            this.grpDelimiter.Controls.Add(this.txtDelimiterOther);
            this.grpDelimiter.Controls.Add(this.optSemiColon);
            this.grpDelimiter.Controls.Add(this.optColon);
            this.grpDelimiter.Controls.Add(this.optComma);
            this.grpDelimiter.Controls.Add(this.optOther);
            this.grpDelimiter.Controls.Add(this.optTab);
            this.grpDelimiter.Location = new System.Drawing.Point(12, 173);
            this.grpDelimiter.Name = "grpDelimiter";
            this.grpDelimiter.Size = new System.Drawing.Size(610, 150);
            this.grpDelimiter.TabIndex = 8;
            this.grpDelimiter.TabStop = false;
            this.grpDelimiter.Text = "Delimiter";
            // 
            // btnAddColumn
            // 
            this.btnAddColumn.Location = new System.Drawing.Point(18, 409);
            this.btnAddColumn.Name = "btnAddColumn";
            this.btnAddColumn.Size = new System.Drawing.Size(75, 23);
            this.btnAddColumn.TabIndex = 9;
            this.btnAddColumn.Text = "&Add";
            this.btnAddColumn.UseVisualStyleBackColor = true;
            this.btnAddColumn.Click += new System.EventHandler(this.btnAddColumn_Click);
            // 
            // chkAutoupdatePositions
            // 
            this.chkAutoupdatePositions.AutoSize = true;
            this.chkAutoupdatePositions.Location = new System.Drawing.Point(107, 414);
            this.chkAutoupdatePositions.Name = "chkAutoupdatePositions";
            this.chkAutoupdatePositions.Size = new System.Drawing.Size(126, 17);
            this.chkAutoupdatePositions.TabIndex = 10;
            this.chkAutoupdatePositions.Text = "Autoupdate Positions";
            this.chkAutoupdatePositions.UseVisualStyleBackColor = true;
            // 
            // chkAddReferences
            // 
            this.chkAddReferences.AutoSize = true;
            this.chkAddReferences.Location = new System.Drawing.Point(326, 410);
            this.chkAddReferences.Name = "chkAddReferences";
            this.chkAddReferences.Size = new System.Drawing.Size(103, 17);
            this.chkAddReferences.TabIndex = 11;
            this.chkAddReferences.Text = "Add References";
            this.chkAddReferences.UseVisualStyleBackColor = true;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 14;
            this.ssStatus.Text = "status";
            // 
            // bgwStatusUpdater
            // 
            this.bgwStatusUpdater.WorkerReportsProgress = true;
            this.bgwStatusUpdater.WorkerSupportsCancellation = true;
            this.bgwStatusUpdater.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwStatusUpdater_DoWork);
            // 
            // FrmInsertCode_TextFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.chkAddReferences);
            this.Controls.Add(this.chkAutoupdatePositions);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.btnAddColumn);
            this.Controls.Add(this.lblVariableName);
            this.Controls.Add(this.grpDelimiter);
            this.Controls.Add(this.txtVariableName);
            this.Controls.Add(this.grpType);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.dgFixedColumns);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.grpDirection);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_TextFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_TextFile_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_TextFile_KeyDown);
            this.grpDirection.ResumeLayout(false);
            this.grpDirection.PerformLayout();
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFixedColumns)).EndInit();
            this.grpDelimiter.ResumeLayout(false);
            this.grpDelimiter.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox grpDirection;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.RadioButton optDirectionWrite;
        private System.Windows.Forms.RadioButton optDirectionRead;
        private System.Windows.Forms.TextBox txtVariableName;
        private System.Windows.Forms.Label lblVariableName;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.OpenFileDialog ofdBrowseOpen;
        private System.Windows.Forms.SaveFileDialog ofdBrowseSave;
        private System.Windows.Forms.RadioButton optFixedFieldLength;
        private System.Windows.Forms.RadioButton optDelimited;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.DataGridView dgFixedColumns;
        private System.Windows.Forms.TextBox txtDelimiterOther;
        private System.Windows.Forms.RadioButton optSemiColon;
        private System.Windows.Forms.RadioButton optColon;
        private System.Windows.Forms.RadioButton optComma;
        private System.Windows.Forms.RadioButton optOther;
        private System.Windows.Forms.RadioButton optTab;
        private System.Windows.Forms.GroupBox grpDelimiter;
        private System.Windows.Forms.Button btnAddColumn;
        private System.Windows.Forms.CheckBox chkAutoupdatePositions;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColStartChar;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSize;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDataType;
        private System.Windows.Forms.CheckBox chkAddReferences;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.ComponentModel.BackgroundWorker bgwStatusUpdater;
    }
}