namespace KodeMagd
{
    partial class FrmInsertCode_Rst
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_Rst));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.lblName = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.txtConnectionString = new System.Windows.Forms.TextBox();
            this.btnBuildConnectionString = new System.Windows.Forms.Button();
            this.grpSourceType = new System.Windows.Forms.GroupBox();
            this.optSql = new System.Windows.Forms.RadioButton();
            this.optStoreProcedure = new System.Windows.Forms.RadioButton();
            this.dgParameters = new System.Windows.Forms.DataGridView();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColAssignVariable = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colValue = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.lblParameters = new System.Windows.Forms.Label();
            this.lblSource = new System.Windows.Forms.Label();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.lblWarning = new System.Windows.Forms.Label();
            this.btnOptions = new System.Windows.Forms.Button();
            this.chkMultipleRstReturned = new System.Windows.Forms.CheckBox();
            this.btnRecent = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.grpDestinatin = new System.Windows.Forms.GroupBox();
            this.optEmptyLoop = new System.Windows.Forms.RadioButton();
            this.optListboxCombo = new System.Windows.Forms.RadioButton();
            this.optRange = new System.Windows.Forms.RadioButton();
            this.btnDestinationDetails = new System.Windows.Forms.Button();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnRemove = new System.Windows.Forms.Button();
            this.chkAddReference = new System.Windows.Forms.CheckBox();
            this.grpSourceType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).BeginInit();
            this.grpDestinatin.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(568, 395);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 31);
            this.btnCancel.TabIndex = 20;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(500, 395);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(62, 31);
            this.btnGenerate.TabIndex = 19;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(15, 10);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(70, 13);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "Name (Suffix)";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(119, 10);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(177, 20);
            this.txtName.TabIndex = 1;
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(12, 37);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 2;
            this.lblConnectionString.Text = "Connection String";
            // 
            // txtConnectionString
            // 
            this.txtConnectionString.Location = new System.Drawing.Point(119, 37);
            this.txtConnectionString.Multiline = true;
            this.txtConnectionString.Name = "txtConnectionString";
            this.txtConnectionString.Size = new System.Drawing.Size(424, 56);
            this.txtConnectionString.TabIndex = 3;
            // 
            // btnBuildConnectionString
            // 
            this.btnBuildConnectionString.Location = new System.Drawing.Point(549, 37);
            this.btnBuildConnectionString.Name = "btnBuildConnectionString";
            this.btnBuildConnectionString.Size = new System.Drawing.Size(75, 23);
            this.btnBuildConnectionString.TabIndex = 4;
            this.btnBuildConnectionString.Text = "&Build";
            this.btnBuildConnectionString.UseVisualStyleBackColor = true;
            this.btnBuildConnectionString.Click += new System.EventHandler(this.btnBuildConnectionString_Click);
            // 
            // grpSourceType
            // 
            this.grpSourceType.Controls.Add(this.optSql);
            this.grpSourceType.Controls.Add(this.optStoreProcedure);
            this.grpSourceType.Location = new System.Drawing.Point(12, 99);
            this.grpSourceType.Name = "grpSourceType";
            this.grpSourceType.Size = new System.Drawing.Size(124, 68);
            this.grpSourceType.TabIndex = 7;
            this.grpSourceType.TabStop = false;
            this.grpSourceType.Text = "Source";
            // 
            // optSql
            // 
            this.optSql.AutoSize = true;
            this.optSql.Location = new System.Drawing.Point(6, 42);
            this.optSql.Name = "optSql";
            this.optSql.Size = new System.Drawing.Size(46, 17);
            this.optSql.TabIndex = 1;
            this.optSql.TabStop = true;
            this.optSql.Text = "SQL";
            this.optSql.UseVisualStyleBackColor = true;
            this.optSql.Click += new System.EventHandler(this.optSql_Click);
            // 
            // optStoreProcedure
            // 
            this.optStoreProcedure.AutoSize = true;
            this.optStoreProcedure.Location = new System.Drawing.Point(6, 21);
            this.optStoreProcedure.Name = "optStoreProcedure";
            this.optStoreProcedure.Size = new System.Drawing.Size(102, 17);
            this.optStoreProcedure.TabIndex = 0;
            this.optStoreProcedure.TabStop = true;
            this.optStoreProcedure.Text = "Store Procedure";
            this.optStoreProcedure.UseVisualStyleBackColor = true;
            this.optStoreProcedure.Click += new System.EventHandler(this.optStoreProcedure_Click);
            // 
            // dgParameters
            // 
            this.dgParameters.AllowUserToAddRows = false;
            this.dgParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colName,
            this.colType,
            this.colSize,
            this.ColAssignVariable,
            this.colValue});
            this.dgParameters.Location = new System.Drawing.Point(10, 251);
            this.dgParameters.Name = "dgParameters";
            this.dgParameters.Size = new System.Drawing.Size(614, 102);
            this.dgParameters.TabIndex = 13;
            this.dgParameters.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgParameters_CellEndEdit);
            this.dgParameters.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgParameters_CellValueChanged);
            this.dgParameters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgParameters_RowsAdded);
            // 
            // colName
            // 
            this.colName.HeaderText = "Name";
            this.colName.Name = "colName";
            // 
            // colType
            // 
            this.colType.HeaderText = "Type";
            this.colType.Name = "colType";
            // 
            // colSize
            // 
            this.colSize.HeaderText = "Size";
            this.colSize.Name = "colSize";
            this.colSize.Width = 50;
            // 
            // ColAssignVariable
            // 
            this.ColAssignVariable.HeaderText = "Assign Variable";
            this.ColAssignVariable.Name = "ColAssignVariable";
            this.ColAssignVariable.Width = 50;
            // 
            // colValue
            // 
            this.colValue.HeaderText = "Value";
            this.colValue.Name = "colValue";
            // 
            // lblParameters
            // 
            this.lblParameters.AutoSize = true;
            this.lblParameters.Location = new System.Drawing.Point(12, 235);
            this.lblParameters.Name = "lblParameters";
            this.lblParameters.Size = new System.Drawing.Size(60, 13);
            this.lblParameters.TabIndex = 12;
            this.lblParameters.Text = "Parameters";
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(12, 180);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(35, 13);
            this.lblSource.TabIndex = 10;
            this.lblSource.Text = "label2";
            // 
            // txtSource
            // 
            this.txtSource.Location = new System.Drawing.Point(10, 196);
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(614, 20);
            this.txtSource.TabIndex = 11;
            // 
            // lblWarning
            // 
            this.lblWarning.Location = new System.Drawing.Point(9, 367);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(615, 25);
            this.lblWarning.TabIndex = 14;
            this.lblWarning.Text = "label2";
            // 
            // btnOptions
            // 
            this.btnOptions.Location = new System.Drawing.Point(438, 395);
            this.btnOptions.Name = "btnOptions";
            this.btnOptions.Size = new System.Drawing.Size(56, 31);
            this.btnOptions.TabIndex = 18;
            this.btnOptions.Text = "&Options";
            this.btnOptions.UseVisualStyleBackColor = true;
            this.btnOptions.Click += new System.EventHandler(this.btnOptions_Click);
            // 
            // chkMultipleRstReturned
            // 
            this.chkMultipleRstReturned.AutoSize = true;
            this.chkMultipleRstReturned.Location = new System.Drawing.Point(154, 120);
            this.chkMultipleRstReturned.Name = "chkMultipleRstReturned";
            this.chkMultipleRstReturned.Size = new System.Drawing.Size(166, 17);
            this.chkMultipleRstReturned.TabIndex = 8;
            this.chkMultipleRstReturned.Text = "Multiple Recordsets Returned";
            this.chkMultipleRstReturned.UseVisualStyleBackColor = true;
            this.chkMultipleRstReturned.Visible = false;
            // 
            // btnRecent
            // 
            this.btnRecent.Location = new System.Drawing.Point(549, 66);
            this.btnRecent.Name = "btnRecent";
            this.btnRecent.Size = new System.Drawing.Size(75, 27);
            this.btnRecent.TabIndex = 5;
            this.btnRecent.Text = "R&ecent";
            this.btnRecent.UseVisualStyleBackColor = true;
            this.btnRecent.Click += new System.EventHandler(this.btnRecent_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(10, 395);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(57, 27);
            this.btnAdd.TabIndex = 15;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // grpDestinatin
            // 
            this.grpDestinatin.Controls.Add(this.optEmptyLoop);
            this.grpDestinatin.Controls.Add(this.optListboxCombo);
            this.grpDestinatin.Controls.Add(this.optRange);
            this.grpDestinatin.Location = new System.Drawing.Point(368, 99);
            this.grpDestinatin.Name = "grpDestinatin";
            this.grpDestinatin.Size = new System.Drawing.Size(149, 91);
            this.grpDestinatin.TabIndex = 9;
            this.grpDestinatin.TabStop = false;
            this.grpDestinatin.Text = "Destination";
            // 
            // optEmptyLoop
            // 
            this.optEmptyLoop.AutoSize = true;
            this.optEmptyLoop.Location = new System.Drawing.Point(6, 66);
            this.optEmptyLoop.Name = "optEmptyLoop";
            this.optEmptyLoop.Size = new System.Drawing.Size(81, 17);
            this.optEmptyLoop.TabIndex = 2;
            this.optEmptyLoop.TabStop = true;
            this.optEmptyLoop.Text = "Empty Loop";
            this.optEmptyLoop.UseVisualStyleBackColor = true;
            this.optEmptyLoop.CheckedChanged += new System.EventHandler(this.optEmptyLoop_CheckedChanged);
            // 
            // optListboxCombo
            // 
            this.optListboxCombo.AutoSize = true;
            this.optListboxCombo.Location = new System.Drawing.Point(6, 42);
            this.optListboxCombo.Name = "optListboxCombo";
            this.optListboxCombo.Size = new System.Drawing.Size(106, 17);
            this.optListboxCombo.TabIndex = 1;
            this.optListboxCombo.TabStop = true;
            this.optListboxCombo.Text = "Listbox or Combo";
            this.optListboxCombo.UseVisualStyleBackColor = true;
            this.optListboxCombo.CheckedChanged += new System.EventHandler(this.optListboxCombo_CheckedChanged);
            // 
            // optRange
            // 
            this.optRange.AutoSize = true;
            this.optRange.Location = new System.Drawing.Point(6, 19);
            this.optRange.Name = "optRange";
            this.optRange.Size = new System.Drawing.Size(70, 17);
            this.optRange.TabIndex = 0;
            this.optRange.TabStop = true;
            this.optRange.Text = "On Sheet";
            this.optRange.UseVisualStyleBackColor = true;
            this.optRange.CheckedChanged += new System.EventHandler(this.optRange_CheckedChanged);
            // 
            // btnDestinationDetails
            // 
            this.btnDestinationDetails.Location = new System.Drawing.Point(549, 107);
            this.btnDestinationDetails.Name = "btnDestinationDetails";
            this.btnDestinationDetails.Size = new System.Drawing.Size(75, 42);
            this.btnDestinationDetails.TabIndex = 6;
            this.btnDestinationDetails.Text = "&Destination Details";
            this.btnDestinationDetails.UseVisualStyleBackColor = true;
            this.btnDestinationDetails.Click += new System.EventHandler(this.btnDestinationDetails_Click);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 21;
            this.ssStatus.Text = "status";
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(73, 395);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(57, 27);
            this.btnRemove.TabIndex = 16;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // chkAddReference
            // 
            this.chkAddReference.AutoSize = true;
            this.chkAddReference.Location = new System.Drawing.Point(334, 401);
            this.chkAddReference.Name = "chkAddReference";
            this.chkAddReference.Size = new System.Drawing.Size(98, 17);
            this.chkAddReference.TabIndex = 17;
            this.chkAddReference.Text = "Add Reference";
            this.chkAddReference.UseVisualStyleBackColor = true;
            // 
            // FrmInsertCode_Rst
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.chkAddReference);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnDestinationDetails);
            this.Controls.Add(this.grpDestinatin);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnRecent);
            this.Controls.Add(this.chkMultipleRstReturned);
            this.Controls.Add(this.btnOptions);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.txtSource);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.lblParameters);
            this.Controls.Add(this.dgParameters);
            this.Controls.Add(this.grpSourceType);
            this.Controls.Add(this.btnBuildConnectionString);
            this.Controls.Add(this.txtConnectionString);
            this.Controls.Add(this.lblConnectionString);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_Rst";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmRstOpenLoopClose_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_Rst_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_Rst_Resize);
            this.grpSourceType.ResumeLayout(false);
            this.grpSourceType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).EndInit();
            this.grpDestinatin.ResumeLayout(false);
            this.grpDestinatin.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TextBox txtConnectionString;
        private System.Windows.Forms.Button btnBuildConnectionString;
        private System.Windows.Forms.GroupBox grpSourceType;
        private System.Windows.Forms.DataGridView dgParameters;
        private System.Windows.Forms.Label lblParameters;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Button btnOptions;
        private System.Windows.Forms.CheckBox chkMultipleRstReturned;
        private System.Windows.Forms.Button btnRecent;
        private System.Windows.Forms.RadioButton optSql;
        private System.Windows.Forms.RadioButton optStoreProcedure;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewComboBoxColumn colType;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSize;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColAssignVariable;
        private System.Windows.Forms.DataGridViewComboBoxColumn colValue;
        private System.Windows.Forms.GroupBox grpDestinatin;
        private System.Windows.Forms.RadioButton optEmptyLoop;
        private System.Windows.Forms.RadioButton optListboxCombo;
        private System.Windows.Forms.RadioButton optRange;
        private System.Windows.Forms.Button btnDestinationDetails;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.CheckBox chkAddReference;
    }
}