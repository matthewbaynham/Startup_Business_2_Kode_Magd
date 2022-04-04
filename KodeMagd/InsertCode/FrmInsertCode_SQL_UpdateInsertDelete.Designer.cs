namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_SQL_UpdateInsertDelete
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_SQL_UpdateInsertDelete));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.grpAction = new System.Windows.Forms.GroupBox();
            this.optDelete = new System.Windows.Forms.RadioButton();
            this.optUpdate = new System.Windows.Forms.RadioButton();
            this.optInsert = new System.Windows.Forms.RadioButton();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.optSQL = new System.Windows.Forms.RadioButton();
            this.optViaRecordset = new System.Windows.Forms.RadioButton();
            this.chkAsynchronousWithAuditCheck = new System.Windows.Forms.CheckBox();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.txtConnectionString = new System.Windows.Forms.TextBox();
            this.btnConnectionStringBuild = new System.Windows.Forms.Button();
            this.btnConnectionStringRecent = new System.Windows.Forms.Button();
            this.lblTableName = new System.Windows.Forms.Label();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.dgFields = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColDataType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValueVariable = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColSelection = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColWhere = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColAduitCondition = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.btnAddFields = new System.Windows.Forms.Button();
            this.btnRemoveFields = new System.Windows.Forms.Button();
            this.cmbTableName = new System.Windows.Forms.ComboBox();
            this.chkAdhocTableName = new System.Windows.Forms.CheckBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblName = new System.Windows.Forms.Label();
            this.chkAddReference = new System.Windows.Forms.CheckBox();
            this.btnConnectionStringExpand = new System.Windows.Forms.Button();
            this.lblInstructions = new System.Windows.Forms.Label();
            this.grpAction.SuspendLayout();
            this.grpType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).BeginInit();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 18;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(562, 392);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(58, 33);
            this.btnClose.TabIndex = 17;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(489, 392);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(67, 33);
            this.btnGenerate.TabIndex = 16;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // grpAction
            // 
            this.grpAction.Controls.Add(this.optDelete);
            this.grpAction.Controls.Add(this.optUpdate);
            this.grpAction.Controls.Add(this.optInsert);
            this.grpAction.Location = new System.Drawing.Point(12, 12);
            this.grpAction.Name = "grpAction";
            this.grpAction.Size = new System.Drawing.Size(117, 88);
            this.grpAction.TabIndex = 0;
            this.grpAction.TabStop = false;
            this.grpAction.Text = "Action";
            // 
            // optDelete
            // 
            this.optDelete.AutoSize = true;
            this.optDelete.Location = new System.Drawing.Point(17, 65);
            this.optDelete.Name = "optDelete";
            this.optDelete.Size = new System.Drawing.Size(56, 17);
            this.optDelete.TabIndex = 2;
            this.optDelete.TabStop = true;
            this.optDelete.Text = "Delete";
            this.optDelete.UseVisualStyleBackColor = true;
            this.optDelete.CheckedChanged += new System.EventHandler(this.optDelete_CheckedChanged);
            // 
            // optUpdate
            // 
            this.optUpdate.AutoSize = true;
            this.optUpdate.Location = new System.Drawing.Point(16, 42);
            this.optUpdate.Name = "optUpdate";
            this.optUpdate.Size = new System.Drawing.Size(60, 17);
            this.optUpdate.TabIndex = 1;
            this.optUpdate.TabStop = true;
            this.optUpdate.Text = "Update";
            this.optUpdate.UseVisualStyleBackColor = true;
            this.optUpdate.CheckedChanged += new System.EventHandler(this.optUpdate_CheckedChanged);
            // 
            // optInsert
            // 
            this.optInsert.AutoSize = true;
            this.optInsert.Location = new System.Drawing.Point(16, 19);
            this.optInsert.Name = "optInsert";
            this.optInsert.Size = new System.Drawing.Size(51, 17);
            this.optInsert.TabIndex = 0;
            this.optInsert.TabStop = true;
            this.optInsert.Text = "Insert";
            this.optInsert.UseVisualStyleBackColor = true;
            this.optInsert.CheckedChanged += new System.EventHandler(this.optInsert_CheckedChanged);
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optSQL);
            this.grpType.Controls.Add(this.optViaRecordset);
            this.grpType.Location = new System.Drawing.Point(12, 108);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(117, 63);
            this.grpType.TabIndex = 1;
            this.grpType.TabStop = false;
            this.grpType.Text = "Type";
            // 
            // optSQL
            // 
            this.optSQL.AutoSize = true;
            this.optSQL.Location = new System.Drawing.Point(17, 41);
            this.optSQL.Name = "optSQL";
            this.optSQL.Size = new System.Drawing.Size(46, 17);
            this.optSQL.TabIndex = 1;
            this.optSQL.TabStop = true;
            this.optSQL.Text = "SQL";
            this.optSQL.UseVisualStyleBackColor = true;
            // 
            // optViaRecordset
            // 
            this.optViaRecordset.AutoSize = true;
            this.optViaRecordset.Location = new System.Drawing.Point(16, 19);
            this.optViaRecordset.Name = "optViaRecordset";
            this.optViaRecordset.Size = new System.Drawing.Size(92, 17);
            this.optViaRecordset.TabIndex = 0;
            this.optViaRecordset.TabStop = true;
            this.optViaRecordset.Text = "Via Recordset";
            this.optViaRecordset.UseVisualStyleBackColor = true;
            // 
            // chkAsynchronousWithAuditCheck
            // 
            this.chkAsynchronousWithAuditCheck.Location = new System.Drawing.Point(12, 392);
            this.chkAsynchronousWithAuditCheck.Name = "chkAsynchronousWithAuditCheck";
            this.chkAsynchronousWithAuditCheck.Size = new System.Drawing.Size(197, 33);
            this.chkAsynchronousWithAuditCheck.TabIndex = 13;
            this.chkAsynchronousWithAuditCheck.Text = "Asynchronous with Audit Check";
            this.chkAsynchronousWithAuditCheck.UseVisualStyleBackColor = true;
            this.chkAsynchronousWithAuditCheck.CheckedChanged += new System.EventHandler(this.chkAsynchronousWithAuditCheck_CheckedChanged);
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(215, 12);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 5;
            this.lblConnectionString.Text = "Connection String";
            // 
            // txtConnectionString
            // 
            this.txtConnectionString.Location = new System.Drawing.Point(218, 43);
            this.txtConnectionString.Multiline = true;
            this.txtConnectionString.Name = "txtConnectionString";
            this.txtConnectionString.Size = new System.Drawing.Size(404, 51);
            this.txtConnectionString.TabIndex = 6;
            // 
            // btnConnectionStringBuild
            // 
            this.btnConnectionStringBuild.Location = new System.Drawing.Point(135, 12);
            this.btnConnectionStringBuild.Name = "btnConnectionStringBuild";
            this.btnConnectionStringBuild.Size = new System.Drawing.Size(74, 23);
            this.btnConnectionStringBuild.TabIndex = 2;
            this.btnConnectionStringBuild.Text = "&Build";
            this.btnConnectionStringBuild.UseVisualStyleBackColor = true;
            this.btnConnectionStringBuild.Click += new System.EventHandler(this.btnConnectionStringBuild_Click);
            // 
            // btnConnectionStringRecent
            // 
            this.btnConnectionStringRecent.Location = new System.Drawing.Point(135, 43);
            this.btnConnectionStringRecent.Name = "btnConnectionStringRecent";
            this.btnConnectionStringRecent.Size = new System.Drawing.Size(74, 23);
            this.btnConnectionStringRecent.TabIndex = 3;
            this.btnConnectionStringRecent.Text = "R&ecent";
            this.btnConnectionStringRecent.UseVisualStyleBackColor = true;
            this.btnConnectionStringRecent.Click += new System.EventHandler(this.btnConnectionStringRecent_Click);
            // 
            // lblTableName
            // 
            this.lblTableName.AutoSize = true;
            this.lblTableName.Location = new System.Drawing.Point(144, 169);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(65, 13);
            this.lblTableName.TabIndex = 9;
            this.lblTableName.Text = "Table Name";
            // 
            // txtTableName
            // 
            this.txtTableName.Location = new System.Drawing.Point(218, 169);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(264, 20);
            this.txtTableName.TabIndex = 11;
            // 
            // dgFields
            // 
            this.dgFields.AllowUserToAddRows = false;
            this.dgFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColDataType,
            this.ColSize,
            this.ColValue,
            this.ColValueVariable,
            this.ColSelection,
            this.ColWhere,
            this.ColAduitCondition});
            this.dgFields.Location = new System.Drawing.Point(18, 199);
            this.dgFields.Name = "dgFields";
            this.dgFields.Size = new System.Drawing.Size(602, 183);
            this.dgFields.TabIndex = 12;
            this.dgFields.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgFields_CellEndEdit);
            this.dgFields.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgFields_CellValueChanged);
            this.dgFields.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgFields_RowsAdded);
            // 
            // ColName
            // 
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            this.ColName.ToolTipText = "Field Name";
            // 
            // ColDataType
            // 
            this.ColDataType.HeaderText = "DataType";
            this.ColDataType.Name = "ColDataType";
            this.ColDataType.ToolTipText = "Data Type";
            // 
            // ColSize
            // 
            this.ColSize.HeaderText = "Size";
            this.ColSize.Name = "ColSize";
            this.ColSize.Width = 50;
            // 
            // ColValue
            // 
            this.ColValue.HeaderText = "Value";
            this.ColValue.Name = "ColValue";
            this.ColValue.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ColValue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.ColValue.ToolTipText = "Conditional = True => Variable or Value in Where Condition.";
            // 
            // ColValueVariable
            // 
            this.ColValueVariable.HeaderText = "Val/Var";
            this.ColValueVariable.Name = "ColValueVariable";
            this.ColValueVariable.ToolTipText = "Value or Variable";
            this.ColValueVariable.Width = 60;
            // 
            // ColSelection
            // 
            this.ColSelection.HeaderText = "Selection";
            this.ColSelection.Name = "ColSelection";
            this.ColSelection.Width = 60;
            // 
            // ColWhere
            // 
            this.ColWhere.HeaderText = "Conditional";
            this.ColWhere.Name = "ColWhere";
            this.ColWhere.ToolTipText = "Adds fields to Where condition.";
            this.ColWhere.Width = 60;
            // 
            // ColAduitCondition
            // 
            this.ColAduitCondition.HeaderText = "Audit Condition";
            this.ColAduitCondition.Name = "ColAduitCondition";
            this.ColAduitCondition.Width = 60;
            // 
            // btnAddFields
            // 
            this.btnAddFields.Location = new System.Drawing.Point(191, 392);
            this.btnAddFields.Name = "btnAddFields";
            this.btnAddFields.Size = new System.Drawing.Size(53, 33);
            this.btnAddFields.TabIndex = 14;
            this.btnAddFields.Text = "&Add";
            this.btnAddFields.UseVisualStyleBackColor = true;
            this.btnAddFields.Click += new System.EventHandler(this.btnAddFields_Click);
            // 
            // btnRemoveFields
            // 
            this.btnRemoveFields.Location = new System.Drawing.Point(250, 392);
            this.btnRemoveFields.Name = "btnRemoveFields";
            this.btnRemoveFields.Size = new System.Drawing.Size(56, 33);
            this.btnRemoveFields.TabIndex = 15;
            this.btnRemoveFields.Text = "Remo&ve";
            this.btnRemoveFields.UseVisualStyleBackColor = true;
            this.btnRemoveFields.Click += new System.EventHandler(this.btnRemoveFields_Click);
            // 
            // cmbTableName
            // 
            this.cmbTableName.FormattingEnabled = true;
            this.cmbTableName.Location = new System.Drawing.Point(218, 169);
            this.cmbTableName.Name = "cmbTableName";
            this.cmbTableName.Size = new System.Drawing.Size(264, 21);
            this.cmbTableName.TabIndex = 10;
            this.cmbTableName.SelectedIndexChanged += new System.EventHandler(this.cmbTableName_SelectedIndexChanged);
            // 
            // chkAdhocTableName
            // 
            this.chkAdhocTableName.AutoSize = true;
            this.chkAdhocTableName.Location = new System.Drawing.Point(504, 165);
            this.chkAdhocTableName.Name = "chkAdhocTableName";
            this.chkAdhocTableName.Size = new System.Drawing.Size(118, 17);
            this.chkAdhocTableName.TabIndex = 11;
            this.chkAdhocTableName.Text = "Adhoc Table Name";
            this.chkAdhocTableName.UseVisualStyleBackColor = true;
            this.chkAdhocTableName.CheckedChanged += new System.EventHandler(this.chkAdhocTableName_CheckedChanged);
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(434, 11);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(186, 20);
            this.txtName.TabIndex = 8;
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(358, 12);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(70, 13);
            this.lblName.TabIndex = 7;
            this.lblName.Text = "Name (Suffix)";
            // 
            // chkAddReference
            // 
            this.chkAddReference.AutoSize = true;
            this.chkAddReference.Location = new System.Drawing.Point(384, 401);
            this.chkAddReference.Name = "chkAddReference";
            this.chkAddReference.Size = new System.Drawing.Size(98, 17);
            this.chkAddReference.TabIndex = 19;
            this.chkAddReference.Text = "Add Reference";
            this.chkAddReference.UseVisualStyleBackColor = true;
            // 
            // btnConnectionStringExpand
            // 
            this.btnConnectionStringExpand.Location = new System.Drawing.Point(134, 71);
            this.btnConnectionStringExpand.Name = "btnConnectionStringExpand";
            this.btnConnectionStringExpand.Size = new System.Drawing.Size(75, 23);
            this.btnConnectionStringExpand.TabIndex = 4;
            this.btnConnectionStringExpand.Text = ". . .";
            this.btnConnectionStringExpand.UseVisualStyleBackColor = true;
            this.btnConnectionStringExpand.Click += new System.EventHandler(this.btnConnectionStringExpand_Click);
            // 
            // lblInstructions
            // 
            this.lblInstructions.Location = new System.Drawing.Point(138, 104);
            this.lblInstructions.Name = "lblInstructions";
            this.lblInstructions.Size = new System.Drawing.Size(484, 58);
            this.lblInstructions.TabIndex = 20;
            this.lblInstructions.Text = "Instructions";
            // 
            // FrmInsertCode_SQL_UpdateInsertDelete
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.lblInstructions);
            this.Controls.Add(this.chkAddReference);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.chkAdhocTableName);
            this.Controls.Add(this.btnConnectionStringExpand);
            this.Controls.Add(this.cmbTableName);
            this.Controls.Add(this.btnRemoveFields);
            this.Controls.Add(this.btnAddFields);
            this.Controls.Add(this.dgFields);
            this.Controls.Add(this.txtTableName);
            this.Controls.Add(this.lblTableName);
            this.Controls.Add(this.btnConnectionStringRecent);
            this.Controls.Add(this.btnConnectionStringBuild);
            this.Controls.Add(this.txtConnectionString);
            this.Controls.Add(this.lblConnectionString);
            this.Controls.Add(this.chkAsynchronousWithAuditCheck);
            this.Controls.Add(this.grpType);
            this.Controls.Add(this.grpAction);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_SQL_UpdateInsertDelete";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_SQL_UpdateInsertDelete_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_SQL_UpdateInsertDelete_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_SQL_UpdateInsertDelete_Resize);
            this.grpAction.ResumeLayout(false);
            this.grpAction.PerformLayout();
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgFields)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox grpAction;
        private System.Windows.Forms.RadioButton optDelete;
        private System.Windows.Forms.RadioButton optUpdate;
        private System.Windows.Forms.RadioButton optInsert;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton optSQL;
        private System.Windows.Forms.RadioButton optViaRecordset;
        private System.Windows.Forms.CheckBox chkAsynchronousWithAuditCheck;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TextBox txtConnectionString;
        private System.Windows.Forms.Button btnConnectionStringBuild;
        private System.Windows.Forms.Button btnConnectionStringRecent;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.DataGridView dgFields;
        private System.Windows.Forms.Button btnAddFields;
        private System.Windows.Forms.Button btnRemoveFields;
        private System.Windows.Forms.ComboBox cmbTableName;
        private System.Windows.Forms.CheckBox chkAdhocTableName;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDataType;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValue;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColValueVariable;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColSelection;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColWhere;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColAduitCondition;
        private System.Windows.Forms.CheckBox chkAddReference;
        private System.Windows.Forms.Button btnConnectionStringExpand;
        private System.Windows.Forms.Label lblInstructions;
    }
}