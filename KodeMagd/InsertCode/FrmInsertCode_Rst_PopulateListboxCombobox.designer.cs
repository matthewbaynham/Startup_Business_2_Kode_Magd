namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_Rst_PopulateListboxCombobox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_Rst_PopulateListboxCombobox));
            this.grpType = new System.Windows.Forms.GroupBox();
            this.optComboBox = new System.Windows.Forms.RadioButton();
            this.optListBox = new System.Windows.Forms.RadioButton();
            this.lblControl = new System.Windows.Forms.Label();
            this.cmbControl = new System.Windows.Forms.ComboBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.grpSource = new System.Windows.Forms.GroupBox();
            this.optRangeAddress = new System.Windows.Forms.RadioButton();
            this.optRangeNamed = new System.Windows.Forms.RadioButton();
            this.optArray = new System.Windows.Forms.RadioButton();
            this.optOneValue = new System.Windows.Forms.RadioButton();
            this.optRecordset = new System.Windows.Forms.RadioButton();
            this.btnArrayAdd = new System.Windows.Forms.Button();
            this.btnArrayRemove = new System.Windows.Forms.Button();
            this.lstArray = new System.Windows.Forms.ListBox();
            this.txtRstSql = new System.Windows.Forms.TextBox();
            this.cmbRstCommandType = new System.Windows.Forms.ComboBox();
            this.txtRstConnectionString = new System.Windows.Forms.TextBox();
            this.btnRstConnectionStringBuild = new System.Windows.Forms.Button();
            this.btnRstConnectionStringRecent = new System.Windows.Forms.Button();
            this.cmbRstTableName = new System.Windows.Forms.ComboBox();
            this.btnRstConnectionStringExpand = new System.Windows.Forms.Button();
            this.btnRstSqlExpand = new System.Windows.Forms.Button();
            this.txtOneValue = new System.Windows.Forms.TextBox();
            this.cmbRangeNamed = new System.Windows.Forms.ComboBox();
            this.cmbRangeAddressSheetName = new System.Windows.Forms.ComboBox();
            this.txtRangeAddressAddress = new System.Windows.Forms.TextBox();
            this.lblRangeAddressSheetName = new System.Windows.Forms.Label();
            this.lblOneValue = new System.Windows.Forms.Label();
            this.lblArray = new System.Windows.Forms.Label();
            this.lblRangeAddressAddress = new System.Windows.Forms.Label();
            this.lblRstConnectionString = new System.Windows.Forms.Label();
            this.lblRstCommandType = new System.Windows.Forms.Label();
            this.lblRangeNamed = new System.Windows.Forms.Label();
            this.lblRstSql = new System.Windows.Forms.Label();
            this.lblRstTableName = new System.Windows.Forms.Label();
            this.btnRstParameters = new System.Windows.Forms.Button();
            this.lblRstFieldName = new System.Windows.Forms.Label();
            this.txtRstFieldName = new System.Windows.Forms.TextBox();
            this.lblWarning = new System.Windows.Forms.Label();
            this.cmbRstFieldName = new System.Windows.Forms.ComboBox();
            this.chkAddReference = new System.Windows.Forms.CheckBox();
            this.grpType.SuspendLayout();
            this.grpSource.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optComboBox);
            this.grpType.Controls.Add(this.optListBox);
            this.grpType.Location = new System.Drawing.Point(12, 9);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(92, 63);
            this.grpType.TabIndex = 2;
            this.grpType.TabStop = false;
            this.grpType.Text = "Type";
            // 
            // optComboBox
            // 
            this.optComboBox.AutoSize = true;
            this.optComboBox.Location = new System.Drawing.Point(6, 42);
            this.optComboBox.Name = "optComboBox";
            this.optComboBox.Size = new System.Drawing.Size(79, 17);
            this.optComboBox.TabIndex = 1;
            this.optComboBox.TabStop = true;
            this.optComboBox.Text = "Combo Box";
            this.optComboBox.UseVisualStyleBackColor = true;
            this.optComboBox.CheckedChanged += new System.EventHandler(this.optComboBox_CheckedChanged);
            // 
            // optListBox
            // 
            this.optListBox.AutoSize = true;
            this.optListBox.Location = new System.Drawing.Point(6, 19);
            this.optListBox.Name = "optListBox";
            this.optListBox.Size = new System.Drawing.Size(62, 17);
            this.optListBox.TabIndex = 0;
            this.optListBox.TabStop = true;
            this.optListBox.Text = "List Box";
            this.optListBox.UseVisualStyleBackColor = true;
            this.optListBox.CheckedChanged += new System.EventHandler(this.optListBox_CheckedChanged);
            // 
            // lblControl
            // 
            this.lblControl.AutoSize = true;
            this.lblControl.Location = new System.Drawing.Point(120, 9);
            this.lblControl.Name = "lblControl";
            this.lblControl.Size = new System.Drawing.Size(40, 13);
            this.lblControl.TabIndex = 1;
            this.lblControl.Text = "Control";
            // 
            // cmbControl
            // 
            this.cmbControl.FormattingEnabled = true;
            this.cmbControl.Location = new System.Drawing.Point(166, 9);
            this.cmbControl.Name = "cmbControl";
            this.cmbControl.Size = new System.Drawing.Size(353, 21);
            this.cmbControl.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(535, 388);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 30);
            this.btnClose.TabIndex = 36;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 1;
            this.ssStatus.Text = "status";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(444, 388);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 30);
            this.btnGenerate.TabIndex = 35;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // grpSource
            // 
            this.grpSource.Controls.Add(this.optRangeAddress);
            this.grpSource.Controls.Add(this.optRangeNamed);
            this.grpSource.Controls.Add(this.optArray);
            this.grpSource.Controls.Add(this.optOneValue);
            this.grpSource.Controls.Add(this.optRecordset);
            this.grpSource.Location = new System.Drawing.Point(12, 78);
            this.grpSource.Name = "grpSource";
            this.grpSource.Size = new System.Drawing.Size(148, 137);
            this.grpSource.TabIndex = 4;
            this.grpSource.TabStop = false;
            this.grpSource.Text = "Source";
            // 
            // optRangeAddress
            // 
            this.optRangeAddress.AutoSize = true;
            this.optRangeAddress.Location = new System.Drawing.Point(6, 111);
            this.optRangeAddress.Name = "optRangeAddress";
            this.optRangeAddress.Size = new System.Drawing.Size(126, 17);
            this.optRangeAddress.TabIndex = 4;
            this.optRangeAddress.TabStop = true;
            this.optRangeAddress.Text = "Range (from address)";
            this.optRangeAddress.UseVisualStyleBackColor = true;
            this.optRangeAddress.CheckedChanged += new System.EventHandler(this.optRangeAddress_CheckedChanged);
            // 
            // optRangeNamed
            // 
            this.optRangeNamed.AutoSize = true;
            this.optRangeNamed.Location = new System.Drawing.Point(6, 88);
            this.optRangeNamed.Name = "optRangeNamed";
            this.optRangeNamed.Size = new System.Drawing.Size(100, 17);
            this.optRangeNamed.TabIndex = 3;
            this.optRangeNamed.TabStop = true;
            this.optRangeNamed.Text = "Range (Named)";
            this.optRangeNamed.UseVisualStyleBackColor = true;
            this.optRangeNamed.CheckedChanged += new System.EventHandler(this.optRangeNamed_CheckedChanged);
            // 
            // optArray
            // 
            this.optArray.AutoSize = true;
            this.optArray.Location = new System.Drawing.Point(6, 65);
            this.optArray.Name = "optArray";
            this.optArray.Size = new System.Drawing.Size(49, 17);
            this.optArray.TabIndex = 2;
            this.optArray.TabStop = true;
            this.optArray.Text = "Array";
            this.optArray.UseVisualStyleBackColor = true;
            this.optArray.CheckedChanged += new System.EventHandler(this.optArray_CheckedChanged);
            // 
            // optOneValue
            // 
            this.optOneValue.AutoSize = true;
            this.optOneValue.Location = new System.Drawing.Point(6, 42);
            this.optOneValue.Name = "optOneValue";
            this.optOneValue.Size = new System.Drawing.Size(75, 17);
            this.optOneValue.TabIndex = 1;
            this.optOneValue.TabStop = true;
            this.optOneValue.Text = "One Value";
            this.optOneValue.UseVisualStyleBackColor = true;
            this.optOneValue.CheckedChanged += new System.EventHandler(this.optOneValue_CheckedChanged);
            // 
            // optRecordset
            // 
            this.optRecordset.AutoSize = true;
            this.optRecordset.Location = new System.Drawing.Point(6, 19);
            this.optRecordset.Name = "optRecordset";
            this.optRecordset.Size = new System.Drawing.Size(74, 17);
            this.optRecordset.TabIndex = 0;
            this.optRecordset.TabStop = true;
            this.optRecordset.Text = "Recordset";
            this.optRecordset.UseVisualStyleBackColor = true;
            this.optRecordset.CheckedChanged += new System.EventHandler(this.optRecordset_CheckedChanged);
            // 
            // btnArrayAdd
            // 
            this.btnArrayAdd.Location = new System.Drawing.Point(535, 59);
            this.btnArrayAdd.Name = "btnArrayAdd";
            this.btnArrayAdd.Size = new System.Drawing.Size(75, 23);
            this.btnArrayAdd.TabIndex = 15;
            this.btnArrayAdd.Text = "&Add";
            this.btnArrayAdd.UseVisualStyleBackColor = true;
            this.btnArrayAdd.Click += new System.EventHandler(this.btnArrayAdd_Click);
            // 
            // btnArrayRemove
            // 
            this.btnArrayRemove.Location = new System.Drawing.Point(535, 93);
            this.btnArrayRemove.Name = "btnArrayRemove";
            this.btnArrayRemove.Size = new System.Drawing.Size(75, 23);
            this.btnArrayRemove.TabIndex = 18;
            this.btnArrayRemove.Text = "Re&move";
            this.btnArrayRemove.UseVisualStyleBackColor = true;
            this.btnArrayRemove.Click += new System.EventHandler(this.btnArrayRemove_Click);
            // 
            // lstArray
            // 
            this.lstArray.FormattingEnabled = true;
            this.lstArray.Location = new System.Drawing.Point(166, 64);
            this.lstArray.Name = "lstArray";
            this.lstArray.Size = new System.Drawing.Size(352, 303);
            this.lstArray.TabIndex = 14;
            // 
            // txtRstSql
            // 
            this.txtRstSql.Location = new System.Drawing.Point(166, 230);
            this.txtRstSql.Multiline = true;
            this.txtRstSql.Name = "txtRstSql";
            this.txtRstSql.Size = new System.Drawing.Size(352, 88);
            this.txtRstSql.TabIndex = 27;
            this.txtRstSql.TextChanged += new System.EventHandler(this.txtRstSql_TextChanged);
            // 
            // cmbRstCommandType
            // 
            this.cmbRstCommandType.FormattingEnabled = true;
            this.cmbRstCommandType.Location = new System.Drawing.Point(165, 189);
            this.cmbRstCommandType.Name = "cmbRstCommandType";
            this.cmbRstCommandType.Size = new System.Drawing.Size(352, 21);
            this.cmbRstCommandType.TabIndex = 21;
            this.cmbRstCommandType.SelectedIndexChanged += new System.EventHandler(this.cmbRstCommandType_SelectedIndexChanged);
            // 
            // txtRstConnectionString
            // 
            this.txtRstConnectionString.Location = new System.Drawing.Point(166, 59);
            this.txtRstConnectionString.Multiline = true;
            this.txtRstConnectionString.Name = "txtRstConnectionString";
            this.txtRstConnectionString.Size = new System.Drawing.Size(351, 113);
            this.txtRstConnectionString.TabIndex = 5;
            this.txtRstConnectionString.TextChanged += new System.EventHandler(this.txtRstConnectionString_TextChanged);
            // 
            // btnRstConnectionStringBuild
            // 
            this.btnRstConnectionStringBuild.Location = new System.Drawing.Point(535, 59);
            this.btnRstConnectionStringBuild.Name = "btnRstConnectionStringBuild";
            this.btnRstConnectionStringBuild.Size = new System.Drawing.Size(75, 23);
            this.btnRstConnectionStringBuild.TabIndex = 16;
            this.btnRstConnectionStringBuild.Text = "&Build";
            this.btnRstConnectionStringBuild.UseVisualStyleBackColor = true;
            this.btnRstConnectionStringBuild.Click += new System.EventHandler(this.btnRstConnectionStringBuild_Click);
            // 
            // btnRstConnectionStringRecent
            // 
            this.btnRstConnectionStringRecent.Location = new System.Drawing.Point(535, 89);
            this.btnRstConnectionStringRecent.Name = "btnRstConnectionStringRecent";
            this.btnRstConnectionStringRecent.Size = new System.Drawing.Size(75, 23);
            this.btnRstConnectionStringRecent.TabIndex = 17;
            this.btnRstConnectionStringRecent.Text = "&Recent";
            this.btnRstConnectionStringRecent.UseVisualStyleBackColor = true;
            this.btnRstConnectionStringRecent.Click += new System.EventHandler(this.btnRstConnectionStringRecent_Click);
            // 
            // cmbRstTableName
            // 
            this.cmbRstTableName.FormattingEnabled = true;
            this.cmbRstTableName.Location = new System.Drawing.Point(168, 230);
            this.cmbRstTableName.Name = "cmbRstTableName";
            this.cmbRstTableName.Size = new System.Drawing.Size(351, 21);
            this.cmbRstTableName.TabIndex = 26;
            this.cmbRstTableName.SelectedIndexChanged += new System.EventHandler(this.cmbRstTableName_SelectedIndexChanged);
            // 
            // btnRstConnectionStringExpand
            // 
            this.btnRstConnectionStringExpand.Location = new System.Drawing.Point(536, 119);
            this.btnRstConnectionStringExpand.Name = "btnRstConnectionStringExpand";
            this.btnRstConnectionStringExpand.Size = new System.Drawing.Size(75, 23);
            this.btnRstConnectionStringExpand.TabIndex = 19;
            this.btnRstConnectionStringExpand.Text = ". . .";
            this.btnRstConnectionStringExpand.UseVisualStyleBackColor = true;
            this.btnRstConnectionStringExpand.Click += new System.EventHandler(this.btnRstConnectionStringExpand_Click);
            // 
            // btnRstSqlExpand
            // 
            this.btnRstSqlExpand.Location = new System.Drawing.Point(536, 260);
            this.btnRstSqlExpand.Name = "btnRstSqlExpand";
            this.btnRstSqlExpand.Size = new System.Drawing.Size(75, 23);
            this.btnRstSqlExpand.TabIndex = 32;
            this.btnRstSqlExpand.Text = ". . .";
            this.btnRstSqlExpand.UseVisualStyleBackColor = true;
            this.btnRstSqlExpand.Click += new System.EventHandler(this.btnRstSqlExpand_Click);
            // 
            // txtOneValue
            // 
            this.txtOneValue.Location = new System.Drawing.Point(167, 59);
            this.txtOneValue.Name = "txtOneValue";
            this.txtOneValue.Size = new System.Drawing.Size(352, 20);
            this.txtOneValue.TabIndex = 25;
            // 
            // cmbRangeNamed
            // 
            this.cmbRangeNamed.FormattingEnabled = true;
            this.cmbRangeNamed.Location = new System.Drawing.Point(166, 59);
            this.cmbRangeNamed.Name = "cmbRangeNamed";
            this.cmbRangeNamed.Size = new System.Drawing.Size(351, 21);
            this.cmbRangeNamed.TabIndex = 5;
            // 
            // cmbRangeAddressSheetName
            // 
            this.cmbRangeAddressSheetName.FormattingEnabled = true;
            this.cmbRangeAddressSheetName.Location = new System.Drawing.Point(166, 59);
            this.cmbRangeAddressSheetName.Name = "cmbRangeAddressSheetName";
            this.cmbRangeAddressSheetName.Size = new System.Drawing.Size(351, 21);
            this.cmbRangeAddressSheetName.TabIndex = 14;
            // 
            // txtRangeAddressAddress
            // 
            this.txtRangeAddressAddress.Location = new System.Drawing.Point(168, 209);
            this.txtRangeAddressAddress.Name = "txtRangeAddressAddress";
            this.txtRangeAddressAddress.Size = new System.Drawing.Size(351, 20);
            this.txtRangeAddressAddress.TabIndex = 25;
            // 
            // lblRangeAddressSheetName
            // 
            this.lblRangeAddressSheetName.AutoSize = true;
            this.lblRangeAddressSheetName.Location = new System.Drawing.Point(167, 39);
            this.lblRangeAddressSheetName.Name = "lblRangeAddressSheetName";
            this.lblRangeAddressSheetName.Size = new System.Drawing.Size(66, 13);
            this.lblRangeAddressSheetName.TabIndex = 29;
            this.lblRangeAddressSheetName.Text = "Sheet Name";
            // 
            // lblOneValue
            // 
            this.lblOneValue.AutoSize = true;
            this.lblOneValue.Location = new System.Drawing.Point(167, 39);
            this.lblOneValue.Name = "lblOneValue";
            this.lblOneValue.Size = new System.Drawing.Size(34, 13);
            this.lblOneValue.TabIndex = 8;
            this.lblOneValue.Text = "Value";
            // 
            // lblArray
            // 
            this.lblArray.AutoSize = true;
            this.lblArray.Location = new System.Drawing.Point(169, 39);
            this.lblArray.Name = "lblArray";
            this.lblArray.Size = new System.Drawing.Size(32, 13);
            this.lblArray.TabIndex = 7;
            this.lblArray.Text = "Items";
            // 
            // lblRangeAddressAddress
            // 
            this.lblRangeAddressAddress.AutoSize = true;
            this.lblRangeAddressAddress.Location = new System.Drawing.Point(169, 193);
            this.lblRangeAddressAddress.Name = "lblRangeAddressAddress";
            this.lblRangeAddressAddress.Size = new System.Drawing.Size(45, 13);
            this.lblRangeAddressAddress.TabIndex = 22;
            this.lblRangeAddressAddress.Text = "Address";
            // 
            // lblRstConnectionString
            // 
            this.lblRstConnectionString.AutoSize = true;
            this.lblRstConnectionString.Location = new System.Drawing.Point(170, 38);
            this.lblRstConnectionString.Name = "lblRstConnectionString";
            this.lblRstConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblRstConnectionString.TabIndex = 6;
            this.lblRstConnectionString.Text = "Connection String";
            // 
            // lblRstCommandType
            // 
            this.lblRstCommandType.AutoSize = true;
            this.lblRstCommandType.Location = new System.Drawing.Point(167, 173);
            this.lblRstCommandType.Name = "lblRstCommandType";
            this.lblRstCommandType.Size = new System.Drawing.Size(81, 13);
            this.lblRstCommandType.TabIndex = 20;
            this.lblRstCommandType.Text = "Command Type";
            // 
            // lblRangeNamed
            // 
            this.lblRangeNamed.AutoSize = true;
            this.lblRangeNamed.Location = new System.Drawing.Point(165, 39);
            this.lblRangeNamed.Name = "lblRangeNamed";
            this.lblRangeNamed.Size = new System.Drawing.Size(76, 13);
            this.lblRangeNamed.TabIndex = 13;
            this.lblRangeNamed.Text = "Named Range";
            // 
            // lblRstSql
            // 
            this.lblRstSql.AutoSize = true;
            this.lblRstSql.Location = new System.Drawing.Point(167, 212);
            this.lblRstSql.Name = "lblRstSql";
            this.lblRstSql.Size = new System.Drawing.Size(28, 13);
            this.lblRstSql.TabIndex = 36;
            this.lblRstSql.Text = "SQL";
            // 
            // lblRstTableName
            // 
            this.lblRstTableName.AutoSize = true;
            this.lblRstTableName.Location = new System.Drawing.Point(167, 212);
            this.lblRstTableName.Name = "lblRstTableName";
            this.lblRstTableName.Size = new System.Drawing.Size(65, 13);
            this.lblRstTableName.TabIndex = 24;
            this.lblRstTableName.Text = "Table Name";
            // 
            // btnRstParameters
            // 
            this.btnRstParameters.Location = new System.Drawing.Point(535, 230);
            this.btnRstParameters.Name = "btnRstParameters";
            this.btnRstParameters.Size = new System.Drawing.Size(75, 23);
            this.btnRstParameters.TabIndex = 31;
            this.btnRstParameters.Text = "&Parameters";
            this.btnRstParameters.UseVisualStyleBackColor = true;
            this.btnRstParameters.Click += new System.EventHandler(this.btnParameters_Click);
            // 
            // lblRstFieldName
            // 
            this.lblRstFieldName.AutoSize = true;
            this.lblRstFieldName.Location = new System.Drawing.Point(170, 328);
            this.lblRstFieldName.Name = "lblRstFieldName";
            this.lblRstFieldName.Size = new System.Drawing.Size(60, 13);
            this.lblRstFieldName.TabIndex = 28;
            this.lblRstFieldName.Text = "Field Name";
            // 
            // txtRstFieldName
            // 
            this.txtRstFieldName.Location = new System.Drawing.Point(165, 342);
            this.txtRstFieldName.Name = "txtRstFieldName";
            this.txtRstFieldName.Size = new System.Drawing.Size(352, 20);
            this.txtRstFieldName.TabIndex = 33;
            // 
            // lblWarning
            // 
            this.lblWarning.Location = new System.Drawing.Point(12, 376);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(419, 42);
            this.lblWarning.TabIndex = 0;
            this.lblWarning.Text = "label1";
            // 
            // cmbRstFieldName
            // 
            this.cmbRstFieldName.FormattingEnabled = true;
            this.cmbRstFieldName.Location = new System.Drawing.Point(167, 342);
            this.cmbRstFieldName.Name = "cmbRstFieldName";
            this.cmbRstFieldName.Size = new System.Drawing.Size(350, 21);
            this.cmbRstFieldName.TabIndex = 29;
            // 
            // chkAddReference
            // 
            this.chkAddReference.AutoSize = true;
            this.chkAddReference.Location = new System.Drawing.Point(535, 365);
            this.chkAddReference.Name = "chkAddReference";
            this.chkAddReference.Size = new System.Drawing.Size(98, 17);
            this.chkAddReference.TabIndex = 34;
            this.chkAddReference.Text = "Add Reference";
            this.chkAddReference.UseVisualStyleBackColor = true;
            // 
            // FrmInsertCode_Rst_PopulateListboxCombobox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.chkAddReference);
            this.Controls.Add(this.cmbRstFieldName);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.txtRstFieldName);
            this.Controls.Add(this.lblRstFieldName);
            this.Controls.Add(this.btnRstParameters);
            this.Controls.Add(this.lblRstTableName);
            this.Controls.Add(this.lblRstSql);
            this.Controls.Add(this.lblRangeNamed);
            this.Controls.Add(this.lblRstCommandType);
            this.Controls.Add(this.lblRstConnectionString);
            this.Controls.Add(this.lblRangeAddressAddress);
            this.Controls.Add(this.lblArray);
            this.Controls.Add(this.lblOneValue);
            this.Controls.Add(this.lblRangeAddressSheetName);
            this.Controls.Add(this.txtRangeAddressAddress);
            this.Controls.Add(this.cmbRangeAddressSheetName);
            this.Controls.Add(this.cmbRangeNamed);
            this.Controls.Add(this.txtOneValue);
            this.Controls.Add(this.btnRstSqlExpand);
            this.Controls.Add(this.btnRstConnectionStringExpand);
            this.Controls.Add(this.cmbRstTableName);
            this.Controls.Add(this.btnRstConnectionStringRecent);
            this.Controls.Add(this.btnRstConnectionStringBuild);
            this.Controls.Add(this.txtRstConnectionString);
            this.Controls.Add(this.cmbRstCommandType);
            this.Controls.Add(this.txtRstSql);
            this.Controls.Add(this.lstArray);
            this.Controls.Add(this.btnArrayRemove);
            this.Controls.Add(this.btnArrayAdd);
            this.Controls.Add(this.grpSource);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.cmbControl);
            this.Controls.Add(this.lblControl);
            this.Controls.Add(this.grpType);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_Rst_PopulateListboxCombobox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmInsertCode_PopulateListboxCombobox_FormClosing);
            this.Load += new System.EventHandler(this.FrmInsertCode_PopulateListboxCombobox_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_Rst_PopulateListboxCombobox_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_Rst_PopulateListboxCombobox_Resize);
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            this.grpSource.ResumeLayout(false);
            this.grpSource.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton optComboBox;
        private System.Windows.Forms.RadioButton optListBox;
        private System.Windows.Forms.Label lblControl;
        private System.Windows.Forms.ComboBox cmbControl;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox grpSource;
        private System.Windows.Forms.RadioButton optOneValue;
        private System.Windows.Forms.RadioButton optRecordset;
        private System.Windows.Forms.RadioButton optRangeAddress;
        private System.Windows.Forms.RadioButton optRangeNamed;
        private System.Windows.Forms.RadioButton optArray;
        private System.Windows.Forms.Button btnArrayAdd;
        private System.Windows.Forms.Button btnArrayRemove;
        private System.Windows.Forms.ListBox lstArray;
        private System.Windows.Forms.TextBox txtRstSql;
        private System.Windows.Forms.ComboBox cmbRstCommandType;
        private System.Windows.Forms.TextBox txtRstConnectionString;
        private System.Windows.Forms.Button btnRstConnectionStringBuild;
        private System.Windows.Forms.Button btnRstConnectionStringRecent;
        private System.Windows.Forms.ComboBox cmbRstTableName;
        private System.Windows.Forms.Button btnRstConnectionStringExpand;
        private System.Windows.Forms.Button btnRstSqlExpand;
        private System.Windows.Forms.TextBox txtOneValue;
        private System.Windows.Forms.ComboBox cmbRangeNamed;
        private System.Windows.Forms.ComboBox cmbRangeAddressSheetName;
        private System.Windows.Forms.TextBox txtRangeAddressAddress;
        private System.Windows.Forms.Label lblRangeAddressSheetName;
        private System.Windows.Forms.Label lblOneValue;
        private System.Windows.Forms.Label lblArray;
        private System.Windows.Forms.Label lblRangeAddressAddress;
        private System.Windows.Forms.Label lblRstConnectionString;
        private System.Windows.Forms.Label lblRstCommandType;
        private System.Windows.Forms.Label lblRangeNamed;
        private System.Windows.Forms.Label lblRstSql;
        private System.Windows.Forms.Label lblRstTableName;
        private System.Windows.Forms.Button btnRstParameters;
        private System.Windows.Forms.Label lblRstFieldName;
        private System.Windows.Forms.TextBox txtRstFieldName;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.ComboBox cmbRstFieldName;
        private System.Windows.Forms.CheckBox chkAddReference;
    }
}