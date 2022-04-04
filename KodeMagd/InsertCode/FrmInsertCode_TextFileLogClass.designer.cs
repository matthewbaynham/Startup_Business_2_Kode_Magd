namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_TextFileLogClass
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_TextFileLogClass));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.txtDestinationFolder = new System.Windows.Forms.TextBox();
            this.lblDestinationFolder = new System.Windows.Forms.Label();
            this.dgParameters = new System.Windows.Forms.DataGridView();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColOptional = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colDataType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.grpFileName = new System.Windows.Forms.GroupBox();
            this.optFileNameSpecified = new System.Windows.Forms.RadioButton();
            this.optFileNameAuto = new System.Windows.Forms.RadioButton();
            this.lblClassName = new System.Windows.Forms.Label();
            this.txtClassName = new System.Windows.Forms.TextBox();
            this.lblDateFormatFileName = new System.Windows.Forms.Label();
            this.lblDateFormatFileContents = new System.Windows.Forms.Label();
            this.cmbDateFormatFileName = new System.Windows.Forms.ComboBox();
            this.cmbDateFormatFileContents = new System.Windows.Forms.ComboBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.chkAddReferences = new System.Windows.Forms.CheckBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).BeginInit();
            this.grpFileName.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(548, 404);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(462, 404);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(80, 23);
            this.btnGenerate.TabIndex = 14;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // txtDestinationFolder
            // 
            this.txtDestinationFolder.Location = new System.Drawing.Point(12, 111);
            this.txtDestinationFolder.Name = "txtDestinationFolder";
            this.txtDestinationFolder.Size = new System.Drawing.Size(611, 20);
            this.txtDestinationFolder.TabIndex = 8;
            // 
            // lblDestinationFolder
            // 
            this.lblDestinationFolder.AutoSize = true;
            this.lblDestinationFolder.Location = new System.Drawing.Point(12, 90);
            this.lblDestinationFolder.Name = "lblDestinationFolder";
            this.lblDestinationFolder.Size = new System.Drawing.Size(92, 13);
            this.lblDestinationFolder.TabIndex = 7;
            this.lblDestinationFolder.Text = "Destination Folder";
            // 
            // dgParameters
            // 
            this.dgParameters.AllowUserToAddRows = false;
            this.dgParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colName,
            this.ColOptional,
            this.colDataType});
            this.dgParameters.Location = new System.Drawing.Point(12, 189);
            this.dgParameters.Name = "dgParameters";
            this.dgParameters.Size = new System.Drawing.Size(610, 197);
            this.dgParameters.TabIndex = 10;
            this.dgParameters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgParameters_RowsAdded);
            // 
            // colName
            // 
            this.colName.HeaderText = "Name";
            this.colName.Name = "colName";
            // 
            // ColOptional
            // 
            this.ColOptional.HeaderText = "Optional";
            this.ColOptional.Name = "ColOptional";
            // 
            // colDataType
            // 
            this.colDataType.HeaderText = "Data Type";
            this.colDataType.Name = "colDataType";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 167);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Parameters to be logged";
            // 
            // grpFileName
            // 
            this.grpFileName.Controls.Add(this.optFileNameSpecified);
            this.grpFileName.Controls.Add(this.optFileNameAuto);
            this.grpFileName.Location = new System.Drawing.Point(15, 11);
            this.grpFileName.Name = "grpFileName";
            this.grpFileName.Size = new System.Drawing.Size(215, 69);
            this.grpFileName.TabIndex = 0;
            this.grpFileName.TabStop = false;
            this.grpFileName.Text = "File Name";
            // 
            // optFileNameSpecified
            // 
            this.optFileNameSpecified.AutoSize = true;
            this.optFileNameSpecified.Location = new System.Drawing.Point(16, 42);
            this.optFileNameSpecified.Name = "optFileNameSpecified";
            this.optFileNameSpecified.Size = new System.Drawing.Size(114, 17);
            this.optFileNameSpecified.TabIndex = 1;
            this.optFileNameSpecified.TabStop = true;
            this.optFileNameSpecified.Text = "Specified file name";
            this.optFileNameSpecified.UseVisualStyleBackColor = true;
            this.optFileNameSpecified.CheckedChanged += new System.EventHandler(this.optFileNameSpecified_CheckedChanged);
            // 
            // optFileNameAuto
            // 
            this.optFileNameAuto.AutoSize = true;
            this.optFileNameAuto.Location = new System.Drawing.Point(16, 19);
            this.optFileNameAuto.Name = "optFileNameAuto";
            this.optFileNameAuto.Size = new System.Drawing.Size(163, 17);
            this.optFileNameAuto.TabIndex = 0;
            this.optFileNameAuto.TabStop = true;
            this.optFileNameAuto.Text = "Auto file name based on date";
            this.optFileNameAuto.UseVisualStyleBackColor = true;
            this.optFileNameAuto.CheckedChanged += new System.EventHandler(this.optFileNameAuto_CheckedChanged);
            // 
            // lblClassName
            // 
            this.lblClassName.AutoSize = true;
            this.lblClassName.Location = new System.Drawing.Point(248, 11);
            this.lblClassName.Name = "lblClassName";
            this.lblClassName.Size = new System.Drawing.Size(63, 13);
            this.lblClassName.TabIndex = 1;
            this.lblClassName.Text = "Class Name";
            // 
            // txtClassName
            // 
            this.txtClassName.Location = new System.Drawing.Point(317, 11);
            this.txtClassName.Name = "txtClassName";
            this.txtClassName.Size = new System.Drawing.Size(306, 20);
            this.txtClassName.TabIndex = 2;
            // 
            // lblDateFormatFileName
            // 
            this.lblDateFormatFileName.AutoSize = true;
            this.lblDateFormatFileName.Location = new System.Drawing.Point(255, 44);
            this.lblDateFormatFileName.Name = "lblDateFormatFileName";
            this.lblDateFormatFileName.Size = new System.Drawing.Size(116, 13);
            this.lblDateFormatFileName.TabIndex = 3;
            this.lblDateFormatFileName.Text = "Date Format (file name)";
            // 
            // lblDateFormatFileContents
            // 
            this.lblDateFormatFileContents.AutoSize = true;
            this.lblDateFormatFileContents.Location = new System.Drawing.Point(252, 73);
            this.lblDateFormatFileContents.Name = "lblDateFormatFileContents";
            this.lblDateFormatFileContents.Size = new System.Drawing.Size(135, 13);
            this.lblDateFormatFileContents.TabIndex = 5;
            this.lblDateFormatFileContents.Text = "Date Format (File Contents)";
            // 
            // cmbDateFormatFileName
            // 
            this.cmbDateFormatFileName.FormattingEnabled = true;
            this.cmbDateFormatFileName.Location = new System.Drawing.Point(409, 44);
            this.cmbDateFormatFileName.Name = "cmbDateFormatFileName";
            this.cmbDateFormatFileName.Size = new System.Drawing.Size(121, 21);
            this.cmbDateFormatFileName.TabIndex = 4;
            // 
            // cmbDateFormatFileContents
            // 
            this.cmbDateFormatFileContents.FormattingEnabled = true;
            this.cmbDateFormatFileContents.Location = new System.Drawing.Point(409, 74);
            this.cmbDateFormatFileContents.Name = "cmbDateFormatFileContents";
            this.cmbDateFormatFileContents.Size = new System.Drawing.Size(121, 21);
            this.cmbDateFormatFileContents.TabIndex = 6;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 16;
            this.ssStatus.Text = "status";
            // 
            // chkAddReferences
            // 
            this.chkAddReferences.AutoSize = true;
            this.chkAddReferences.Location = new System.Drawing.Point(317, 404);
            this.chkAddReferences.Name = "chkAddReferences";
            this.chkAddReferences.Size = new System.Drawing.Size(103, 17);
            this.chkAddReferences.TabIndex = 13;
            this.chkAddReferences.Text = "Add References";
            this.chkAddReferences.UseVisualStyleBackColor = true;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(15, 401);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(62, 23);
            this.btnAdd.TabIndex = 11;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(83, 401);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.TabIndex = 12;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // FrmInsertCode_TextFileLogClass
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.chkAddReferences);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.cmbDateFormatFileContents);
            this.Controls.Add(this.cmbDateFormatFileName);
            this.Controls.Add(this.lblDateFormatFileContents);
            this.Controls.Add(this.lblDateFormatFileName);
            this.Controls.Add(this.txtClassName);
            this.Controls.Add(this.lblClassName);
            this.Controls.Add(this.grpFileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgParameters);
            this.Controls.Add(this.lblDestinationFolder);
            this.Controls.Add(this.txtDestinationFolder);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_TextFileLogClass";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_TextFileLogClass_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_TextFileLogClass_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_TextFileLogClass_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).EndInit();
            this.grpFileName.ResumeLayout(false);
            this.grpFileName.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TextBox txtDestinationFolder;
        private System.Windows.Forms.Label lblDestinationFolder;
        private System.Windows.Forms.DataGridView dgParameters;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColOptional;
        private System.Windows.Forms.DataGridViewComboBoxColumn colDataType;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpFileName;
        private System.Windows.Forms.RadioButton optFileNameSpecified;
        private System.Windows.Forms.RadioButton optFileNameAuto;
        private System.Windows.Forms.Label lblClassName;
        private System.Windows.Forms.TextBox txtClassName;
        private System.Windows.Forms.Label lblDateFormatFileName;
        private System.Windows.Forms.Label lblDateFormatFileContents;
        private System.Windows.Forms.ComboBox cmbDateFormatFileName;
        private System.Windows.Forms.ComboBox cmbDateFormatFileContents;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.CheckBox chkAddReferences;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemove;
    }
}