namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_OpenForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_OpenForm));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.dgParameters = new System.Windows.Forms.DataGridView();
            this.ColInternalVariable = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColExternalProperty = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColVariable = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.ColDataType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.lblFormName = new System.Windows.Forms.Label();
            this.txtFormName = new System.Windows.Forms.TextBox();
            this.cmbForms = new System.Windows.Forms.ComboBox();
            this.chkNewForm = new System.Windows.Forms.CheckBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.txtInstanceName = new System.Windows.Forms.TextBox();
            this.lblInstanceName = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 399);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 26);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(451, 399);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 26);
            this.btnGenerate.TabIndex = 8;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // dgParameters
            // 
            this.dgParameters.AllowUserToAddRows = false;
            this.dgParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColInternalVariable,
            this.ColExternalProperty,
            this.ColValue,
            this.ColVariable,
            this.ColDataType});
            this.dgParameters.Location = new System.Drawing.Point(12, 73);
            this.dgParameters.Name = "dgParameters";
            this.dgParameters.Size = new System.Drawing.Size(608, 303);
            this.dgParameters.TabIndex = 5;
            this.dgParameters.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgParameters_CellValueChanged);
            this.dgParameters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgParameters_RowsAdded);
            // 
            // ColInternalVariable
            // 
            this.ColInternalVariable.HeaderText = "Internal Variable";
            this.ColInternalVariable.Name = "ColInternalVariable";
            // 
            // ColExternalProperty
            // 
            this.ColExternalProperty.HeaderText = "External Property Name";
            this.ColExternalProperty.Name = "ColExternalProperty";
            // 
            // ColValue
            // 
            this.ColValue.HeaderText = "Value";
            this.ColValue.Name = "ColValue";
            // 
            // ColVariable
            // 
            this.ColVariable.HeaderText = "Variable";
            this.ColVariable.Name = "ColVariable";
            // 
            // ColDataType
            // 
            this.ColDataType.HeaderText = "Data Type";
            this.ColDataType.Name = "ColDataType";
            // 
            // lblFormName
            // 
            this.lblFormName.AutoSize = true;
            this.lblFormName.Location = new System.Drawing.Point(329, 9);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(61, 13);
            this.lblFormName.TabIndex = 3;
            this.lblFormName.Text = "Form Name";
            // 
            // txtFormName
            // 
            this.txtFormName.Location = new System.Drawing.Point(328, 35);
            this.txtFormName.Name = "txtFormName";
            this.txtFormName.Size = new System.Drawing.Size(292, 20);
            this.txtFormName.TabIndex = 4;
            // 
            // cmbForms
            // 
            this.cmbForms.FormattingEnabled = true;
            this.cmbForms.Location = new System.Drawing.Point(328, 34);
            this.cmbForms.Name = "cmbForms";
            this.cmbForms.Size = new System.Drawing.Size(292, 21);
            this.cmbForms.TabIndex = 4;
            // 
            // chkNewForm
            // 
            this.chkNewForm.AutoSize = true;
            this.chkNewForm.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkNewForm.Location = new System.Drawing.Point(248, 38);
            this.chkNewForm.Name = "chkNewForm";
            this.chkNewForm.Size = new System.Drawing.Size(74, 17);
            this.chkNewForm.TabIndex = 2;
            this.chkNewForm.Text = "New Form";
            this.chkNewForm.UseVisualStyleBackColor = true;
            this.chkNewForm.CheckedChanged += new System.EventHandler(this.chkNewForm_CheckedChanged);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 10;
            this.ssStatus.Text = "status";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(12, 396);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 26);
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(107, 396);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 26);
            this.btnRemove.TabIndex = 7;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // txtInstanceName
            // 
            this.txtInstanceName.Location = new System.Drawing.Point(13, 36);
            this.txtInstanceName.Name = "txtInstanceName";
            this.txtInstanceName.Size = new System.Drawing.Size(201, 20);
            this.txtInstanceName.TabIndex = 1;
            // 
            // lblInstanceName
            // 
            this.lblInstanceName.AutoSize = true;
            this.lblInstanceName.Location = new System.Drawing.Point(12, 9);
            this.lblInstanceName.Name = "lblInstanceName";
            this.lblInstanceName.Size = new System.Drawing.Size(70, 13);
            this.lblInstanceName.TabIndex = 0;
            this.lblInstanceName.Text = "Name (Suffix)";
            // 
            // FrmInsertCode_OpenForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.lblInstanceName);
            this.Controls.Add(this.txtInstanceName);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.chkNewForm);
            this.Controls.Add(this.cmbForms);
            this.Controls.Add(this.txtFormName);
            this.Controls.Add(this.lblFormName);
            this.Controls.Add(this.dgParameters);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_OpenForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_OpenForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_OpenForm_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_OpenForm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.DataGridView dgParameters;
        private System.Windows.Forms.Label lblFormName;
        private System.Windows.Forms.TextBox txtFormName;
        private System.Windows.Forms.ComboBox cmbForms;
        private System.Windows.Forms.CheckBox chkNewForm;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColInternalVariable;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColExternalProperty;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValue;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColVariable;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDataType;
        private System.Windows.Forms.TextBox txtInstanceName;
        private System.Windows.Forms.Label lblInstanceName;
    }
}