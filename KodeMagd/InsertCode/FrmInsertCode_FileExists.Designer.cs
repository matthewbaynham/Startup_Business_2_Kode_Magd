namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_FileExists
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_FileExists));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.chkAddReference = new System.Windows.Forms.CheckBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.optHardcoded = new System.Windows.Forms.RadioButton();
            this.optVariable = new System.Windows.Forms.RadioButton();
            this.cmbVariable = new System.Windows.Forms.ComboBox();
            this.lblVariable = new System.Windows.Forms.Label();
            this.ofdBrowseOpen = new System.Windows.Forms.OpenFileDialog();
            this.lblName = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.grpType.SuspendLayout();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 10;
            this.ssStatus.Text = "statusStrip1";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(520, 378);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(85, 29);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(541, 196);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 6;
            this.btnBrowse.Text = "&Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(9, 152);
            this.txtPath.Multiline = true;
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(607, 38);
            this.txtPath.TabIndex = 5;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(12, 128);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(29, 13);
            this.lblPath.TabIndex = 4;
            this.lblPath.Text = "Path";
            // 
            // chkAddReference
            // 
            this.chkAddReference.AutoSize = true;
            this.chkAddReference.Location = new System.Drawing.Point(318, 385);
            this.chkAddReference.Name = "chkAddReference";
            this.chkAddReference.Size = new System.Drawing.Size(98, 17);
            this.chkAddReference.TabIndex = 7;
            this.chkAddReference.Text = "Add Reference";
            this.chkAddReference.UseVisualStyleBackColor = true;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(427, 378);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(87, 29);
            this.btnGenerate.TabIndex = 8;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optHardcoded);
            this.grpType.Controls.Add(this.optVariable);
            this.grpType.Location = new System.Drawing.Point(21, 16);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(114, 72);
            this.grpType.TabIndex = 0;
            this.grpType.TabStop = false;
            // 
            // optHardcoded
            // 
            this.optHardcoded.AutoSize = true;
            this.optHardcoded.Location = new System.Drawing.Point(6, 42);
            this.optHardcoded.Name = "optHardcoded";
            this.optHardcoded.Size = new System.Drawing.Size(78, 17);
            this.optHardcoded.TabIndex = 1;
            this.optHardcoded.TabStop = true;
            this.optHardcoded.Text = "Hardcoded";
            this.optHardcoded.UseVisualStyleBackColor = true;
            this.optHardcoded.CheckedChanged += new System.EventHandler(this.optHardcoded_CheckedChanged);
            // 
            // optVariable
            // 
            this.optVariable.AutoSize = true;
            this.optVariable.Location = new System.Drawing.Point(6, 19);
            this.optVariable.Name = "optVariable";
            this.optVariable.Size = new System.Drawing.Size(63, 17);
            this.optVariable.TabIndex = 0;
            this.optVariable.TabStop = true;
            this.optVariable.Text = "Variable";
            this.optVariable.UseVisualStyleBackColor = true;
            this.optVariable.CheckedChanged += new System.EventHandler(this.optVariable_CheckedChanged);
            // 
            // cmbVariable
            // 
            this.cmbVariable.FormattingEnabled = true;
            this.cmbVariable.Location = new System.Drawing.Point(9, 152);
            this.cmbVariable.Name = "cmbVariable";
            this.cmbVariable.Size = new System.Drawing.Size(308, 21);
            this.cmbVariable.TabIndex = 3;
            // 
            // lblVariable
            // 
            this.lblVariable.AutoSize = true;
            this.lblVariable.Location = new System.Drawing.Point(12, 128);
            this.lblVariable.Name = "lblVariable";
            this.lblVariable.Size = new System.Drawing.Size(45, 13);
            this.lblVariable.TabIndex = 4;
            this.lblVariable.Text = "Variable";
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Location = new System.Drawing.Point(278, 21);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(70, 13);
            this.lblName.TabIndex = 1;
            this.lblName.Text = "Name (Suffix)";
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(354, 18);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(262, 20);
            this.txtName.TabIndex = 2;
            // 
            // FrmInsertCode_FileExists
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.txtName);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.lblVariable);
            this.Controls.Add(this.cmbVariable);
            this.Controls.Add(this.grpType);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.chkAddReference);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_FileExists";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmInsertCode_FileExists";
            this.Load += new System.EventHandler(this.FrmInsertCode_FileExists_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_FileExists_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_FileExists_Resize);
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.CheckBox chkAddReference;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton optHardcoded;
        private System.Windows.Forms.RadioButton optVariable;
        private System.Windows.Forms.ComboBox cmbVariable;
        private System.Windows.Forms.Label lblVariable;
        private System.Windows.Forms.OpenFileDialog ofdBrowseOpen;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.TextBox txtName;
    }
}