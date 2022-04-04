namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_CommandBarClass_NodeEdit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_CommandBarClass_NodeEdit));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtCaption = new System.Windows.Forms.TextBox();
            this.lblCaption = new System.Windows.Forms.Label();
            this.lblMacroToRun = new System.Windows.Forms.Label();
            this.lblToolTipText = new System.Windows.Forms.Label();
            this.txtToolTipText = new System.Windows.Forms.TextBox();
            this.cmbControlType = new System.Windows.Forms.ComboBox();
            this.lblControlType = new System.Windows.Forms.Label();
            this.lstValues = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnMinus = new System.Windows.Forms.Button();
            this.lblMessage = new System.Windows.Forms.Label();
            this.cmbMacroToRun = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(403, 298);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(55, 23);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(342, 298);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(55, 23);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtCaption
            // 
            this.txtCaption.Location = new System.Drawing.Point(17, 25);
            this.txtCaption.Name = "txtCaption";
            this.txtCaption.Size = new System.Drawing.Size(443, 20);
            this.txtCaption.TabIndex = 1;
            // 
            // lblCaption
            // 
            this.lblCaption.AutoSize = true;
            this.lblCaption.Location = new System.Drawing.Point(12, 9);
            this.lblCaption.Name = "lblCaption";
            this.lblCaption.Size = new System.Drawing.Size(43, 13);
            this.lblCaption.TabIndex = 0;
            this.lblCaption.Text = "Caption";
            // 
            // lblMacroToRun
            // 
            this.lblMacroToRun.AutoSize = true;
            this.lblMacroToRun.Location = new System.Drawing.Point(12, 59);
            this.lblMacroToRun.Name = "lblMacroToRun";
            this.lblMacroToRun.Size = new System.Drawing.Size(72, 13);
            this.lblMacroToRun.TabIndex = 2;
            this.lblMacroToRun.Text = "Macro to Run";
            // 
            // lblToolTipText
            // 
            this.lblToolTipText.AutoSize = true;
            this.lblToolTipText.Location = new System.Drawing.Point(14, 111);
            this.lblToolTipText.Name = "lblToolTipText";
            this.lblToolTipText.Size = new System.Drawing.Size(63, 13);
            this.lblToolTipText.TabIndex = 4;
            this.lblToolTipText.Text = "Tooltip Text";
            // 
            // txtToolTipText
            // 
            this.txtToolTipText.AcceptsReturn = true;
            this.txtToolTipText.Location = new System.Drawing.Point(17, 127);
            this.txtToolTipText.Multiline = true;
            this.txtToolTipText.Name = "txtToolTipText";
            this.txtToolTipText.Size = new System.Drawing.Size(439, 70);
            this.txtToolTipText.TabIndex = 5;
            // 
            // cmbControlType
            // 
            this.cmbControlType.FormattingEnabled = true;
            this.cmbControlType.Location = new System.Drawing.Point(18, 228);
            this.cmbControlType.Name = "cmbControlType";
            this.cmbControlType.Size = new System.Drawing.Size(121, 21);
            this.cmbControlType.TabIndex = 7;
            this.cmbControlType.TextChanged += new System.EventHandler(this.cmbControlType_TextChanged);
            this.cmbControlType.Validating += new System.ComponentModel.CancelEventHandler(this.cmbControlType_Validating);
            // 
            // lblControlType
            // 
            this.lblControlType.AutoSize = true;
            this.lblControlType.Location = new System.Drawing.Point(17, 211);
            this.lblControlType.Name = "lblControlType";
            this.lblControlType.Size = new System.Drawing.Size(67, 13);
            this.lblControlType.TabIndex = 6;
            this.lblControlType.Text = "Control Type";
            // 
            // lstValues
            // 
            this.lstValues.FormattingEnabled = true;
            this.lstValues.Location = new System.Drawing.Point(155, 210);
            this.lstValues.Name = "lstValues";
            this.lstValues.Size = new System.Drawing.Size(120, 95);
            this.lstValues.TabIndex = 8;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(281, 211);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(23, 23);
            this.btnAdd.TabIndex = 9;
            this.btnAdd.Text = "+";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnMinus
            // 
            this.btnMinus.Location = new System.Drawing.Point(281, 240);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(24, 25);
            this.btnMinus.TabIndex = 10;
            this.btnMinus.Text = "-";
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(12, 308);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(60, 13);
            this.lblMessage.TabIndex = 13;
            this.lblMessage.Text = "lblMessage";
            // 
            // cmbMacroToRun
            // 
            this.cmbMacroToRun.FormattingEnabled = true;
            this.cmbMacroToRun.Location = new System.Drawing.Point(17, 78);
            this.cmbMacroToRun.Name = "cmbMacroToRun";
            this.cmbMacroToRun.Size = new System.Drawing.Size(442, 21);
            this.cmbMacroToRun.TabIndex = 3;
            // 
            // FrmInsertCode_CommandBarClass_NodeEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 338);
            this.Controls.Add(this.cmbMacroToRun);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lstValues);
            this.Controls.Add(this.lblControlType);
            this.Controls.Add(this.cmbControlType);
            this.Controls.Add(this.txtToolTipText);
            this.Controls.Add(this.lblToolTipText);
            this.Controls.Add(this.lblMacroToRun);
            this.Controls.Add(this.lblCaption);
            this.Controls.Add(this.txtCaption);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmInsertCode_CommandBarClass_NodeEdit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_CommandBarClass_NodeEdit_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_CommandBarClass_NodeEdit_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_CommandBarClass_NodeEdit_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtCaption;
        private System.Windows.Forms.Label lblCaption;
        private System.Windows.Forms.Label lblMacroToRun;
        private System.Windows.Forms.Label lblToolTipText;
        private System.Windows.Forms.TextBox txtToolTipText;
        private System.Windows.Forms.ComboBox cmbControlType;
        private System.Windows.Forms.Label lblControlType;
        private System.Windows.Forms.ListBox lstValues;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.ComboBox cmbMacroToRun;
    }
}