namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_ConnectionString
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_ConnectionString));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnBuild = new System.Windows.Forms.Button();
            this.txtConnectionString = new System.Windows.Forms.TextBox();
            this.btnRecent = new System.Windows.Forms.Button();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.txtControl = new System.Windows.Forms.TextBox();
            this.cmbControl = new System.Windows.Forms.ComboBox();
            this.lblVariable = new System.Windows.Forms.Label();
            this.chkControlFromList = new System.Windows.Forms.CheckBox();
            this.chkAddReference = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 9;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(530, 380);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 8;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(440, 380);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 7;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(15, 38);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(75, 23);
            this.btnBuild.TabIndex = 1;
            this.btnBuild.Text = "&Build";
            this.btnBuild.UseVisualStyleBackColor = true;
            this.btnBuild.Click += new System.EventHandler(this.btnBuild_Click);
            // 
            // txtConnectionString
            // 
            this.txtConnectionString.Location = new System.Drawing.Point(100, 40);
            this.txtConnectionString.Multiline = true;
            this.txtConnectionString.Name = "txtConnectionString";
            this.txtConnectionString.Size = new System.Drawing.Size(520, 306);
            this.txtConnectionString.TabIndex = 3;
            // 
            // btnRecent
            // 
            this.btnRecent.Location = new System.Drawing.Point(15, 67);
            this.btnRecent.Name = "btnRecent";
            this.btnRecent.Size = new System.Drawing.Size(75, 23);
            this.btnRecent.TabIndex = 2;
            this.btnRecent.Text = "&Recent";
            this.btnRecent.UseVisualStyleBackColor = true;
            this.btnRecent.Click += new System.EventHandler(this.btnRecent_Click);
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(97, 9);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 0;
            this.lblConnectionString.Text = "Connection String";
            // 
            // txtControl
            // 
            this.txtControl.Location = new System.Drawing.Point(58, 377);
            this.txtControl.Name = "txtControl";
            this.txtControl.Size = new System.Drawing.Size(212, 20);
            this.txtControl.TabIndex = 7;
            // 
            // cmbControl
            // 
            this.cmbControl.FormattingEnabled = true;
            this.cmbControl.Location = new System.Drawing.Point(58, 376);
            this.cmbControl.Name = "cmbControl";
            this.cmbControl.Size = new System.Drawing.Size(212, 21);
            this.cmbControl.TabIndex = 5;
            // 
            // lblVariable
            // 
            this.lblVariable.AutoSize = true;
            this.lblVariable.Location = new System.Drawing.Point(12, 379);
            this.lblVariable.Name = "lblVariable";
            this.lblVariable.Size = new System.Drawing.Size(45, 13);
            this.lblVariable.TabIndex = 4;
            this.lblVariable.Text = "Variable";
            // 
            // chkControlFromList
            // 
            this.chkControlFromList.AutoSize = true;
            this.chkControlFromList.Location = new System.Drawing.Point(50, 410);
            this.chkControlFromList.Name = "chkControlFromList";
            this.chkControlFromList.Size = new System.Drawing.Size(104, 17);
            this.chkControlFromList.TabIndex = 6;
            this.chkControlFromList.Text = "Control From List";
            this.chkControlFromList.UseVisualStyleBackColor = true;
            this.chkControlFromList.CheckedChanged += new System.EventHandler(this.chkControlFromList_CheckedChanged);
            // 
            // chkAddReference
            // 
            this.chkAddReference.AutoSize = true;
            this.chkAddReference.Location = new System.Drawing.Point(336, 380);
            this.chkAddReference.Name = "chkAddReference";
            this.chkAddReference.Size = new System.Drawing.Size(98, 17);
            this.chkAddReference.TabIndex = 10;
            this.chkAddReference.Text = "Add Reference";
            this.chkAddReference.UseVisualStyleBackColor = true;
            // 
            // FrmInsertCode_ConnectionString
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.chkAddReference);
            this.Controls.Add(this.chkControlFromList);
            this.Controls.Add(this.lblVariable);
            this.Controls.Add(this.cmbControl);
            this.Controls.Add(this.txtControl);
            this.Controls.Add(this.lblConnectionString);
            this.Controls.Add(this.btnRecent);
            this.Controls.Add(this.txtConnectionString);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_ConnectionString";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_ConnectionString_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_ConnectionString_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_ConnectionString_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.TextBox txtConnectionString;
        private System.Windows.Forms.Button btnRecent;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TextBox txtControl;
        private System.Windows.Forms.ComboBox cmbControl;
        private System.Windows.Forms.Label lblVariable;
        private System.Windows.Forms.CheckBox chkControlFromList;
        private System.Windows.Forms.CheckBox chkAddReference;
    }
}