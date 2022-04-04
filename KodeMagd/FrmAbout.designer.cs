namespace KodeMagd
{
    partial class FrmAbout
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAbout));
            this.btnClose = new System.Windows.Forms.Button();
            this.lblKodeMagd = new System.Windows.Forms.LinkLabel();
            this.pnlKodeMagdImage = new System.Windows.Forms.Panel();
            this.rtbInfo = new System.Windows.Forms.RichTextBox();
            this.txtMachineID = new System.Windows.Forms.TextBox();
            this.lblMachineID = new System.Windows.Forms.Label();
            this.lblLicenseID = new System.Windows.Forms.Label();
            this.txtLicenseID = new System.Windows.Forms.TextBox();
            this.lblVersion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(536, 401);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblKodeMagd
            // 
            this.lblKodeMagd.AutoSize = true;
            this.lblKodeMagd.Location = new System.Drawing.Point(12, 350);
            this.lblKodeMagd.Name = "lblKodeMagd";
            this.lblKodeMagd.Size = new System.Drawing.Size(104, 13);
            this.lblKodeMagd.TabIndex = 3;
            this.lblKodeMagd.TabStop = true;
            this.lblKodeMagd.Text = "Kode Magd Website";
            this.lblKodeMagd.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblKodeMagd_LinkClicked);
            // 
            // pnlKodeMagdImage
            // 
            this.pnlKodeMagdImage.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pnlKodeMagdImage.BackgroundImage")));
            this.pnlKodeMagdImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pnlKodeMagdImage.Location = new System.Drawing.Point(3, 22);
            this.pnlKodeMagdImage.Name = "pnlKodeMagdImage";
            this.pnlKodeMagdImage.Size = new System.Drawing.Size(326, 314);
            this.pnlKodeMagdImage.TabIndex = 4;
            this.pnlKodeMagdImage.Click += new System.EventHandler(this.pnlKodeMagdImage_Click);
            // 
            // rtbInfo
            // 
            this.rtbInfo.Location = new System.Drawing.Point(335, 22);
            this.rtbInfo.Name = "rtbInfo";
            this.rtbInfo.ReadOnly = true;
            this.rtbInfo.Size = new System.Drawing.Size(287, 373);
            this.rtbInfo.TabIndex = 5;
            this.rtbInfo.Text = "Big words little words bold words underlined";
            // 
            // txtMachineID
            // 
            this.txtMachineID.Location = new System.Drawing.Point(80, 375);
            this.txtMachineID.Name = "txtMachineID";
            this.txtMachineID.ReadOnly = true;
            this.txtMachineID.Size = new System.Drawing.Size(249, 20);
            this.txtMachineID.TabIndex = 8;
            // 
            // lblMachineID
            // 
            this.lblMachineID.AutoSize = true;
            this.lblMachineID.Location = new System.Drawing.Point(0, 382);
            this.lblMachineID.Name = "lblMachineID";
            this.lblMachineID.Size = new System.Drawing.Size(62, 13);
            this.lblMachineID.TabIndex = 9;
            this.lblMachineID.Text = "Machine ID";
            // 
            // lblLicenseID
            // 
            this.lblLicenseID.AutoSize = true;
            this.lblLicenseID.Location = new System.Drawing.Point(0, 412);
            this.lblLicenseID.Name = "lblLicenseID";
            this.lblLicenseID.Size = new System.Drawing.Size(58, 13);
            this.lblLicenseID.TabIndex = 10;
            this.lblLicenseID.Text = "License ID";
            // 
            // txtLicenseID
            // 
            this.txtLicenseID.Location = new System.Drawing.Point(80, 409);
            this.txtLicenseID.Name = "txtLicenseID";
            this.txtLicenseID.ReadOnly = true;
            this.txtLicenseID.Size = new System.Drawing.Size(249, 20);
            this.txtLicenseID.TabIndex = 11;
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(122, 350);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(52, 13);
            this.lblVersion.TabIndex = 12;
            this.lblVersion.Text = "lblVersion";
            // 
            // FrmAbout
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.txtLicenseID);
            this.Controls.Add(this.lblLicenseID);
            this.Controls.Add(this.lblMachineID);
            this.Controls.Add(this.txtMachineID);
            this.Controls.Add(this.rtbInfo);
            this.Controls.Add(this.pnlKodeMagdImage);
            this.Controls.Add(this.lblKodeMagd);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmAbout";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "About";
            this.Load += new System.EventHandler(this.FrmAbout_Load);
            this.Resize += new System.EventHandler(this.FrmAbout_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.LinkLabel lblKodeMagd;
        private System.Windows.Forms.Panel pnlKodeMagdImage;
        private System.Windows.Forms.RichTextBox rtbInfo;
        private System.Windows.Forms.TextBox txtMachineID;
        private System.Windows.Forms.Label lblMachineID;
        private System.Windows.Forms.Label lblLicenseID;
        private System.Windows.Forms.TextBox txtLicenseID;
        private System.Windows.Forms.Label lblVersion;
    }
}