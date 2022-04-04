namespace KodeMagd.License
{
    partial class FrmLoadingLicense
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmLoadingLicense));
            this.btnClose = new System.Windows.Forms.Button();
            this.txtDosCopy = new System.Windows.Forms.TextBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblWarning = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(365, 276);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(84, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txtDosCopy
            // 
            this.txtDosCopy.Location = new System.Drawing.Point(15, 186);
            this.txtDosCopy.Multiline = true;
            this.txtDosCopy.Name = "txtDosCopy";
            this.txtDosCopy.Size = new System.Drawing.Size(445, 80);
            this.txtDosCopy.TabIndex = 1;
            this.txtDosCopy.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDosCopy_KeyDown);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(15, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(196, 25);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Installing License...";
            // 
            // lblWarning
            // 
            this.lblWarning.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWarning.Location = new System.Drawing.Point(14, 54);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(446, 82);
            this.lblWarning.TabIndex = 3;
            this.lblWarning.Text = "Warning: Different security settings can be troublesome.  If this automated proce" +
    "sses for installing the license fails then please deal with your own security se" +
    "ttings and run the DOS below.";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 314);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(472, 22);
            this.ssStatus.TabIndex = 4;
            // 
            // FrmLoadingLicense
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 336);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.txtDosCopy);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmLoadingLicense";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmLoadingLicense_Load);
            this.Resize += new System.EventHandler(this.FrmLoadingLicense_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.TextBox txtDosCopy;
    }
}