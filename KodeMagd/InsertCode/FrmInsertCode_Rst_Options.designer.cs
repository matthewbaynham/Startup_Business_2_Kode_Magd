namespace KodeMagd
{
    partial class FrmRstOpenLoopClose_Options
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRstOpenLoopClose_Options));
            this.btnClose = new System.Windows.Forms.Button();
            this.cmbLockType = new System.Windows.Forms.ComboBox();
            this.cmbCursorType = new System.Windows.Forms.ComboBox();
            this.lblLockType = new System.Windows.Forms.Label();
            this.lblCursorType = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(357, 266);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // cmbLockType
            // 
            this.cmbLockType.FormattingEnabled = true;
            this.cmbLockType.Location = new System.Drawing.Point(85, 12);
            this.cmbLockType.Name = "cmbLockType";
            this.cmbLockType.Size = new System.Drawing.Size(215, 21);
            this.cmbLockType.TabIndex = 1;
            // 
            // cmbCursorType
            // 
            this.cmbCursorType.FormattingEnabled = true;
            this.cmbCursorType.Location = new System.Drawing.Point(85, 51);
            this.cmbCursorType.Name = "cmbCursorType";
            this.cmbCursorType.Size = new System.Drawing.Size(215, 21);
            this.cmbCursorType.TabIndex = 3;
            // 
            // lblLockType
            // 
            this.lblLockType.AutoSize = true;
            this.lblLockType.Location = new System.Drawing.Point(15, 12);
            this.lblLockType.Name = "lblLockType";
            this.lblLockType.Size = new System.Drawing.Size(58, 13);
            this.lblLockType.TabIndex = 0;
            this.lblLockType.Text = "Lock Type";
            // 
            // lblCursorType
            // 
            this.lblCursorType.AutoSize = true;
            this.lblCursorType.Location = new System.Drawing.Point(15, 51);
            this.lblCursorType.Name = "lblCursorType";
            this.lblCursorType.Size = new System.Drawing.Size(64, 13);
            this.lblCursorType.TabIndex = 2;
            this.lblCursorType.Text = "Cursor Type";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 316);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(474, 22);
            this.ssStatus.TabIndex = 5;
            this.ssStatus.Text = "status";
            // 
            // FrmRstOpenLoopClose_Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 338);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblCursorType);
            this.Controls.Add(this.lblLockType);
            this.Controls.Add(this.cmbCursorType);
            this.Controls.Add(this.cmbLockType);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmRstOpenLoopClose_Options";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Options";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmRstOpenLoopClose_Options_FormClosing);
            this.Load += new System.EventHandler(this.FrmRstOpenLoopClose_Options_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmRstOpenLoopClose_Options_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ComboBox cmbLockType;
        private System.Windows.Forms.ComboBox cmbCursorType;
        private System.Windows.Forms.Label lblLockType;
        private System.Windows.Forms.Label lblCursorType;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}