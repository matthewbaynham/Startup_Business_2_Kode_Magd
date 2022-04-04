namespace KodeMagd.WorkbookAnalysis
{
    partial class FrmCodeInColour
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCodeInColour));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOuputCodeInColour = new System.Windows.Forms.Button();
            this.lstFunctions = new System.Windows.Forms.CheckedListBox();
            this.lblFunctions = new System.Windows.Forms.Label();
            this.lstModules = new System.Windows.Forms.CheckedListBox();
            this.lblModules = new System.Windows.Forms.Label();
            this.btnSettings = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 431);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 394);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOuputCodeInColour
            // 
            this.btnOuputCodeInColour.Location = new System.Drawing.Point(406, 394);
            this.btnOuputCodeInColour.Name = "btnOuputCodeInColour";
            this.btnOuputCodeInColour.Size = new System.Drawing.Size(133, 23);
            this.btnOuputCodeInColour.TabIndex = 11;
            this.btnOuputCodeInColour.Text = "&Output Code in Colour";
            this.btnOuputCodeInColour.UseVisualStyleBackColor = true;
            this.btnOuputCodeInColour.Click += new System.EventHandler(this.btnOuputCodeInColour_Click);
            // 
            // lstFunctions
            // 
            this.lstFunctions.CheckOnClick = true;
            this.lstFunctions.FormattingEnabled = true;
            this.lstFunctions.Location = new System.Drawing.Point(265, 37);
            this.lstFunctions.Name = "lstFunctions";
            this.lstFunctions.Size = new System.Drawing.Size(346, 319);
            this.lstFunctions.TabIndex = 9;
            this.lstFunctions.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstFunctions_KeyUp);
            this.lstFunctions.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFunctions_MouseUp);
            // 
            // lblFunctions
            // 
            this.lblFunctions.AutoSize = true;
            this.lblFunctions.Location = new System.Drawing.Point(262, 10);
            this.lblFunctions.Name = "lblFunctions";
            this.lblFunctions.Size = new System.Drawing.Size(146, 13);
            this.lblFunctions.TabIndex = 8;
            this.lblFunctions.Text = "Functions / Subs / Properties";
            // 
            // lstModules
            // 
            this.lstModules.CheckOnClick = true;
            this.lstModules.FormattingEnabled = true;
            this.lstModules.Location = new System.Drawing.Point(13, 35);
            this.lstModules.Name = "lstModules";
            this.lstModules.Size = new System.Drawing.Size(243, 379);
            this.lstModules.TabIndex = 7;
            this.lstModules.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstModules_KeyUp);
            this.lstModules.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstModules_MouseUp);
            // 
            // lblModules
            // 
            this.lblModules.AutoSize = true;
            this.lblModules.Location = new System.Drawing.Point(14, 10);
            this.lblModules.Name = "lblModules";
            this.lblModules.Size = new System.Drawing.Size(47, 13);
            this.lblModules.TabIndex = 6;
            this.lblModules.Text = "Modules";
            // 
            // btnSettings
            // 
            this.btnSettings.Location = new System.Drawing.Point(274, 392);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(75, 23);
            this.btnSettings.TabIndex = 12;
            this.btnSettings.Text = "Settings";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // FrmCodeInColour
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 453);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.btnOuputCodeInColour);
            this.Controls.Add(this.lstFunctions);
            this.Controls.Add(this.lblFunctions);
            this.Controls.Add(this.lstModules);
            this.Controls.Add(this.lblModules);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmCodeInColour";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmCodeInColour_Load);
            this.Resize += new System.EventHandler(this.FrmCodeInColour_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOuputCodeInColour;
        private System.Windows.Forms.CheckedListBox lstFunctions;
        private System.Windows.Forms.Label lblFunctions;
        private System.Windows.Forms.CheckedListBox lstModules;
        private System.Windows.Forms.Label lblModules;
        private System.Windows.Forms.Button btnSettings;
    }
}