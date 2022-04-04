namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_ErrorHandler
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_ErrorHandler));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.lblModules = new System.Windows.Forms.Label();
            this.lstModules = new System.Windows.Forms.CheckedListBox();
            this.lblFunctions = new System.Windows.Forms.Label();
            this.lstFunctions = new System.Windows.Forms.CheckedListBox();
            this.lblWarning = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnAddHandler = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 7;
            this.ssStatus.Text = "statusStrip1";
            // 
            // lblModules
            // 
            this.lblModules.AutoSize = true;
            this.lblModules.Location = new System.Drawing.Point(16, 11);
            this.lblModules.Name = "lblModules";
            this.lblModules.Size = new System.Drawing.Size(47, 13);
            this.lblModules.TabIndex = 0;
            this.lblModules.Text = "Modules";
            // 
            // lstModules
            // 
            this.lstModules.CheckOnClick = true;
            this.lstModules.FormattingEnabled = true;
            this.lstModules.Location = new System.Drawing.Point(15, 36);
            this.lstModules.Name = "lstModules";
            this.lstModules.Size = new System.Drawing.Size(243, 379);
            this.lstModules.TabIndex = 1;
            this.lstModules.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstModules_KeyUp);
            this.lstModules.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstModules_MouseUp);
            // 
            // lblFunctions
            // 
            this.lblFunctions.AutoSize = true;
            this.lblFunctions.Location = new System.Drawing.Point(264, 11);
            this.lblFunctions.Name = "lblFunctions";
            this.lblFunctions.Size = new System.Drawing.Size(146, 13);
            this.lblFunctions.TabIndex = 2;
            this.lblFunctions.Text = "Functions / Subs / Properties";
            // 
            // lstFunctions
            // 
            this.lstFunctions.CheckOnClick = true;
            this.lstFunctions.FormattingEnabled = true;
            this.lstFunctions.Location = new System.Drawing.Point(267, 38);
            this.lstFunctions.Name = "lstFunctions";
            this.lstFunctions.Size = new System.Drawing.Size(346, 319);
            this.lstFunctions.TabIndex = 3;
            this.lstFunctions.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstFunctions_KeyUp);
            this.lstFunctions.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFunctions_MouseUp);
            // 
            // lblWarning
            // 
            this.lblWarning.AutoSize = true;
            this.lblWarning.Location = new System.Drawing.Point(264, 360);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(221, 13);
            this.lblWarning.TabIndex = 4;
            this.lblWarning.Text = "(*) Function has / sub / property on error goto";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(528, 395);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(85, 23);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnAddHandler
            // 
            this.btnAddHandler.Location = new System.Drawing.Point(431, 395);
            this.btnAddHandler.Name = "btnAddHandler";
            this.btnAddHandler.Size = new System.Drawing.Size(91, 23);
            this.btnAddHandler.TabIndex = 5;
            this.btnAddHandler.Text = "&Add Handler";
            this.btnAddHandler.UseVisualStyleBackColor = true;
            this.btnAddHandler.Click += new System.EventHandler(this.btnAddHandler_Click);
            // 
            // FrmInsertCode_ErrorHandler
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.btnAddHandler);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.lstFunctions);
            this.Controls.Add(this.lblFunctions);
            this.Controls.Add(this.lstModules);
            this.Controls.Add(this.lblModules);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_ErrorHandler";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmInsertCode_ErrorHandler";
            this.Load += new System.EventHandler(this.FrmInsertCode_ErrorHandler_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_ErrorHandler_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_ErrorHandler_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Label lblModules;
        private System.Windows.Forms.CheckedListBox lstModules;
        private System.Windows.Forms.Label lblFunctions;
        private System.Windows.Forms.CheckedListBox lstFunctions;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnAddHandler;
    }
}