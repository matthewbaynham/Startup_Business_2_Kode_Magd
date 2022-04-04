namespace KodeMagd.Format
{
    partial class FrmRemoveLineNo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmRemoveLineNo));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.lblModules = new System.Windows.Forms.Label();
            this.lblFunctions = new System.Windows.Forms.Label();
            this.lstModules = new System.Windows.Forms.CheckedListBox();
            this.lstFunctions = new System.Windows.Forms.CheckedListBox();
            this.lblWarning = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 0;
            this.ssStatus.Text = "ssStatus";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(533, 394);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(441, 394);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.TabIndex = 2;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // lblModules
            // 
            this.lblModules.AutoSize = true;
            this.lblModules.Location = new System.Drawing.Point(21, 21);
            this.lblModules.Name = "lblModules";
            this.lblModules.Size = new System.Drawing.Size(47, 13);
            this.lblModules.TabIndex = 3;
            this.lblModules.Text = "Modules";
            // 
            // lblFunctions
            // 
            this.lblFunctions.AutoSize = true;
            this.lblFunctions.Location = new System.Drawing.Point(254, 21);
            this.lblFunctions.Name = "lblFunctions";
            this.lblFunctions.Size = new System.Drawing.Size(53, 13);
            this.lblFunctions.TabIndex = 4;
            this.lblFunctions.Text = "Functions";
            // 
            // lstModules
            // 
            this.lstModules.CheckOnClick = true;
            this.lstModules.FormattingEnabled = true;
            this.lstModules.Location = new System.Drawing.Point(19, 49);
            this.lstModules.Name = "lstModules";
            this.lstModules.Size = new System.Drawing.Size(214, 319);
            this.lstModules.TabIndex = 0;
            this.lstModules.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstModules_KeyUp);
            this.lstModules.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstModules_MouseUp);
            // 
            // lstFunctions
            // 
            this.lstFunctions.CheckOnClick = true;
            this.lstFunctions.FormattingEnabled = true;
            this.lstFunctions.Location = new System.Drawing.Point(257, 52);
            this.lstFunctions.Name = "lstFunctions";
            this.lstFunctions.Size = new System.Drawing.Size(351, 319);
            this.lstFunctions.TabIndex = 1;
            this.lstFunctions.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstFunctions_KeyUp);
            this.lstFunctions.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFunctions_MouseUp);
            // 
            // lblWarning
            // 
            this.lblWarning.AutoSize = true;
            this.lblWarning.Location = new System.Drawing.Point(16, 385);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(0, 13);
            this.lblWarning.TabIndex = 7;
            // 
            // FrmRemoveLineNo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.lstFunctions);
            this.Controls.Add(this.lstModules);
            this.Controls.Add(this.lblFunctions);
            this.Controls.Add(this.lblModules);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmRemoveLineNo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmRemoveLineNo_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmRemoveLineNo_KeyDown);
            this.Resize += new System.EventHandler(this.FrmRemoveLineNo_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Label lblModules;
        private System.Windows.Forms.Label lblFunctions;
        private System.Windows.Forms.CheckedListBox lstModules;
        private System.Windows.Forms.CheckedListBox lstFunctions;
        private System.Windows.Forms.Label lblWarning;
    }
}