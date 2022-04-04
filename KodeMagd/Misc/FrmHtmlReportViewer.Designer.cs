namespace KodeMagd.Misc
{
    partial class FrmHtmlReportViewer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmHtmlReportViewer));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.pnlHtml = new System.Windows.Forms.Panel();
            this.webHtml = new System.Windows.Forms.WebBrowser();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnSaveOpenClose = new System.Windows.Forms.Button();
            this.lblSaveOpenClose = new System.Windows.Forms.Label();
            this.pnlHtml.SuspendLayout();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 434);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 3;
            this.ssStatus.Text = "statusStrip1";
            // 
            // pnlHtml
            // 
            this.pnlHtml.Controls.Add(this.webHtml);
            this.pnlHtml.Location = new System.Drawing.Point(4, 4);
            this.pnlHtml.Name = "pnlHtml";
            this.pnlHtml.Size = new System.Drawing.Size(632, 370);
            this.pnlHtml.TabIndex = 1;
            // 
            // webHtml
            // 
            this.webHtml.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webHtml.Location = new System.Drawing.Point(0, 0);
            this.webHtml.Margin = new System.Windows.Forms.Padding(0);
            this.webHtml.MinimumSize = new System.Drawing.Size(20, 20);
            this.webHtml.Name = "webHtml";
            this.webHtml.ScriptErrorsSuppressed = true;
            this.webHtml.Size = new System.Drawing.Size(632, 370);
            this.webHtml.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(531, 380);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(24, 380);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnSaveOpenClose
            // 
            this.btnSaveOpenClose.Location = new System.Drawing.Point(105, 380);
            this.btnSaveOpenClose.Name = "btnSaveOpenClose";
            this.btnSaveOpenClose.Size = new System.Drawing.Size(75, 23);
            this.btnSaveOpenClose.TabIndex = 1;
            this.btnSaveOpenClose.Text = "S.&O.C.";
            this.btnSaveOpenClose.UseVisualStyleBackColor = true;
            this.btnSaveOpenClose.Click += new System.EventHandler(this.btnSaveOpenClose_Click);
            // 
            // lblSaveOpenClose
            // 
            this.lblSaveOpenClose.AutoSize = true;
            this.lblSaveOpenClose.Location = new System.Drawing.Point(21, 415);
            this.lblSaveOpenClose.Name = "lblSaveOpenClose";
            this.lblSaveOpenClose.Size = new System.Drawing.Size(350, 13);
            this.lblSaveOpenClose.TabIndex = 4;
            this.lblSaveOpenClose.Text = "S.O.C. = Save as HTML file, Open HTML in Browser and Close this form.";
            // 
            // FrmHtmlReportViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 456);
            this.Controls.Add(this.lblSaveOpenClose);
            this.Controls.Add(this.btnSaveOpenClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.pnlHtml);
            this.Controls.Add(this.ssStatus);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmHtmlReportViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmHtmlReportViewer_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmHtmlReportViewer_KeyDown);
            this.Resize += new System.EventHandler(this.FrmHtmlReportViewer_Resize);
            this.pnlHtml.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Panel pnlHtml;
        private System.Windows.Forms.WebBrowser webHtml;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnSaveOpenClose;
        private System.Windows.Forms.Label lblSaveOpenClose;
    }
}