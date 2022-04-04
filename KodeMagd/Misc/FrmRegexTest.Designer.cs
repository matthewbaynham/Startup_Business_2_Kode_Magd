namespace KodeMagd.Misc
{
    partial class FrmRegexTest
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
            this.lblPattern = new System.Windows.Forms.Label();
            this.lblText = new System.Windows.Forms.Label();
            this.txtPattern = new System.Windows.Forms.TextBox();
            this.txtText = new System.Windows.Forms.TextBox();
            this.lblResult = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblPattern
            // 
            this.lblPattern.AutoSize = true;
            this.lblPattern.Location = new System.Drawing.Point(12, 63);
            this.lblPattern.Name = "lblPattern";
            this.lblPattern.Size = new System.Drawing.Size(41, 13);
            this.lblPattern.TabIndex = 0;
            this.lblPattern.Text = "Pattern";
            // 
            // lblText
            // 
            this.lblText.AutoSize = true;
            this.lblText.Location = new System.Drawing.Point(12, 186);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(28, 13);
            this.lblText.TabIndex = 1;
            this.lblText.Text = "Text";
            // 
            // txtPattern
            // 
            this.txtPattern.Location = new System.Drawing.Point(12, 90);
            this.txtPattern.Name = "txtPattern";
            this.txtPattern.Size = new System.Drawing.Size(481, 20);
            this.txtPattern.TabIndex = 2;
            this.txtPattern.TextChanged += new System.EventHandler(this.txtPattern_TextChanged);
            // 
            // txtText
            // 
            this.txtText.Location = new System.Drawing.Point(12, 217);
            this.txtText.Name = "txtText";
            this.txtText.Size = new System.Drawing.Size(481, 20);
            this.txtText.TabIndex = 3;
            this.txtText.TextChanged += new System.EventHandler(this.txtText_TextChanged);
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(82, 277);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(0, 13);
            this.lblResult.TabIndex = 4;
            // 
            // FrmRegexTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(505, 425);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.txtText);
            this.Controls.Add(this.txtPattern);
            this.Controls.Add(this.lblText);
            this.Controls.Add(this.lblPattern);
            this.Name = "FrmRegexTest";
            this.Text = "FrmRegexTest";
            this.Load += new System.EventHandler(this.FrmRegexTest_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblPattern;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.TextBox txtPattern;
        private System.Windows.Forms.TextBox txtText;
        private System.Windows.Forms.Label lblResult;
    }
}