namespace KodeMagd.Misc
{
    partial class FrmDoubleInputBox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmDoubleInputBox));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtAnswer1 = new System.Windows.Forms.TextBox();
            this.lblQuestion = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.lblLabel1 = new System.Windows.Forms.Label();
            this.lblLabel2 = new System.Windows.Forms.Label();
            this.txtAnswer2 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(407, 269);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(55, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(337, 269);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(55, 23);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtAnswer1
            // 
            this.txtAnswer1.Location = new System.Drawing.Point(16, 120);
            this.txtAnswer1.Name = "txtAnswer1";
            this.txtAnswer1.Size = new System.Drawing.Size(448, 20);
            this.txtAnswer1.TabIndex = 2;
            this.txtAnswer1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAnswer1_KeyPress);
            // 
            // lblQuestion
            // 
            this.lblQuestion.Location = new System.Drawing.Point(13, 9);
            this.lblQuestion.Name = "lblQuestion";
            this.lblQuestion.Size = new System.Drawing.Size(446, 58);
            this.lblQuestion.TabIndex = 0;
            this.lblQuestion.Text = "lblQuestion";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 314);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(474, 22);
            this.ssStatus.TabIndex = 7;
            this.ssStatus.Text = "status";
            // 
            // lblLabel1
            // 
            this.lblLabel1.Location = new System.Drawing.Point(16, 86);
            this.lblLabel1.Name = "lblLabel1";
            this.lblLabel1.Size = new System.Drawing.Size(446, 23);
            this.lblLabel1.TabIndex = 1;
            this.lblLabel1.Text = "label1";
            // 
            // lblLabel2
            // 
            this.lblLabel2.Location = new System.Drawing.Point(16, 162);
            this.lblLabel2.Name = "lblLabel2";
            this.lblLabel2.Size = new System.Drawing.Size(448, 19);
            this.lblLabel2.TabIndex = 3;
            this.lblLabel2.Text = "label2";
            // 
            // txtAnswer2
            // 
            this.txtAnswer2.Location = new System.Drawing.Point(16, 192);
            this.txtAnswer2.Name = "txtAnswer2";
            this.txtAnswer2.Size = new System.Drawing.Size(448, 20);
            this.txtAnswer2.TabIndex = 4;
            this.txtAnswer2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAnswer2_KeyPress);
            // 
            // FrmDoubleInputBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 336);
            this.Controls.Add(this.txtAnswer2);
            this.Controls.Add(this.lblLabel2);
            this.Controls.Add(this.lblLabel1);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblQuestion);
            this.Controls.Add(this.txtAnswer1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmDoubleInputBox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmDoubleInputBox_KeyDown);
            this.Resize += new System.EventHandler(this.FrmDoubleInputBox_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtAnswer1;
        private System.Windows.Forms.Label lblQuestion;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Label lblLabel1;
        private System.Windows.Forms.Label lblLabel2;
        private System.Windows.Forms.TextBox txtAnswer2;
    }
}