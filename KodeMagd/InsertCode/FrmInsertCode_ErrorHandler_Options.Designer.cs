namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_ErrorHandler_Options
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_ErrorHandler_Options));
            this.btnClose = new System.Windows.Forms.Button();
            this.grpExistingHandlers = new System.Windows.Forms.GroupBox();
            this.optDoNothing = new System.Windows.Forms.RadioButton();
            this.optReplaceAll = new System.Windows.Forms.RadioButton();
            this.optOneInMiddle = new System.Windows.Forms.RadioButton();
            this.optOneAtTop = new System.Windows.Forms.RadioButton();
            this.optIgnore = new System.Windows.Forms.RadioButton();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnOk = new System.Windows.Forms.Button();
            this.grpExistingHandlers.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(385, 274);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // grpExistingHandlers
            // 
            this.grpExistingHandlers.Controls.Add(this.optDoNothing);
            this.grpExistingHandlers.Controls.Add(this.optReplaceAll);
            this.grpExistingHandlers.Controls.Add(this.optOneInMiddle);
            this.grpExistingHandlers.Controls.Add(this.optOneAtTop);
            this.grpExistingHandlers.Controls.Add(this.optIgnore);
            this.grpExistingHandlers.Location = new System.Drawing.Point(12, 30);
            this.grpExistingHandlers.Name = "grpExistingHandlers";
            this.grpExistingHandlers.Size = new System.Drawing.Size(450, 147);
            this.grpExistingHandlers.TabIndex = 0;
            this.grpExistingHandlers.TabStop = false;
            this.grpExistingHandlers.Text = "Dealing with Existing Error Handlers";
            // 
            // optDoNothing
            // 
            this.optDoNothing.AutoSize = true;
            this.optDoNothing.Location = new System.Drawing.Point(19, 114);
            this.optDoNothing.Name = "optDoNothing";
            this.optDoNothing.Size = new System.Drawing.Size(296, 17);
            this.optDoNothing.TabIndex = 4;
            this.optDoNothing.TabStop = true;
            this.optDoNothing.Text = "DO NOTHING, in the routines with existing Error Handlers";
            this.optDoNothing.UseVisualStyleBackColor = true;
            // 
            // optReplaceAll
            // 
            this.optReplaceAll.AutoSize = true;
            this.optReplaceAll.Location = new System.Drawing.Point(19, 91);
            this.optReplaceAll.Name = "optReplaceAll";
            this.optReplaceAll.Size = new System.Drawing.Size(313, 17);
            this.optReplaceAll.TabIndex = 3;
            this.optReplaceAll.TabStop = true;
            this.optReplaceAll.Text = "If MANY error handlers are found in the code REPLACE ALL.";
            this.optReplaceAll.UseVisualStyleBackColor = true;
            // 
            // optOneInMiddle
            // 
            this.optOneInMiddle.AutoSize = true;
            this.optOneInMiddle.Location = new System.Drawing.Point(19, 68);
            this.optOneInMiddle.Name = "optOneInMiddle";
            this.optOneInMiddle.Size = new System.Drawing.Size(362, 17);
            this.optOneInMiddle.TabIndex = 2;
            this.optOneInMiddle.TabStop = true;
            this.optOneInMiddle.Text = "If only ONE error handler is found ANYWHERE in some code replace it.";
            this.optOneInMiddle.UseVisualStyleBackColor = true;
            // 
            // optOneAtTop
            // 
            this.optOneAtTop.AutoSize = true;
            this.optOneAtTop.Location = new System.Drawing.Point(19, 45);
            this.optOneAtTop.Name = "optOneAtTop";
            this.optOneAtTop.Size = new System.Drawing.Size(390, 17);
            this.optOneAtTop.TabIndex = 1;
            this.optOneAtTop.TabStop = true;
            this.optOneAtTop.Text = "If only ONE error handler is found at the BEGINNING of some code replace it.";
            this.optOneAtTop.UseVisualStyleBackColor = true;
            // 
            // optIgnore
            // 
            this.optIgnore.AutoSize = true;
            this.optIgnore.Location = new System.Drawing.Point(19, 22);
            this.optIgnore.Name = "optIgnore";
            this.optIgnore.Size = new System.Drawing.Size(299, 17);
            this.optIgnore.TabIndex = 0;
            this.optIgnore.TabStop = true;
            this.optIgnore.Text = "Ignore Existing Error Handlers and just add new regardless";
            this.optIgnore.UseVisualStyleBackColor = true;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 314);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(472, 22);
            this.ssStatus.TabIndex = 3;
            this.ssStatus.Text = "ssStatus";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(295, 274);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "&OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // FrmInsertCode_ErrorHandler_Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 336);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.grpExistingHandlers);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmInsertCode_ErrorHandler_Options";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_ErrorHandler_Options_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_ErrorHandler_Options_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_ErrorHandler_Options_Resize);
            this.grpExistingHandlers.ResumeLayout(false);
            this.grpExistingHandlers.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox grpExistingHandlers;
        private System.Windows.Forms.RadioButton optReplaceAll;
        private System.Windows.Forms.RadioButton optOneInMiddle;
        private System.Windows.Forms.RadioButton optOneAtTop;
        private System.Windows.Forms.RadioButton optIgnore;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.RadioButton optDoNothing;
        private System.Windows.Forms.Button btnOk;

    }
}