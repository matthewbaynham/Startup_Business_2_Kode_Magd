namespace KodeMagd.WorkbookAnalysis
{
    partial class FrmCodeInColour_Options
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmCodeInColour_Options));
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnClose = new System.Windows.Forms.Button();
            this.dlgColour = new System.Windows.Forms.ColorDialog();
            this.btnDeclaringVariable = new System.Windows.Forms.Button();
            this.btnAssigningValues = new System.Windows.Forms.Button();
            this.btnIfStatements = new System.Windows.Forms.Button();
            this.btnLoops = new System.Windows.Forms.Button();
            this.btnFunctions = new System.Windows.Forms.Button();
            this.lblDeclaringVariable = new System.Windows.Forms.Label();
            this.lblAssigningValues = new System.Windows.Forms.Label();
            this.lblIfStatements = new System.Windows.Forms.Label();
            this.lblLoops = new System.Windows.Forms.Label();
            this.lblFunctions = new System.Windows.Forms.Label();
            this.lblComments = new System.Windows.Forms.Label();
            this.btnComments = new System.Windows.Forms.Button();
            this.lblErrorCode = new System.Windows.Forms.Label();
            this.btnErrorCode = new System.Windows.Forms.Button();
            this.lblWith = new System.Windows.Forms.Label();
            this.btnWith = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnResetDefault = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 313);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(474, 22);
            this.ssStatus.TabIndex = 0;
            this.ssStatus.Text = "statusStrip1";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(380, 272);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnDeclaringVariable
            // 
            this.btnDeclaringVariable.Location = new System.Drawing.Point(128, 33);
            this.btnDeclaringVariable.Name = "btnDeclaringVariable";
            this.btnDeclaringVariable.Size = new System.Drawing.Size(75, 23);
            this.btnDeclaringVariable.TabIndex = 3;
            this.btnDeclaringVariable.UseVisualStyleBackColor = true;
            this.btnDeclaringVariable.Click += new System.EventHandler(this.btnDeclaringVariable_Click);
            // 
            // btnAssigningValues
            // 
            this.btnAssigningValues.Location = new System.Drawing.Point(128, 62);
            this.btnAssigningValues.Name = "btnAssigningValues";
            this.btnAssigningValues.Size = new System.Drawing.Size(75, 23);
            this.btnAssigningValues.TabIndex = 4;
            this.btnAssigningValues.UseVisualStyleBackColor = true;
            this.btnAssigningValues.Click += new System.EventHandler(this.btnAssigningValues_Click);
            // 
            // btnIfStatements
            // 
            this.btnIfStatements.Location = new System.Drawing.Point(128, 91);
            this.btnIfStatements.Name = "btnIfStatements";
            this.btnIfStatements.Size = new System.Drawing.Size(75, 23);
            this.btnIfStatements.TabIndex = 5;
            this.btnIfStatements.UseVisualStyleBackColor = true;
            this.btnIfStatements.Click += new System.EventHandler(this.btnIfStatements_Click);
            // 
            // btnLoops
            // 
            this.btnLoops.Location = new System.Drawing.Point(128, 120);
            this.btnLoops.Name = "btnLoops";
            this.btnLoops.Size = new System.Drawing.Size(75, 23);
            this.btnLoops.TabIndex = 6;
            this.btnLoops.UseVisualStyleBackColor = true;
            this.btnLoops.Click += new System.EventHandler(this.btnLoops_Click);
            // 
            // btnFunctions
            // 
            this.btnFunctions.Location = new System.Drawing.Point(128, 149);
            this.btnFunctions.Name = "btnFunctions";
            this.btnFunctions.Size = new System.Drawing.Size(75, 23);
            this.btnFunctions.TabIndex = 7;
            this.btnFunctions.UseVisualStyleBackColor = true;
            this.btnFunctions.Click += new System.EventHandler(this.btnFunctions_Click);
            // 
            // lblDeclaringVariable
            // 
            this.lblDeclaringVariable.AutoSize = true;
            this.lblDeclaringVariable.Location = new System.Drawing.Point(10, 33);
            this.lblDeclaringVariable.Name = "lblDeclaringVariable";
            this.lblDeclaringVariable.Size = new System.Drawing.Size(93, 13);
            this.lblDeclaringVariable.TabIndex = 8;
            this.lblDeclaringVariable.Text = "Declaring Variable";
            // 
            // lblAssigningValues
            // 
            this.lblAssigningValues.AutoSize = true;
            this.lblAssigningValues.Location = new System.Drawing.Point(10, 62);
            this.lblAssigningValues.Name = "lblAssigningValues";
            this.lblAssigningValues.Size = new System.Drawing.Size(87, 13);
            this.lblAssigningValues.TabIndex = 9;
            this.lblAssigningValues.Text = "Assigning Values";
            // 
            // lblIfStatements
            // 
            this.lblIfStatements.AutoSize = true;
            this.lblIfStatements.Location = new System.Drawing.Point(10, 91);
            this.lblIfStatements.Name = "lblIfStatements";
            this.lblIfStatements.Size = new System.Drawing.Size(69, 13);
            this.lblIfStatements.TabIndex = 10;
            this.lblIfStatements.Text = "If Statements";
            // 
            // lblLoops
            // 
            this.lblLoops.AutoSize = true;
            this.lblLoops.Location = new System.Drawing.Point(10, 120);
            this.lblLoops.Name = "lblLoops";
            this.lblLoops.Size = new System.Drawing.Size(36, 13);
            this.lblLoops.TabIndex = 11;
            this.lblLoops.Text = "Loops";
            // 
            // lblFunctions
            // 
            this.lblFunctions.AutoSize = true;
            this.lblFunctions.Location = new System.Drawing.Point(10, 149);
            this.lblFunctions.Name = "lblFunctions";
            this.lblFunctions.Size = new System.Drawing.Size(53, 13);
            this.lblFunctions.TabIndex = 12;
            this.lblFunctions.Text = "Functions";
            // 
            // lblComments
            // 
            this.lblComments.AutoSize = true;
            this.lblComments.Location = new System.Drawing.Point(10, 181);
            this.lblComments.Name = "lblComments";
            this.lblComments.Size = new System.Drawing.Size(56, 13);
            this.lblComments.TabIndex = 13;
            this.lblComments.Text = "Comments";
            // 
            // btnComments
            // 
            this.btnComments.Location = new System.Drawing.Point(128, 181);
            this.btnComments.Name = "btnComments";
            this.btnComments.Size = new System.Drawing.Size(75, 23);
            this.btnComments.TabIndex = 14;
            this.btnComments.UseVisualStyleBackColor = true;
            this.btnComments.Click += new System.EventHandler(this.btnComments_Click);
            // 
            // lblErrorCode
            // 
            this.lblErrorCode.AutoSize = true;
            this.lblErrorCode.Location = new System.Drawing.Point(10, 210);
            this.lblErrorCode.Name = "lblErrorCode";
            this.lblErrorCode.Size = new System.Drawing.Size(57, 13);
            this.lblErrorCode.TabIndex = 15;
            this.lblErrorCode.Text = "Error Code";
            // 
            // btnErrorCode
            // 
            this.btnErrorCode.Location = new System.Drawing.Point(128, 210);
            this.btnErrorCode.Name = "btnErrorCode";
            this.btnErrorCode.Size = new System.Drawing.Size(75, 23);
            this.btnErrorCode.TabIndex = 16;
            this.btnErrorCode.UseVisualStyleBackColor = true;
            this.btnErrorCode.Click += new System.EventHandler(this.btnErrorCode_Click);
            // 
            // lblWith
            // 
            this.lblWith.AutoSize = true;
            this.lblWith.Location = new System.Drawing.Point(10, 242);
            this.lblWith.Name = "lblWith";
            this.lblWith.Size = new System.Drawing.Size(29, 13);
            this.lblWith.TabIndex = 17;
            this.lblWith.Text = "With";
            // 
            // btnWith
            // 
            this.btnWith.Location = new System.Drawing.Point(128, 242);
            this.btnWith.Name = "btnWith";
            this.btnWith.Size = new System.Drawing.Size(75, 23);
            this.btnWith.TabIndex = 18;
            this.btnWith.UseVisualStyleBackColor = true;
            this.btnWith.Click += new System.EventHandler(this.btnWith_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(292, 272);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 19;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(11, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(27, 13);
            this.lblTitle.TabIndex = 20;
            this.lblTitle.Text = "Title";
            // 
            // btnResetDefault
            // 
            this.btnResetDefault.Location = new System.Drawing.Point(364, 33);
            this.btnResetDefault.Name = "btnResetDefault";
            this.btnResetDefault.Size = new System.Drawing.Size(91, 23);
            this.btnResetDefault.TabIndex = 21;
            this.btnResetDefault.Text = "Reset Defaults";
            this.btnResetDefault.UseVisualStyleBackColor = true;
            this.btnResetDefault.Click += new System.EventHandler(this.btnResetDefault_Click);
            // 
            // FrmCodeInColour_Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 333);
            this.Controls.Add(this.btnResetDefault);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnWith);
            this.Controls.Add(this.lblWith);
            this.Controls.Add(this.btnErrorCode);
            this.Controls.Add(this.lblErrorCode);
            this.Controls.Add(this.btnComments);
            this.Controls.Add(this.lblComments);
            this.Controls.Add(this.lblFunctions);
            this.Controls.Add(this.lblLoops);
            this.Controls.Add(this.lblIfStatements);
            this.Controls.Add(this.lblAssigningValues);
            this.Controls.Add(this.lblDeclaringVariable);
            this.Controls.Add(this.btnFunctions);
            this.Controls.Add(this.btnLoops);
            this.Controls.Add(this.btnIfStatements);
            this.Controls.Add(this.btnAssigningValues);
            this.Controls.Add(this.btnDeclaringVariable);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.ssStatus);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmCodeInColour_Options";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmCodeInColour_Options_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ColorDialog dlgColour;
        private System.Windows.Forms.Button btnDeclaringVariable;
        private System.Windows.Forms.Button btnAssigningValues;
        private System.Windows.Forms.Button btnIfStatements;
        private System.Windows.Forms.Button btnLoops;
        private System.Windows.Forms.Button btnFunctions;
        private System.Windows.Forms.Label lblDeclaringVariable;
        private System.Windows.Forms.Label lblAssigningValues;
        private System.Windows.Forms.Label lblIfStatements;
        private System.Windows.Forms.Label lblLoops;
        private System.Windows.Forms.Label lblFunctions;
        private System.Windows.Forms.Label lblComments;
        private System.Windows.Forms.Button btnComments;
        private System.Windows.Forms.Label lblErrorCode;
        private System.Windows.Forms.Button btnErrorCode;
        private System.Windows.Forms.Label lblWith;
        private System.Windows.Forms.Button btnWith;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnResetDefault;
    }
}