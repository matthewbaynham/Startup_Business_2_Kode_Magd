namespace KodeMagd
{
    partial class FrmOptions
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmOptions));
            this.label3 = new System.Windows.Forms.Label();
            this.chkFocusActivePane = new System.Windows.Forms.CheckBox();
            this.chkIndentFirstLevel = new System.Windows.Forms.CheckBox();
            this.lblSplitLinesDescription = new System.Windows.Forms.Label();
            this.chkSplitLines = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtIndentSize = new System.Windows.Forms.NumericUpDown();
            this.chkIncludeErrorHandler = new System.Windows.Forms.CheckBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.chkUseWith = new System.Windows.Forms.CheckBox();
            this.cmbFormatLineCutMethodology = new System.Windows.Forms.ComboBox();
            this.txtFormatCutTextChar = new System.Windows.Forms.NumericUpDown();
            this.lblTxtFormattingLineBreakMethodology = new System.Windows.Forms.Label();
            this.lblTextFormattingLineBreakCutAfterNumberOfChar = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.btnRegistryFolders = new System.Windows.Forms.Button();
            this.btnDefault = new System.Windows.Forms.Button();
            this.cmbFormatOfDimStatement = new System.Windows.Forms.ComboBox();
            this.lblFormatOfDimStatement = new System.Windows.Forms.Label();
            this.chkTips = new System.Windows.Forms.CheckBox();
            this.btnCodeInColourDefaults = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.txtIndentSize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFormatCutTextChar)).BeginInit();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(219, 89);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label3.Size = new System.Drawing.Size(400, 39);
            this.label3.TabIndex = 5;
            this.label3.Text = "Fixes a \"feature\" of the VBE editor window: When the VBE Editor is first opened a" +
    "nd the Code Panel appears to have focus but if you look at the Project Explorer " +
    "window there’s an inconsistency.";
            // 
            // chkFocusActivePane
            // 
            this.chkFocusActivePane.AutoSize = true;
            this.chkFocusActivePane.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkFocusActivePane.Location = new System.Drawing.Point(8, 89);
            this.chkFocusActivePane.Name = "chkFocusActivePane";
            this.chkFocusActivePane.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkFocusActivePane.Size = new System.Drawing.Size(150, 17);
            this.chkFocusActivePane.TabIndex = 4;
            this.chkFocusActivePane.Text = "Set Focus on Active Pane";
            this.chkFocusActivePane.UseVisualStyleBackColor = true;
            // 
            // chkIndentFirstLevel
            // 
            this.chkIndentFirstLevel.AutoSize = true;
            this.chkIndentFirstLevel.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkIndentFirstLevel.Location = new System.Drawing.Point(8, 37);
            this.chkIndentFirstLevel.Name = "chkIndentFirstLevel";
            this.chkIndentFirstLevel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkIndentFirstLevel.Size = new System.Drawing.Size(107, 17);
            this.chkIndentFirstLevel.TabIndex = 2;
            this.chkIndentFirstLevel.Text = "Indent First Level";
            this.chkIndentFirstLevel.UseVisualStyleBackColor = true;
            // 
            // lblSplitLinesDescription
            // 
            this.lblSplitLinesDescription.Location = new System.Drawing.Point(219, 141);
            this.lblSplitLinesDescription.Name = "lblSplitLinesDescription";
            this.lblSplitLinesDescription.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lblSplitLinesDescription.Size = new System.Drawing.Size(400, 39);
            this.lblSplitLinesDescription.TabIndex = 7;
            this.lblSplitLinesDescription.Text = "When a colon enables multiple lines to be written on the same line, split it on t" +
    "o multiple lines.";
            this.lblSplitLinesDescription.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // chkSplitLines
            // 
            this.chkSplitLines.AutoSize = true;
            this.chkSplitLines.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkSplitLines.Location = new System.Drawing.Point(8, 141);
            this.chkSplitLines.Name = "chkSplitLines";
            this.chkSplitLines.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkSplitLines.Size = new System.Drawing.Size(74, 17);
            this.chkSplitLines.TabIndex = 6;
            this.chkSplitLines.Text = "Split Lines";
            this.chkSplitLines.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(67, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Indent Size";
            // 
            // txtIndentSize
            // 
            this.txtIndentSize.Location = new System.Drawing.Point(8, 11);
            this.txtIndentSize.Name = "txtIndentSize";
            this.txtIndentSize.Size = new System.Drawing.Size(53, 20);
            this.txtIndentSize.TabIndex = 0;
            // 
            // chkIncludeErrorHandler
            // 
            this.chkIncludeErrorHandler.AutoSize = true;
            this.chkIncludeErrorHandler.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkIncludeErrorHandler.Location = new System.Drawing.Point(8, 63);
            this.chkIncludeErrorHandler.Name = "chkIncludeErrorHandler";
            this.chkIncludeErrorHandler.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.chkIncludeErrorHandler.Size = new System.Drawing.Size(126, 17);
            this.chkIncludeErrorHandler.TabIndex = 3;
            this.chkIncludeErrorHandler.Text = "Include Error Handler";
            this.chkIncludeErrorHandler.UseVisualStyleBackColor = true;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 385);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 20;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(450, 385);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(85, 22);
            this.btnOK.TabIndex = 19;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // chkUseWith
            // 
            this.chkUseWith.AutoSize = true;
            this.chkUseWith.Location = new System.Drawing.Point(8, 167);
            this.chkUseWith.Name = "chkUseWith";
            this.chkUseWith.Size = new System.Drawing.Size(70, 17);
            this.chkUseWith.TabIndex = 8;
            this.chkUseWith.Text = "Use With";
            this.chkUseWith.UseVisualStyleBackColor = true;
            // 
            // cmbFormatLineCutMethodology
            // 
            this.cmbFormatLineCutMethodology.FormattingEnabled = true;
            this.cmbFormatLineCutMethodology.Location = new System.Drawing.Point(8, 250);
            this.cmbFormatLineCutMethodology.Name = "cmbFormatLineCutMethodology";
            this.cmbFormatLineCutMethodology.Size = new System.Drawing.Size(200, 21);
            this.cmbFormatLineCutMethodology.TabIndex = 10;
            this.cmbFormatLineCutMethodology.SelectedIndexChanged += new System.EventHandler(this.cmbFormatLineCutMethodology_SelectedIndexChanged);
            // 
            // txtFormatCutTextChar
            // 
            this.txtFormatCutTextChar.Location = new System.Drawing.Point(219, 250);
            this.txtFormatCutTextChar.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.txtFormatCutTextChar.Minimum = new decimal(new int[] {
            80,
            0,
            0,
            0});
            this.txtFormatCutTextChar.Name = "txtFormatCutTextChar";
            this.txtFormatCutTextChar.Size = new System.Drawing.Size(53, 20);
            this.txtFormatCutTextChar.TabIndex = 12;
            this.txtFormatCutTextChar.Value = new decimal(new int[] {
            80,
            0,
            0,
            0});
            // 
            // lblTxtFormattingLineBreakMethodology
            // 
            this.lblTxtFormattingLineBreakMethodology.AutoSize = true;
            this.lblTxtFormattingLineBreakMethodology.Location = new System.Drawing.Point(8, 230);
            this.lblTxtFormattingLineBreakMethodology.Name = "lblTxtFormattingLineBreakMethodology";
            this.lblTxtFormattingLineBreakMethodology.Size = new System.Drawing.Size(199, 13);
            this.lblTxtFormattingLineBreakMethodology.TabIndex = 11;
            this.lblTxtFormattingLineBreakMethodology.Text = "Text Formatting: Line break methodology";
            // 
            // lblTextFormattingLineBreakCutAfterNumberOfChar
            // 
            this.lblTextFormattingLineBreakCutAfterNumberOfChar.Location = new System.Drawing.Point(219, 216);
            this.lblTextFormattingLineBreakCutAfterNumberOfChar.Name = "lblTextFormattingLineBreakCutAfterNumberOfChar";
            this.lblTextFormattingLineBreakCutAfterNumberOfChar.Size = new System.Drawing.Size(170, 30);
            this.lblTextFormattingLineBreakCutAfterNumberOfChar.TabIndex = 13;
            this.lblTextFormattingLineBreakCutAfterNumberOfChar.Text = "Text Formatting: Line break, cut after number of char";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 431);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 21;
            this.ssStatus.Text = "status";
            // 
            // btnRegistryFolders
            // 
            this.btnRegistryFolders.Location = new System.Drawing.Point(340, 11);
            this.btnRegistryFolders.Name = "btnRegistryFolders";
            this.btnRegistryFolders.Size = new System.Drawing.Size(133, 23);
            this.btnRegistryFolders.TabIndex = 18;
            this.btnRegistryFolders.Text = "&Registry Folders";
            this.btnRegistryFolders.UseVisualStyleBackColor = true;
            this.btnRegistryFolders.Click += new System.EventHandler(this.btnRegistryFolders_Click);
            // 
            // btnDefault
            // 
            this.btnDefault.Location = new System.Drawing.Point(487, 11);
            this.btnDefault.Name = "btnDefault";
            this.btnDefault.Size = new System.Drawing.Size(133, 23);
            this.btnDefault.TabIndex = 17;
            this.btnDefault.Text = "Reset to &Default";
            this.btnDefault.UseVisualStyleBackColor = true;
            this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
            // 
            // cmbFormatOfDimStatement
            // 
            this.cmbFormatOfDimStatement.FormattingEnabled = true;
            this.cmbFormatOfDimStatement.Location = new System.Drawing.Point(8, 300);
            this.cmbFormatOfDimStatement.Name = "cmbFormatOfDimStatement";
            this.cmbFormatOfDimStatement.Size = new System.Drawing.Size(200, 21);
            this.cmbFormatOfDimStatement.TabIndex = 14;
            // 
            // lblFormatOfDimStatement
            // 
            this.lblFormatOfDimStatement.AutoSize = true;
            this.lblFormatOfDimStatement.Location = new System.Drawing.Point(8, 280);
            this.lblFormatOfDimStatement.Name = "lblFormatOfDimStatement";
            this.lblFormatOfDimStatement.Size = new System.Drawing.Size(123, 13);
            this.lblFormatOfDimStatement.TabIndex = 15;
            this.lblFormatOfDimStatement.Text = "Format of Dim Statement";
            // 
            // chkTips
            // 
            this.chkTips.AutoSize = true;
            this.chkTips.Location = new System.Drawing.Point(8, 193);
            this.chkTips.Name = "chkTips";
            this.chkTips.Size = new System.Drawing.Size(167, 17);
            this.chkTips.TabIndex = 16;
            this.chkTips.Text = "Tips (extra comments in code)";
            this.chkTips.UseVisualStyleBackColor = true;
            // 
            // btnCodeInColourDefaults
            // 
            this.btnCodeInColourDefaults.Location = new System.Drawing.Point(193, 11);
            this.btnCodeInColourDefaults.Name = "btnCodeInColourDefaults";
            this.btnCodeInColourDefaults.Size = new System.Drawing.Size(133, 23);
            this.btnCodeInColourDefaults.TabIndex = 22;
            this.btnCodeInColourDefaults.Text = "Code In Colour Defaults";
            this.btnCodeInColourDefaults.UseVisualStyleBackColor = true;
            this.btnCodeInColourDefaults.Click += new System.EventHandler(this.btnCodeInColourDefaults_Click);
            // 
            // FrmOptions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(632, 453);
            this.Controls.Add(this.btnCodeInColourDefaults);
            this.Controls.Add(this.chkTips);
            this.Controls.Add(this.lblFormatOfDimStatement);
            this.Controls.Add(this.cmbFormatOfDimStatement);
            this.Controls.Add(this.btnDefault);
            this.Controls.Add(this.btnRegistryFolders);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblTextFormattingLineBreakCutAfterNumberOfChar);
            this.Controls.Add(this.lblTxtFormattingLineBreakMethodology);
            this.Controls.Add(this.txtFormatCutTextChar);
            this.Controls.Add(this.cmbFormatLineCutMethodology);
            this.Controls.Add(this.chkUseWith);
            this.Controls.Add(this.chkIncludeErrorHandler);
            this.Controls.Add(this.lblSplitLinesDescription);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.chkSplitLines);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkFocusActivePane);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.chkIndentFirstLevel);
            this.Controls.Add(this.txtIndentSize);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmOptions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmOptions_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmOptions_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.txtIndentSize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtFormatCutTextChar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblSplitLinesDescription;
        private System.Windows.Forms.CheckBox chkSplitLines;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown txtIndentSize;
        private System.Windows.Forms.CheckBox chkIndentFirstLevel;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.CheckBox chkFocusActivePane;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkIncludeErrorHandler;
        private System.Windows.Forms.CheckBox chkUseWith;
        private System.Windows.Forms.ComboBox cmbFormatLineCutMethodology;
        private System.Windows.Forms.NumericUpDown txtFormatCutTextChar;
        private System.Windows.Forms.Label lblTxtFormattingLineBreakMethodology;
        private System.Windows.Forms.Label lblTextFormattingLineBreakCutAfterNumberOfChar;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.Button btnRegistryFolders;
        private System.Windows.Forms.Button btnDefault;
        private System.Windows.Forms.ComboBox cmbFormatOfDimStatement;
        private System.Windows.Forms.Label lblFormatOfDimStatement;
        private System.Windows.Forms.CheckBox chkTips;
        private System.Windows.Forms.Button btnCodeInColourDefaults;
    }
}