namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_SparkLines
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_SparkLines));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnSource = new System.Windows.Forms.Button();
            this.txtSourceRange = new System.Windows.Forms.TextBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.chkSourceNamedRange = new System.Windows.Forms.CheckBox();
            this.cmbSourceNamedRange = new System.Windows.Forms.ComboBox();
            this.lblDestination = new System.Windows.Forms.Label();
            this.chkDestinationNamedRange = new System.Windows.Forms.CheckBox();
            this.btnDestination = new System.Windows.Forms.Button();
            this.txtDestinationRange = new System.Windows.Forms.TextBox();
            this.cmbDestinationNameRange = new System.Windows.Forms.ComboBox();
            this.grpDirection = new System.Windows.Forms.GroupBox();
            this.optLine = new System.Windows.Forms.RadioButton();
            this.optColumnStacked100 = new System.Windows.Forms.RadioButton();
            this.optColumn = new System.Windows.Forms.RadioButton();
            this.lblWarning = new System.Windows.Forms.Label();
            this.btnColour = new System.Windows.Forms.Button();
            this.lblSourceShtName = new System.Windows.Forms.Label();
            this.txtSourceShtName = new System.Windows.Forms.TextBox();
            this.lblSourceRange = new System.Windows.Forms.Label();
            this.shapeContainer1 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
            this.lineShape1 = new Microsoft.VisualBasic.PowerPacks.LineShape();
            this.lblDestinationShtName = new System.Windows.Forms.Label();
            this.txtDestinationShtName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkNewSheet = new System.Windows.Forms.CheckBox();
            this.cmbDestinationShtName = new System.Windows.Forms.ComboBox();
            this.chkDestinationCreateNamedRange = new System.Windows.Forms.CheckBox();
            this.lblSourceNamedRange = new System.Windows.Forms.Label();
            this.lblDestinationNamedRange = new System.Windows.Forms.Label();
            this.txtDestinationNameRange = new System.Windows.Forms.TextBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.grpDirection.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 396);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(464, 396);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 1;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // btnSource
            // 
            this.btnSource.Location = new System.Drawing.Point(257, 95);
            this.btnSource.Name = "btnSource";
            this.btnSource.Size = new System.Drawing.Size(75, 23);
            this.btnSource.TabIndex = 2;
            this.btnSource.Text = "Source";
            this.btnSource.UseVisualStyleBackColor = true;
            this.btnSource.Click += new System.EventHandler(this.btnSource_Click);
            // 
            // txtSourceRange
            // 
            this.txtSourceRange.Location = new System.Drawing.Point(20, 95);
            this.txtSourceRange.Name = "txtSourceRange";
            this.txtSourceRange.Size = new System.Drawing.Size(233, 20);
            this.txtSourceRange.TabIndex = 3;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(20, 9);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 4;
            this.lblSource.Text = "Source";
            // 
            // chkSourceNamedRange
            // 
            this.chkSourceNamedRange.AutoSize = true;
            this.chkSourceNamedRange.Location = new System.Drawing.Point(257, 146);
            this.chkSourceNamedRange.Name = "chkSourceNamedRange";
            this.chkSourceNamedRange.Size = new System.Drawing.Size(95, 17);
            this.chkSourceNamedRange.TabIndex = 5;
            this.chkSourceNamedRange.Text = "Named Range";
            this.chkSourceNamedRange.UseVisualStyleBackColor = true;
            this.chkSourceNamedRange.CheckedChanged += new System.EventHandler(this.chkNamedRange_CheckedChanged);
            // 
            // cmbSourceNamedRange
            // 
            this.cmbSourceNamedRange.FormattingEnabled = true;
            this.cmbSourceNamedRange.Location = new System.Drawing.Point(20, 144);
            this.cmbSourceNamedRange.Name = "cmbSourceNamedRange";
            this.cmbSourceNamedRange.Size = new System.Drawing.Size(233, 21);
            this.cmbSourceNamedRange.TabIndex = 6;
            this.cmbSourceNamedRange.SelectedIndexChanged += new System.EventHandler(this.cmbSource_SelectedIndexChanged);
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(20, 191);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(60, 13);
            this.lblDestination.TabIndex = 7;
            this.lblDestination.Text = "Destination";
            // 
            // chkDestinationNamedRange
            // 
            this.chkDestinationNamedRange.AutoSize = true;
            this.chkDestinationNamedRange.Location = new System.Drawing.Point(257, 337);
            this.chkDestinationNamedRange.Name = "chkDestinationNamedRange";
            this.chkDestinationNamedRange.Size = new System.Drawing.Size(95, 17);
            this.chkDestinationNamedRange.TabIndex = 8;
            this.chkDestinationNamedRange.Text = "Named Range";
            this.chkDestinationNamedRange.UseVisualStyleBackColor = true;
            this.chkDestinationNamedRange.CheckedChanged += new System.EventHandler(this.chkDestinationNamedRange_CheckedChanged);
            // 
            // btnDestination
            // 
            this.btnDestination.Location = new System.Drawing.Point(257, 283);
            this.btnDestination.Name = "btnDestination";
            this.btnDestination.Size = new System.Drawing.Size(75, 23);
            this.btnDestination.TabIndex = 9;
            this.btnDestination.Text = "Destination";
            this.btnDestination.UseVisualStyleBackColor = true;
            this.btnDestination.Click += new System.EventHandler(this.btnDestination_Click);
            // 
            // txtDestinationRange
            // 
            this.txtDestinationRange.Cursor = System.Windows.Forms.Cursors.No;
            this.txtDestinationRange.Location = new System.Drawing.Point(20, 286);
            this.txtDestinationRange.Name = "txtDestinationRange";
            this.txtDestinationRange.Size = new System.Drawing.Size(233, 20);
            this.txtDestinationRange.TabIndex = 10;
            // 
            // cmbDestinationNameRange
            // 
            this.cmbDestinationNameRange.FormattingEnabled = true;
            this.cmbDestinationNameRange.Location = new System.Drawing.Point(20, 335);
            this.cmbDestinationNameRange.Name = "cmbDestinationNameRange";
            this.cmbDestinationNameRange.Size = new System.Drawing.Size(233, 21);
            this.cmbDestinationNameRange.TabIndex = 11;
            this.cmbDestinationNameRange.SelectedIndexChanged += new System.EventHandler(this.cmbDestinationNameRange_SelectedIndexChanged);
            this.cmbDestinationNameRange.TextChanged += new System.EventHandler(this.cmbDestinationNameRange_TextChanged);
            // 
            // grpDirection
            // 
            this.grpDirection.Controls.Add(this.optLine);
            this.grpDirection.Controls.Add(this.optColumnStacked100);
            this.grpDirection.Controls.Add(this.optColumn);
            this.grpDirection.Location = new System.Drawing.Point(354, 9);
            this.grpDirection.Name = "grpDirection";
            this.grpDirection.Size = new System.Drawing.Size(265, 92);
            this.grpDirection.TabIndex = 12;
            this.grpDirection.TabStop = false;
            this.grpDirection.Text = "Direction";
            // 
            // optLine
            // 
            this.optLine.AutoSize = true;
            this.optLine.Location = new System.Drawing.Point(7, 65);
            this.optLine.Name = "optLine";
            this.optLine.Size = new System.Drawing.Size(45, 17);
            this.optLine.TabIndex = 2;
            this.optLine.TabStop = true;
            this.optLine.Text = "Line";
            this.optLine.UseVisualStyleBackColor = true;
            // 
            // optColumnStacked100
            // 
            this.optColumnStacked100.AutoSize = true;
            this.optColumnStacked100.Location = new System.Drawing.Point(7, 42);
            this.optColumnStacked100.Name = "optColumnStacked100";
            this.optColumnStacked100.Size = new System.Drawing.Size(124, 17);
            this.optColumnStacked100.TabIndex = 1;
            this.optColumnStacked100.TabStop = true;
            this.optColumnStacked100.Text = "Column Stacked 100";
            this.optColumnStacked100.UseVisualStyleBackColor = true;
            // 
            // optColumn
            // 
            this.optColumn.AutoSize = true;
            this.optColumn.Location = new System.Drawing.Point(7, 19);
            this.optColumn.Name = "optColumn";
            this.optColumn.Size = new System.Drawing.Size(60, 17);
            this.optColumn.TabIndex = 0;
            this.optColumn.TabStop = true;
            this.optColumn.Text = "Column";
            this.optColumn.UseVisualStyleBackColor = true;
            // 
            // lblWarning
            // 
            this.lblWarning.Location = new System.Drawing.Point(358, 122);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(260, 67);
            this.lblWarning.TabIndex = 13;
            this.lblWarning.Text = "Warning";
            // 
            // btnColour
            // 
            this.btnColour.Location = new System.Drawing.Point(20, 362);
            this.btnColour.Name = "btnColour";
            this.btnColour.Size = new System.Drawing.Size(75, 23);
            this.btnColour.TabIndex = 14;
            this.btnColour.Text = "Colour";
            this.btnColour.UseVisualStyleBackColor = true;
            this.btnColour.Click += new System.EventHandler(this.btnColour_Click);
            // 
            // lblSourceShtName
            // 
            this.lblSourceShtName.AutoSize = true;
            this.lblSourceShtName.Location = new System.Drawing.Point(20, 30);
            this.lblSourceShtName.Name = "lblSourceShtName";
            this.lblSourceShtName.Size = new System.Drawing.Size(66, 13);
            this.lblSourceShtName.TabIndex = 15;
            this.lblSourceShtName.Text = "Sheet Name";
            // 
            // txtSourceShtName
            // 
            this.txtSourceShtName.Enabled = false;
            this.txtSourceShtName.Location = new System.Drawing.Point(20, 46);
            this.txtSourceShtName.Name = "txtSourceShtName";
            this.txtSourceShtName.Size = new System.Drawing.Size(233, 20);
            this.txtSourceShtName.TabIndex = 16;
            // 
            // lblSourceRange
            // 
            this.lblSourceRange.AutoSize = true;
            this.lblSourceRange.Location = new System.Drawing.Point(21, 76);
            this.lblSourceRange.Name = "lblSourceRange";
            this.lblSourceRange.Size = new System.Drawing.Size(39, 13);
            this.lblSourceRange.TabIndex = 17;
            this.lblSourceRange.Text = "Range";
            // 
            // shapeContainer1
            // 
            this.shapeContainer1.Location = new System.Drawing.Point(0, 0);
            this.shapeContainer1.Margin = new System.Windows.Forms.Padding(0);
            this.shapeContainer1.Name = "shapeContainer1";
            this.shapeContainer1.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {
            this.lineShape1});
            this.shapeContainer1.Size = new System.Drawing.Size(634, 458);
            this.shapeContainer1.TabIndex = 18;
            this.shapeContainer1.TabStop = false;
            // 
            // lineShape1
            // 
            this.lineShape1.Name = "lineShape1";
            this.lineShape1.X1 = 13;
            this.lineShape1.X2 = 332;
            this.lineShape1.Y1 = 175;
            this.lineShape1.Y2 = 175;
            // 
            // lblDestinationShtName
            // 
            this.lblDestinationShtName.AutoSize = true;
            this.lblDestinationShtName.Location = new System.Drawing.Point(20, 210);
            this.lblDestinationShtName.Name = "lblDestinationShtName";
            this.lblDestinationShtName.Size = new System.Drawing.Size(66, 13);
            this.lblDestinationShtName.TabIndex = 19;
            this.lblDestinationShtName.Text = "Sheet Name";
            // 
            // txtDestinationShtName
            // 
            this.txtDestinationShtName.Location = new System.Drawing.Point(20, 229);
            this.txtDestinationShtName.Name = "txtDestinationShtName";
            this.txtDestinationShtName.Size = new System.Drawing.Size(233, 20);
            this.txtDestinationShtName.TabIndex = 20;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 270);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Range";
            // 
            // chkNewSheet
            // 
            this.chkNewSheet.AutoSize = true;
            this.chkNewSheet.Location = new System.Drawing.Point(259, 229);
            this.chkNewSheet.Name = "chkNewSheet";
            this.chkNewSheet.Size = new System.Drawing.Size(79, 17);
            this.chkNewSheet.TabIndex = 22;
            this.chkNewSheet.Text = "New Sheet";
            this.chkNewSheet.UseVisualStyleBackColor = true;
            this.chkNewSheet.CheckedChanged += new System.EventHandler(this.chkNewSheet_CheckedChanged);
            // 
            // cmbDestinationShtName
            // 
            this.cmbDestinationShtName.FormattingEnabled = true;
            this.cmbDestinationShtName.Location = new System.Drawing.Point(20, 229);
            this.cmbDestinationShtName.Name = "cmbDestinationShtName";
            this.cmbDestinationShtName.Size = new System.Drawing.Size(233, 21);
            this.cmbDestinationShtName.TabIndex = 23;
            // 
            // chkDestinationCreateNamedRange
            // 
            this.chkDestinationCreateNamedRange.AutoSize = true;
            this.chkDestinationCreateNamedRange.Location = new System.Drawing.Point(258, 361);
            this.chkDestinationCreateNamedRange.Name = "chkDestinationCreateNamedRange";
            this.chkDestinationCreateNamedRange.Size = new System.Drawing.Size(92, 17);
            this.chkDestinationCreateNamedRange.TabIndex = 24;
            this.chkDestinationCreateNamedRange.Text = "Create Range";
            this.chkDestinationCreateNamedRange.UseVisualStyleBackColor = true;
            this.chkDestinationCreateNamedRange.CheckedChanged += new System.EventHandler(this.chkDestinationCreateNamedRange_CheckedChanged);
            // 
            // lblSourceNamedRange
            // 
            this.lblSourceNamedRange.AutoSize = true;
            this.lblSourceNamedRange.Location = new System.Drawing.Point(21, 118);
            this.lblSourceNamedRange.Name = "lblSourceNamedRange";
            this.lblSourceNamedRange.Size = new System.Drawing.Size(76, 13);
            this.lblSourceNamedRange.TabIndex = 25;
            this.lblSourceNamedRange.Text = "Named Range";
            // 
            // lblDestinationNamedRange
            // 
            this.lblDestinationNamedRange.AutoSize = true;
            this.lblDestinationNamedRange.Location = new System.Drawing.Point(24, 317);
            this.lblDestinationNamedRange.Name = "lblDestinationNamedRange";
            this.lblDestinationNamedRange.Size = new System.Drawing.Size(132, 13);
            this.lblDestinationNamedRange.TabIndex = 26;
            this.lblDestinationNamedRange.Text = "Destination Named Range";
            // 
            // txtDestinationNameRange
            // 
            this.txtDestinationNameRange.Location = new System.Drawing.Point(20, 335);
            this.txtDestinationNameRange.Name = "txtDestinationNameRange";
            this.txtDestinationNameRange.Size = new System.Drawing.Size(233, 20);
            this.txtDestinationNameRange.TabIndex = 27;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 28;
            this.ssStatus.Text = "status";
            // 
            // FrmInsertCode_SparkLines
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.txtDestinationNameRange);
            this.Controls.Add(this.lblDestinationNamedRange);
            this.Controls.Add(this.lblSourceNamedRange);
            this.Controls.Add(this.chkDestinationCreateNamedRange);
            this.Controls.Add(this.cmbDestinationShtName);
            this.Controls.Add(this.chkNewSheet);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDestinationShtName);
            this.Controls.Add(this.lblDestinationShtName);
            this.Controls.Add(this.lblSourceRange);
            this.Controls.Add(this.txtSourceShtName);
            this.Controls.Add(this.lblSourceShtName);
            this.Controls.Add(this.btnColour);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.grpDirection);
            this.Controls.Add(this.cmbDestinationNameRange);
            this.Controls.Add(this.txtDestinationRange);
            this.Controls.Add(this.btnDestination);
            this.Controls.Add(this.chkDestinationNamedRange);
            this.Controls.Add(this.lblDestination);
            this.Controls.Add(this.cmbSourceNamedRange);
            this.Controls.Add(this.chkSourceNamedRange);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.txtSourceRange);
            this.Controls.Add(this.btnSource);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.shapeContainer1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_SparkLines";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_SparkLines_Load);
            this.grpDirection.ResumeLayout(false);
            this.grpDirection.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnSource;
        private System.Windows.Forms.TextBox txtSourceRange;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.CheckBox chkSourceNamedRange;
        private System.Windows.Forms.ComboBox cmbSourceNamedRange;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.CheckBox chkDestinationNamedRange;
        private System.Windows.Forms.Button btnDestination;
        private System.Windows.Forms.TextBox txtDestinationRange;
        private System.Windows.Forms.ComboBox cmbDestinationNameRange;
        private System.Windows.Forms.GroupBox grpDirection;
        private System.Windows.Forms.RadioButton optLine;
        private System.Windows.Forms.RadioButton optColumnStacked100;
        private System.Windows.Forms.RadioButton optColumn;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Button btnColour;
        private System.Windows.Forms.Label lblSourceShtName;
        private System.Windows.Forms.TextBox txtSourceShtName;
        private System.Windows.Forms.Label lblSourceRange;
        private Microsoft.VisualBasic.PowerPacks.ShapeContainer shapeContainer1;
        private Microsoft.VisualBasic.PowerPacks.LineShape lineShape1;
        private System.Windows.Forms.Label lblDestinationShtName;
        private System.Windows.Forms.TextBox txtDestinationShtName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkNewSheet;
        private System.Windows.Forms.ComboBox cmbDestinationShtName;
        private System.Windows.Forms.CheckBox chkDestinationCreateNamedRange;
        private System.Windows.Forms.Label lblSourceNamedRange;
        private System.Windows.Forms.Label lblDestinationNamedRange;
        private System.Windows.Forms.TextBox txtDestinationNameRange;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}