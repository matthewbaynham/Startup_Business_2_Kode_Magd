namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_Rst_Range
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_Rst_Range));
            this.btnClose = new System.Windows.Forms.Button();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.optCoordinates = new System.Windows.Forms.RadioButton();
            this.optNamedRange = new System.Windows.Forms.RadioButton();
            this.txtWrkName = new System.Windows.Forms.TextBox();
            this.txtShtName = new System.Windows.Forms.TextBox();
            this.txtNamedRange = new System.Windows.Forms.TextBox();
            this.lblWrkName = new System.Windows.Forms.Label();
            this.lblShtName = new System.Windows.Forms.Label();
            this.lblNamedRange = new System.Windows.Forms.Label();
            this.lblRow = new System.Windows.Forms.Label();
            this.txtRow = new System.Windows.Forms.TextBox();
            this.lblColumn = new System.Windows.Forms.Label();
            this.txtColumn = new System.Windows.Forms.TextBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.grpType.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(382, 271);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optCoordinates);
            this.grpType.Controls.Add(this.optNamedRange);
            this.grpType.Location = new System.Drawing.Point(12, 12);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(114, 67);
            this.grpType.TabIndex = 0;
            this.grpType.TabStop = false;
            this.grpType.Text = "Type";
            // 
            // optCoordinates
            // 
            this.optCoordinates.AutoSize = true;
            this.optCoordinates.Location = new System.Drawing.Point(6, 42);
            this.optCoordinates.Name = "optCoordinates";
            this.optCoordinates.Size = new System.Drawing.Size(81, 17);
            this.optCoordinates.TabIndex = 1;
            this.optCoordinates.TabStop = true;
            this.optCoordinates.Text = "Coordinates";
            this.optCoordinates.UseVisualStyleBackColor = true;
            this.optCoordinates.CheckedChanged += new System.EventHandler(this.optCoordinates_CheckedChanged);
            // 
            // optNamedRange
            // 
            this.optNamedRange.AutoSize = true;
            this.optNamedRange.Location = new System.Drawing.Point(6, 19);
            this.optNamedRange.Name = "optNamedRange";
            this.optNamedRange.Size = new System.Drawing.Size(94, 17);
            this.optNamedRange.TabIndex = 0;
            this.optNamedRange.TabStop = true;
            this.optNamedRange.Text = "Named Range";
            this.optNamedRange.UseVisualStyleBackColor = true;
            this.optNamedRange.CheckedChanged += new System.EventHandler(this.optNamedRange_CheckedChanged);
            // 
            // txtWrkName
            // 
            this.txtWrkName.Location = new System.Drawing.Point(12, 117);
            this.txtWrkName.Name = "txtWrkName";
            this.txtWrkName.Size = new System.Drawing.Size(180, 20);
            this.txtWrkName.TabIndex = 2;
            // 
            // txtShtName
            // 
            this.txtShtName.Location = new System.Drawing.Point(12, 172);
            this.txtShtName.Name = "txtShtName";
            this.txtShtName.Size = new System.Drawing.Size(180, 20);
            this.txtShtName.TabIndex = 4;
            // 
            // txtNamedRange
            // 
            this.txtNamedRange.Location = new System.Drawing.Point(235, 31);
            this.txtNamedRange.Name = "txtNamedRange";
            this.txtNamedRange.Size = new System.Drawing.Size(180, 20);
            this.txtNamedRange.TabIndex = 4;
            // 
            // lblWrkName
            // 
            this.lblWrkName.AutoSize = true;
            this.lblWrkName.Location = new System.Drawing.Point(9, 101);
            this.lblWrkName.Name = "lblWrkName";
            this.lblWrkName.Size = new System.Drawing.Size(88, 13);
            this.lblWrkName.TabIndex = 1;
            this.lblWrkName.Text = "Workbook Name";
            // 
            // lblShtName
            // 
            this.lblShtName.AutoSize = true;
            this.lblShtName.Location = new System.Drawing.Point(14, 156);
            this.lblShtName.Name = "lblShtName";
            this.lblShtName.Size = new System.Drawing.Size(90, 13);
            this.lblShtName.TabIndex = 3;
            this.lblShtName.Text = "Worksheet Name";
            // 
            // lblNamedRange
            // 
            this.lblNamedRange.AutoSize = true;
            this.lblNamedRange.Location = new System.Drawing.Point(232, 12);
            this.lblNamedRange.Name = "lblNamedRange";
            this.lblNamedRange.Size = new System.Drawing.Size(76, 13);
            this.lblNamedRange.TabIndex = 7;
            this.lblNamedRange.Text = "Named Range";
            // 
            // lblRow
            // 
            this.lblRow.AutoSize = true;
            this.lblRow.Location = new System.Drawing.Point(332, 12);
            this.lblRow.Name = "lblRow";
            this.lblRow.Size = new System.Drawing.Size(29, 13);
            this.lblRow.TabIndex = 6;
            this.lblRow.Text = "Row";
            // 
            // txtRow
            // 
            this.txtRow.Location = new System.Drawing.Point(335, 31);
            this.txtRow.Name = "txtRow";
            this.txtRow.Size = new System.Drawing.Size(80, 20);
            this.txtRow.TabIndex = 8;
            // 
            // lblColumn
            // 
            this.lblColumn.AutoSize = true;
            this.lblColumn.Location = new System.Drawing.Point(232, 12);
            this.lblColumn.Name = "lblColumn";
            this.lblColumn.Size = new System.Drawing.Size(42, 13);
            this.lblColumn.TabIndex = 5;
            this.lblColumn.Text = "Column";
            // 
            // txtColumn
            // 
            this.txtColumn.Location = new System.Drawing.Point(235, 31);
            this.txtColumn.Name = "txtColumn";
            this.txtColumn.Size = new System.Drawing.Size(80, 20);
            this.txtColumn.TabIndex = 7;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 316);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(474, 22);
            this.ssStatus.TabIndex = 10;
            this.ssStatus.Text = "status";
            // 
            // FrmInsertCode_Rst_Range
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 338);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.txtColumn);
            this.Controls.Add(this.lblColumn);
            this.Controls.Add(this.txtRow);
            this.Controls.Add(this.lblRow);
            this.Controls.Add(this.lblNamedRange);
            this.Controls.Add(this.lblShtName);
            this.Controls.Add(this.lblWrkName);
            this.Controls.Add(this.txtNamedRange);
            this.Controls.Add(this.txtShtName);
            this.Controls.Add(this.txtWrkName);
            this.Controls.Add(this.grpType);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmInsertCode_Rst_Range";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmInsertCode_Rst_Range";
            this.Load += new System.EventHandler(this.FrmInsertCode_Rst_Range_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_Rst_Range_KeyDown);
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton optCoordinates;
        private System.Windows.Forms.RadioButton optNamedRange;
        private System.Windows.Forms.TextBox txtWrkName;
        private System.Windows.Forms.TextBox txtShtName;
        private System.Windows.Forms.TextBox txtNamedRange;
        private System.Windows.Forms.Label lblWrkName;
        private System.Windows.Forms.Label lblShtName;
        private System.Windows.Forms.Label lblNamedRange;
        private System.Windows.Forms.Label lblRow;
        private System.Windows.Forms.TextBox txtRow;
        private System.Windows.Forms.Label lblColumn;
        private System.Windows.Forms.TextBox txtColumn;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}