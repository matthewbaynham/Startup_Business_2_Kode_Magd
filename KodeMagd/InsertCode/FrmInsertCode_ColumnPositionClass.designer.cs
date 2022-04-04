namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_ColumnPositionClass
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_ColumnPositionClass));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.lblClassName = new System.Windows.Forms.Label();
            this.txtClassName = new System.Windows.Forms.TextBox();
            this.grpReferencingObjects = new System.Windows.Forms.GroupBox();
            this.optPassingObjects = new System.Windows.Forms.RadioButton();
            this.optUsingNames = new System.Windows.Forms.RadioButton();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.grpReferencingObjects.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 395);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(464, 395);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 3;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // lblClassName
            // 
            this.lblClassName.AutoSize = true;
            this.lblClassName.Location = new System.Drawing.Point(158, 22);
            this.lblClassName.Name = "lblClassName";
            this.lblClassName.Size = new System.Drawing.Size(70, 13);
            this.lblClassName.TabIndex = 1;
            this.lblClassName.Text = "Name (Suffix)";
            // 
            // txtClassName
            // 
            this.txtClassName.Location = new System.Drawing.Point(161, 41);
            this.txtClassName.Name = "txtClassName";
            this.txtClassName.Size = new System.Drawing.Size(459, 20);
            this.txtClassName.TabIndex = 2;
            // 
            // grpReferencingObjects
            // 
            this.grpReferencingObjects.Controls.Add(this.optPassingObjects);
            this.grpReferencingObjects.Controls.Add(this.optUsingNames);
            this.grpReferencingObjects.Location = new System.Drawing.Point(12, 12);
            this.grpReferencingObjects.Name = "grpReferencingObjects";
            this.grpReferencingObjects.Size = new System.Drawing.Size(130, 67);
            this.grpReferencingObjects.TabIndex = 0;
            this.grpReferencingObjects.TabStop = false;
            this.grpReferencingObjects.Text = "Referencing Objects";
            // 
            // optPassingObjects
            // 
            this.optPassingObjects.AutoSize = true;
            this.optPassingObjects.Location = new System.Drawing.Point(6, 42);
            this.optPassingObjects.Name = "optPassingObjects";
            this.optPassingObjects.Size = new System.Drawing.Size(101, 17);
            this.optPassingObjects.TabIndex = 1;
            this.optPassingObjects.TabStop = true;
            this.optPassingObjects.Text = "Passing Objects";
            this.optPassingObjects.UseVisualStyleBackColor = true;
            // 
            // optUsingNames
            // 
            this.optUsingNames.AutoSize = true;
            this.optUsingNames.Location = new System.Drawing.Point(6, 19);
            this.optUsingNames.Name = "optUsingNames";
            this.optUsingNames.Size = new System.Drawing.Size(88, 17);
            this.optUsingNames.TabIndex = 0;
            this.optUsingNames.TabStop = true;
            this.optUsingNames.Text = "Using Names";
            this.optUsingNames.UseVisualStyleBackColor = true;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 5;
            this.ssStatus.Text = "status";
            // 
            // FrmInsertCode_ColumnPositionClass
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.grpReferencingObjects);
            this.Controls.Add(this.txtClassName);
            this.Controls.Add(this.lblClassName);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_ColumnPositionClass";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_ColumnPositionClass_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_ColumnPositionClass_KeyDown);
            this.grpReferencingObjects.ResumeLayout(false);
            this.grpReferencingObjects.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label lblClassName;
        private System.Windows.Forms.TextBox txtClassName;
        private System.Windows.Forms.GroupBox grpReferencingObjects;
        private System.Windows.Forms.RadioButton optPassingObjects;
        private System.Windows.Forms.RadioButton optUsingNames;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}