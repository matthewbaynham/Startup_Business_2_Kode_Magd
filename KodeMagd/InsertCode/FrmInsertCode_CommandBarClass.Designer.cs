namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_CommandBarClass
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_CommandBarClass));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.treMenu = new System.Windows.Forms.TreeView();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.lblClassName = new System.Windows.Forms.Label();
            this.txtClassName = new System.Windows.Forms.TextBox();
            this.chkPutSampleCallInOwnNewMod = new System.Windows.Forms.CheckBox();
            this.grpType = new System.Windows.Forms.GroupBox();
            this.optRightClick = new System.Windows.Forms.RadioButton();
            this.optRibbonAddin = new System.Windows.Forms.RadioButton();
            this.btnEdit = new System.Windows.Forms.Button();
            this.lblWarning = new System.Windows.Forms.Label();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.grpType.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(540, 399);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 10;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(448, 399);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(86, 23);
            this.btnGenerate.TabIndex = 9;
            this.btnGenerate.Text = "&Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // treMenu
            // 
            this.treMenu.AllowDrop = true;
            this.treMenu.Location = new System.Drawing.Point(14, 12);
            this.treMenu.Name = "treMenu";
            this.treMenu.ShowNodeToolTips = true;
            this.treMenu.Size = new System.Drawing.Size(373, 337);
            this.treMenu.TabIndex = 0;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(14, 399);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(176, 399);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.TabIndex = 8;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // lblClassName
            // 
            this.lblClassName.AutoSize = true;
            this.lblClassName.Location = new System.Drawing.Point(393, 9);
            this.lblClassName.Name = "lblClassName";
            this.lblClassName.Size = new System.Drawing.Size(63, 13);
            this.lblClassName.TabIndex = 1;
            this.lblClassName.Text = "Class Name";
            // 
            // txtClassName
            // 
            this.txtClassName.Location = new System.Drawing.Point(396, 30);
            this.txtClassName.Name = "txtClassName";
            this.txtClassName.Size = new System.Drawing.Size(219, 20);
            this.txtClassName.TabIndex = 2;
            // 
            // chkPutSampleCallInOwnNewMod
            // 
            this.chkPutSampleCallInOwnNewMod.AutoSize = true;
            this.chkPutSampleCallInOwnNewMod.Location = new System.Drawing.Point(396, 56);
            this.chkPutSampleCallInOwnNewMod.Name = "chkPutSampleCallInOwnNewMod";
            this.chkPutSampleCallInOwnNewMod.Size = new System.Drawing.Size(184, 17);
            this.chkPutSampleCallInOwnNewMod.TabIndex = 3;
            this.chkPutSampleCallInOwnNewMod.Text = "Sample Code Call In New Module";
            this.chkPutSampleCallInOwnNewMod.UseVisualStyleBackColor = true;
            // 
            // grpType
            // 
            this.grpType.Controls.Add(this.optRightClick);
            this.grpType.Controls.Add(this.optRibbonAddin);
            this.grpType.Location = new System.Drawing.Point(397, 88);
            this.grpType.Name = "grpType";
            this.grpType.Size = new System.Drawing.Size(217, 73);
            this.grpType.TabIndex = 4;
            this.grpType.TabStop = false;
            this.grpType.Text = "Type";
            // 
            // optRightClick
            // 
            this.optRightClick.AutoSize = true;
            this.optRightClick.Location = new System.Drawing.Point(15, 45);
            this.optRightClick.Name = "optRightClick";
            this.optRightClick.Size = new System.Drawing.Size(76, 17);
            this.optRightClick.TabIndex = 1;
            this.optRightClick.TabStop = true;
            this.optRightClick.Text = "Right Click";
            this.optRightClick.UseVisualStyleBackColor = true;
            this.optRightClick.CheckedChanged += new System.EventHandler(this.optRightClick_CheckedChanged);
            // 
            // optRibbonAddin
            // 
            this.optRibbonAddin.AutoSize = true;
            this.optRibbonAddin.Location = new System.Drawing.Point(15, 22);
            this.optRibbonAddin.Name = "optRibbonAddin";
            this.optRibbonAddin.Size = new System.Drawing.Size(92, 17);
            this.optRibbonAddin.TabIndex = 0;
            this.optRibbonAddin.TabStop = true;
            this.optRibbonAddin.Text = "Ribbon Add-in";
            this.optRibbonAddin.UseVisualStyleBackColor = true;
            this.optRibbonAddin.CheckedChanged += new System.EventHandler(this.optRibbonAddin_CheckedChanged);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(95, 399);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(75, 23);
            this.btnEdit.TabIndex = 7;
            this.btnEdit.Text = "&Edit";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // lblWarning
            // 
            this.lblWarning.Location = new System.Drawing.Point(402, 176);
            this.lblWarning.Name = "lblWarning";
            this.lblWarning.Size = new System.Drawing.Size(212, 173);
            this.lblWarning.TabIndex = 5;
            this.lblWarning.Text = "Warning";
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 436);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(634, 22);
            this.ssStatus.TabIndex = 11;
            this.ssStatus.Text = "status";
            // 
            // FrmInsertCode_CommandBarClass
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(634, 458);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.lblWarning);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.grpType);
            this.Controls.Add(this.chkPutSampleCallInOwnNewMod);
            this.Controls.Add(this.txtClassName);
            this.Controls.Add(this.lblClassName);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.treMenu);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmInsertCode_CommandBarClass";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmInsertCode_CommandBarClass_FormClosing);
            this.Load += new System.EventHandler(this.FrmInsertCode_CommandBarClass_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_CommandBarClass_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_CommandBarClass_Resize);
            this.grpType.ResumeLayout(false);
            this.grpType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.TreeView treMenu;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Label lblClassName;
        private System.Windows.Forms.TextBox txtClassName;
        private System.Windows.Forms.CheckBox chkPutSampleCallInOwnNewMod;
        private System.Windows.Forms.GroupBox grpType;
        private System.Windows.Forms.RadioButton optRightClick;
        private System.Windows.Forms.RadioButton optRibbonAddin;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}