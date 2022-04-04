namespace KodeMagd.WorkbookAnalysis
{
    partial class FrmObjectModel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmObjectModel));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.ofdBrowse_old = new System.Windows.Forms.OpenFileDialog();
            this.ofdBrowse = new System.Windows.Forms.SaveFileDialog();
            this.chkLstModules = new System.Windows.Forms.CheckedListBox();
            this.chkLstOnlyPublicFunctions = new System.Windows.Forms.CheckBox();
            this.chkIncludeMemberVariables = new System.Windows.Forms.CheckBox();
            this.chkIncludeMemberVariablePublic = new System.Windows.Forms.CheckBox();
            this.lblMemberVariables = new System.Windows.Forms.Label();
            this.chkIncludeMemberVariablePrivate = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(545, 395);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(464, 395);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 1;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 431);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(632, 22);
            this.ssStatus.TabIndex = 4;
            this.ssStatus.Text = "statusStrip1";
            // 
            // ofdBrowse_old
            // 
            this.ofdBrowse_old.CheckFileExists = false;
            this.ofdBrowse_old.DefaultExt = "*.html";
            this.ofdBrowse_old.FileName = "map.html";
            // 
            // ofdBrowse
            // 
            this.ofdBrowse.DefaultExt = "html";
            this.ofdBrowse.Filter = "HTML files (*.html)|*.html";
            // 
            // chkLstModules
            // 
            this.chkLstModules.FormattingEnabled = true;
            this.chkLstModules.Location = new System.Drawing.Point(12, 24);
            this.chkLstModules.Name = "chkLstModules";
            this.chkLstModules.Size = new System.Drawing.Size(262, 349);
            this.chkLstModules.TabIndex = 5;
            this.chkLstModules.KeyUp += new System.Windows.Forms.KeyEventHandler(this.chkLstModules_KeyUp);
            this.chkLstModules.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkLstModules_MouseUp);
            // 
            // chkLstOnlyPublicFunctions
            // 
            this.chkLstOnlyPublicFunctions.AutoSize = true;
            this.chkLstOnlyPublicFunctions.Location = new System.Drawing.Point(299, 26);
            this.chkLstOnlyPublicFunctions.Name = "chkLstOnlyPublicFunctions";
            this.chkLstOnlyPublicFunctions.Size = new System.Drawing.Size(195, 17);
            this.chkLstOnlyPublicFunctions.TabIndex = 6;
            this.chkLstOnlyPublicFunctions.Text = "Only Public Function Sub Properties";
            this.chkLstOnlyPublicFunctions.UseVisualStyleBackColor = true;
            // 
            // chkIncludeMemberVariables
            // 
            this.chkIncludeMemberVariables.AutoSize = true;
            this.chkIncludeMemberVariables.Location = new System.Drawing.Point(299, 149);
            this.chkIncludeMemberVariables.Name = "chkIncludeMemberVariables";
            this.chkIncludeMemberVariables.Size = new System.Drawing.Size(148, 17);
            this.chkIncludeMemberVariables.TabIndex = 7;
            this.chkIncludeMemberVariables.Text = "Include Member Variables";
            this.chkIncludeMemberVariables.UseVisualStyleBackColor = true;
            this.chkIncludeMemberVariables.CheckedChanged += new System.EventHandler(this.chkIncludeMemberVariables_CheckedChanged);
            // 
            // chkIncludeMemberVariablePublic
            // 
            this.chkIncludeMemberVariablePublic.AutoSize = true;
            this.chkIncludeMemberVariablePublic.Location = new System.Drawing.Point(347, 179);
            this.chkIncludeMemberVariablePublic.Name = "chkIncludeMemberVariablePublic";
            this.chkIncludeMemberVariablePublic.Size = new System.Drawing.Size(55, 17);
            this.chkIncludeMemberVariablePublic.TabIndex = 8;
            this.chkIncludeMemberVariablePublic.Text = "Public";
            this.chkIncludeMemberVariablePublic.UseVisualStyleBackColor = true;
            // 
            // lblMemberVariables
            // 
            this.lblMemberVariables.AutoSize = true;
            this.lblMemberVariables.Location = new System.Drawing.Point(296, 122);
            this.lblMemberVariables.Name = "lblMemberVariables";
            this.lblMemberVariables.Size = new System.Drawing.Size(91, 13);
            this.lblMemberVariables.TabIndex = 9;
            this.lblMemberVariables.Text = "Member Variables";
            // 
            // chkIncludeMemberVariablePrivate
            // 
            this.chkIncludeMemberVariablePrivate.AutoSize = true;
            this.chkIncludeMemberVariablePrivate.Location = new System.Drawing.Point(347, 202);
            this.chkIncludeMemberVariablePrivate.Name = "chkIncludeMemberVariablePrivate";
            this.chkIncludeMemberVariablePrivate.Size = new System.Drawing.Size(59, 17);
            this.chkIncludeMemberVariablePrivate.TabIndex = 10;
            this.chkIncludeMemberVariablePrivate.Text = "Private";
            this.chkIncludeMemberVariablePrivate.UseVisualStyleBackColor = true;
            // 
            // FrmObjectModel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 453);
            this.Controls.Add(this.chkIncludeMemberVariablePrivate);
            this.Controls.Add(this.lblMemberVariables);
            this.Controls.Add(this.chkIncludeMemberVariablePublic);
            this.Controls.Add(this.chkIncludeMemberVariables);
            this.Controls.Add(this.chkLstOnlyPublicFunctions);
            this.Controls.Add(this.chkLstModules);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(640, 480);
            this.Name = "FrmObjectModel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmModuleMap_Load);
            this.Resize += new System.EventHandler(this.FrmObjectModel_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.OpenFileDialog ofdBrowse_old;
        private System.Windows.Forms.SaveFileDialog ofdBrowse;
        private System.Windows.Forms.CheckedListBox chkLstModules;
        private System.Windows.Forms.CheckBox chkLstOnlyPublicFunctions;
        private System.Windows.Forms.CheckBox chkIncludeMemberVariables;
        private System.Windows.Forms.CheckBox chkIncludeMemberVariablePublic;
        private System.Windows.Forms.Label lblMemberVariables;
        private System.Windows.Forms.CheckBox chkIncludeMemberVariablePrivate;
    }
}