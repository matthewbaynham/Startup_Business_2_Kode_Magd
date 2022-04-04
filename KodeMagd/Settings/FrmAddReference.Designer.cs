namespace KodeMagd.Settings
{
    partial class FrmAddReference
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmAddReference));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnAddReference = new System.Windows.Forms.Button();
            this.lstVwAssembliesTypeLib = new System.Windows.Forms.ListView();
            this.colAssName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAssVersion = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAssPath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAssGUID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAssWinXX = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblComments = new System.Windows.Forms.Label();
            this.chkFiltered = new System.Windows.Forms.CheckBox();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(383, 277);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnAddReference
            // 
            this.btnAddReference.Location = new System.Drawing.Point(302, 277);
            this.btnAddReference.Name = "btnAddReference";
            this.btnAddReference.Size = new System.Drawing.Size(75, 23);
            this.btnAddReference.TabIndex = 3;
            this.btnAddReference.Text = "&Add";
            this.btnAddReference.UseVisualStyleBackColor = true;
            this.btnAddReference.Click += new System.EventHandler(this.btnAddReference_Click);
            // 
            // lstVwAssembliesTypeLib
            // 
            this.lstVwAssembliesTypeLib.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colAssName,
            this.colAssVersion,
            this.colAssPath,
            this.colAssGUID,
            this.colAssWinXX});
            this.lstVwAssembliesTypeLib.Location = new System.Drawing.Point(13, 12);
            this.lstVwAssembliesTypeLib.Name = "lstVwAssembliesTypeLib";
            this.lstVwAssembliesTypeLib.Size = new System.Drawing.Size(444, 247);
            this.lstVwAssembliesTypeLib.TabIndex = 0;
            this.lstVwAssembliesTypeLib.UseCompatibleStateImageBehavior = false;
            this.lstVwAssembliesTypeLib.View = System.Windows.Forms.View.Details;
            this.lstVwAssembliesTypeLib.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lstVwAssembliesTypeLib_ColumnClick);
            this.lstVwAssembliesTypeLib.DoubleClick += new System.EventHandler(this.lstVwAssembliesTypeLib_DoubleClick);
            // 
            // colAssName
            // 
            this.colAssName.Text = "Name";
            this.colAssName.Width = 120;
            // 
            // colAssVersion
            // 
            this.colAssVersion.Text = "Version";
            // 
            // colAssPath
            // 
            this.colAssPath.Text = "Path";
            this.colAssPath.Width = 120;
            // 
            // colAssGUID
            // 
            this.colAssGUID.Text = "GUID";
            // 
            // colAssWinXX
            // 
            this.colAssWinXX.Text = "WinXX";
            // 
            // lblComments
            // 
            this.lblComments.Location = new System.Drawing.Point(15, 274);
            this.lblComments.Name = "lblComments";
            this.lblComments.Size = new System.Drawing.Size(281, 23);
            this.lblComments.TabIndex = 1;
            this.lblComments.Text = "Comments";
            // 
            // chkFiltered
            // 
            this.chkFiltered.AutoSize = true;
            this.chkFiltered.Location = new System.Drawing.Point(13, 291);
            this.chkFiltered.Name = "chkFiltered";
            this.chkFiltered.Size = new System.Drawing.Size(60, 17);
            this.chkFiltered.TabIndex = 2;
            this.chkFiltered.Text = "Filtered";
            this.chkFiltered.UseVisualStyleBackColor = true;
            this.chkFiltered.CheckedChanged += new System.EventHandler(this.chkFiltered_CheckedChanged);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 311);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(472, 22);
            this.ssStatus.TabIndex = 5;
            this.ssStatus.Text = "Status";
            // 
            // FrmAddReference
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 333);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.chkFiltered);
            this.Controls.Add(this.lblComments);
            this.Controls.Add(this.lstVwAssembliesTypeLib);
            this.Controls.Add(this.btnAddReference);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.Name = "FrmAddReference";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmAddReference_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmAddReference_KeyDown);
            this.Resize += new System.EventHandler(this.FrmAddReference_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnAddReference;
        private System.Windows.Forms.ListView lstVwAssembliesTypeLib;
        private System.Windows.Forms.ColumnHeader colAssName;
        private System.Windows.Forms.ColumnHeader colAssVersion;
        private System.Windows.Forms.ColumnHeader colAssPath;
        private System.Windows.Forms.ColumnHeader colAssGUID;
        private System.Windows.Forms.ColumnHeader colAssWinXX;
        private System.Windows.Forms.Label lblComments;
        private System.Windows.Forms.CheckBox chkFiltered;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}