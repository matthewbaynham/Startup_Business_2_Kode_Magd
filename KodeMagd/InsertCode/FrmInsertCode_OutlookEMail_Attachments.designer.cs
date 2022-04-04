namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_OutlookEMail_Attachments
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_OutlookEMail_Attachments));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.dgAttachments = new System.Windows.Forms.DataGridView();
            this.ColFullPath = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColBrowse = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ColDelete = new System.Windows.Forms.DataGridViewButtonColumn();
            this.btnAdd = new System.Windows.Forms.Button();
            this.ofdAdd = new System.Windows.Forms.OpenFileDialog();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            ((System.ComponentModel.ISupportInitialize)(this.dgAttachments)).BeginInit();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(387, 285);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(306, 285);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // dgAttachments
            // 
            this.dgAttachments.AllowUserToAddRows = false;
            this.dgAttachments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgAttachments.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColFullPath,
            this.ColBrowse,
            this.ColDelete});
            this.dgAttachments.Location = new System.Drawing.Point(13, 13);
            this.dgAttachments.Name = "dgAttachments";
            this.dgAttachments.RowHeadersVisible = false;
            this.dgAttachments.Size = new System.Drawing.Size(447, 266);
            this.dgAttachments.TabIndex = 0;
            this.dgAttachments.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgAttachments_CellClick);
            // 
            // ColFullPath
            // 
            this.ColFullPath.HeaderText = "Full Path";
            this.ColFullPath.Name = "ColFullPath";
            this.ColFullPath.Width = 200;
            // 
            // ColBrowse
            // 
            this.ColBrowse.HeaderText = "";
            this.ColBrowse.Name = "ColBrowse";
            this.ColBrowse.ReadOnly = true;
            this.ColBrowse.Text = "Browse";
            this.ColBrowse.UseColumnTextForButtonValue = true;
            this.ColBrowse.Width = 50;
            // 
            // ColDelete
            // 
            this.ColDelete.HeaderText = "";
            this.ColDelete.Name = "ColDelete";
            this.ColDelete.Text = "Delete";
            this.ColDelete.UseColumnTextForButtonValue = true;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(12, 285);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // ofdAdd
            // 
            this.ofdAdd.Multiselect = true;
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 314);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(472, 22);
            this.ssStatus.TabIndex = 4;
            this.ssStatus.Text = "status";
            // 
            // FrmInsertCode_OutlookEMail_Attachments
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 336);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.dgAttachments);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmInsertCode_OutlookEMail_Attachments";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_OutlookEMail_Attachments_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_OutlookEMail_Attachments_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_OutlookEMail_Attachments_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgAttachments)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.DataGridView dgAttachments;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColFullPath;
        private System.Windows.Forms.DataGridViewButtonColumn ColBrowse;
        private System.Windows.Forms.DataGridViewButtonColumn ColDelete;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.OpenFileDialog ofdAdd;
        private System.Windows.Forms.StatusStrip ssStatus;
    }
}