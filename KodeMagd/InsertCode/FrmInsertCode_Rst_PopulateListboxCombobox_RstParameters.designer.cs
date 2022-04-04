namespace KodeMagd.InsertCode
{
    partial class FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.ssStatus = new System.Windows.Forms.StatusStrip();
            this.dgParameters = new System.Windows.Forms.DataGridView();
            this.ColName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColType = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColDirection = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ColSize = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.lblRememberQuestionMarks = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(407, 277);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(55, 23);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(346, 277);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(55, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "&OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // ssStatus
            // 
            this.ssStatus.Location = new System.Drawing.Point(0, 316);
            this.ssStatus.Name = "ssStatus";
            this.ssStatus.Size = new System.Drawing.Size(474, 22);
            this.ssStatus.TabIndex = 4;
            this.ssStatus.Text = "status";
            // 
            // dgParameters
            // 
            this.dgParameters.AllowUserToAddRows = false;
            this.dgParameters.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgParameters.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColName,
            this.ColType,
            this.ColDirection,
            this.ColSize,
            this.ColValue});
            this.dgParameters.Location = new System.Drawing.Point(12, 12);
            this.dgParameters.MultiSelect = false;
            this.dgParameters.Name = "dgParameters";
            this.dgParameters.RowHeadersVisible = false;
            this.dgParameters.Size = new System.Drawing.Size(450, 229);
            this.dgParameters.TabIndex = 5;
            this.dgParameters.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dgParameters_RowsAdded);
            // 
            // ColName
            // 
            this.ColName.HeaderText = "Name";
            this.ColName.Name = "ColName";
            // 
            // ColType
            // 
            this.ColType.HeaderText = "Data Type";
            this.ColType.Name = "ColType";
            // 
            // ColDirection
            // 
            this.ColDirection.HeaderText = "Direction";
            this.ColDirection.Name = "ColDirection";
            // 
            // ColSize
            // 
            this.ColSize.FillWeight = 60F;
            this.ColSize.HeaderText = "Size";
            this.ColSize.Name = "ColSize";
            this.ColSize.Width = 60;
            // 
            // ColValue
            // 
            this.ColValue.HeaderText = "Value";
            this.ColValue.Name = "ColValue";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(12, 277);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Text = "&Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(93, 277);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(75, 23);
            this.btnRemove.TabIndex = 7;
            this.btnRemove.Text = "&Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
            // 
            // lblRememberQuestionMarks
            // 
            this.lblRememberQuestionMarks.AutoSize = true;
            this.lblRememberQuestionMarks.ForeColor = System.Drawing.Color.Red;
            this.lblRememberQuestionMarks.Location = new System.Drawing.Point(14, 253);
            this.lblRememberQuestionMarks.Name = "lblRememberQuestionMarks";
            this.lblRememberQuestionMarks.Size = new System.Drawing.Size(406, 13);
            this.lblRememberQuestionMarks.TabIndex = 8;
            this.lblRememberQuestionMarks.Text = "Remember to add ? charactors in the SQL statement.  Each ? relates to a parameter" +
    ".";
            // 
            // FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 338);
            this.Controls.Add(this.lblRememberQuestionMarks);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.dgParameters);
            this.Controls.Add(this.ssStatus);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_KeyDown);
            this.Resize += new System.EventHandler(this.FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgParameters)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.StatusStrip ssStatus;
        private System.Windows.Forms.DataGridView dgParameters;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColName;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColType;
        private System.Windows.Forms.DataGridViewComboBoxColumn ColDirection;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColSize;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColValue;
        private System.Windows.Forms.Label lblRememberQuestionMarks;
    }
}