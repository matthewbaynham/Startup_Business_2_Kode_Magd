using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.InsertCode;
using KodeMagd.Misc;
using KodeMagd.Settings;
using System.Threading;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    //public delegate void refreshStatusBarDelegate(); //This line might be preventing the cConfigReporter initialisation being recognised.
    
    public partial class FrmInsertCode_TextFile : Form
    {
        //private ClsStatusBarMultiThreadUpdater workerObject;
        //private Thread workerThread;
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();

        private ClsInsertCode_Files.strFileFormat objFileFormat;

        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_TextFile()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try 
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void FrmInsertCode_TextFile_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;
                txtVariableName.Text = ClsDefaults.defaultName;

                ClsDefaults.FormatControl(ref btnCancel);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref grpDirection);
                ClsDefaults.FormatControl(ref optDirectionRead);
                ClsDefaults.FormatControl(ref optDirectionWrite);

                ClsDefaults.FormatControl(ref btnAddColumn);

                ClsDefaults.FormatControl(ref grpType);
                ClsDefaults.FormatControl(ref optDelimited);
                ClsDefaults.FormatControl(ref optFixedFieldLength);

                ClsDefaults.FormatControl(ref grpDelimiter);
                ClsDefaults.FormatControl(ref optColon);
                ClsDefaults.FormatControl(ref optComma);
                ClsDefaults.FormatControl(ref optOther);
                ClsDefaults.FormatControl(ref optSemiColon);
                ClsDefaults.FormatControl(ref optTab);
                ClsDefaults.FormatControl(ref txtDelimiterOther);

                ClsDefaults.FormatControl(ref dgFixedColumns);

                ClsDefaults.FormatControl(ref btnAddColumn);
                ClsDefaults.FormatControl(ref chkAutoupdatePositions);

                ClsDefaults.FormatControl(ref ssStatus);

                optDirectionRead.Checked = true;
                optDirectionWrite.Checked = false;

                optFixedFieldLength.Checked = false;
                optDelimited.Checked = true;

                optTab.Checked = true;
                chkAddReferences.Checked = true;

                objFileFormat.cDelimiter = '\t';
                objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eDelimitedFile;
                objFileFormat.lstColumns = new List<ClsInsertCode_Files.strFileFormat_FixedColumn>();

                changeType();

                cCodeMapper.readCode();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                generate();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void generate()
        {
            try
            {
                //http://msdn.microsoft.com/en-us/library/ms171728.aspx

                //http://msdn.microsoft.com/en-us/library/ms173178.aspx
                //http://msdn.microsoft.com/en-us/library/3e8s7xdd%28v=vs.80%29.aspx
                //http://msdn.microsoft.com/en-us/library/ck8bc5c6.aspx
                //http://stackoverflow.com/questions/142003/cross-thread-operation-not-valid-control-accessed-from-a-thread-other-than-the
                //http://stackoverflow.com/questions/10775367/cross-thread-operation-not-valid-control-textbox1-accessed-from-a-thread-othe
                //http://msdn.microsoft.com/en-us/library/ms171728.aspx

                ClsDefaults.changeStatusStrip_ProgressBar(ref this.ssStatus, true);
                ClsDefaults.changeStatusStrip_ProgressBar(ref this.ssStatus);

                if (chkAddReferences.Checked == true)
                {
                    FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_Scripting, ref ssStatus);

                    if (!frmReference.referenceAlreadySet)
                    { frmReference.ShowDialog(this); }

                    frmReference = null;
                }

                ClsInsertCode_Files cInsertCode_Files = new ClsInsertCode_Files();
                bool bIsOk = true;
                string sErrorMessage = "";

                this.setObjFileFormat();

                string sVariableName = ClsMiscString.makeValidVarName(txtVariableName.Text);

                cInsertCode_Files.Name = sVariableName;
                cInsertCode_Files.FullFilePath = txtPath.Text;
                cInsertCode_Files.fileFormat = objFileFormat;

                if (bIsOk)
                {
                    if (optDirectionRead.Checked & !optDirectionWrite.Checked)
                    { cInsertCode_Files.direction = ClsInsertCode_Files.enumDirection.eDir_Read; }
                    else if (!optDirectionRead.Checked & optDirectionWrite.Checked)
                    { cInsertCode_Files.direction = ClsInsertCode_Files.enumDirection.eDir_Write; }
                    else
                    {
                        cInsertCode_Files.direction = ClsInsertCode_Files.enumDirection.eDir_Unknown;
                        bIsOk = false;
                        sErrorMessage = "Much select either Read or Write";
                    }
                }

                if (bIsOk) 
                {
                    cInsertCode_Files.generateCode(ref cCodeMapper);
                
                    configHtmlSummary(ref cInsertCode_Files);
                    displayHtmlSummary();

                    cInsertCode_Files = null;

                    ClsDefaults.changeStatusStrip_ProgressBar(ref this.ssStatus, false);

                    this.Close(); 
                }
                else
                {
                    cInsertCode_Files = null;
                    ClsDefaults.changeStatusStrip_ProgressBar(ref this.ssStatus, false);
                    MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); 
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            try
            {
                browse();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void browse()
        {
            try
            {
                string sFullPath = "";

                if (optDirectionRead.Checked & !optDirectionWrite.Checked)
                {
                    ofdBrowseOpen.Multiselect = false;

                    DialogResult result = ofdBrowseOpen.ShowDialog(this);

                    if (result == DialogResult.OK)
                    {
                        sFullPath = ofdBrowseOpen.FileName;
                        txtPath.Text = sFullPath;
                    }
                }
                else if (!optDirectionRead.Checked & optDirectionWrite.Checked)
                {
                    DialogResult result = ofdBrowseSave.ShowDialog(this);

                    if (result == DialogResult.OK)
                    {
                        sFullPath = ofdBrowseSave.FileName;
                        txtPath.Text = sFullPath;
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optDirectionRead_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //directionChange();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optDirectionWrite_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //directionChange();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optDelimited_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                changeType();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optFixedFieldLength_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                changeType();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void changeType()
        {
            try
            {
                if (optDelimited.Checked & !optFixedFieldLength.Checked)
                {
                    this.objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eDelimitedFile;
                    grpDelimiter.Visible = true;
                    dgFixedColumns.Visible = false;
                    btnAddColumn.Visible = false;
                    chkAutoupdatePositions.Visible = false;
                }
                else if (!optDelimited.Checked & optFixedFieldLength.Checked)
                {
                    this.objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eFixedColumnLengthFile;
                    grpDelimiter.Visible = false;
                    dgFixedColumns.Visible = true;
                    btnAddColumn.Visible = true;
                    chkAutoupdatePositions.Visible = true;
                }
                else
                {
                    this.objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eUnknown;
                    grpDelimiter.Visible = false;
                    dgFixedColumns.Visible = false;
                    btnAddColumn.Visible = false;
                    chkAutoupdatePositions.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optOther_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optOther.Checked)
                { txtDelimiterOther.Enabled = true; }
                else
                { txtDelimiterOther.Enabled = false; }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optSemiColon_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optOther.Checked)
                { txtDelimiterOther.Enabled = true; }
                else
                { txtDelimiterOther.Enabled = false; }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optTab_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optOther.Checked)
                { txtDelimiterOther.Enabled = true; }
                else
                { txtDelimiterOther.Enabled = false; }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optComma_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optOther.Checked)
                { txtDelimiterOther.Enabled = true; }
                else
                { txtDelimiterOther.Enabled = false; }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void optColon_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optOther.Checked)
                { txtDelimiterOther.Enabled = true; }
                else
                { txtDelimiterOther.Enabled = false; }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            try
            {
                add();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void add()
        {
            try
            {
                string sName = FrmInputBox.GetString(ClsDefaults.formTitle,"Please enter the name of the field");

                if (sName.Trim() != "")
                {
                    int iRow = dgFixedColumns.Rows.Add();

                    dgFixedColumns[ColName.Index, iRow].Value = sName;
                    dgFixedColumns[ColSize.Index, iRow].Value = 0;
                    dgFixedColumns[ColStartChar.Index, iRow].Value = 0;
                    dgFixedColumns[ColDataType.Index, iRow].Value = "";
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void dgFixedColumns_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == ColSize.Index || e.ColumnIndex == ColStartChar.Index)
                {
                    if (dgFixedColumns[e.ColumnIndex, e.RowIndex].Value != null)
                    {
                        string sTemp = dgFixedColumns[e.ColumnIndex, e.RowIndex].Value.ToString();

                        int iTemp;
                        if (int.TryParse(sTemp, out iTemp))
                        {
                            if (chkAutoupdatePositions.Checked)
                            {
                                int iPos = 1;
                                for (int iRow = 0; iRow < dgFixedColumns.RowCount; iRow++)
                                {
                                    dgFixedColumns[ColStartChar.Index, iRow].Value = iPos;
                                    int iSize;
                                    if (dgFixedColumns[ColSize.Index, iRow].Value == null)
                                    { iSize = 0; }
                                    else
                                    {
                                        if (!int.TryParse(dgFixedColumns[ColSize.Index, iRow].Value.ToString(), out iSize))
                                        { iSize = 0; }
                                    }
                                    iPos += iSize;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("The Column " + dgFixedColumns.Columns[e.ColumnIndex].HeaderText + " must be an integer.", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            dgFixedColumns[e.ColumnIndex, e.RowIndex].Value = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void setObjFileFormat()
        {
            try
            {
                if (optDelimited.Checked & !optFixedFieldLength.Checked)
                {
                    objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eDelimitedFile;
                    objFileFormat.lstColumns = new List<ClsInsertCode_Files.strFileFormat_FixedColumn>();
                    if (optOther.Checked)
                    {
                        string sTemp = txtDelimiterOther.Text;
                        char cTemp = (char)sTemp[0];

                        objFileFormat.cDelimiter = cTemp;
                    }
                    else if (optTab.Checked)
                    { objFileFormat.cDelimiter = '\t'; }
                    else if (optSemiColon.Checked)
                    { objFileFormat.cDelimiter = ';'; }
                    else if (optColon.Checked)
                    { objFileFormat.cDelimiter = ':'; }
                    else if (optComma.Checked)
                    { objFileFormat.cDelimiter = ','; }
                    else
                    { objFileFormat.cDelimiter = ' '; }
                }
                else if (!optDelimited.Checked & optFixedFieldLength.Checked)
                {
                    ClsDataTypes cDataTypes = new ClsDataTypes();

                    objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eFixedColumnLengthFile;
                    objFileFormat.cDelimiter = ' ';
                    objFileFormat.lstColumns = new List<ClsInsertCode_Files.strFileFormat_FixedColumn>();

                    //int iPos = 0;
                    for (int iRow = 0; iRow < dgFixedColumns.RowCount; iRow++)
                    {
                        ClsInsertCode_Files.strFileFormat_FixedColumn objColumn;

                        objColumn.sName = dgFixedColumns[ColName.Index, iRow].Value.ToString();
                        //objColumn.bEnabled = (bool)dgFixedColumns[ciParamCol_Enabled, iRow].Value;

                        string sDataType;
                        if (dgFixedColumns[ColDataType.Index, iRow].Value == null)
                        { sDataType = ClsDataTypes.vbVarType.vbUnknown.ToString(); }
                        else
                        {
                            sDataType = dgFixedColumns[ColDataType.Index, iRow].Value.ToString();

                            if (sDataType == "")
                            { sDataType = ClsDataTypes.vbVarType.vbUnknown.ToString(); }
                        }

                        objColumn.eDataType = cDataTypes.getDataType(sDataType);

                        int iStartPos;
                        if (!int.TryParse(dgFixedColumns[ColStartChar.Index, iRow].Value.ToString(), out iStartPos))
                        { iStartPos = 0; }
                        objColumn.iPosStart = iStartPos;
                        
                        int iSize;
                        if (!int.TryParse(dgFixedColumns[ColSize.Index, iRow].Value.ToString(), out iSize))
                        { iSize = 0; }
                        objColumn.iSize = iSize;

                        //iPos = iPos + iSize;

                        this.objFileFormat.lstColumns.Add(objColumn);
                    }

                }
                else
                { objFileFormat.eFileType = ClsInsertCode_Files.enumFileType.eUnknown; }



            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void dgFixedColumns_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();

                dgFixedColumns[ColSize.Index, e.RowIndex].Value = 0;
                dgFixedColumns[ColStartChar.Index, e.RowIndex].Value = 0;

                DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgFixedColumns[ColDataType.Index, e.RowIndex];

                //Array arrDataType = Enum.GetValues(typeof(ADODB.DataTypeEnum));
                //Array.Sort(arrDataType);

                List<string> lstDataType = cDataTypes.allDataTypes();

                //foreach (ADODB.DataTypeEnum eTemp in arrDataType)
                //{ lstDataType.Add(eTemp.ToString()); }
                lstDataType.Sort();

                foreach (string sTemp in lstDataType)
                { cellDataType.Items.Add(sTemp); }

                //arrDataType = null;
                lstDataType = null;
                cDataTypes = null;
                cellDataType = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void bgwStatusUpdater_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ssStatus.BackColor = Color.Red;
                ssStatus.Refresh();
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void configHtmlSummary(ref ClsInsertCode_Files cInsertCode_Files)
        {
            try
            {
                cConfigReporter = new ClsConfigReporter();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId;
                int iRowId = 0;


                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, 2, "Auto generated code is located.");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Description";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_Files.moduleName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_Files.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //cInsertCode_Files.fileFormat

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 5 }, "Details");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Text file full path";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_Files.FullFilePath.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Direction";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                switch (cInsertCode_Files.direction) 
                {
                    case ClsInsertCode_Files.enumDirection.eDir_Read:
                        objCell.sText = "Read file";
                        objCell.sHiddenText = "";
                        break;
                    case ClsInsertCode_Files.enumDirection.eDir_Write:
                        objCell.sText = "Write file";
                        objCell.sHiddenText = "";
                        break;
                    case ClsInsertCode_Files.enumDirection.eDir_Unknown:
                        objCell.sText = "Unknown";
                        objCell.sHiddenText = "Should be either Read or Write.";
                        break;
                }
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                switch (cInsertCode_Files.fileFormat.eFileType)
                {
                    case ClsInsertCode_Files.enumFileType.eDelimitedFile:
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "File Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Delimited File";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Delimiter";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (cInsertCode_Files.fileFormat.cDelimiter == ' ')
                        { objCell.sText = "<Space>"; }
                        else if (cInsertCode_Files.fileFormat.cDelimiter == '\t')
                        { objCell.sText = "<Tab>"; }
                        else
                        { objCell.sText = cInsertCode_Files.fileFormat.cDelimiter.ToString(); }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        if (cInsertCode_Files.fileFormat.cDelimiter == ' ' || cInsertCode_Files.fileFormat.cDelimiter == '\t')
                        {
                            objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                            objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Maroon);
                        }

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        break;
                    case ClsInsertCode_Files.enumFileType.eFixedColumnLengthFile:
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "File Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Fixed length fields";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        
                        if (cInsertCode_Files.fileFormat.lstColumns.Count == 0)
                        {
                            //Add Row
                            cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Fields";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "No fields specified";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        }
                        else
                        {
                            /***************
                             *   A table   *
                             ***************/
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 1, 2 }, "Fields");

                            //Add Row
                            cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Name";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Datatype";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Start Position";
                            objCell.sHiddenText = "Number of charactors along line to the start of the field.";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Size";
                            objCell.sHiddenText = "Number of charactors in one line of text file for this field.";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);


                            foreach (ClsInsertCode_Files.strFileFormat_FixedColumn cField in cInsertCode_Files.fileFormat.lstColumns)
                            { 
                                //Add Row
                                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                                objCell.iOrder = 0;
                                objCell.bPropHtml = true;
                                objCell.sText = cField.sName.Trim();
                                objCell.sHiddenText = "";
                                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                                objCell.iOrder = 0;
                                objCell.bPropHtml = true;
                                objCell.sText = cDataTypes.getName(cField.eDataType);
                                objCell.sHiddenText = "";
                                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                                objCell.iOrder = 0;
                                objCell.bPropHtml = true;
                                objCell.sText = cField.iPosStart.ToString();
                                objCell.sHiddenText = "";
                                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                                objCell.iOrder = 0;
                                objCell.bPropHtml = true;
                                objCell.sText = cField.iSize.ToString();
                                objCell.sHiddenText = "";
                                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                            }
                        }
                        break;
                    case ClsInsertCode_Files.enumFileType.eUnknown:
                        break;
                    default:
                        break;
                }

                cDataTypes = null;
                objCell.lstFormatDetails = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void displayHtmlSummary()
        {
            try
            {
                string sHtml = cConfigReporter.getHtml();

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Text_File");

                frm.ShowDialog(this);

                frm = null;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private void FrmInsertCode_TextFile_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.B)
                    { browse(); }

                    if (e.KeyCode == Keys.G)
                    { generate(); }

                    if (e.KeyCode == Keys.C)
                    { this.Close(); }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
    }
}