using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Misc;
using KodeMagd.Settings;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_TextFileLogClass : Form
    {
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_TextFileLogClass()
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

        private void FrmInsertCode_TextFileLogClass_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref lblDateFormatFileContents);
                ClsDefaults.FormatControl(ref cmbDateFormatFileContents);

                ClsDefaults.FormatControl(ref lblDateFormatFileName);
                ClsDefaults.FormatControl(ref cmbDateFormatFileName);

                ClsDefaults.FormatControl(ref lblClassName);
                ClsDefaults.FormatControl(ref txtClassName);

                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref chkAddReferences);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref txtDestinationFolder);
                ClsDefaults.FormatControl(ref lblDestinationFolder, ClsDefaults.enumLabelState.eLbl_normal);

                ClsDefaults.FormatControl(ref grpFileName);
                ClsDefaults.FormatControl(ref optFileNameAuto);
                ClsDefaults.FormatControl(ref optFileNameSpecified);

                ClsDefaults.FormatControl(ref dgParameters);

                ClsDefaults.FormatControl(ref ssStatus);


                cControlPosition.setControl(lblDateFormatFileContents, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbDateFormatFileContents, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblDateFormatFileName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbDateFormatFileName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblClassName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtClassName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(chkAddReferences, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(txtDestinationFolder, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblDestinationFolder, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(grpFileName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optFileNameAuto, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optFileNameSpecified, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(dgParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                string sClassNamePrefix = "ClsLog";
                string sClassName = sClassNamePrefix;
                int iAttempts = 1;
                while (ClsMisc.moduleExists(ClsMiscString.makeValidVarName(sClassName))) 
                {
                    iAttempts++;
                    sClassName = sClassNamePrefix + iAttempts.ToString();
                }
                txtClassName.Text = sClassName;
                txtDestinationFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                optFileNameAuto.Checked = true;
                optFileNameSpecified.Checked = false;

                foreach (string sTemp in ClsMisc.CommonDateFormats().Distinct().OrderBy(x => x))
                {
                    cmbDateFormatFileName.Items.Add(sTemp.Trim());
                    cmbDateFormatFileName.Text = sTemp.Trim();
                }

                foreach (string sTemp in ClsMisc.CommonDateFormats().Distinct().OrderBy(x => x))
                {
                    cmbDateFormatFileContents.Items.Add(sTemp.Trim());
                    cmbDateFormatFileContents.Text = sTemp.Trim();
                }


                VBA.VBComponent VBComp = ClsMisc.ActiveVBComponent();

                if (VBComp != null)
                {
                    cCodeMapper = new ClsCodeMapper();
                    cCodeMapper.readCode(ClsMisc.ActiveVBComponent());
                }
                else
                { cCodeMapper = null; }

                VBComp = null;

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

        private void btnClose_Click(object sender, EventArgs e)
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
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();
                ClsInsertCode_TextFileLogClass cText = new ClsInsertCode_TextFileLogClass();

                string sPath = txtDestinationFolder.Text;
                bool bIsOK = true;
                string sErrorMessage = "";
                string sClsName = "";

                if (bIsOK)
                {
                    if (optFileNameAuto.Checked == true)
                    {
                        if (!fso.FolderExists(sPath))
                        {
                            sErrorMessage = "Destination Folder is not found";
                            bIsOK = false;
                        }

                        if (!bIsOK)
                        { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                    }
                    else
                    {
                        string sPathFolder = ClsMisc.GetFolder(sPath);

                        if (!fso.FolderExists(sPathFolder))
                        {
                            sErrorMessage = "Destination Folder is not found";
                            bIsOK = false;
                        }

                        if (fso.FileExists(sPath))
                        {
                            sErrorMessage = "There is already a file in that location";
                            bIsOK = false;
                        }
                    }
                }

                sClsName = ClsMiscString.makeValidVarName(txtClassName.Text);

                if (ClsMisc.moduleExists(sClsName))
                {
                    bIsOK = false;
                    sErrorMessage = "The Module Name already exists";
                }

                if (string.IsNullOrEmpty(sClsName.Trim())) 
                {
                    sErrorMessage = "Please enter a valid Class Name";
                    bIsOK = false;
                }
                
                if (bIsOK)
                {
                    for (int lRow = 0; lRow < dgParameters.RowCount; lRow++) 
                    {
                        ClsInsertCode_TextFileLogClass.strVariablesToLog objPara = new ClsInsertCode_TextFileLogClass.strVariablesToLog();

                        if (string.IsNullOrEmpty(dgParameters[colName.Index, lRow].Value.ToString().Trim()))
                        { objPara.sName = ""; }
                        else
                        { objPara.sName = ClsMiscString.makeValidVarName(dgParameters[colName.Index, lRow].Value.ToString()); }

                        if (string.IsNullOrEmpty(objPara.sName.Trim()))
                        {
                            sErrorMessage = "Parameters have to have a name";
                            bIsOK = false;
                        }

                        if (string.IsNullOrEmpty(dgParameters[colDataType.Index, lRow].Value.ToString().Trim()))
                        { 
                            sErrorMessage = "No data type was selected";
                            bIsOK = false;
                        }
                        else
                        { objPara.sDataType = dgParameters[colDataType.Index, lRow].Value.ToString(); }

                        DataGridViewCheckBoxCell objTemp = (DataGridViewCheckBoxCell)dgParameters[ColOptional.Index, lRow];

                        if (objTemp.Value == null) 
                        { objPara.bOptional = false; }
                        else
                        {
                            if ((bool)objTemp.Value)
                            { objPara.bOptional = true; }
                            else
                            { objPara.bOptional = false; }
                        }

                        cText.addParameter(objPara);
                    }
                }

                if (bIsOK)
                {
                    if (chkAddReferences.Checked)
                    {
                        FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_Scripting, ref ssStatus);

                        if (!frmReference.referenceAlreadySet)
                        { frmReference.ShowDialog(this); }

                        frmReference = null;
                    }

                    cText.Path = txtDestinationFolder.Text;
                    cText.ClassName = sClsName;
                    //cText.moduleNameCallLog = cCodeMapper.ModuleDetails.sName;
                    cText.DateFormat_FileContents = cmbDateFormatFileContents.Text;
                    cText.DateFormat_FileName = cmbDateFormatFileName.Text;

                    cText.CallLog(ref cCodeMapper);
                    cText.generateClass(ref cCodeMapper);

                    configHtmlSummary(ref cText);
                    displayHtmlSummary();
                    cText = null;

                    fso = null;
                    cText = null;

                    this.Close();
                }
                else
                {
                    fso = null;
                    cText = null;
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

        private void dgParameters_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try 
            {
                DataGridViewComboBoxCell cellValues = (DataGridViewComboBoxCell)dgParameters[colDataType.Index, e.RowIndex];
                ClsDataTypes cDataTypes = new ClsDataTypes();

                cellValues.Items.Clear();

                foreach (string sVariable in cDataTypes.commonDataTypes())
                { cellValues.Items.Add(sVariable); }

                cellValues = null;
                cDataTypes = null;
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

        private void optFileNameAuto_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                whatsVisible();
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

        private void optFileNameSpecified_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                whatsVisible();
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

        private void whatsVisible()
        {
            try
            {
                if (optFileNameAuto.Checked)
                {
                    lblDateFormatFileName.Visible = true;
                    cmbDateFormatFileName.Visible = true;
                }
                else
                {
                    lblDateFormatFileName.Visible = false;
                    cmbDateFormatFileName.Visible = false;
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

        private void FrmInsertCode_TextFileLogClass_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblDateFormatFileContents);
                cControlPosition.positionControl(ref cmbDateFormatFileContents);

                cControlPosition.positionControl(ref lblDateFormatFileName);
                cControlPosition.positionControl(ref cmbDateFormatFileName);

                cControlPosition.positionControl(ref lblClassName);
                cControlPosition.positionControl(ref txtClassName);

                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnRemove);
                cControlPosition.positionControl(ref chkAddReferences);
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);

                cControlPosition.positionControl(ref txtDestinationFolder);
                cControlPosition.positionControl(ref lblDestinationFolder);

                cControlPosition.positionControl(ref grpFileName);
                cControlPosition.positionControl(ref optFileNameAuto);
                cControlPosition.positionControl(ref optFileNameSpecified);

                cControlPosition.positionControl(ref dgParameters);
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

        private void btnAdd_Click(object sender, EventArgs e)
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
                string sNewItem = FrmInputBox.GetString("Parameter Name", "Please enter the parameter name you wish to add.");

                if (sNewItem.Trim() != "")
                {
                    int iRow = dgParameters.Rows.Add();

                    dgParameters[colName.Index, iRow].Value = sNewItem;

                    this.ActiveControl = dgParameters;
                    dgParameters.CurrentCell = dgParameters[colName.Index, iRow];
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

        private void btnRemove_Click(object sender, EventArgs e)
        {
            try
            {
                remove();
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

        private void remove()
        {
            try
            {
                if (dgParameters.CurrentRow != null)
                {
                    int iRow = dgParameters.CurrentRow.Index;

                    string sQuestion = "Are you sure you want to remove " + dgParameters[colName.Index, iRow].Value.ToString().Trim() + "?";
                    DialogResult dlgRemove = MessageBox.Show(sQuestion, ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dlgRemove == System.Windows.Forms.DialogResult.Yes)
                    { dgParameters.Rows.RemoveAt(iRow); }
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

        private void configHtmlSummary(ref ClsInsertCode_TextFileLogClass cText)
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId;
                int iRowId = 0;

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 2 }, "Auto generated code is located.");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

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
                objCell.sText = cText.ClassName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Class Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cText.moduleNameCallLog;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module where the code class is called from.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cText.functionNameCallLog;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function / Sub / Property where the code class is called from";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cText.Path;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Path of text log file";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (cText.variablesToLog.Count == 0)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "No variables to log";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Variables";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {
                    /***************
                     *   A table   *
                     ***************/
                    cConfigReporter.TableAddNew(out iTableId, new List<int> { 4, 1, 1 }, "Variables to log");

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
                    objCell.sText = "Is Optional?";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Data Type";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    foreach (ClsInsertCode_TextFileLogClass.strVariablesToLog objVar in cText.variablesToLog)
                    {
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objVar.sName;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (objVar.bOptional)
                        { objCell.sText = "Optional"; }
                        else
                        { objCell.sText = ""; }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objVar.sDataType;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }
                }

                //cDataTypes = null;
                //objCell.lstFormatDetails = null;
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Text File Log Class");

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

        private void FrmInsertCode_TextFileLogClass_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.R)
                    { remove(); }

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
