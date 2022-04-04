using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_OpenForm : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_OpenForm()
        {
            try 
            {
                InitializeComponent();

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
                bool bIsOk = true;
                string sErrorMessage = "";
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsInsertCode_OpenForm cInsertCode_OpenForm = new ClsInsertCode_OpenForm();
                List<string> lstParameterNames = new List<string>();

                for (int iRow = 0; iRow < dgParameters.RowCount; iRow++)
                {
                    ClsInsertCode_OpenForm.strParameter objParameter = new ClsInsertCode_OpenForm.strParameter();

                    if (dgParameters[ColInternalVariable.Index, iRow].Value == null)
                    {
                        objParameter.sNamePrivatelyInForm = "";
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else if (string.IsNullOrEmpty(dgParameters[ColInternalVariable.Index, iRow].Value.ToString()))
                    {
                        objParameter.sNamePrivatelyInForm = "";
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else
                    { 
                        objParameter.sNamePrivatelyInForm = ClsMiscString.makeValidVarName(dgParameters[ColInternalVariable.Index, iRow].Value.ToString());
                        lstParameterNames.Add(objParameter.sNamePrivatelyInForm.ToString().ToLower().Trim());
                    }

                    if (dgParameters[ColExternalProperty.Index, iRow].Value == null)
                    { 
                        objParameter.sNamePublicOutsideForm = "";
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else if (string.IsNullOrEmpty(dgParameters[ColExternalProperty.Index, iRow].Value.ToString()))
                    {
                        objParameter.sNamePublicOutsideForm = "";
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else
                    { objParameter.sNamePublicOutsideForm = ClsMiscString.makeValidVarName(dgParameters[ColExternalProperty.Index, iRow].Value.ToString()); }

                    if (dgParameters[ColVariable.Index, iRow].Value == null)
                    { objParameter.bIsVariable = false; }
                    else
                    {
                        string sIsVariable = dgParameters[ColVariable.Index, iRow].Value.ToString();
                        bool bIsVariable = false;

                        if (!bool.TryParse(sIsVariable, out bIsVariable))
                        { bIsVariable = false; }

                        objParameter.bIsVariable = bIsVariable;
                    }

                    if (dgParameters[ColDataType.Index, iRow].Value == null)
                    { 
                        objParameter.eDataType = ClsDataTypes.vbVarType.vbUnknown;
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else if (string.IsNullOrEmpty(dgParameters[ColDataType.Index, iRow].Value.ToString()))
                    {
                        objParameter.eDataType = ClsDataTypes.vbVarType.vbUnknown;
                        bIsOk = false;
                        sErrorMessage = "Please make sure that each parameter has every field entered.";
                    }
                    else
                    { objParameter.eDataType = cDataTypes.getDataType(dgParameters[ColDataType.Index, iRow].Value.ToString()); }

                    if (bIsOk)
                    {
                        if (dgParameters[ColValue.Index, iRow].Value == null)
                        { 
                            objParameter.sValueGiveToParameter = "";
                            bIsOk = false;
                            sErrorMessage = "Please make sure that each parameter has every field entered.";
                        }
                        else if (string.IsNullOrEmpty(dgParameters[ColValue.Index, iRow].Value.ToString()))
                        {
                            objParameter.sValueGiveToParameter = "";
                            bIsOk = false;
                            sErrorMessage = "Please make sure that each parameter has every field entered.";
                        }
                        else
                        {
                            objParameter.sValueGiveToParameter = ClsMiscString.makeValidVarName(dgParameters[ColValue.Index, iRow].Value.ToString());
                            /*
                            if (objParameter.bIsVariable)
                            { objParameter.sValueGiveToParameter = ClsMiscString.makeValidVarName(dgParameters[ColValue.Index, iRow].Value.ToString()); }
                            else
                            {
                                switch (cDataTypes.getGeneralType(objParameter.eDataType))
                                {
                                    case ClsDataTypes.enumGeneralDateType.eBool:
                                    case ClsDataTypes.enumGeneralDateType.eNumber:
                                    case ClsDataTypes.enumGeneralDateType.eUnknown:
                                        objParameter.sValueGiveToParameter = dgParameters[ColValue.Index, iRow].Value.ToString();
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eDate:
                                        objParameter.sValueGiveToParameter = "#" + dgParameters[ColValue.Index, iRow].Value.ToString() + "#";
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eString:
                                        objParameter.sValueGiveToParameter = ClsMiscString.addQuotes(dgParameters[ColValue.Index, iRow].Value.ToString());
                                        break;
                                    default:
                                        objParameter.sValueGiveToParameter = dgParameters[ColValue.Index, iRow].Value.ToString();
                                        break;
                                }
                            }
                            */
                        }
                    }

                    if (bIsOk)
                    {
                        if (objParameter.sNamePrivatelyInForm == null || objParameter.sNamePublicOutsideForm == null || objParameter.sValueGiveToParameter == null)
                        {
                            bIsOk = false;
                            sErrorMessage = "Please make sure that each parameter has every field entered.";
                        } 
                        else if (string.IsNullOrEmpty(objParameter.sNamePrivatelyInForm.Trim()) || string.IsNullOrEmpty(objParameter.sNamePublicOutsideForm.Trim()) || string.IsNullOrEmpty(objParameter.sValueGiveToParameter.Trim()))
                        {
                            bIsOk = false;
                            sErrorMessage = "Please make sure that each parameter has every field entered.";
                        }
                    }

                    if (bIsOk)
                    { cInsertCode_OpenForm.addParameter(objParameter); }
                }

                if (lstParameterNames.Count > lstParameterNames.Distinct().Count())
                {
                    bIsOk = false;
                    sErrorMessage = "Please make sure all the Parameters have unique names";
                }

                if (bIsOk) 
                {
                    string sFormName;
                    bool bIsNewForm;

                    if (chkNewForm.Checked == null)
                    { bIsNewForm = false; }
                    else
                    { bIsNewForm = chkNewForm.Checked; }

                    cInsertCode_OpenForm.isNewForm = bIsNewForm;

                    if (bIsNewForm)
                    {
                        sFormName = txtFormName.Text;
                        sFormName = ClsMiscString.makeValidVarName(sFormName);

                        if (ClsMisc.isExistsForm(sFormName))
                        {
                            bIsOk = false;
                            sErrorMessage = "Form \"" + sFormName + "\"already exists, please select a different name";
                        }
                    }
                    else
                    { sFormName = cmbForms.Text; }

                    if (string.IsNullOrEmpty(sFormName.Trim()))
                    {
                        bIsOk = false;
                        sErrorMessage = "Please make sure you give the form a valid name.";
                    }
                    else
                    { cInsertCode_OpenForm.FormName = sFormName; }
                }

                if (bIsOk)
                {
                    if (txtInstanceName.Text == null)
                    { cInsertCode_OpenForm.InstanceName = ""; }
                    else
                    { cInsertCode_OpenForm.InstanceName = txtInstanceName.Text; }
                }
                
                cInsertCode_OpenForm.fixAmbiguousFieldNames();

                if (bIsOk)
                {
                    cInsertCode_OpenForm.openForm(ref cCodeMapper);
                    cInsertCode_OpenForm.addForm(ref cCodeMapper);

                    configHtmlSummary(ref cInsertCode_OpenForm);
                    displayHtmlSummary();

                    this.Close();
                }
                else
                { MessageBox.Show(this, sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

                cDataTypes = null;
                cInsertCode_OpenForm = null;
                lstParameterNames = null;
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

        private void FrmInsertCode_OpenForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref txtFormName);
                ClsDefaults.FormatControl(ref lblFormName);
                ClsDefaults.FormatControl(ref txtInstanceName);
                ClsDefaults.FormatControl(ref lblInstanceName);
                ClsDefaults.FormatControl(ref dgParameters);
                ClsDefaults.FormatControl(ref chkNewForm);
                ClsDefaults.FormatControl(ref cmbForms);

                ClsDefaults.FormatControl(ref ssStatus);

                chkNewForm.Checked = false;
                fillCmbForm();
                chkNewForm_Change();
                txtInstanceName.Text = ClsDefaults.defaultName;

                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(txtFormName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblFormName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtInstanceName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblInstanceName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(dgParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(chkNewForm, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbForms, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
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

        private void fillCmbForm() 
        {
            try 
            {
                cmbForms.Items.Clear();

                foreach (string sFormName in ClsMisc.listForms()) 
                { cmbForms.Items.Add(sFormName); }
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
                ClsDataTypes cDataTypes = new ClsDataTypes();
                DataGridViewComboBoxCell cellValues = (DataGridViewComboBoxCell)dgParameters[ColDataType.Index, e.RowIndex];

                cellValues.Items.Clear();

                foreach (string sTemp in cDataTypes.commonDataTypes()) 
                { cellValues.Items.Add(sTemp); }

                cDataTypes = null;
                cellValues = null;
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

        private void chkNewForm_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                chkNewForm_Change();
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

        private void chkNewForm_Change()
        {
            try {
                bool bNew;

                if (chkNewForm.Checked == null)
                { bNew = false; }
                else
                {
                    if (chkNewForm.Checked == true)
                    {bNew = true;} 
                    else 
                    {bNew = false;}
                }

                if (bNew)
                {
                    txtFormName.Visible = true;
                    cmbForms.Visible = false;
                }
                else
                {
                    txtFormName.Visible = false;
                    cmbForms.Visible = true;
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
                string sName = FrmInputBox.GetString("Name", "Please enter the name of the Variable");

                if (sName.Trim() != "")
                {
                    sName = ClsMiscString.makeValidVarName(sName);
                    
                    int iRowNew = dgParameters.Rows.Add();

                    dgParameters[ColInternalVariable.DisplayIndex, iRowNew].Value = sName;
                    dgParameters[ColExternalProperty.DisplayIndex, iRowNew].Value = sName;

                    this.ActiveControl = dgParameters;
                    dgParameters.CurrentCell = dgParameters[ColInternalVariable.Index, iRowNew];
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
                bool bIsOk = true;

                if (dgParameters.CurrentRow == null)
                {
                    MessageBox.Show(this, "Please select one row to remove.", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    int iRow = dgParameters.CurrentRow.Index;

                    DialogResult drAreYouSure = MessageBox.Show("Are you sure you want to remove that row?", ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (drAreYouSure == DialogResult.Yes)
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


        private void configHtmlSummary(ref ClsInsertCode_OpenForm cInsertCode_OpenForm)
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
                objCell.sText = cInsertCode_OpenForm.FormName; //ClsMisc.ActiveVBComponent().Name; //cInsertCode_CommandBarClass.className;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Form name.";
                if (cInsertCode_OpenForm.isNewForm)
                { objCell.sHiddenText = "New Form being opened by this code."; }
                else
                { objCell.sHiddenText = "Existing Form being opened by this code."; }
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_OpenForm.ModuleCallForm;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module where form is opened from.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_OpenForm.FunctionCallForm; // ClsMisc.ActiveVBComponent().Name; //cInsertCode_CommandBarClass.SampleCodeModulePrefix.Trim() + cInsertCode_CommandBarClass.className.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (cInsertCode_OpenForm.parameters.Count == 0)
                {
                    /***************
                     *   A table   *
                     ***************/
                    cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 4 }, "Details");

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameters";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "No parameters used.";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {
                    /************************
                     *   Parameters table   *
                     ************************/
                    cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 3, 1, 1 }, "Parameters");

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Name internally in the form.";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Name outside form";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Value assigned to parameter";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Datatype";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    foreach (ClsInsertCode_OpenForm.strParameter objParameter in cInsertCode_OpenForm.parameters.Distinct().OrderBy(x => x.sNamePublicOutsideForm))
                    {
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.sNamePrivatelyInForm;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.sNamePublicOutsideForm;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.sValueGiveToParameter;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cDataTypes.getName(objParameter.eDataType);
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Open_Form");

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

        private void dgParameters_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == ColVariable.Index)
                {
                    bool bAssignVariable;
                    int iRow = e.RowIndex;

                    if (iRow >= 0)
                    {
                        string sAssignVariable = "";
                        if (dgParameters[ColVariable.Index, iRow].Value == null)
                        { sAssignVariable = ""; }
                        else
                        { sAssignVariable = dgParameters[ColVariable.Index, iRow].Value.ToString(); }

                        if (!bool.TryParse(sAssignVariable, out bAssignVariable))
                        { bAssignVariable = false; }

                        if (bAssignVariable)
                        {
                            DataGridViewComboBoxCell ComboCellString = new DataGridViewComboBoxCell();

                            //List<string> lst = cCodeMapper.variableNames();

                            //lst.Sort();
                            ComboCellString.Items.Clear();

                            //foreach (string sVarName in lst)
                            //{ ComboCellString.Items.Add(sVarName); }

                            foreach (ClsCodeMapper.strVariables objTemp in cCodeMapper.lstVariablesInCurrentScope().OrderBy(x => x.sName))
                            { ComboCellString.Items.Add(objTemp.sName.Trim()); }


                            dgParameters[ColValue.Index, iRow] = ComboCellString;

                            //guess the data type column from the varibles data type

                            //ClsMisc.
                        }
                        else
                        {
                            DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                            dgParameters[ColValue.Index, iRow] = TextCellString;
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

        private void FrmInsertCode_OpenForm_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref txtFormName);
                cControlPosition.positionControl(ref lblFormName);
                cControlPosition.positionControl(ref txtInstanceName);
                cControlPosition.positionControl(ref lblInstanceName);
                cControlPosition.positionControl(ref dgParameters);
                cControlPosition.positionControl(ref chkNewForm);
                cControlPosition.positionControl(ref cmbForms);
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnRemove);
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

        private void FrmInsertCode_OpenForm_KeyDown(object sender, KeyEventArgs e)
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
