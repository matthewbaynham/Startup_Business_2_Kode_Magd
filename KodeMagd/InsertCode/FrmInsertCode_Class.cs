using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using Office = Microsoft.Office.Core;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_Class : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_Class()
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

        private void FrmInsertCode_Class_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;
                
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnRemove);

                ClsDefaults.FormatControl(ref dgProperties);

                ClsDefaults.FormatControl(ref lblClassName);
                ClsDefaults.FormatControl(ref txtClassName);

                ClsDefaults.FormatControl(ref chkSampleCodeInNewModule);

                ClsDefaults.FormatControl(ref ssStatus);

                chkSampleCodeInNewModule.Checked = true;
                string sClassNamePrefix = "Cls";
                string sClassName = sClassNamePrefix;
                int iAttempts = 1;

                while (ClsMisc.moduleExists(ClsMiscString.makeValidVarName(sClassName)))
                {
                    iAttempts++;
                    sClassName = sClassNamePrefix + iAttempts.ToString();
                }
                txtClassName.Text = sClassName;

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                
                cControlPosition.setControl(dgProperties, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                
                cControlPosition.setControl(lblClassName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtClassName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkSampleCodeInNewModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

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

        private void fillLstParameters(int iRow) 
        { 
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();

                DataGridViewComboBoxCell cmbType = (DataGridViewComboBoxCell)dgProperties[colDataType.Index, iRow];
                
                cmbType.Items.Clear();
                
                foreach (string sType in cDataTypes.commonDataTypes()) 
                { cmbType.Items.Add(sType); }

                cDataTypes = null;
                cmbType = null;
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
                generateCode();
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

        private void generateCode()
        {
            try
            {
                bool bReadOnly = false;
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsInsertCode_Class cClass = new ClsInsertCode_Class();
                ClsDataTypes.vbVarType eType = new ClsDataTypes.vbVarType();
                
                foreach (DataGridViewRow objRow in dgProperties.Rows)
                {
                    int iRowIndex = objRow.Index;
                    string sName = "";
                    string sReadOnly = "FALSE";
                    string sType = "";
                    string sDefaultValue = "";

                    if (dgProperties[ColName.Index, iRowIndex].Value != null)
                    { sName = ClsMiscString.makeValidVarName(dgProperties[ColName.Index, iRowIndex].Value.ToString().Trim()); }

                    string sTemp = "";
                    if (!ClsMisc.validVariableNameCheck(sName, out sTemp))
                    {
                        bIsOk = false;
                        { sErrorMessage += "\"" + sName + "\": " + sTemp + Environment.NewLine; }
                    }

                    if (dgProperties[colReadOnly.Index, iRowIndex].Value != null)
                    { sReadOnly = dgProperties[colReadOnly.Index, iRowIndex].Value.ToString(); }

                    if (dgProperties[colDataType.Index, iRowIndex].Value != null)
                    { sType = dgProperties[colDataType.Index, iRowIndex].Value.ToString(); }

                    if (dgProperties[colDefaultValue.Index, iRowIndex].Value != null)
                    { sDefaultValue = dgProperties[colDefaultValue.Index, iRowIndex].Value.ToString(); }

                    switch (sReadOnly.Trim().ToUpper())
                    {
                        case "TRUE":
                            bReadOnly = true;
                            break;
                        case "FALSE":
                            bReadOnly = false;
                            break;
                        default:
                            sErrorMessage += "true or false Value expected in Boolean Column." + Environment.NewLine;
                            bIsOk = false;
                            break;
                    }

                    eType = cDataTypes.getDataType(sType);

                    if (eType == ClsDataTypes.vbVarType.vbUnknown) 
                    {
                        bIsOk = false;
                        sErrorMessage += "Unknown Datatype." + Environment.NewLine;
                    }

                    if (bIsOk)
                    { cClass.addParameter(sName, bReadOnly, eType, sDefaultValue);}
                }


                if (cClass.hasDupicateParameters) 
                {
                    bIsOk = false;
                    sErrorMessage += "Dupicate parameter names." + Environment.NewLine;
                }


                cClass.PutSampleCallInOwnNewMod = chkSampleCodeInNewModule.Checked;

                if (txtClassName.Text == null)
                { cClass.className = ""; }
                else
                { cClass.className = txtClassName.Text.Trim(); }

                string sTemp2 = "";
                if (!ClsMisc.validVariableNameCheck(cClass.className, out sTemp2))
                {
                    bIsOk = false;
                    sErrorMessage += "Class Name: " + sTemp2 + Environment.NewLine;
                }

                cClass.PutSampleCallInOwnNewMod = chkSampleCodeInNewModule.Checked;


                if (ClsMisc.isExistsModule(cClass.className))
                {
                    bIsOk = false;
                    sErrorMessage += "There is already a module called " + cClass.className + Environment.NewLine;
                }

                if (cClass.PutSampleCallInOwnNewMod)
                {
                    cClass.SampleModuleName = ClsMiscString.nextFunctionName(ref cCodeMapper, ClsDefaults.sampleCodeModulePrefix + cClass.className);

                    if (ClsMisc.isExistsModule(cClass.SampleModuleName))
                    {
                        bIsOk = false;
                        sErrorMessage += "(For the code sample) there is already a module called " + cClass.SampleModuleName + Environment.NewLine;
                    }
                }

                if (bIsOk) 
                {
                    configHtmlSummary(ref cClass);

                    cClass.generateSampleCode();
                    cClass.generateClass();

                    displayHtmlSummary();

                    this.Close();
                }
                else
                {
                    sErrorMessage += Environment.NewLine + "Please resolve these issues.";
                    MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); 
                }

                cDataTypes = null;
                cClass = null;
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

        private void dgProperties_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                fillLstParameters(e.RowIndex);
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
                addProperty();
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

        private void addProperty()
        {
            try
            {
                string sNewItem = FrmInputBox.GetString("Property Name", "Please enter the property name you wish to add.");

                if (sNewItem.Trim() != "")
                {
                    int iRow = dgProperties.Rows.Add();

                    dgProperties[ColName.Index, iRow].Value = sNewItem;

                    this.ActiveControl = dgProperties;
                    dgProperties.CurrentCell = dgProperties[ColName.Index, iRow];
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
                removeProperty();
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

        private void removeProperty()
        {
            try
            {
                if (dgProperties.CurrentRow != null)
                {
                    int iRow = dgProperties.CurrentRow.Index;

                    string sQuestion = "Are you sure you want to remove " + dgProperties[ColName.Index, iRow].Value.ToString().Trim() + "?";
                    DialogResult dlgRemove = MessageBox.Show(sQuestion, ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dlgRemove == System.Windows.Forms.DialogResult.Yes)
                    { dgProperties.Rows.RemoveAt(iRow); }
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

        private void configHtmlSummary(ref ClsInsertCode_Class cClass) 
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
                objCell.sText = cClass.className;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Class";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cClass.SampleModuleName.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Sample Module Name";
                objCell.sHiddenText = "Sample code is an example of working with the Class";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 6, 1, 1, 3 }, "Parameters");

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
                objCell.sText = "Data Type";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Read only";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Default Value";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                foreach (ClsInsertCode_Class.strParameter objParameter in cClass.parameters)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objParameter.sNameExtermal;
                    objCell.sHiddenText = "Inside the class the value is stored in the variable: " + objParameter.sNameInternal;
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = cDataTypes.getName(objParameter.eType);
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    if (objParameter.bReadOnly)
                    { objCell.sText = "Read Only"; }
                    else
                    { objCell.sText = "Read/Write"; }
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objParameter.sDefaultValue;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Class");

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

        private void FrmInsertCode_Class_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnRemove);
                
                cControlPosition.positionControl(ref dgProperties);
                
                cControlPosition.positionControl(ref lblClassName);
                cControlPosition.positionControl(ref txtClassName);
                cControlPosition.positionControl(ref chkSampleCodeInNewModule);
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

        private void FrmInsertCode_Class_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { addProperty(); }

                    if (e.KeyCode == Keys.R)
                    { removeProperty(); }

                    if (e.KeyCode == Keys.G)
                    { generateCode(); }

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
