using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using KodeMagd.Misc;
using System.Reflection;
using KodeMagd.Settings;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_FileExists : Form
    {
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_FileExists()
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

        private void FrmInsertCode_FileExists_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref lblPath);
                ClsDefaults.FormatControl(ref txtPath);
                ClsDefaults.FormatControl(ref btnBrowse);

                ClsDefaults.FormatControl(ref lblName);
                ClsDefaults.FormatControl(ref txtName);

                ClsDefaults.FormatControl(ref grpType);
                ClsDefaults.FormatControl(ref optHardcoded);
                ClsDefaults.FormatControl(ref optVariable);

                ClsDefaults.FormatControl(ref lblVariable);
                ClsDefaults.FormatControl(ref cmbVariable);

                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref chkAddReference);


                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(lblPath, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtPath, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnBrowse, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(grpType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optHardcoded, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optVariable, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblVariable, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbVariable, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(chkAddReference, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cCodeMapper.readCode();

                fillCmbVariableNames();
                optVariable.Checked = false;
                optHardcoded.Checked = true;
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

        private void FrmInsertCode_FileExists_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblPath);
                cControlPosition.positionControl(ref txtPath);
                cControlPosition.positionControl(ref btnBrowse);

                cControlPosition.positionControl(ref lblName);
                cControlPosition.positionControl(ref txtName);

                cControlPosition.positionControl(ref grpType);
                cControlPosition.positionControl(ref optHardcoded);
                cControlPosition.positionControl(ref optVariable);

                cControlPosition.positionControl(ref lblVariable);
                cControlPosition.positionControl(ref cmbVariable);

                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref chkAddReference);
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

        private void optVariable_CheckedChanged(object sender, EventArgs e)
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

        private void optHardcoded_CheckedChanged(object sender, EventArgs e)
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
                bool bIsHardcoded;
                bool bIsVariable;

                if (optHardcoded.Checked == null)
                { bIsHardcoded = false; }
                else
                { bIsHardcoded = optHardcoded.Checked;}

                if (optVariable.Checked == null)
                { bIsVariable = false; }
                else
                { bIsVariable = optVariable.Checked; }

                if (!bIsHardcoded && bIsVariable)
                {
                    lblPath.Visible = false;
                    txtPath.Visible = false;
                    btnBrowse.Visible = false;

                    lblVariable.Visible = true;
                    cmbVariable.Visible = true;
                }
                else if (bIsHardcoded && !bIsVariable)
                {
                    lblPath.Visible = true;
                    txtPath.Visible = true;
                    btnBrowse.Visible = true;

                    lblVariable.Visible = false;
                    cmbVariable.Visible = false;
                }
                else
                {
                    lblPath.Visible = false;
                    txtPath.Visible = false;
                    btnBrowse.Visible = false;

                    lblVariable.Visible = false;
                    cmbVariable.Visible = false;
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

                DialogResult result = ofdBrowseOpen.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    sFullPath = ofdBrowseOpen.FileName;
                    txtPath.Text = sFullPath;
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
                ClsInsertCode_FileExists cInsertCode_FileExists = new ClsInsertCode_FileExists();
                string sErrorMessage = "";
                bool bIsOk = true;

                if (chkAddReference.Checked)
                {
                    FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_Scripting, ref ssStatus);

                    if (!frmReference.referenceAlreadySet)
                    { frmReference.ShowDialog(this); }

                    frmReference = null;
                }

                if (optHardcoded.Checked == false && optVariable.Checked == true)
                { cInsertCode_FileExists.type = ClsInsertCode_FileExists.enumType.eTyp_Variable; }
                else if (optHardcoded.Checked == true && optVariable.Checked == false)
                { cInsertCode_FileExists.type = ClsInsertCode_FileExists.enumType.eTyp_HardCoded; }
                else
                { 
                    cInsertCode_FileExists.type = ClsInsertCode_FileExists.enumType.eTyp_Unknown;
                    bIsOk = false;
                    sErrorMessage = "";
                }

                switch (cInsertCode_FileExists.type)
                {
                    case ClsInsertCode_FileExists.enumType.eTyp_HardCoded:
                        cInsertCode_FileExists.variableName = "";
                        if (txtPath.Text == null)
                        { cInsertCode_FileExists.path = ""; }
                        else
                        { cInsertCode_FileExists.path = txtPath.Text; }
                        break;
                    case ClsInsertCode_FileExists.enumType.eTyp_Variable:
                        cInsertCode_FileExists.path = "";
                        if (cmbVariable.Text == null)
                        { cInsertCode_FileExists.variableName = ""; }
                        else
                        { cInsertCode_FileExists.variableName = cmbVariable.Text; }
                        break;
                    case ClsInsertCode_FileExists.enumType.eTyp_Unknown:
                        break;
                }

                if (bIsOk)
                {
                    cInsertCode_FileExists.generateCode(ref cCodeMapper);

                    configHtmlSummary(ref cInsertCode_FileExists);
                    displayHtmlSummary();

                    this.Close();
                }
                else
                {
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

        private void configHtmlSummary(ref ClsInsertCode_FileExists cInsertCode_FileExists)
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
                objCell.sText = cInsertCode_FileExists.moduleName;
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
                objCell.sText = cInsertCode_FileExists.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 5 }, "Details");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Type";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                switch (cInsertCode_FileExists.type)
                {
                    case ClsInsertCode_FileExists.enumType.eTyp_HardCoded:
                        objCell.sText = "Hard Coded Path";
                        objCell.sHiddenText = "";
                        break;
                    case ClsInsertCode_FileExists.enumType.eTyp_Variable:
                        objCell.sText = "Variable Name for Path";
                        objCell.sHiddenText = "";
                        break;
                    default:
                        objCell.sText = "Unknown";
                        objCell.sHiddenText = "Please check the correct options have been selected.";
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Bold);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);
                        break;
                }

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                switch (cInsertCode_FileExists.type)
                {
                    case ClsInsertCode_FileExists.enumType.eTyp_HardCoded:
                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Path";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_FileExists.path;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        break;
                    case ClsInsertCode_FileExists.enumType.eTyp_Variable:
                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Variable";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_FileExists.variableName;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "File_Exists");

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

        private void fillCmbVariableNames() 
        {
            try
            {
                foreach (ClsCodeMapper.strVariables objTemp in cCodeMapper.lstVariablesInCurrentScope().OrderBy(x => x.sName))
                { cmbVariable.Items.Add(objTemp.sName.Trim()); }
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

        private void FrmInsertCode_FileExists_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
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
