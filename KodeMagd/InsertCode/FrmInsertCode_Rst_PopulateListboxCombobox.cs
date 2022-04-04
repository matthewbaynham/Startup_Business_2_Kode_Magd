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
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Misc;
using System.Diagnostics;
using KodeMagd.Settings;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_Rst_PopulateListboxCombobox : Form
    {
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsControlPosition cControlPosition = new ClsControlPosition();

        public enum enumControlType 
        {
            enumCtrlType_Listbox,
            enumCtrlType_Combobox,
            enumCtrlType_Unknown
        }

        private ClsInsertCode_PopulateListboxCombobox.enumSource eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Array;
        private List<ClsInsertCode_PopulateListboxCombobox.strParameter> lstParameters = new List<ClsInsertCode_PopulateListboxCombobox.strParameter>();

        public enumControlType ControlType 
        {
            get 
            {
                try
                {
                    enumControlType eType = enumControlType.enumCtrlType_Unknown;

                    if (!optComboBox.Checked & optListBox.Checked)
                    { eType = enumControlType.enumCtrlType_Listbox; }

                    if (optComboBox.Checked & !optListBox.Checked)
                    { eType = enumControlType.enumCtrlType_Combobox; }

                    return eType;
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
                    
                    return enumControlType.enumCtrlType_Unknown;
                }
            }
            set 
            {
                try
                {
                    switch (value)
                    {
                        case enumControlType.enumCtrlType_Combobox:
                            optComboBox.Checked = true;
                            optListBox.Checked = false;
                            break;
                        case enumControlType.enumCtrlType_Listbox:
                            optComboBox.Checked = false;
                            optListBox.Checked = true;
                            break;
                        case enumControlType.enumCtrlType_Unknown:
                            optComboBox.Checked = false;
                            optListBox.Checked = false;
                            break;
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

        public string controlName 
        {
            get {
                try
                {
                    string sTemp;

                    if (string.IsNullOrEmpty(cmbControl.Text))
                    { sTemp = ""; }
                    else
                    { sTemp = cmbControl.Text; }

                    return sTemp;
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

                    return string.Empty;
                }
            }
            set 
            {
                try
                {
                    cmbControl.Text = value;
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

        public string fieldName
        {
            get
            {
                try
                {
                    string sTemp;

                    if (string.IsNullOrEmpty(txtOneValue.Text))
                    { sTemp = ""; }
                    else
                    { sTemp = txtOneValue.Text; }

                    return sTemp;
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

                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    txtOneValue.Text = value;
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

        public FrmInsertCode_Rst_PopulateListboxCombobox()
        {
            try
            {
                InitializeComponent();

                cCodeMapper = new ClsCodeMapper();
                cCodeMapper.readCode(ClsMisc.ActiveVBComponent());

                this.BackColor = ClsDefaults.FormColour; 
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

        private void FrmInsertCode_PopulateListboxCombobox_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref lblRstSql);
                ClsDefaults.FormatControl(ref txtRstSql);
                ClsDefaults.FormatControl(ref lblRstFieldName);
                ClsDefaults.FormatControl(ref txtRstFieldName);
                ClsDefaults.FormatControl(ref lblRstTableName);
                ClsDefaults.FormatControl(ref cmbRstTableName);
                ClsDefaults.FormatControl(ref lblRstConnectionString);
                ClsDefaults.FormatControl(ref txtRstConnectionString);
                ClsDefaults.FormatControl(ref lblRstCommandType);
                ClsDefaults.FormatControl(ref cmbRstCommandType);
                ClsDefaults.FormatControl(ref btnRstSqlExpand);
                ClsDefaults.FormatControl(ref btnRstParameters);
                ClsDefaults.FormatControl(ref btnRstConnectionStringExpand);
                ClsDefaults.FormatControl(ref btnRstConnectionStringBuild);
                ClsDefaults.FormatControl(ref btnRstConnectionStringRecent);

                ClsDefaults.FormatControl(ref lblOneValue);
                ClsDefaults.FormatControl(ref txtOneValue);

                ClsDefaults.FormatControl(ref btnArrayAdd);
                ClsDefaults.FormatControl(ref btnArrayRemove);
                ClsDefaults.FormatControl(ref lblArray);
                ClsDefaults.FormatControl(ref lstArray);

                ClsDefaults.FormatControl(ref lblRangeNamed);
                ClsDefaults.FormatControl(ref cmbRangeNamed);

                ClsDefaults.FormatControl(ref lblRangeAddressSheetName);
                ClsDefaults.FormatControl(ref cmbRangeAddressSheetName);
                ClsDefaults.FormatControl(ref lblRangeAddressAddress);
                ClsDefaults.FormatControl(ref txtRangeAddressAddress);

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref chkAddReference);

                ClsDefaults.FormatControl(ref grpSource);
                ClsDefaults.FormatControl(ref grpType);

                ClsDefaults.FormatControl(ref optArray);
                ClsDefaults.FormatControl(ref optComboBox);
                ClsDefaults.FormatControl(ref optListBox);
                ClsDefaults.FormatControl(ref optOneValue);
                ClsDefaults.FormatControl(ref optRangeAddress);
                ClsDefaults.FormatControl(ref optRangeNamed);
                ClsDefaults.FormatControl(ref optRecordset);

                ClsDefaults.FormatControl(ref lblControl);
                ClsDefaults.FormatControl(ref cmbControl);

                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(lblRstSql, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtRstSql, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(lblRstFieldName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(txtRstFieldName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(lblRstTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbRstTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblRstConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtRstConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblRstCommandType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbRstCommandType, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnRstSqlExpand, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnRstParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnRstConnectionStringExpand, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnRstConnectionStringBuild, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnRstConnectionStringRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblOneValue, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtOneValue, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnArrayAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnArrayRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblArray, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstArray, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(cmbRangeNamed, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblRangeAddressSheetName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbRangeAddressSheetName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblRangeAddressAddress, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtRangeAddressAddress, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(chkAddReference, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(grpSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(grpType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(optArray, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optComboBox, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optListBox, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optOneValue, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRangeAddress, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRangeNamed, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRecordset, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblControl, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbControl, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                //cControlPosition.setControl(lblFieldName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                //cControlPosition.setControl(txtFieldName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                if (ClsMisc.ActiveVBComponent().Type == Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm)
                {
                    optComboBox.Checked = false;
                    optListBox.Checked = true;

                    //fillFormCombo();
                    fillControlsCombo();
                    //fillValuesCombo();
                    fillCmbRstCommandType();
                    fillCmbRstTableName();
                    fillCmbRangeNamed();
                    fillCmbRangeAddressSheetName();

                    optRecordset.Checked = false;
                    optArray.Checked = true;
                    optOneValue.Checked = false;
                    optRangeNamed.Checked = false;
                    optRangeAddress.Checked = false;
                    chkAddReference.Checked = false;

                    cCodeMapper.readCode();
                }
                else 
                {
                    AllControlsVisibility(false);
                    MessageBox.Show("The active code window is not a form.\n\rPlease select a form that contains the correct controls.",
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    this.Close();
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

        private void fillControlsCombo() 
        {
            try
            {
                bool bIsCombo;
                bool bIsList;
                bool bIsOk = true;
                string sErrorMessage = "";

                List<string> lstControls = new List<string>();

                if (optComboBox.Checked)
                { bIsCombo = true; }
                else
                { bIsCombo = false; }

                if (optListBox.Checked)
                { bIsList = true; }
                else
                { bIsList = false; }

                if (bIsCombo == true & bIsList == true)
                {
                    sErrorMessage = "You can't select both ComboBox and ListBox it's one or the other";
                    bIsOk = false;
                }
                else if (bIsCombo == false & bIsList == false)
                {
                    sErrorMessage = "You have to select either ComboBox or ListBox";
                    bIsOk = false;
                }
                else
                {
                    bIsOk = true;
                    if (bIsCombo == true & bIsList == false)
                    { lstControls = ClsMisc.getComboBoxesNames(); }
                    if (bIsCombo == false & bIsList == true)
                    { lstControls = ClsMisc.getListboxesNames(); }

                    lstControls.Sort();

                    cmbControl.Items.Clear();
                    foreach (string sTemp in lstControls)
                    { cmbControl.Items.Add(sTemp); }
                }

                if (!bIsOk)
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

                lstControls = null;
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

        private void optListBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optListBox.Checked == true)
                { optComboBox.Checked = false; }
                else
                { optComboBox.Checked = true; }

                fillControlsCombo();
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

        private void optComboBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (optComboBox.Checked == true)
                { optListBox.Checked = false; }
                else
                { optListBox.Checked = true; }

                fillControlsCombo();
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

        private void FrmInsertCode_PopulateListboxCombobox_FormClosing(object sender, FormClosingEventArgs e)
        {
            try 
            {
                cCodeMapper = null;
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

        private void FrmInsertCode_Rst_PopulateListboxCombobox_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { arrayAdd(); }

                    if (e.KeyCode == Keys.M)
                    { arrayRemove(); }

                    if (e.KeyCode == Keys.B)
                    { build(); }

                    if (e.KeyCode == Keys.R)
                    { recent(); }

                    if (e.KeyCode == Keys.P)
                    { lstParameters = FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters.GetParameters("", lstParameters); }

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
                switch (eSourceType)
                {
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Array:
                        generateArray();
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_OneValue:
                        generateOneValue();
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeByAddress:
                        generateRangeAddressed();
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeNamed:
                        generateRangeNamed();
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Recordset:
                        generateRecordSet();
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Unknown:


                        break;
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

        private void generateArray()
        {
            try
            {
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsInsertCode_PopulateListboxCombobox cInsertCode_PopulateListboxCombobox = new ClsInsertCode_PopulateListboxCombobox();

                cInsertCode_PopulateListboxCombobox.ControlName = cmbControl.Text;
                cInsertCode_PopulateListboxCombobox.sourceType = eSourceType;

                cInsertCode_PopulateListboxCombobox.arrayClear();
                foreach(object objItems in lstArray.Items)
                { cInsertCode_PopulateListboxCombobox.arrayAdd(objItems.ToString()); }

                if (bIsOk == true)
                {
                    cInsertCode_PopulateListboxCombobox.generateCode(ref cCodeMapper);
                    configHtmlSummary(ref cInsertCode_PopulateListboxCombobox);
                    displayHtmlSummary();
                    this.Close();
                }
                else
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information); }

                cInsertCode_PopulateListboxCombobox = null;
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

        private void generateOneValue()
        {
            try
            {
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsInsertCode_PopulateListboxCombobox cInsertCode_PopulateListboxCombobox = new ClsInsertCode_PopulateListboxCombobox();

                cInsertCode_PopulateListboxCombobox.ControlName = cmbControl.Text;
                cInsertCode_PopulateListboxCombobox.sourceType = eSourceType;
                cInsertCode_PopulateListboxCombobox.Value = txtOneValue.Text;

                if (bIsOk == true)
                {
                    cInsertCode_PopulateListboxCombobox.generateCode(ref cCodeMapper);
                    configHtmlSummary(ref cInsertCode_PopulateListboxCombobox);
                    displayHtmlSummary();
                    this.Close();
                }
                else
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }

                cInsertCode_PopulateListboxCombobox = null;
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

        private void generateRecordSet()
        {
            try
            {
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsInsertCode_PopulateListboxCombobox cInsertCode_PopulateListboxCombobox = new ClsInsertCode_PopulateListboxCombobox();

                if (cmbRstFieldName.Items.Count == 0)
                { cInsertCode_PopulateListboxCombobox.fieldName = txtRstFieldName.Text; }
                else
                { cInsertCode_PopulateListboxCombobox.fieldName = cmbRstFieldName.Text; }
                
                if (cmbRstCommandType.Text == null)
                { cInsertCode_PopulateListboxCombobox.cmdType = ADODB.CommandTypeEnum.adCmdUnknown; }
                else
                { cInsertCode_PopulateListboxCombobox.cmdType = ClsMisc.getAdoCommandTypeEnum(cmbRstCommandType.Text); }

                cInsertCode_PopulateListboxCombobox.ControlName = cmbControl.Text;
                cInsertCode_PopulateListboxCombobox.sourceType = eSourceType;
                cInsertCode_PopulateListboxCombobox.connectionString = txtRstConnectionString.Text;
                if (cInsertCode_PopulateListboxCombobox.cmdType == ADODB.CommandTypeEnum.adCmdTable || cInsertCode_PopulateListboxCombobox.cmdType == ADODB.CommandTypeEnum.adCmdTableDirect)
                { cInsertCode_PopulateListboxCombobox.sql = cmbRstTableName.Text; }
                else
                { cInsertCode_PopulateListboxCombobox.sql = txtRstSql.Text; }
                
                cInsertCode_PopulateListboxCombobox.parameters = lstParameters;

                if (chkAddReference.Checked)
                {
                    FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_ADO, ref ssStatus);

                    if (!frmReference.referenceAlreadySet)
                    { frmReference.ShowDialog(this); }

                    frmReference = null;
                }

                if (bIsOk)
                {
                    ClsSettings cSettings = new ClsSettings();
                    cSettings.addUsedConnectionString(txtRstConnectionString.Text.Trim());
                    cSettings = null;
                }

                if (bIsOk == true)
                {
                    cInsertCode_PopulateListboxCombobox.generateCode(ref cCodeMapper);
                    configHtmlSummary(ref cInsertCode_PopulateListboxCombobox);
                    displayHtmlSummary();
                    this.Close();
                }
                else
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }

                cInsertCode_PopulateListboxCombobox = null;
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

        private void generateRangeNamed()
        {
            try
            {
                bool bIsOk = true;
                String sErrorMessage = "";

                ClsInsertCode_PopulateListboxCombobox cInsertCode_PopulateListboxCombobox = new ClsInsertCode_PopulateListboxCombobox();

                cInsertCode_PopulateListboxCombobox.sourceType = eSourceType;
                cInsertCode_PopulateListboxCombobox.ControlName = cmbControl.Text;
                cInsertCode_PopulateListboxCombobox.NamedRange = cmbRangeNamed.Text;

                if (bIsOk == true)
                {
                    cInsertCode_PopulateListboxCombobox.generateCode(ref cCodeMapper);
                    configHtmlSummary(ref cInsertCode_PopulateListboxCombobox);
                    displayHtmlSummary();
                    this.Close();
                }
                else
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }

                cInsertCode_PopulateListboxCombobox = null;
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

        private void generateRangeAddressed()
        {
            try
            {
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsInsertCode_PopulateListboxCombobox cInsertCode_PopulateListboxCombobox = new ClsInsertCode_PopulateListboxCombobox();

                cInsertCode_PopulateListboxCombobox.sourceType = eSourceType;
                cInsertCode_PopulateListboxCombobox.ControlName = cmbControl.Text;
                cInsertCode_PopulateListboxCombobox.sheetName = cmbRangeAddressSheetName.Text;
                cInsertCode_PopulateListboxCombobox.address = txtRangeAddressAddress.Text;

                if (bIsOk == true)
                {
                    cInsertCode_PopulateListboxCombobox.generateCode(ref cCodeMapper);
                    configHtmlSummary(ref cInsertCode_PopulateListboxCombobox);
                    displayHtmlSummary();
                    this.Close();
                }
                else
                { MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }

                cInsertCode_PopulateListboxCombobox = null;
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

        private void fillCmbRstCommandType()
        {
            try
            {
                List<string> lstTemp = new List<string>();

                foreach (ADODB.CommandTypeEnum eTemp in Enum.GetValues(typeof(ADODB.CommandTypeEnum)))
                { lstTemp.Add(eTemp.ToString()); }

                lstTemp.Sort();

                cmbRstCommandType.Items.Clear();

                foreach (string sTemp in lstTemp)
                { cmbRstCommandType.Items.Add(sTemp); }

                lstTemp = null;
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

        private void fillCmbRangeAddressSheetName()
        {
            try
            {
                List<string> lstTemp = new List<string>();

                foreach (Excel.Worksheet sht in ClsMisc.ActiveWorkBook().Worksheets)
                { lstTemp.Add(sht.Name); }

                lstTemp.Sort();

                cmbRangeAddressSheetName.Items.Clear();
                foreach (string sName in lstTemp)
                { cmbRangeAddressSheetName.Items.Add(sName); }

                lstTemp = null;
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

        private void fillCmbRangeNamed()
        {
            try
            {
                List<string> lstTemp = new List<string>();

                foreach (Excel.Name objName in ClsMisc.ActiveWorkBook().Names)
                { lstTemp.Add(objName.Name); }

                lstTemp.Sort();

                cmbRangeNamed.Items.Clear();
                foreach (string sName in lstTemp)
                { cmbRangeNamed.Items.Add(sName); }

                lstTemp = null;
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

        private void makeControlsVisible()
        {
            try
            {
                if (optRecordset.Checked && !optOneValue.Checked && !optArray.Checked && !optRangeNamed.Checked && !optRangeAddress.Checked)
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Recordset; }
                else if (!optRecordset.Checked && optOneValue.Checked && !optArray.Checked && !optRangeNamed.Checked && !optRangeAddress.Checked)
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_OneValue; }
                else if (!optRecordset.Checked && !optOneValue.Checked && optArray.Checked && !optRangeNamed.Checked && !optRangeAddress.Checked)
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Array; }
                else if (!optRecordset.Checked && !optOneValue.Checked && !optArray.Checked && optRangeNamed.Checked && !optRangeAddress.Checked)
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeNamed; }
                else if (!optRecordset.Checked && !optOneValue.Checked && !optArray.Checked && !optRangeNamed.Checked && optRangeAddress.Checked)
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeByAddress; }
                else
                { eSourceType = ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Unknown; }

                switch( eSourceType)
                {
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Recordset:
                        if (cmbRstCommandType.Text == null)
                        {
                            lblRstSql.Visible = true;
                            txtRstSql.Visible = true;
                            btnRstSqlExpand.Visible = true;
                            btnRstParameters.Visible = true;
                            lblRstTableName.Visible = false;
                            cmbRstTableName.Visible = false;
                        }
                        else
                        {
                            if (cmbRstCommandType.Text == ADODB.CommandTypeEnum.adCmdTable.ToString() || cmbRstCommandType.Text == ADODB.CommandTypeEnum.adCmdTableDirect.ToString())
                            {
                                lblRstSql.Visible = false;
                                txtRstSql.Visible = false;
                                btnRstSqlExpand.Visible = false;
                                btnRstParameters.Visible = false;
                                lblRstTableName.Visible = true;
                                cmbRstTableName.Visible = true;
                            }
                            else
                            {
                                lblRstSql.Visible = true;
                                txtRstSql.Visible = true;
                                btnRstSqlExpand.Visible = true;
                                btnRstParameters.Visible = true;
                                lblRstTableName.Visible = false;
                                cmbRstTableName.Visible = false;
                            }
                        }

                        lblRstFieldName.Visible = true;
                        if (cmbRstFieldName.Items.Count == 0)
                        {
                            txtRstFieldName.Visible = true;
                            cmbRstFieldName.Visible = false;
                        }
                        else
                        {
                            txtRstFieldName.Visible = false;
                            cmbRstFieldName.Visible = true;
                        }

                        lblRstConnectionString.Visible = true;
                        txtRstConnectionString.Visible = true;
                        lblRstCommandType.Visible = true;
                        cmbRstCommandType.Visible = true;
                        btnRstConnectionStringExpand.Visible = true;
                        btnRstConnectionStringBuild.Visible = true;
                        btnRstConnectionStringRecent.Visible = true;

                        lblOneValue.Visible = false;
                        txtOneValue.Visible = false;

                        btnArrayAdd.Visible = false;
                        btnArrayRemove.Visible = false;
                        lblArray.Visible = false;
                        lstArray.Visible = false;

                        lblRangeNamed.Visible = false;
                        cmbRangeNamed.Visible = false;

                        cmbRangeAddressSheetName.Visible = false;
                        lblRangeAddressSheetName.Visible = false;
                        lblRangeAddressAddress.Visible = false;
                        txtRangeAddressAddress.Visible = false;

                        chkAddReference.Visible = true;
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_OneValue:
                        lblRstSql.Visible = false;
                        txtRstSql.Visible = false;
                        btnRstParameters.Visible = false;
                        lblRstTableName.Visible = false;
                        cmbRstTableName.Visible = false;
                        lblRstConnectionString.Visible = false;
                        txtRstConnectionString.Visible = false;
                        lblRstCommandType.Visible = false;
                        cmbRstCommandType.Visible = false;
                        btnRstSqlExpand.Visible = false;
                        btnRstConnectionStringExpand.Visible = false;
                        btnRstConnectionStringBuild.Visible = false;
                        btnRstConnectionStringRecent.Visible = false;
                        lblRstFieldName.Visible = false;
                        txtRstFieldName.Visible = false;
                        cmbRstFieldName.Visible = false;

                        lblOneValue.Visible = true;
                        txtOneValue.Visible = true;

                        btnArrayAdd.Visible = false;
                        btnArrayRemove.Visible = false;
                        lblArray.Visible = false;
                        lstArray.Visible = false;

                        lblRangeNamed.Visible = false;
                        cmbRangeNamed.Visible = false;

                        cmbRangeAddressSheetName.Visible = false;
                        lblRangeAddressSheetName.Visible = false;
                        lblRangeAddressAddress.Visible = false;
                        txtRangeAddressAddress.Visible = false;

                        chkAddReference.Visible = false;
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Array:
                        lblRstSql.Visible = false;
                        txtRstSql.Visible = false;
                        btnRstParameters.Visible = false;
                        lblRstTableName.Visible = false;
                        cmbRstTableName.Visible = false;
                        lblRstConnectionString.Visible = false;
                        txtRstConnectionString.Visible = false;
                        lblRstCommandType.Visible = false;
                        cmbRstCommandType.Visible = false;
                        btnRstSqlExpand.Visible = false;
                        btnRstConnectionStringExpand.Visible = false;
                        btnRstConnectionStringBuild.Visible = false;
                        btnRstConnectionStringRecent.Visible = false;
                        lblRstFieldName.Visible = false;
                        txtRstFieldName.Visible = false;
                        cmbRstFieldName.Visible = false;

                        lblOneValue.Visible = false;
                        txtOneValue.Visible = false;

                        btnArrayAdd.Visible = true;
                        btnArrayRemove.Visible = true;
                        lblArray.Visible = true;
                        lstArray.Visible = true;

                        lblRangeNamed.Visible = false;
                        cmbRangeNamed.Visible = false;

                        cmbRangeAddressSheetName.Visible = false;
                        lblRangeAddressSheetName.Visible = false;
                        txtRangeAddressAddress.Visible = false;
                        lblRangeAddressAddress.Visible = false;

                        chkAddReference.Visible = false;
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeNamed:
                        lblRstSql.Visible = false;
                        txtRstSql.Visible = false;
                        btnRstParameters.Visible = false;
                        lblRstTableName.Visible = false;
                        cmbRstTableName.Visible = false;
                        lblRstConnectionString.Visible = false;
                        txtRstConnectionString.Visible = false;
                        lblRstCommandType.Visible = false;
                        cmbRstCommandType.Visible = false;
                        btnRstSqlExpand.Visible = false;
                        btnRstConnectionStringExpand.Visible = false;
                        btnRstConnectionStringBuild.Visible = false;
                        btnRstConnectionStringRecent.Visible = false;
                        lblRstFieldName.Visible = false;
                        txtRstFieldName.Visible = false;
                        cmbRstFieldName.Visible = false;

                        lblOneValue.Visible = false;
                        txtOneValue.Visible = false;

                        btnArrayAdd.Visible = false;
                        btnArrayRemove.Visible = false;
                        lblArray.Visible = false;
                        lstArray.Visible = false;

                        lblRangeNamed.Visible = true;
                        cmbRangeNamed.Visible = true;

                        cmbRangeAddressSheetName.Visible = false;
                        lblRangeAddressSheetName.Visible = false;
                        txtRangeAddressAddress.Visible = false;
                        lblRangeAddressAddress.Visible = false;

                        chkAddReference.Visible = false;
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeByAddress:
                        lblRstSql.Visible = false;
                        txtRstSql.Visible = false;
                        btnRstParameters.Visible = false;
                        lblRstTableName.Visible = false;
                        cmbRstTableName.Visible = false;
                        lblRstConnectionString.Visible = false;
                        txtRstConnectionString.Visible = false;
                        lblRstCommandType.Visible = false;
                        cmbRstCommandType.Visible = false;
                        btnRstSqlExpand.Visible = false;
                        btnRstConnectionStringExpand.Visible = false;
                        btnRstConnectionStringBuild.Visible = false;
                        btnRstConnectionStringRecent.Visible = false;
                        lblRstFieldName.Visible = false;
                        txtRstFieldName.Visible = false;
                        cmbRstFieldName.Visible = false;

                        lblOneValue.Visible = false;
                        txtOneValue.Visible = false;

                        btnArrayAdd.Visible = false;
                        btnArrayRemove.Visible = false;
                        lblArray.Visible = false;
                        lstArray.Visible = false;

                        lblRangeNamed.Visible = false;
                        cmbRangeNamed.Visible = false;

                        cmbRangeAddressSheetName.Visible = true;
                        lblRangeAddressSheetName.Visible = true;
                        txtRangeAddressAddress.Visible = true;
                        lblRangeAddressAddress.Visible = true;

                        chkAddReference.Visible = false;
                        break;
                    default:
                        lblRstSql.Visible = false;
                        txtRstSql.Visible = false;
                        btnRstParameters.Visible = false;
                        lblRstTableName.Visible = false;
                        cmbRstTableName.Visible = false;
                        lblRstConnectionString.Visible = false;
                        txtRstConnectionString.Visible = false;
                        lblRstCommandType.Visible = false;
                        cmbRstCommandType.Visible = false;
                        btnRstSqlExpand.Visible = false;
                        btnRstConnectionStringExpand.Visible = false;
                        btnRstConnectionStringBuild.Visible = false;
                        btnRstConnectionStringRecent.Visible = false;
                        lblRstFieldName.Visible = false;
                        txtRstFieldName.Visible = false;
                        cmbRstFieldName.Visible = false;

                        lblOneValue.Visible = false;
                        txtOneValue.Visible = false;

                        btnArrayAdd.Visible = false;
                        btnArrayRemove.Visible = false;
                        
                        lblArray.Visible = false;
                        lstArray.Visible = false;

                        lblRangeNamed.Visible = false;
                        cmbRangeNamed.Visible = false;

                        cmbRangeAddressSheetName.Visible = false;
                        lblRangeAddressSheetName.Visible = false;
                        txtRangeAddressAddress.Visible = false;
                        lblRangeAddressAddress.Visible = false;

                        chkAddReference.Visible = false;
                        break;
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

        private void btnArrayAdd_Click(object sender, EventArgs e)
        {
            try
            {
                arrayAdd();
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

        private void arrayAdd()
        {
            try
            {
                string sNewItem = FrmInputBox.GetString("Item", "Please enter the item you wish to add.");

                //bool bIsOk = true;
                bool bAddItem;

                List<string> lstItems = new List<string>();

                foreach (object objItem in lstArray.Items)
                { lstItems.Add(objItem.ToString());  }

                //if (lstArray.Items.Contains(sNewItem))
                if (lstItems.Exists(x => x.ToLower().Trim() == sNewItem.ToLower().Trim()))
                {
                    DialogResult dlgIsOk = MessageBox.Show("This item is already in the list are you sure you want to add it?", ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (dlgIsOk == System.Windows.Forms.DialogResult.Yes)
                    { bAddItem = true; }
                    else
                    { bAddItem = false; }
                }
                else
                { bAddItem = true; }

                if (bAddItem == true)
                { lstArray.Items.Add(sNewItem); }

                lstItems = null;
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

        private void btnArrayRemove_Click(object sender, EventArgs e)
        {
            try
            {
                arrayRemove();
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

        private void arrayRemove()
        {
            try
            {
                if (lstArray.SelectedIndex != null)
                {
                    if (lstArray.SelectedIndex != -1)
                    { lstArray.Items.RemoveAt(lstArray.SelectedIndex); }
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

        private void optRecordset_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
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

        private void optOneValue_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
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

        private void optArray_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
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

        private void optRangeNamed_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
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

        private void optRangeAddress_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
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

        private void cmbRstCommandType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                makeControlsVisible();
                fillCmbRstColumnNames();
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

        private void btnRstSqlExpand_Click(object sender, EventArgs e)
        {
            try
            {
                string sOrig;

                if (txtRstSql.Text == null)
                { sOrig = ""; }
                else
                { sOrig = txtRstSql.Text; }

                string sSource = FrmLargeTextBox.GetString(ClsDefaults.formTitle, "SQL", sOrig, false);

                if (!string.IsNullOrEmpty(sSource))
                { txtRstSql.Text = sSource; }
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

        private void btnRstConnectionStringBuild_Click(object sender, EventArgs e)
        {
            try
            {
                build();
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

        private void build()
        {
            try
            {
                FrmConnectionString frm = new FrmConnectionString();

                frm.ShowDialog(this);

                string sResult = frm.Result;

                if (sResult != "")
                { txtRstConnectionString.Text = sResult; }

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

        private void btnRstConnectionStringRecent_Click(object sender, EventArgs e)
        {
            try
            {
                recent();
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

        private void recent()
        {
            try
            {
                string sTemp = FrmRecentConnectionStrings.GetString();

                if (sTemp != "")
                { txtRstConnectionString.Text = sTemp; }
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

        private void btnRstConnectionStringExpand_Click(object sender, EventArgs e)
        {
            try
            {
                string sOrig;

                if (txtRstConnectionString.Text == null)
                { sOrig = ""; }
                else
                { sOrig = txtRstConnectionString.Text; }

                string sSource = FrmLargeTextBox.GetString(ClsDefaults.formTitle, "Connection String", sOrig, false);

                if (!string.IsNullOrEmpty(sSource))
                { txtRstConnectionString.Text = sSource; }
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

        private void FrmInsertCode_Rst_PopulateListboxCombobox_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblRstSql);
                cControlPosition.positionControl(ref txtRstSql);
                cControlPosition.positionControl(ref lblRstFieldName);
                cControlPosition.positionControl(ref txtRstFieldName);
                cControlPosition.positionControl(ref lblRstTableName);
                cControlPosition.positionControl(ref cmbRstTableName);
                cControlPosition.positionControl(ref lblRstConnectionString);
                cControlPosition.positionControl(ref txtRstConnectionString);
                cControlPosition.positionControl(ref lblRstCommandType);
                cControlPosition.positionControl(ref cmbRstCommandType);
                cControlPosition.positionControl(ref btnRstSqlExpand);
                cControlPosition.positionControl(ref btnRstParameters);
                cControlPosition.positionControl(ref btnRstConnectionStringExpand);
                cControlPosition.positionControl(ref btnRstConnectionStringBuild);
                cControlPosition.positionControl(ref btnRstConnectionStringRecent);

                cControlPosition.positionControl(ref lblOneValue);
                cControlPosition.positionControl(ref txtOneValue);

                cControlPosition.positionControl(ref btnArrayAdd);
                cControlPosition.positionControl(ref btnArrayRemove);
                cControlPosition.positionControl(ref lblArray);
                cControlPosition.positionControl(ref lstArray);

                cControlPosition.positionControl(ref lblRangeNamed);
                cControlPosition.positionControl(ref cmbRangeNamed);

                cControlPosition.positionControl(ref lblRangeAddressSheetName);
                cControlPosition.positionControl(ref cmbRangeAddressSheetName);
                cControlPosition.positionControl(ref lblRangeAddressAddress);
                cControlPosition.positionControl(ref txtRangeAddressAddress);

                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref chkAddReference);

                cControlPosition.positionControl(ref grpSource);
                cControlPosition.positionControl(ref grpType);

                cControlPosition.positionControl(ref optArray);
                cControlPosition.positionControl(ref optComboBox);
                cControlPosition.positionControl(ref optListBox);
                cControlPosition.positionControl(ref optOneValue);
                cControlPosition.positionControl(ref optRangeAddress);
                cControlPosition.positionControl(ref optRangeNamed);
                cControlPosition.positionControl(ref optRecordset);

                cControlPosition.positionControl(ref lblControl);
                cControlPosition.positionControl(ref cmbControl);

                cControlPosition.positionControl(ref lblWarning);

                //cControlPosition.positionControl(ref lblFieldName);
                //cControlPosition.positionControl(ref txtFieldName);
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

        private void btnParameters_Click(object sender, EventArgs e)
        {
            try
            {
                lstParameters = FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters.GetParameters("", lstParameters);

                //MessageBox.Show(lstParameters.Count.ToString());
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

        private void fillCmbRstColumnNames()
        {
            try
            {
                bool bWithSqlStatement;
                bool bIsOk = true;
                bool bSkipFill = false;
                string sMessage = "";

                cmbRstFieldName.Items.Clear();

                if (txtRstConnectionString.Text == null)
                { bSkipFill = true; }
                else if (txtRstConnectionString.Text == "")
                { bSkipFill = true; }

                if (cmbRstCommandType.Text == null)
                { bSkipFill = true; }
                else if (cmbRstCommandType.Text == "")
                { bSkipFill = true; }

                if (cmbRstCommandType.Text == null)
                { bWithSqlStatement = true; }
                else
                {
                    if (cmbRstCommandType.Text == ADODB.CommandTypeEnum.adCmdTable.ToString() || cmbRstCommandType.Text == ADODB.CommandTypeEnum.adCmdTableDirect.ToString())
                    { bWithSqlStatement = false; }
                    else
                    { bWithSqlStatement = true; }
                }

                ADODB.Connection con = new ADODB.Connection();

                if (!bSkipFill)
                {
                    try { con.Open(txtRstConnectionString.Text); }
                    catch
                    {
                        bIsOk = false;
                        sMessage = "Can't open connection, please check Connection String";
                    }
                }

                if (bWithSqlStatement)
                {
                    if (txtRstSql.Text == null)
                    { bSkipFill = true; }
                    else if (txtRstSql.Text == "")
                    { bSkipFill = true; }

                    if (!bSkipFill)
                    {
                        ADODB.Command cmd = new ADODB.Command();
                        ADODB.Recordset rst = new ADODB.Recordset();

                        if (bIsOk)
                        {
                            try { cmd.ActiveConnection = con; }
                            catch
                            {
                                bIsOk = false;
                                sMessage = "Can't open connection, please check Connection String";
                            }
                        }

                        ADODB.CommandTypeEnum eCmdType = ClsMisc.getAdoCommandTypeEnum(cmbRstCommandType.Text);

                        if (eCmdType == ADODB.CommandTypeEnum.adCmdUnknown)
                        {
                            bIsOk = false;
                            sMessage = "Unknown Command Type.";
                        }
                        else
                        { cmd.CommandType = eCmdType; }

                        if (bIsOk)
                        {
                            try { cmd.CommandText = txtRstSql.Text; }
                            catch
                            {
                                bIsOk = false;
                                sMessage = "Can't use the SQL, please check SQL statement is valid.";
                            }
                        }

                        if (bIsOk)
                        {
                            try { rst.Open(cmd, Type.Missing, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly); }
                            catch
                            {
                                bIsOk = false;
                                sMessage = "Can't use the SQL, please check SQL statement is valid.";
                            }
                        }

                        if (bIsOk)
                        {
                            for (int iCol = 0; iCol < rst.Fields.Count; iCol++)
                            { cmbRstFieldName.Items.Add(rst.Fields[iCol].Name); }
                        }

                        try
                        {
                            if (rst != null)
                            {
                                if (rst.State != (int)ADODB.ObjectStateEnum.adStateClosed)
                                { rst.Close(); }
                            }
                        }
                        catch
                        {
                            bIsOk = false;
                            sMessage = "Can't close the data.";
                        }

                        rst = null;
                        cmd = null;
                    }
                }
                else
                {
                    if (cmbRstTableName.Text == null)
                    { bSkipFill = true; }
                    else if (cmbRstTableName.Text == "")
                    { bSkipFill = true; }

                    if (!bSkipFill)
                    {
                        ADOX.Catalog cat = new ADOX.Catalog();
                        ADOX.Table tbl = new ADOX.Table();

                        try { cat.ActiveConnection = con; }
                        catch
                        {
                            bIsOk = false;
                            sMessage = "Can't open connection, please check Connection String";
                        }

                        if (bIsOk)
                        {
                            try { tbl = cat.Tables[cmbRstTableName.Text]; }
                            catch
                            {
                                bIsOk = false;
                                sMessage = "Can't open table, please check table name";
                            }
                        }

                        if (bIsOk)
                        {
                            List<string> lst = new List<string>();

                            foreach (ADOX.Column col in tbl.Columns)
                            { lst.Add(col.Name); }

                            lst.Sort();

                            cmbRstFieldName.Items.Clear();
                            foreach (string sItem in lst)
                            { cmbRstFieldName.Items.Add(sItem); }

                            lst = null;
                        }

                        tbl = null;
                        cat = null;
                    }
                }

                con = null;

                if (bSkipFill || bIsOk)
                { lblWarning.Visible = false; }
                else
                { lblWarning.Text = sMessage; }

                if (cmbRstFieldName.Items.Count == 0)
                {
                    txtRstFieldName.Visible = true;
                    cmbRstFieldName.Visible = false;
                }
                else
                {
                    txtRstFieldName.Visible = false;
                    cmbRstFieldName.Visible = true;
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

        private void fillCmbRstTableName()
        {
            try
            {
                ADOX.Catalog cat = new ADOX.Catalog();
                bool bIsOk = true;
                string sMessage = "";

                ADODB.Connection con = new ADODB.Connection();

                try { con.Open(txtRstConnectionString.Text); }
                catch
                {
                    bIsOk = false;
                    sMessage = "Can't open connection, please check Connection String";
                }

                try { cat.ActiveConnection = con; }
                catch
                {
                    bIsOk = false;
                    sMessage = "Can't open connection, please check Connection String";
                }

                if (bIsOk)
                {
                    List<string> lst = new List<string>();

                    foreach (ADOX.Table tbl in cat.Tables)
                    { lst.Add(tbl.Name); }

                    lst.Sort();

                    cmbRstTableName.Items.Clear();
                    foreach (string sItem in lst)
                    { cmbRstTableName.Items.Add(sItem); }
                    lblWarning.Text = "Please check your Connection String";

                    lst = null;
                }
                else
                { lblWarning.Text = sMessage; }

                cat = null;
                con = null;
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

        private void txtRstConnectionString_TextChanged(object sender, EventArgs e)
        {
            try
            {
                fillCmbRstTableName();
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

        private void cmbRstTableName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillCmbRstColumnNames();
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

        private void txtRstSql_TextChanged(object sender, EventArgs e)
        {
            try
            {
                fillCmbRstColumnNames();
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

        private void configHtmlSummary(ref ClsInsertCode_PopulateListboxCombobox cInsertCode_PopLstCmb)
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
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 3 }, "Auto generated code is located.");

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
                objCell.sText = "Value";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module where form is opened from.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PopLstCmb.moduleName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PopLstCmb.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Control Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PopLstCmb.ControlName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                switch (cInsertCode_PopLstCmb.sourceType)
                {
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Array:
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Array";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Count Of Items";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.array.Count.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        if (cInsertCode_PopLstCmb.array.Count > 0)
                        {
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 1 }, "Items");

                            foreach (string sItem in cInsertCode_PopLstCmb.array)
                            {
                                //Add Row
                                cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                                objCell.iOrder = 0;
                                objCell.bPropHtml = true;
                                objCell.sText = sItem;
                                objCell.sHiddenText = "";
                                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                            }
                        }
                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_OneValue:

                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1 }, "Items");
                    
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "One Value";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Value";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.Value;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeByAddress:
                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 3 }, "Details");

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Range by Address";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Sheet Name";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.sheetName;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Address";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.address;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_RangeNamed:
                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 3 }, "Details");

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Named Range";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Name";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.NamedRange;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        break;
                    case ClsInsertCode_PopulateListboxCombobox.enumSource.eSource_Recordset:
                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 3 }, "Details");
                        
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Recordset";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Connection String";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.connectionString;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Command Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.cmdType.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "SQL";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PopLstCmb.sql;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Fill Control ");

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

        private void AllControlsVisibility(bool bVisible)
        {
            try
            {
                foreach (System.Windows.Forms.Control cntl in this.Controls)
                { cntl.Visible = bVisible; }
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

        private bool checkNoParameters()
        {
            try
            {
                bool bIsOk = true;

                int iCountQuestionMarks = cmbRstTableName.Text.Count(x => x=='?');

                if (iCountQuestionMarks == lstParameters.Count)
                { bIsOk = true; }
                else
                { bIsOk = false; }

                return bIsOk;
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

                return false;
            }
        }

        //private void cmbRstTableName_Leave(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        checkForWarnings();
        //    }            
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void checkForWarnings()
        //{
        //    try
        //    {
        //        bool bIsOk = true;
        //        string sMessage = "";

        //        if (!checkNoParameters())
        //        {
        //            bIsOk = false;
        //            sMessage = "Please check the number of parameters equals the number of ? charactors in the SQL statement.  Each ? relates to a parameter.";
        //        }

        //        if (bIsOk)
        //        { lblWarning.Visible =false; }
        //        else
        //        {
        //            lblWarning.Text = sMessage;
        //            lblWarning.Visible = true;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

    }
}
