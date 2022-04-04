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

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_CommandBarClass_NodeEdit : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        private const string csCmbType_Button = "Button";
        private const string csCmbType_SubMenu = "Sub Menu";
        private const string csCmbType_ComboBox = "Combo Box";

        public FrmInsertCode_CommandBarClass_NodeEdit(string sCaption, string sMarcoToRun, string sTooltipText, Microsoft.Office.Core.MsoControlType eType, List<string> lstItems)
        {
            try
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;

                txtCaption.Text = sCaption;
                cmbMacroToRun.Text = sMarcoToRun;
                txtToolTipText.Text = sTooltipText;

                switch (eType)
                {
                    case Microsoft.Office.Core.MsoControlType.msoControlButton:
                        cmbControlType.Text = csCmbType_Button;
                        break;
                    case Microsoft.Office.Core.MsoControlType.msoControlComboBox:
                        cmbControlType.Text = csCmbType_ComboBox;
                        break;
                    case Microsoft.Office.Core.MsoControlType.msoControlPopup:
                        cmbControlType.Text = csCmbType_SubMenu;
                        break;
                    default:
                        cmbControlType.Text = csCmbType_Button;
                        break;
                }

                lstValues.Items.Clear();
                foreach (string sItem in lstItems)
                { lstValues.Items.Add(sItem); }

                ClsDefaults.FormatControl(ref lblCaption);
                ClsDefaults.FormatControl(ref txtCaption);
                ClsDefaults.FormatControl(ref lblMacroToRun);
                ClsDefaults.FormatControl(ref cmbMacroToRun);
                ClsDefaults.FormatControl(ref lblToolTipText);
                ClsDefaults.FormatControl(ref txtToolTipText);

                ClsDefaults.FormatControl(ref lblControlType);
                ClsDefaults.FormatControl(ref cmbControlType);
                ClsDefaults.FormatControl(ref lstValues);

                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnMinus);
                ClsDefaults.FormatControl(ref btnCancel);
                ClsDefaults.FormatControl(ref btnOK);
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

        public ClsInsertCode_CommandBarClass.strCommandControl ResultList { get; set; }

        public static ClsInsertCode_CommandBarClass.strCommandControl GetResults(string sCaption, string sMarcoToRun, string sTooltipText, Microsoft.Office.Core.MsoControlType eType, List<string> lstItems)
        {
            try
            {
                FrmInsertCode_CommandBarClass_NodeEdit box = new FrmInsertCode_CommandBarClass_NodeEdit(sCaption, sMarcoToRun, sTooltipText, eType, lstItems);

                if (box.ShowDialog() == DialogResult.OK)
                { return box.ResultList; }
                else
                {
                    ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                    objNode.sCaption = "";
                    objNode.sFullPath = "";
                    objNode.sFullPathParent = "";
                    objNode.sOnAction = "";
                    objNode.sTooltipText = "";
                    objNode.eType = Microsoft.Office.Core.MsoControlType.msoControlCustom;

                    return objNode; 
                }

                box = null;
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

                ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                objNode.sCaption = "";
                objNode.sFullPath = "";
                objNode.sFullPathParent = "";
                objNode.sOnAction = "";
                objNode.sTooltipText = "";
                objNode.eType = Microsoft.Office.Core.MsoControlType.msoControlCustom;

                return objNode; 
            }
        }

        public static ClsInsertCode_CommandBarClass.strCommandControl GetResults()
        {
            try
            {
                FrmInsertCode_CommandBarClass_NodeEdit box = new FrmInsertCode_CommandBarClass_NodeEdit("", "", "", Microsoft.Office.Core.MsoControlType.msoControlButton, new List<string>());

                if (box.ShowDialog() == DialogResult.OK)
                { return box.ResultList; }
                else
                { 
                    ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                    objNode.lstCmbValues = new List<string>();
                    objNode.lstCmbValues.Clear();
                    objNode.sCaption = "";
                    objNode.sFullPath = "";
                    objNode.sFullPathParent = "";
                    objNode.sOnAction = "";
                    objNode.sTooltipText = "";
                    objNode.sVariableName = "";
                    objNode.eType = Microsoft.Office.Core.MsoControlType.msoControlCustom;

                    return objNode; 
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

                ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                objNode.sCaption = "";
                objNode.sFullPath = "";
                objNode.sFullPathParent = "";
                objNode.sOnAction = "";
                objNode.sTooltipText = "";
                objNode.eType = Microsoft.Office.Core.MsoControlType.msoControlCustom;

                return objNode; 
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                ok();
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

        private void ok()
        { 
            try
            {
                bool bIsOk = true;
                string sMessage = "";
                
                if (txtCaption.Text.Trim() == "")
                {
                    bIsOk = false;
                    sMessage = "Caption is blank.";
                }

                if (bIsOk)
                {
                    ClsInsertCode_CommandBarClass.strCommandControl objResult = new ClsInsertCode_CommandBarClass.strCommandControl();

                    objResult.sCaption = txtCaption.Text;
                    objResult.sOnAction = cmbMacroToRun.Text;
                    objResult.sTooltipText = txtToolTipText.Text;
                    objResult.sFullPath = "";
                    objResult.sFullPathParent = "";
                    switch (cmbControlType.Text)
                    {
                        case csCmbType_Button:
                            objResult.eType = Microsoft.Office.Core.MsoControlType.msoControlButton;
                            break;
                        case csCmbType_ComboBox:
                            objResult.eType = Microsoft.Office.Core.MsoControlType.msoControlComboBox;
                            break;
                        case csCmbType_SubMenu:
                            objResult.eType = Microsoft.Office.Core.MsoControlType.msoControlPopup;
                            break;
                        default:
                            objResult.eType = Microsoft.Office.Core.MsoControlType.msoControlCustom;
                            break;
                    }
                    objResult.lstCmbValues = new List<string>();
                    objResult.lstCmbValues.Clear();
                    foreach (string sTemp in lstValues.Items)
                    { objResult.lstCmbValues.Add(sTemp); }

                    this.ResultList = objResult;
                    this.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private void FrmInsertCode_CommandBarClass_NodeEdit_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;

                ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();

                cmbControlType.Items.Clear();
                cmbControlType.Items.Add(csCmbType_Button);
                cmbControlType.Items.Add(csCmbType_SubMenu);
                cmbControlType.Items.Add(csCmbType_ComboBox);
                
                //cmbControlType.Text = csCmbType_Button;
                cmbControlTypeValueCheck();
                lblMessage.Visible = false;

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();
                cmbMacroToRun.Items.Clear();

                foreach (string sMarcoName in cCodeMapperWrk.getLstFunctionNames(true)) 
                { cmbMacroToRun.Items.Add(sMarcoName); }

                cControlPosition.setControl(lblCaption, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtCaption, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblMacroToRun, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbMacroToRun, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblToolTipText, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtToolTipText, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblMessage, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblControlType, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbControlType, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstValues, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnMinus, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnCancel, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOK, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
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

        private void cmbControlTypeValueCheck() 
        {
            try
            {
                switch (cmbControlType.Text)
                {
                    case csCmbType_Button:
                    case csCmbType_SubMenu:
                        lstValues.Visible = false;
                        btnAdd.Visible = false;
                        btnMinus.Visible = false;
                        break;
                    case csCmbType_ComboBox:
                        lstValues.Visible = true;
                        btnAdd.Visible = true;
                        btnMinus.Visible = true;
                        break;
                    default:
                        lstValues.Visible = true;
                        btnAdd.Visible = true;
                        btnMinus.Visible = true;
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

        private void cmbControlType_TextChanged(object sender, EventArgs e)
        {
            try
            {
                cmbControlTypeValueCheck();
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

        private void cmbControlType_Validating(object sender, CancelEventArgs e)
        {
            try 
            {
                if (string.IsNullOrEmpty(cmbControlType.Text))
                { 
                    e.Cancel = true;
                    lblMessage.Text = "Please make sure a value is entered in the " + lblControlType.Text;
                    lblMessage.Visible = true;
                    ClsDefaults.FormatControl(ref lblMessage, ClsDefaults.enumLabelState.eLbl_Warning);
                }
                else
                {
                    if (!cmbControlType.Items.Contains(cmbControlType.Text))
                    { 
                        e.Cancel = true;
                        lblMessage.Text = "Please make sure value is in the list";
                        lblMessage.Visible = true;
                        ClsDefaults.FormatControl(ref lblMessage, ClsDefaults.enumLabelState.eLbl_Warning);
                    } 
                    else
                    {
                        lblMessage.Visible= false;
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
                string sTemp = FrmInputBox.GetString("Value", "Please enter a value");

                if (!string.IsNullOrEmpty(sTemp))
                { lstValues.Items.Add(sTemp); }
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

        private void btnMinus_Click(object sender, EventArgs e)
        {
            try
            {
                subtract();
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

        private void subtract()
        {
            try
            {
                if (lstValues.SelectedItem != null)
                { lstValues.Items.RemoveAt(lstValues.SelectedIndex);}
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

        private void FrmInsertCode_CommandBarClass_NodeEdit_Resize(object sender, EventArgs e)
        {
            try{
                cControlPosition.positionControl(ref lblCaption);
                cControlPosition.positionControl(ref txtCaption);
                cControlPosition.positionControl(ref lblMacroToRun);
                cControlPosition.positionControl(ref cmbMacroToRun);
                cControlPosition.positionControl(ref lblToolTipText);
                cControlPosition.positionControl(ref txtToolTipText);
                cControlPosition.positionControl(ref lblMessage);

                cControlPosition.positionControl(ref lblControlType);
                cControlPosition.positionControl(ref cmbControlType);
                cControlPosition.positionControl(ref lstValues);

                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnMinus);
                cControlPosition.positionControl(ref btnCancel);
                cControlPosition.positionControl(ref btnOK);
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

        private void FrmInsertCode_CommandBarClass_NodeEdit_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.Subtract)
                    { subtract(); }

                    if (e.KeyCode == Keys.Add)
                    { add(); }

                    if (e.KeyCode == Keys.O)
                    { ok(); }

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
