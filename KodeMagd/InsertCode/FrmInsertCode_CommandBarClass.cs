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
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_CommandBarClass : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsInsertCode_CommandBarClass cInsertCode_CommandBarClass;
        private ClsSettings cSettings = new ClsSettings();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        private enum enumWarningAction
        {
            eWarn_MessageBox,
            eWarn_Label,
            eWarn_MessageBoxAndLabel,
            eWarn_Unknown
        }

        public FrmInsertCode_CommandBarClass()
        {
            try 
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref treMenu);

                ClsDefaults.FormatControl(ref chkPutSampleCallInOwnNewMod);

                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Warning);
                ClsDefaults.FormatControl(ref lblClassName);
                ClsDefaults.FormatControl(ref txtClassName);
                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnEdit);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref ssStatus);

                addRightClickMenu();

                cInsertCode_CommandBarClass = new ClsInsertCode_CommandBarClass();

                ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                cControlPosition.setControl(treMenu, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left , ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnEdit, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left , ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left , ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                
                cControlPosition.setControl(chkPutSampleCallInOwnNewMod, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right , ClsControlPosition.enumAnchorVertical.eAnchor_Top                    );
                cControlPosition.setControl(lblClassName,ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtClassName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right , ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(grpType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRibbonAddin, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRightClick, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
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

        private void FrmInsertCode_CommandBarClass_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;

                string sClassNamePrefix = "ClsCommandBar";
                string sClassName = sClassNamePrefix;
                int iAttempts = 1;
                while (ClsMisc.moduleExists(ClsMiscString.makeValidVarName(sClassName)))
                {
                    iAttempts++;
                    sClassName = sClassNamePrefix + iAttempts.ToString();
                }
                txtClassName.Text = sClassName;

                chkPutSampleCallInOwnNewMod.Checked = true;
                optRibbonAddin.Checked = true;
                optRightClick.Checked = false;

                ClsInsertCode_CommandBarClass.strCommandControl objNode = new ClsInsertCode_CommandBarClass.strCommandControl();

                objNode.sFullPath = "Toolbar";// nodNew.FullPath;
                objNode.sFullPathParent = "";
                objNode.sCaption = "Toolbar";
                objNode.sOnAction = "";
                objNode.sTooltipText = "Custom Menu Bar";
                objNode.eType = Microsoft.Office.Core.MsoControlType.msoControlPopup; //Not really custom it's not a control object it's a commandBar object

                addNode(objNode);

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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                addNode();
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

        private void editNode()
        {
            try 
            {
                bool bIsOk = true;
                string sMessage = "";

                if (treMenu.SelectedNode == null)
                {
                    bIsOk = false;
                    sMessage="Please select a nod before trying to edit it.";
                }

                if (bIsOk)
                {
                    string sFullPath = treMenu.SelectedNode.FullPath;

                    ClsInsertCode_CommandBarClass.strCommandControl con = cInsertCode_CommandBarClass.findControl(sFullPath);
                    FrmInsertCode_CommandBarClass_NodeEdit frm = new FrmInsertCode_CommandBarClass_NodeEdit(con.sCaption, con.sOnAction, con.sTooltipText, con.eType, con.lstCmbValues);
                    ClsInsertCode_CommandBarClass.strCommandControl objNode = FrmInsertCode_CommandBarClass_NodeEdit.GetResults(con.sCaption, con.sOnAction, con.sTooltipText, con.eType, con.lstCmbValues);
                    string sCaption = "";
                    string sToolTipText = "";

                    sCaption = objNode.sCaption;

                    if (sCaption.Trim() != "")
                    {
                        sToolTipText = "Caption: " + objNode.sCaption + Environment.NewLine;
                        sToolTipText += "Macro: " + objNode.sOnAction + Environment.NewLine;
                        sToolTipText += "Tooltip Text: " + objNode.sTooltipText;

                        if (objNode.eType == Microsoft.Office.Core.MsoControlType.msoControlPopup)
                        {
                            sToolTipText += Environment.NewLine;
                            sToolTipText += "Values: ";

                            string sValue = "";
                            for (int iCounter = 0; iCounter < objNode.lstCmbValues.Count; iCounter++)
                            {
                                if (iCounter == 0)
                                { sValue = objNode.lstCmbValues[iCounter]; }
                                else if (iCounter == objNode.lstCmbValues.Count - 1)
                                { sValue = " and " + objNode.lstCmbValues[iCounter] + "."; }
                                else
                                { sValue = ", " + objNode.lstCmbValues[iCounter]; }

                                sToolTipText += sValue;
                            }
                        }

                        TreeNode nodEdit = treMenu.SelectedNode;

                        treMenu.SelectedNode.Text = objNode.sCaption;
                        treMenu.SelectedNode.Name = objNode.sCaption;
                        treMenu.SelectedNode.ToolTipText = sToolTipText;

                        ClsDefaults.FormatControl(ref nodEdit);

                        if (treMenu.SelectedNode.Parent != null)
                        { treMenu.SelectedNode.Parent.ExpandAll(); }

                        objNode.sFullPath = sFullPath;
                        if (treMenu.SelectedNode == null)
                        { objNode.sFullPathParent = ""; }
                        else
                        { objNode.sFullPathParent = treMenu.SelectedNode.FullPath; }

                        cInsertCode_CommandBarClass.editControl(objNode);

                        nodEdit = null;
                    }
                    frm = null;
                }
                else
                {
                    MessageBox.Show(sMessage,
                                    ClsDefaults.messageBoxTitle(),
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Exclamation);
                }

                bool bDummy = checkInput(enumWarningAction.eWarn_Label);
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

        private void deleteNode()
        {
            try 
            {
                if (treMenu.SelectedNode == null)
                { MessageBox.Show("Please select a nod before trying to edit it.",
                                  ClsDefaults.messageBoxTitle(), 
                                  MessageBoxButtons.OK, 
                                  MessageBoxIcon.Exclamation); }
                else
                {
                    if (treMenu.SelectedNode.Level != 1)
                    {
                        DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete '" + treMenu.SelectedNode.Text + "'",
                                                                    ClsDefaults.messageBoxTitle(), 
                                                                    MessageBoxButtons.YesNo, 
                                                                    MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            cInsertCode_CommandBarClass.deleteControl(treMenu.SelectedNode.FullPath);

                            treMenu.SelectedNode.Remove();
                        }
                    }
                }

                bool bDummy = checkInput(enumWarningAction.eWarn_Label);
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

        private void addNode(ClsInsertCode_CommandBarClass.strCommandControl objNode)
        {
            try 
            {
                TreeNode nodNew;

                if (treMenu.SelectedNode == null)
                { nodNew = treMenu.Nodes.Add(objNode.sCaption); }
                else
                { nodNew = treMenu.SelectedNode.Nodes.Add(objNode.sCaption); }

                nodNew.Text = objNode.sCaption;
                nodNew.Name = objNode.sCaption;
                nodNew.ToolTipText = this.nodeTooltipText(objNode);

                switch (objNode.eType)
                {
                    case Microsoft.Office.Core.MsoControlType.msoControlButton:
                        ClsDefaults.FormatControl(ref nodNew, ClsDefaults.enumStyle.eStyle3);
                        break;
                    case Microsoft.Office.Core.MsoControlType.msoControlComboBox:
                        ClsDefaults.FormatControl(ref nodNew, ClsDefaults.enumStyle.eStyle5);
                        break;
                    case Microsoft.Office.Core.MsoControlType.msoControlPopup:
                        ClsDefaults.FormatControl(ref nodNew, ClsDefaults.enumStyle.eStyle2);
                        break;
                    case Microsoft.Office.Core.MsoControlType.msoControlCustom:
                        ClsDefaults.FormatControl(ref nodNew, ClsDefaults.enumStyle.eStyle4);
                        break;
                    default:
                        ClsDefaults.FormatControl(ref nodNew, ClsDefaults.enumStyle.eStyle5);
                        break;
                }

                if (nodNew.Parent != null)
                { nodNew.Parent.ExpandAll(); }

                objNode.sFullPath = nodNew.FullPath;

                cInsertCode_CommandBarClass.addControl(objNode);

                if (treMenu.SelectedNode == null)
                { treMenu.SelectedNode = nodNew; }

                nodNew = null;
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

        private void addNode()
        {
            try 
            {
                ClsInsertCode_CommandBarClass.strCommandControl objNode = FrmInsertCode_CommandBarClass_NodeEdit.GetResults();
                string sCaption = "";
                string sToolTipText = "";
                bool bIsOk = true;
                string sMessage = "";

                sCaption = objNode.sCaption;

                if (sCaption.Trim() == "")
                {
                    bIsOk = false;
                    sMessage = "";
                }

                if (bIsOk)
                {
                    sToolTipText += "Control Type: ";
                    switch (objNode.eType)
                    {
                        case Microsoft.Office.Core.MsoControlType.msoControlButton:
                            sToolTipText += "Button";
                            break;
                        case Microsoft.Office.Core.MsoControlType.msoControlComboBox:
                            sToolTipText += "Combo Box";
                            break;
                        case Microsoft.Office.Core.MsoControlType.msoControlPopup:
                            sToolTipText += "Sub Menu";
                            break;
                        default:
                            sToolTipText += "Unknown";
                            break;
                    }
                    sToolTipText += Environment.NewLine;
                    sToolTipText = "Caption: " + objNode.sCaption + Environment.NewLine;
                    sToolTipText += "Macro: " + objNode.sOnAction + Environment.NewLine;
                    sToolTipText += "Tooltip Text: " + objNode.sTooltipText;

                    objNode.sFullPath = "";
                    if (treMenu.SelectedNode == null)
                    { objNode.sFullPathParent = ""; }
                    else
                    { objNode.sFullPathParent = treMenu.SelectedNode.FullPath; }

                    if (!(objNode.sCaption == ""
                        & objNode.sOnAction == ""
                        & objNode.sTooltipText == ""
                        & objNode.sFullPath == ""
                        & objNode.lstCmbValues.Count == 0
                        & objNode.sVariableName == ""
                        & objNode.eType == Microsoft.Office.Core.MsoControlType.msoControlCustom))
                    { addNode(objNode); }

                    bool bDummy = checkInput(enumWarningAction.eWarn_Label);
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

        private string nodeTooltipText(ClsInsertCode_CommandBarClass.strCommandControl objNode) 
        { 
            try
            {
                string sResult = "";

                sResult = "Caption: " + objNode.sCaption + Environment.NewLine;
                sResult += "Macro: " + objNode.sOnAction + Environment.NewLine;
                sResult += "Tooltip Text: " + objNode.sTooltipText;

                if (objNode.eType == Microsoft.Office.Core.MsoControlType.msoControlComboBox) 
                {
                    if (objNode.lstCmbValues.Count > 0)
                    {
                        sResult += Environment.NewLine;
                        sResult += "Items in List:" + Environment.NewLine;
                        sResult += ClsMiscString.LstToText(objNode.lstCmbValues, 80);
                    }
                }

                return sResult;
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
                return "";
            }
        }

        private bool checkInput(enumWarningAction eAction) 
        {
            try
            {
                string sErrorMessage = "";
                bool bIsOk = checkInput(out sErrorMessage, eAction);

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

        private bool checkInput(out string sErrorMessage, enumWarningAction eAction) 
        {
            try
            {
                bool bIsOk = true;
                sErrorMessage = "";

                if (ClsMisc.isExistsModule(txtClassName.Text))
                {
                    bIsOk = false;
                    sErrorMessage += "There is all ready a code module called " + txtClassName.Text + Environment.NewLine;
                }

                if (chkPutSampleCallInOwnNewMod.Checked)
                {
                    if (ClsMisc.isExistsModule(cInsertCode_CommandBarClass.SampleCodeModulePrefix + txtClassName.Text))
                    {
                        bIsOk = false;
                        sErrorMessage += "There is all ready a code module called " + cInsertCode_CommandBarClass.SampleCodeModulePrefix + txtClassName.Text + Environment.NewLine;
                    }
                }

                if (optRibbonAddin.Checked & !optRightClick.Checked)
                { this.cInsertCode_CommandBarClass.menuType = ClsInsertCode_CommandBarClass.enumMenuType.eMenuRibbonAddin; }
                else if (!optRibbonAddin.Checked & optRightClick.Checked)
                { this.cInsertCode_CommandBarClass.menuType = ClsInsertCode_CommandBarClass.enumMenuType.eMenuRightClick; }
                else
                {
                    bIsOk = false;
                    sErrorMessage += "Can't select both types of menu at the same time." + Environment.NewLine;
                }

                if (optRightClick.Checked)
                {
                    if (cInsertCode_CommandBarClass.isExistsComboBox())
                    {
                        bIsOk = false;
                        sErrorMessage += "ComboBoxes don't appear in Right Click Menus." + Environment.NewLine;
                    }
                }

                if (!cInsertCode_CommandBarClass.isAllSubMenusAreSubMenus())
                {
                    bIsOk = false;
                    sErrorMessage += "Sub Menu's need to be given the control type of sub menu." + Environment.NewLine;
                }

                if (eAction == enumWarningAction.eWarn_MessageBoxAndLabel || eAction == enumWarningAction.eWarn_Label)
                {
                    if (bIsOk)
                    { ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible); }
                    else
                    {
                        lblWarning.Text = sErrorMessage;
                        ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Warning);
                    }
                }
                else if (eAction == enumWarningAction.eWarn_MessageBoxAndLabel || eAction == enumWarningAction.eWarn_MessageBox)
                { MessageBox.Show (sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

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
 
                sErrorMessage = "Unexpected Error";
                return false;
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
                bool bIsOk = checkInput(enumWarningAction.eWarn_MessageBoxAndLabel);

                this.cInsertCode_CommandBarClass.className = txtClassName.Text;
                this.cInsertCode_CommandBarClass.PutSampleCallInOwnNewMod = chkPutSampleCallInOwnNewMod.Checked;

                if (optRibbonAddin.Checked & !optRightClick.Checked)
                { this.cInsertCode_CommandBarClass.menuType = ClsInsertCode_CommandBarClass.enumMenuType.eMenuRibbonAddin; }
                else if (!optRibbonAddin.Checked & optRightClick.Checked)
                { this.cInsertCode_CommandBarClass.menuType = ClsInsertCode_CommandBarClass.enumMenuType.eMenuRightClick; }

                if (bIsOk)
                {
                    this.cInsertCode_CommandBarClass.generateSampleCode(ref cCodeMapper);
                    this.cInsertCode_CommandBarClass.generateToolbarClass();

                    configHtmlSummary();
                    displayHtmlSummary();

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

        private void addRightClickMenu() 
        {
            try
            {
                ContextMenuStrip mnu = new ContextMenuStrip();
                ToolStripMenuItem mnuAdd = new ToolStripMenuItem("Add");
                ToolStripMenuItem mnuEdit = new ToolStripMenuItem("Edit");
                ToolStripMenuItem mnuDelete = new ToolStripMenuItem("Delete");
                
                //Assign event handlers
                mnuAdd.Click += new EventHandler(mnuAdd_Click);
                mnuEdit.Click += new EventHandler(mnuEdit_Click);
                mnuDelete.Click += new EventHandler(mnuDelete_Click);
                
                //Add to main context menu
                mnu.Items.AddRange(new ToolStripItem[] { mnuAdd, mnuEdit, mnuDelete});
                
                //Assign to datagridview
                treMenu.ContextMenuStrip = mnu;

                mnu = null;
                mnuAdd = null;
                mnuEdit = null;
                mnuDelete = null;
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

        private void mnuAdd_Click(object sender, System.EventArgs e)
        {
            try
            {
                addNode();
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

        private void mnuEdit_Click(object sender, System.EventArgs e)
        {
            try
            {
                editNode();
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

        private void mnuDelete_Click(object sender, System.EventArgs e)
        {
            try
            {
                deleteNode();
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

        private void FrmInsertCode_CommandBarClass_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                cInsertCode_CommandBarClass = null;
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

        private void btnEdit_Click(object sender, EventArgs e)
        {
            try
            {
                editNode();
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

        private void optRibbonAddin_CheckedChanged(object sender, EventArgs e)
        {
            try 
            {
                bool bDummy = checkInput(enumWarningAction.eWarn_Label);
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

        private void optRightClick_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                bool bDummy = checkInput(enumWarningAction.eWarn_Label);
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

        private void configHtmlSummary()
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
                objCell.sText = cInsertCode_CommandBarClass.className;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Class";
                objCell.sHiddenText = "Code to manage the Toolbar";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_CommandBarClass.SampleCodeModulePrefix.Trim() + cInsertCode_CommandBarClass.className.Trim();
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
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 4 }, "Controls added to toolbar");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Caption";
                objCell.sHiddenText = "Text on Button";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Type";
                objCell.sHiddenText = "Type of control\nButton\nCombo Box\nSub Menu";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "On Action";
                objCell.sHiddenText = "Code called by button";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                List<ClsInsertCode_CommandBarClass.strCommandControl> lstControls = cInsertCode_CommandBarClass.CommandBarControls;

                recursiveLoopThroughControls(iTableId, out iRowId, ref lstControls, "", 0);

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

        private void recursiveLoopThroughControls(int iTableId, out int iRowId, ref List<ClsInsertCode_CommandBarClass.strCommandControl> lstControls, string sPath, int iIndent)
        {             
            try
            {
                iRowId = 0;
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();

                foreach (ClsInsertCode_CommandBarClass.strCommandControl objControl in lstControls.FindAll(x => x.sFullPathParent == sPath))
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = cSettings.Indent(iIndent) + objControl.sCaption;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    switch (objControl.eType)
                    {
                        case Microsoft.Office.Core.MsoControlType.msoControlButton:
                            objCell.sText = "Button";
                            break;
                        case Microsoft.Office.Core.MsoControlType.msoControlPopup:
                            objCell.sText = "Sub Menu";
                            break;
                        case Microsoft.Office.Core.MsoControlType.msoControlComboBox:
                            objCell.sText = "Combo Box";
                            break;
                        default:
                            objCell.sText = objControl.eType.ToString();
                            break;
                    }
                    
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = objControl.sOnAction;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    if (lstControls.Count > 0)
                    { recursiveLoopThroughControls(iTableId, out iRowId, ref lstControls, objControl.sFullPath, iIndent + 1); }
                }

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
                iRowId = 0;
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

        private void FrmInsertCode_CommandBarClass_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref treMenu);
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnEdit);
                cControlPosition.positionControl(ref btnRemove);

                cControlPosition.positionControl(ref chkPutSampleCallInOwnNewMod);
                cControlPosition.positionControl(ref lblClassName);
                cControlPosition.positionControl(ref txtClassName);
                cControlPosition.positionControl(ref lblWarning);

                cControlPosition.positionControl(ref grpType);
                cControlPosition.positionControl(ref optRibbonAddin);
                cControlPosition.positionControl(ref optRightClick);

                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnClose);
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

        private void FrmInsertCode_CommandBarClass_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { addNode(); }

                    if (e.KeyCode == Keys.E)
                    { editNode(); }

                    if (e.KeyCode == Keys.R)
                    { deleteNode(); }

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

        private void btnRemove_Click(object sender, EventArgs e)
        {
            try
            {
                deleteNode();
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
