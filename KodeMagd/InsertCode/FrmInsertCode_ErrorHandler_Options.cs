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
    public partial class FrmInsertCode_ErrorHandler_Options : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eAction = new FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions();

        public FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions ResultEnum { get; set; }


        public static FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions GetEnum(FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eAction)
        {
            try
            {
                FrmInsertCode_ErrorHandler_Options box = new FrmInsertCode_ErrorHandler_Options { action = eAction };

                if (box.ShowDialog() == DialogResult.OK)
                { return box.ResultEnum; }
                else
                { return FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown; }

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
                return FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown;
            }
        }

        public FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions action
        {
            get 
            {
                try
                {
                    if (optDoNothing.Checked && !optIgnore.Checked && !optOneAtTop.Checked && !optOneInMiddle.Checked & !optReplaceAll.Checked)
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_DoNothingIfExists; }
                    else if (!optDoNothing.Checked && optIgnore.Checked && !optOneAtTop.Checked && !optOneInMiddle.Checked & !optReplaceAll.Checked)
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_IngoreOldAddNewRegardless; }
                    else if (!optDoNothing.Checked && !optIgnore.Checked && optOneAtTop.Checked && !optOneInMiddle.Checked & !optReplaceAll.Checked)
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneAtTop_ThenReplace; }
                    else if (!optDoNothing.Checked && !optIgnore.Checked && !optOneAtTop.Checked && optOneInMiddle.Checked & !optReplaceAll.Checked)
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneAnywhere_ThenReplace; }
                    else if (!optDoNothing.Checked && !optIgnore.Checked && !optOneAtTop.Checked && !optOneInMiddle.Checked & optReplaceAll.Checked)
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneOrMany_ThenReplace; }
                    else
                    { eAction = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown; }

                    return eAction;
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
                    return FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown;
                }
            }
            set 
            { 
                try
                {
                    eAction = value;
                    switch (eAction) 
                    {
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_DoNothingIfExists:
                            optDoNothing.Checked = true;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = false;
                            break;
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_IngoreOldAddNewRegardless:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = true;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = false;
                            break;
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneOrMany_ThenReplace:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = true;
                            break;
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneAnywhere_ThenReplace:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = true;
                            optReplaceAll.Checked = false;
                            break;
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneAtTop_ThenReplace:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = true;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = false;
                            break;
                        case FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = false;
                            break;
                        default:
                            optDoNothing.Checked = false;
                            optIgnore.Checked = false;
                            optOneAtTop.Checked = false;
                            optOneInMiddle.Checked = false;
                            optReplaceAll.Checked = false;
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
        
        public FrmInsertCode_ErrorHandler_Options()
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

        private void FrmInsertCode_ErrorHandler_Options_Load(object sender, EventArgs e)
        {
            try
            {
                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnOk);

                ClsDefaults.FormatControl(ref grpExistingHandlers);

                ClsDefaults.FormatControl(ref optIgnore);
                ClsDefaults.FormatControl(ref optOneAtTop);
                ClsDefaults.FormatControl(ref optOneInMiddle);
                ClsDefaults.FormatControl(ref optReplaceAll);
                ClsDefaults.FormatControl(ref optDoNothing);
                
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOk, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(grpExistingHandlers, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(optIgnore, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optOneAtTop, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optOneInMiddle, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optReplaceAll, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optDoNothing, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
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

        private void FrmInsertCode_ErrorHandler_Options_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnOk);

                cControlPosition.positionControl(ref grpExistingHandlers);

                cControlPosition.positionControl(ref optIgnore);
                cControlPosition.positionControl(ref optOneAtTop);
                cControlPosition.positionControl(ref optOneInMiddle);
                cControlPosition.positionControl(ref optReplaceAll);
                cControlPosition.positionControl(ref optDoNothing);
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
                this.ResultEnum = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_Unknown; 
                
                this.DialogResult = DialogResult.Cancel;
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

        private void btnOk_Click(object sender, EventArgs e)
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
                this.ResultEnum = this.action;
                this.DialogResult = DialogResult.OK;
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

        private void FrmInsertCode_ErrorHandler_Options_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
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
