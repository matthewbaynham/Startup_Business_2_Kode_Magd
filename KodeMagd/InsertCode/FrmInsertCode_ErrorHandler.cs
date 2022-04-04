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
    public partial class FrmInsertCode_ErrorHandler : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        enumReplaceErrorHandlerActions eActions = enumReplaceErrorHandlerActions.eErrHdl_DoNothingIfExists;

        public enum enumReplaceErrorHandlerActions
        {
            eErrHdl_IngoreOldAddNewRegardless,
            eErrHdl_OneAtTop_ThenReplace,
            eErrHdl_OneAnywhere_ThenReplace,
            eErrHdl_OneOrMany_ThenReplace,
            eErrHdl_DoNothingIfExists,
            eErrHdl_Unknown
        }

        public struct strFnMod
        {
            public string sModuleName;
            public string sFunctionName;
        }

        public FrmInsertCode_ErrorHandler()
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

        private void FrmInsertCode_ErrorHandler_Load(object sender, EventArgs e)
        {
            try
            {
                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref btnAddHandler);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref lblModules);
                ClsDefaults.FormatControl(ref lstModules);

                ClsDefaults.FormatControl(ref lblFunctions);
                ClsDefaults.FormatControl(ref lstFunctions);

                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnAddHandler, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                this.Text = ClsDefaults.formTitle;

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillLstModule();
                enumReplaceErrorHandlerActions eActions = FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions.eErrHdl_OneOrMany_ThenReplace;
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

        private void fillLstModule()
        {
            try
            {
                List<string> lst = new List<string>();

                foreach (ClsCodeMapper.strModuleDetails cModuleDetails in cCodeMapperWrk.getLstModuleDetails())
                { lst.Add(cModuleDetails.sName.Trim()); }
                lst.Sort();

                lstModules.Items.Clear();
                lstModules.Items.Add(ClsDefaults.textAll);
                lstModules.SetSelected(lstModules.Items.IndexOf(ClsDefaults.textAll), true);
                foreach (string sTemp in lst)
                {
                    lstModules.Items.Add(sTemp);
                    lstModules.SetSelected(lstModules.Items.IndexOf(sTemp), false);
                }

                lstModules.SetItemChecked(lstModules.Items.IndexOf(ClsDefaults.textAll), true);
                fillLstFunctions();

                lst = null;
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

        public void fillLstFunctions()
        {
            try
            {
                bool bIncludePrefixModName;
                List<string> lstText = new List<string>();

                if (lstModules.CheckedItems.Count != 1 | lstModules.CheckedItems.Contains(ClsDefaults.textAll))
                { bIncludePrefixModName = true; }
                else
                { bIncludePrefixModName = false; }

                lstFunctions.Text = null;

                List<ClsCodeMapper.strModuleDetails> lst = cCodeMapperWrk.getLstModuleDetails().Distinct().ToList();

                lst = lst.OrderBy(x => x.sName).ToList();

                lstFunctions.Items.Clear();

                bool bFilter;

                if (lstModules.CheckedItems.Count == 0)
                { bFilter = false; }
                else
                {
                    if (lstModules.CheckedItems.Contains(ClsDefaults.textAll) | lstModules.CheckedItems.Count == lstModules.Items.Count)
                    { bFilter = false; }
                    else
                    { bFilter = true; }
                }

                int iAllIndex = lstFunctions.Items.Add(ClsDefaults.textAll);

                foreach (ClsCodeMapper.strModuleDetails objMod in lst)
                {
                    bool bAddModule = false;

                    if (bFilter)
                    {
                        if (lstModules.CheckedItems.Contains(objMod.sName))
                        { bAddModule = true; }
                    }
                    else
                    { bAddModule = true; }

                    if (bAddModule)
                    {
                        foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(objMod.sName).Distinct().OrderBy(x => x.sName))
                        {
                            string sTemp = "";

                            if (bIncludePrefixModName == true)
                            { sTemp +=  objFunction.sModuleName + ".";}

                            sTemp += objFunction.sName;

                            if (objFunction.bHasErrorHandler == true)
                            { sTemp += " (*)"; }
                            lstText.Add(sTemp);
                        }
                    }
                }

                lstText = lstText.Distinct().ToList();
                lstText.Sort();

                foreach (string sText in lstText)
                { lstFunctions.Items.Add(sText); }

                lstFunctions.SetItemChecked(iAllIndex, true);

                lstText = null;
                lst = null;
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

        private void checkAllSelectedFunction()
        {
            try
            {
                if (lstFunctions.SelectedItem == lstFunctions.Items[lstModules.Items.IndexOf(ClsDefaults.textAll)])
                {
                    if (lstFunctions.CheckedItems.Contains(ClsDefaults.textAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < lstFunctions.Items.Count; iIndex++)
                        {
                            if (iIndex != lstFunctions.Items.IndexOf(ClsDefaults.textAll))
                            { lstFunctions.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (lstFunctions.CheckedItems.Count == 0 | (lstFunctions.CheckedItems.Count == 1 & lstFunctions.CheckedItems.Contains(ClsDefaults.textAll)))
                    { bAnySelected = false; }
                    else
                    { bAnySelected = true; }

                    if (bAnySelected)
                    { lstFunctions.SetItemChecked(lstFunctions.Items.IndexOf(ClsDefaults.textAll), false); }
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

        private void checkAllSelectedModules()
        {
            try
            {
                if (lstModules.SelectedItem == lstModules.Items[lstModules.Items.IndexOf(ClsDefaults.textAll)])
                {
                    if (lstModules.CheckedItems.Contains(ClsDefaults.textAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < lstModules.Items.Count; iIndex++)
                        {
                            if (iIndex != lstModules.Items.IndexOf(ClsDefaults.textAll))
                            { lstModules.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (lstModules.CheckedItems.Count == 0 | (lstModules.CheckedItems.Count == 1 & lstModules.CheckedItems.Contains(ClsDefaults.textAll)))
                    { bAnySelected = false; }
                    else
                    { bAnySelected = true; }

                    if (bAnySelected)
                    { lstModules.SetItemChecked(lstModules.Items.IndexOf(ClsDefaults.textAll), false); }
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

        private void lstModules_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelectedModules();
                fillLstFunctions();
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

        private void lstModules_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelectedModules();
                fillLstFunctions();
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

        private void lstFunctions_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelectedFunction();
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

        private void lstFunctions_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelectedFunction();
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

        private void FrmInsertCode_ErrorHandler_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnAddHandler);
                cControlPosition.positionControl(ref btnClose);

                cControlPosition.positionControl(ref lblModules);
                cControlPosition.positionControl(ref lstModules);

                cControlPosition.positionControl(ref lblFunctions);
                cControlPosition.positionControl(ref lstFunctions);

                cControlPosition.positionControl(ref lblWarning);
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

        private void btnAddHandler_Click(object sender, EventArgs e)
        {
            try
            {
                addHandler();
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

        private void addHandler()
        {
            try
            {
                enumReplaceErrorHandlerActions eActionsTemp = FrmInsertCode_ErrorHandler_Options.GetEnum(eActions);

                if (eActionsTemp != enumReplaceErrorHandlerActions.eErrHdl_Unknown)
                {
                    eActions = eActionsTemp;
                    editErrorHandlers();
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

        private void editErrorHandlers()
        {
            try
            {
                List<strFnMod> lstFnMod = new List<strFnMod>();
                List<strFnMod> lstEffectedFn = new List<strFnMod>();
                List<string> lstModuleNames = new List<string>();
                ClsInsertCode_ErrorHandler cInsertCode_ErrorHandler = new ClsInsertCode_ErrorHandler();

                foreach (string sFunction in lstFunctions.CheckedItems)
                {
                    strFnMod objFnMod;

                    if (sFunction.Contains("."))
                    {
                        int iPos = sFunction.IndexOf('.');

                        objFnMod.sModuleName = ClsMiscString.Left(sFunction, iPos);
                        objFnMod.sFunctionName = ClsMiscString.Right(sFunction, sFunction.Length - iPos - 1);

                        if (ClsMiscString.Right(ref objFnMod.sFunctionName, 4) == " (*)")
                        { objFnMod.sFunctionName = ClsMiscString.Left(ref objFnMod.sFunctionName, objFnMod.sFunctionName.Length - 4); }
                    }
                    else
                    {
                        objFnMod.sModuleName = "";
                        objFnMod.sFunctionName = sFunction;

                        if (lstModules.CheckedItems.Count == 1)
                        { objFnMod.sModuleName = lstModules.CheckedItems[0].ToString(); }

                        if (ClsMiscString.Right(ref objFnMod.sFunctionName, 4) == " (*)")
                        { objFnMod.sFunctionName = ClsMiscString.Left(ref objFnMod.sFunctionName, objFnMod.sFunctionName.Length - 4); }
                    }

                    lstFnMod.Add(objFnMod);
                    lstModuleNames.Add(objFnMod.sModuleName);
                }

                if (lstModuleNames.Contains(ClsDefaults.textAll))
                {
                    foreach (ClsCodeMapper.strModuleDetails objModuleDetails in cCodeMapperWrk.getLstModuleDetails())
                    { editErrorHandlers_ModuleLevel(objModuleDetails.sName, ref cInsertCode_ErrorHandler, ref lstFnMod, ref lstEffectedFn); }
                }
                else
                {
                    foreach (string sModuleName in lstModuleNames)
                    { editErrorHandlers_ModuleLevel(sModuleName, ref cInsertCode_ErrorHandler, ref lstFnMod, ref lstEffectedFn); }
                }

                cInsertCode_ErrorHandler.makeUniqueLstLog();
                cInsertCode_ErrorHandler.generateHtmlFile(ref cConfigReporter, ref lstEffectedFn);
                displayHtmlSummary();

                lstFnMod = null;
                lstModuleNames = null;
                cInsertCode_ErrorHandler = null;
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

        private void editErrorHandlers_ModuleLevel(string sModuleName, ref ClsInsertCode_ErrorHandler cInsertCode_ErrorHandler, ref List<strFnMod> lstFnMod, ref List<strFnMod> lstEffectedFn)
        {
            try
            {
                List<string> lstFn = new List<string>();

                if ((lstFnMod.Exists(y => y.sModuleName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower() 
                            || y.sModuleName.Trim().ToLower() == sModuleName.Trim().ToLower() 
                                    && y.sFunctionName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower())))
                {
                    foreach (string sFunctionName in cCodeMapperWrk.getLstFunctionNames(sModuleName, false))
                    {
                        lstFn.Add(sFunctionName);
                        lstFn = lstFn.Distinct().ToList();
                    }
                }
                else
                {
                    foreach (strFnMod objFn in lstFnMod.FindAll(x => x.sModuleName.Trim().ToLower() == sModuleName.Trim().ToLower()))
                    {
                        lstFn.Add(objFn.sFunctionName);
                        lstFn = lstFn.Distinct().ToList();
                    }
                }

                switch (eActions)
                {
                    case enumReplaceErrorHandlerActions.eErrHdl_DoNothingIfExists:
                        cInsertCode_ErrorHandler.replaceErrorRoutines_DoNothingIfErrorHandlerExists(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                        break;
                    case enumReplaceErrorHandlerActions.eErrHdl_IngoreOldAddNewRegardless:
                        cInsertCode_ErrorHandler.replaceErrorRoutines_IgnoreOldAddNewRegardless(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                        break;
                    case enumReplaceErrorHandlerActions.eErrHdl_OneAnywhere_ThenReplace:
                        cInsertCode_ErrorHandler.replaceErrorRoutines_OneAnywhereThenReplace(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                        break;
                    case enumReplaceErrorHandlerActions.eErrHdl_OneAtTop_ThenReplace:
                        cInsertCode_ErrorHandler.replaceErrorRoutines_OneAtTopThenReplace(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                        break;
                    case enumReplaceErrorHandlerActions.eErrHdl_OneOrMany_ThenReplace:
                        cInsertCode_ErrorHandler.replaceErrorRoutines_OneOrManyThenReplace(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                        break;
                    case enumReplaceErrorHandlerActions.eErrHdl_Unknown:
                        MessageBox.Show("Code not yet written");
                        break;
                    //cInsertCode_ErrorHandler.replaceErrorRoutines_OneOrManyThenReplace(ref cCodeMapperWrk, sModuleName, lstFn, eActions);
                }

                lstFn = null;
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Error_Handler_Modifications");

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

        private void FrmInsertCode_ErrorHandler_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { addHandler(); }

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
