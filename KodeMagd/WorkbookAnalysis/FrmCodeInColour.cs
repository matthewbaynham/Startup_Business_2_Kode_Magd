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
using KodeMagd.Settings;

namespace KodeMagd.WorkbookAnalysis
{
    public partial class FrmCodeInColour : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        List<ClsConfigReporter.strCss> lstColourSettings = new List<ClsConfigReporter.strCss>();
        //private const string sTextCodeOutsideFunctions = "<Code Outside functions>";

        public struct strFnMod
        {
            public string sModuleName;
            public string sFunctionName;
        }

        public FrmCodeInColour()
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

        private void FrmCodeInColour_Load(object sender, EventArgs e)
        {
            try
            {
                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref btnSettings);
                ClsDefaults.FormatControl(ref btnOuputCodeInColour);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref lblModules);
                ClsDefaults.FormatControl(ref lstModules);

                ClsDefaults.FormatControl(ref lblFunctions);
                ClsDefaults.FormatControl(ref lstFunctions);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnSettings, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOuputCodeInColour, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                this.Text = ClsDefaults.formTitle;

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillLstModule();
                fillLstColourSettings();
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
                lstFunctions.Items.Add(ClsDefaults.textCodeOutsideFunctions);
                
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
                cControlPosition.positionControl(ref btnOuputCodeInColour);
                cControlPosition.positionControl(ref btnClose);

                cControlPosition.positionControl(ref lblModules);
                cControlPosition.positionControl(ref lstModules);

                cControlPosition.positionControl(ref lblFunctions);
                cControlPosition.positionControl(ref lstFunctions);
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
                    if (e.KeyCode == Keys.O)
                    { OutputCodeInColour(); }

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

        private void btnOuputCodeInColour_Click(object sender, EventArgs e)
        {
            try
            {
                OutputCodeInColour();
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

        private void btnSettings_Click(object sender, EventArgs e)
        {
            try
            {
                FrmCodeInColour_Options frm = new FrmCodeInColour_Options(lstColourSettings);

                frm.ShowDialog();

                if (frm.OK)
                { lstColourSettings = frm.result; }

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

        private void fillLstColourSettings()
        {
            try
            {
                ClsSettings_CodeInColour cSettings_CodeInColour = new ClsSettings_CodeInColour();

                fillLstColourSettings(ClsCodeInColour.sCssName_DeclareVariables, cSettings_CodeInColour.lineColour_DeclareVariables);
                fillLstColourSettings(ClsCodeInColour.sCssName_AssignVariables, cSettings_CodeInColour.lineColour_AssignVariables);
                fillLstColourSettings(ClsCodeInColour.sCssName_IfStatements, cSettings_CodeInColour.lineColour_If);
                fillLstColourSettings(ClsCodeInColour.sCssName_Loops, cSettings_CodeInColour.lineColour_Loops);
                fillLstColourSettings(ClsCodeInColour.sCssName_DeclareFunctions, cSettings_CodeInColour.lineColour_DeclareFunctions);
                fillLstColourSettings(ClsCodeInColour.sCssName_Comments, cSettings_CodeInColour.lineColour_Comments);
                fillLstColourSettings(ClsCodeInColour.sCssName_Errors, cSettings_CodeInColour.lineColour_Errors);
                fillLstColourSettings(ClsCodeInColour.sCssName_With, cSettings_CodeInColour.lineColour_With);

                cSettings_CodeInColour = null;
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

        private void fillLstColourSettings(string sName, string sColour)
        {
            try
            {
                ClsConfigReporter.strCss objCss = new ClsConfigReporter.strCss();
                ClsConfigReporter.strCssStyle objCssItem = new ClsConfigReporter.strCssStyle();
                objCss.sName = sName;
                objCss.lstCssStyles = new List<ClsConfigReporter.strCssStyle>();
                objCssItem.sName = "color";
                objCssItem.sValue = sColour;
                objCss.lstCssStyles.Add(objCssItem);
                lstColourSettings.Add(objCss);
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

        private void FrmCodeInColour_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnSettings);
                cControlPosition.positionControl(ref btnOuputCodeInColour);
                cControlPosition.positionControl(ref btnClose);

                cControlPosition.positionControl(ref lblModules);
                cControlPosition.positionControl(ref lstModules);

                cControlPosition.positionControl(ref lblFunctions);
                cControlPosition.positionControl(ref lstFunctions);
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

        public void OutputCodeInColour()
        {
            try
            {
                ClsConfigReporterCodeInColour cConfigReporterCodeInColour = new ClsConfigReporterCodeInColour();
                List<string> lstMod = new List<string>();
                List<ClsCodeMapper.strFunctions> lstFn = new List<ClsCodeMapper.strFunctions>();
                bool bIsOk = true;
                string sMessage = "";

                foreach (ClsConfigReporter.strCss objCss in this.lstColourSettings)
                { cConfigReporterCodeInColour.setCssColour(objCss); }

                bool bAllFunctions = false;

                if (lstFunctions.CheckedItems.Count == 1)
                {
                    if (lstFunctions.CheckedItems[0] == ClsDefaults.textAll)
                    { bAllFunctions = true; }
                }

                if (bAllFunctions)
                {
                    foreach (string sFunctionName in lstFunctions.Items)
                    {
                        if (sFunctionName != ClsDefaults.textAll && sFunctionName != ClsDefaults.textCodeOutsideFunctions)
                        {
                            ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();

                            if (sFunctionName.Contains('.'))
                            {
                                int iPos = sFunctionName.IndexOf('.');

                                objFn.sModuleName = ClsMiscString.Left(sFunctionName, iPos);
                                objFn.sName = ClsMiscString.Right(sFunctionName, sFunctionName.Length - iPos - 1);
                            }
                            else
                            {
                                objFn.sModuleName = "";
                                if (lstModules.CheckedItems.Count == 1)
                                { objFn.sModuleName = lstModules.CheckedItems[0].ToString(); }
                                else
                                {
                                    if (lstModules.CheckedItems.Count == 2)
                                    {
                                        if (lstModules.CheckedItems[0].ToString() == ClsDefaults.textAll)
                                        { objFn.sModuleName = lstModules.CheckedItems[1].ToString(); }
                                    }
                                }
                                objFn.sName = sFunctionName;
                            }
                            
                            lstFn.Add(objFn);
                        }
                    }

                    //Code outside functions
                    if (lstModules.CheckedItems.Contains(ClsDefaults.textAll))
                    {
                        foreach (string sModuleName in lstModules.Items)
                        {
                            if (sModuleName != ClsDefaults.textAll)
                            {
                                ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();
                            
                                objFn.sModuleName = sModuleName;
                                objFn.sName = "";

                                lstFn.Add(objFn);
                            }
                        }
                    }
                    else
                    {
                        foreach (string sModuleName in lstModules.CheckedItems)
                        {
                            if (sModuleName != ClsDefaults.textAll)
                            {
                                ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();

                                objFn.sModuleName = sModuleName;
                                objFn.sName = "";

                                lstFn.Add(objFn);
                            }
                        }
                    }
                }
                else
                {
                    foreach (string sFunctionName in lstFunctions.CheckedItems)
                    {
                        if (sFunctionName != ClsDefaults.textCodeOutsideFunctions)
                        {
                            ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();

                            if (sFunctionName.Contains('.'))
                            {
                                int iPos = sFunctionName.IndexOf('.');

                                objFn.sModuleName = ClsMiscString.Left(sFunctionName, iPos);
                                objFn.sName = ClsMiscString.Right(sFunctionName, sFunctionName.Length - iPos - 1);
                            }
                            else
                            {
                                objFn.sModuleName = "";
                                if (lstModules.CheckedItems.Count == 1)
                                { objFn.sModuleName = lstModules.CheckedItems[0].ToString(); }

                                objFn.sName = sFunctionName;
                            }

                            lstFn.Add(objFn);
                        }
                    }

                    if (lstFunctions.CheckedItems.Contains(ClsDefaults.textCodeOutsideFunctions))
                    {

                        if (lstModules.CheckedItems.Contains (ClsDefaults.textAll))
                        {
                            foreach (string sModuleName in lstModules.Items)
                            {
                                ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();

                                objFn.sName="";
                                objFn.sModuleName= sModuleName;

                                lstFn.Add(objFn);
                            }
                        }
                        else
                        {
                            foreach (string sModuleName in lstModules.CheckedItems)
                            {
                                ClsCodeMapper.strFunctions objFn = new ClsCodeMapper.strFunctions();

                                objFn.sName = "";
                                objFn.sModuleName = sModuleName;

                                lstFn.Add(objFn);
                            }
                        }
                    
                    }

                }

                /*
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
                 */


                if (bIsOk)
                {
                    createObjectModelHtml(ref cConfigReporterCodeInColour, lstFn);

                    string sHtml = cConfigReporterCodeInColour.getHtml();

                    //foreach (string sText in lstText)
                    //{ sHtml += sText; }

                    FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Code in Colour");

                    frm.ShowDialog(this);

                    frm = null;

                    this.Close();
                }
                else
                { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                 
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

        public void createObjectModelHtml(ref ClsConfigReporterCodeInColour cConfigReporterCodeInColour, List<ClsCodeMapper.strFunctions> lstFunctions)
        {
            try
            {
                cConfigReporterCodeInColour = new ClsConfigReporterCodeInColour();

                cConfigReporterCodeInColour.lstCssExtra = lstColourSettings;

                ClsConfigReporterCodeInColour.strTableCell objCell = new ClsConfigReporterCodeInColour.strTableCell();
                int iTableId = 0;
                int iRowId = 0;

                foreach (ClsCodeMapper.strFunctions objFn in lstFunctions.OrderBy(x => x.sModuleName).ThenBy(y => y.sName).ThenBy(y => y.ePropertyType))//Note: only the module name and function name are set.
                {
                    List<ClsCodeMapper.strLine> lstLists = cCodeMapperWrk.getLines(objFn.sModuleName, new List<string> { objFn.sName });

                    ClsMisc.removeEmptyLinesAtBeginningAndEnd(ref lstLists);

                    if (lstLists.Count > 0)
                    {
                        string sFunctionType = "";
                        //string sPropertyType = "";

                        switch (lstLists[0].eFunctionType)
                        {
                            case ClsCodeMapper.enumFunctionType.eFnType_Function:
                                sFunctionType = "Function";
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                                sFunctionType = "Sub";
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Property:
                                sFunctionType = "Property";
                                //switch (lstLists[0].ePropertyType)
                                //{
                                //    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                //        sPropertyType = "Get";
                                //        break;
                                //    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                //        sPropertyType = "Let";
                                //        break;
                                //    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                //        sPropertyType = "Set";
                                //        break;
                                //}
                                break;
                            default:
                                sFunctionType = "Unknown";
                                break;
                        }

                        string sSubTitle = objFn.sModuleName + " - ";

                        if (objFn.sName.Trim() == "")
                        { sSubTitle += ClsDefaults.textCodeOutsideFunctions; }
                        else
                        { sSubTitle += sFunctionType + ": " + objFn.sName; }

                        //if (sFunctionType == "Property")
                        //{ sSubTitle += " (" + sPropertyType + ")"; }

                        cConfigReporterCodeInColour.TableAddNew(out iTableId, new List<int> { 1 }, sSubTitle);

                        //Add Row
                        cConfigReporterCodeInColour.TableAddNewRow(iTableId, out iRowId, false);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = false;
                        objCell.sText = cConfigReporterCodeInColour.createColouredHtmlText(ref lstLists, ref lstColourSettings);
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporterCodeInColour.TableAddNewCell(iTableId, iRowId, objCell);
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
    }
}
