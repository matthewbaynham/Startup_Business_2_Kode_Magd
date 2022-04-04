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

namespace KodeMagd.Dependencies
{
    public partial class FrmDependenciesModule : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();

        public FrmDependenciesModule()
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

        private void FrmDependenciesModule_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                cControlPosition.setControl(lblModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                ClsDefaults.FormatControl(ref lblModule);
                ClsDefaults.FormatControl(ref cmbModule);

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref ssStatus);

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModules();
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
                bool bIsOk = true;
                string sMessage = "";

                if (string.IsNullOrEmpty(cmbModule.Text))
                {
                    bIsOk = false;
                    sMessage = "Please select a Module Name.";
                }


                if (bIsOk)
                {
                    //List<ClsCodeMapper.strLine> lstLines = cCodeMapperWrk.findModuleReferences(cmbModule.Text);
                    generate();

                    this.Close();
                }
                else
                { MessageBox.Show(sMessage, "Operation Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

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

        private void fillCmbModules()
        {
            try
            {
                cmbModule.Items.Clear();
                
                foreach (ClsCodeMapper.strModuleDetails cModuleDetails in cCodeMapperWrk.getLstModuleDetails().OrderBy(x => x.sName)) 
                { cmbModule.Items.Add(cModuleDetails.sName); }
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

        private void FrmDependenciesModule_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblModule);
                cControlPosition.positionControl(ref cmbModule);

                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);
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
                ClsConfigReporter cConfigReporter = new ClsConfigReporter();

                ClsCodeMapper.strModuleDetails objModuleDetails = cCodeMapperWrk.getLstModuleDetails().Find(x => x.sName.ToLower().Trim() == cmbModule.Text.ToLower().Trim());
                List<ClsCodeMapperWrk.strLinesInModule> lstModuleInfo = cCodeMapperWrk.findModuleReferences(cmbModule.Text);

                foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(cmbModule.Text).FindAll(x => x.eScope != ClsCodeMapper.enumScopeFn.eScopeFn_Private))
                {
                    ClsCodeMapper.strFunctions objFunctionTemp = objFunction;

                    ClsCodeMapperWrk.strLinesInModule objModuleInfoTemp = new ClsCodeMapperWrk.strLinesInModule();
                    List<ClsCodeMapper.strLine> lstTemp = ClsDependenciesFunction.searchFunctionCalls(ref cCodeMapperWrk, ref objFunctionTemp);

                    foreach (ClsCodeMapper.strLine objLine in lstTemp)
                    {
                        if (lstModuleInfo.Exists(x => x.objModuleDetails.sName.ToUpper().Trim() == objLine.sModuleName.ToUpper().Trim()))
                        {
                            int iModuleIndex = lstModuleInfo.FindIndex(x => x.objModuleDetails.sName.ToUpper().Trim() == objLine.sModuleName.ToUpper().Trim());

                            ClsCodeMapperWrk.strLinesInModule objModuleInfoNew = lstModuleInfo[iModuleIndex];

                            objModuleInfoNew.lstLines.Add(objLine);
                            objModuleInfoNew.lstLines = objModuleInfoNew.lstLines.Distinct().OrderBy(x => x.sLineNo).ToList<ClsCodeMapper.strLine>();

                            lstModuleInfo[iModuleIndex] = objModuleInfoNew;
                        }
                        else
                        {
                            ClsCodeMapperWrk.strLinesInModule objModuleInfoNew =new ClsCodeMapperWrk.strLinesInModule();
                            objModuleInfoNew.lstLines = new List<ClsCodeMapper.strLine>();
                            objModuleInfoNew.lstLines.Add(objLine);
                            objModuleInfoNew.objModuleDetails.sName = objLine.sModuleName;
                            List<ClsCodeMapper.strModuleDetails> lstMod = cCodeMapperWrk.getLstModuleDetails().FindAll(x => x.sName.ToUpper().Trim() == objLine.sModuleName.ToUpper().Trim());
                            if (lstMod.Count == 1)
                            { objModuleInfoNew.objModuleDetails.eType = lstMod[0].eType; }
                            else
                            {objModuleInfoNew.objModuleDetails.eType =  Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule; }

                            lstModuleInfo.Add(objModuleInfoNew);
                        }

                        //lstModuleInfo
                    }
                    /*
                    objModuleInfoTemp.objModuleDetails = objModuleDetails;
                    objModuleInfoTemp.lstLines = ClsDependenciesFunction.searchFunctionCalls(ref cCodeMapperWrk, ref objFunctionTemp);

                    lstModuleInfo.Add(objModuleInfoTemp);
                     */
                }

                ClsDependenciesModule.buildHtml(ref cConfigReporter, ref lstModuleInfo, objModuleDetails);
                ClsDependenciesModule.displayHtmlSummary(ref cConfigReporter, this);

                //foreach()


/*
 * -===-====-==-================-
 *  Put this in ClsCodeMapperWrk
 * -===-====-==-================-
 * Search for module name through all text ( .contains )
 * Exclude anything where it is in a string or a comment
 * Look for variables where the variable name is the same as the name of the module and then exclude were its a variable
 */
                /*
                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Dependences");

                frm.ShowDialog(this);

                frm = null;
                */

                
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
