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

namespace KodeMagd.WorkbookAnalysis
{
    public partial class FrmFlowDiagram : Form
    {
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        ClsControlPosition cControlPosition = new ClsControlPosition();

        public FrmFlowDiagram()
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
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();
                List<string> lstText = new List<string>();
                bool bIsOk = true;
                string sErrorMessage = "";

                ClsGenerateFlowDiagramme cGenerateFlowDiagramme = new ClsGenerateFlowDiagramme();
                List<string> lstFn = new List<string>();
                lstFn.Add(cmbFn.Text);

                if (string.IsNullOrEmpty(cmbModule.Text))
                {
                    bIsOk = false;
                    sErrorMessage = "Module/Form/Class Name can not be blank.";
                }
                else if (!cCodeMapperWrk.moduleExists(cmbModule.Text.Trim()))
                {
                    bIsOk = false;
                    sErrorMessage = "No such Module/Form/Class as '" + cmbModule.Text.Trim() + "'.";
                }

                if (bIsOk)
                {
                    if (string.IsNullOrEmpty(cmbFn.Text))
                    {
                        bIsOk = false;
                        sErrorMessage = "Function/Sub Routine/Property Name can not be blank.";
                    }
                    else if (!cCodeMapperWrk.functionNameExists(cmbFn.Text.Trim()))
                    {
                        bIsOk = false;
                        sErrorMessage = "No such Function/Sub Routine/Property as '" + cmbFn.Text.Trim() + "'.";
                    }
                }

                if (bIsOk)
                {
                    cGenerateFlowDiagramme.moduleName = cmbModule.Text;
                    cGenerateFlowDiagramme.functionName = cmbFn.Text;

                    cGenerateFlowDiagramme.generate(ref lstText, cCodeMapperWrk.getLines(cmbModule.Text, lstFn));

                    string sHtml = "";
                    foreach (string sText in lstText)
                    { sHtml += sText + "\n"; }

                    FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Flow Diagramme");

                    frm.ShowDialog(this);

                    frm = null;
                }
                else
                { MessageBox.Show(sErrorMessage, "Ooops", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

                fso = null;
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

        private void FrmFlowDiagram_Load(object sender, EventArgs e)
        {
            try
            {
                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref lblFnName);
                ClsDefaults.FormatControl(ref lblModule);
                
                ClsDefaults.FormatControl(ref cmbFn);
                ClsDefaults.FormatControl(ref cmbModule);
                
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(lblModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblFnName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(cmbModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbFn, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();
                fillComboModule();
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

        private void fillComboModule()
        {
            try
            {
                cmbModule.Items.Clear();

                foreach (ClsCodeMapper.strModuleDetails objDetails in cCodeMapperWrk.getLstModuleDetails().OrderBy(x => x.sName))
                { cmbModule.Items.Add(objDetails.sName); }
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

        private void fillComboFunctions()
        {
            try
            {
                cmbFn.Items.Clear();

                if (!string.IsNullOrEmpty(cmbModule.Text))
                {
                    string sModuleName = cmbModule.Text;

                    foreach (ClsCodeMapper.strFunctions objFn in cCodeMapperWrk.getCodeMapper(sModuleName).getLstFunctions().OrderBy(x => x.sName))
                    {
                        if (objFn.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        { 
                            switch(objFn.ePropertyType)
                            {
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                    cmbFn.Items.Add(objFn.sName + " - (Get)");
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                    cmbFn.Items.Add(objFn.sName + " - (Let)");
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                    cmbFn.Items.Add(objFn.sName + " - (Set)");
                                    break;
                                default:
                                    cmbFn.Items.Add(objFn.sName + " - (Unknown Property Type)");
                                    break;
                            }
                        }
                        else
                        { cmbFn.Items.Add(objFn.sName); }
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

        private void cmbModule_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillComboFunctions();
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
