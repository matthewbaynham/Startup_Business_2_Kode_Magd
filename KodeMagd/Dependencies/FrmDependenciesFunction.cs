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

namespace KodeMagd.Dependencies
{
    public partial class FrmDependenciesFunction : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();

        public FrmDependenciesFunction()
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

        private void FrmDependenciesFunction_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                cControlPosition.setControl(lblModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblFunction, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbFunction, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                ClsDefaults.FormatControl(ref lblModule);
                ClsDefaults.FormatControl(ref cmbModule);

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref ssStatus);

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModules();
                fillCmbFunctions();
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

        private void fillCmbFunctions()
        {
            try
            {
                cmbFunction.Items.Clear();

                string sModuleName = cmbModule.Text;
                cmbFunction.Text = "";

                if (string.IsNullOrEmpty(sModuleName))
                {
                    cmbFunction.Enabled = false;
                }
                else
                {
                    cmbFunction.Enabled = true;

                    if (cCodeMapperWrk.moduleExists(sModuleName))
                    {
                        foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(sModuleName))
                        {
                            string sItem = objFunction.sName;

                            switch (objFunction.eFunctionType)
                            {
                                case ClsCodeMapper.enumFunctionType.eFnType_Function:
                                    sItem += " | Function";
                                    break;
                                case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                                    sItem += " | Sub routine";
                                    break;
                                case ClsCodeMapper.enumFunctionType.eFnType_Property:
                                    sItem += " | Property";
                                    switch (objFunction.ePropertyType)
                                    {
                                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                            sItem += " | (Get)";
                                            break;
                                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                            sItem += " | (Let)";
                                            break;
                                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                            sItem += " | (Set)";
                                            break;
                                        default:
                                            sItem += " | (Unknow type)";
                                            break;
                                    }
                                    break;
                                default:
                                    sItem += " | Unknown type";
                                    break;
                            }

                            cmbFunction.Items.Add(sItem);
                        }
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
                string sModuleName = "";
                string sFunctionName="";
                ClsCodeMapper.enumFunctionPropertyType ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;
                ClsCodeMapper.enumFunctionType eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_None;
                bool bIsOk = true;
                string sMessage = "";

                if (string.IsNullOrWhiteSpace( cmbModule.Text ))
                {
                    bIsOk = false;
                    sMessage = "You must select a Module, Class or form.";
                }
                else
                { sModuleName = cmbModule.Text; }

                if (string.IsNullOrWhiteSpace(cmbFunction.Text))
                {
                    bIsOk = false;
                    sMessage = "You must select a Function, Sub routine or Property.";
                }
                else
                {
                    string sTempFunctionName = cmbFunction.Text;
                    List<string> lstFunction = sTempFunctionName.Split('|').ToList<string>();

                    switch (lstFunction.Count)
                    {
                        case 2:
                            sFunctionName = lstFunction[0].Trim();
                            switch(lstFunction[1].ToLower().Trim())
                            {
                                case "function":
                                    eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_Function;
                                    break;
                                case "sub routine":
                                    eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_Sub;
                                    break;
                                case "property":
                                    eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_Property;
                                    break;
                                default:
                                    eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_Error;
                                    break;
                            }
                            break;
                        case 3:
                            sFunctionName = lstFunction[0].Trim();
                            eFuntionType = ClsCodeMapper.enumFunctionType.eFnType_Property;
                            if (lstFunction[2].ToUpper().Trim() == "(GET)")
                            { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get; }
                            else if (lstFunction[2].ToUpper().Trim() == "(LET)")
                            { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Let; }
                            else if (lstFunction[2].ToUpper().Trim() == "(SET)")
                            { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Set; }
                            else
                            {
                                bIsOk = false;
                                sMessage = "Having problems recognising the Function, Sub routine or Property name";
                            }
                            break;
                        default:
                            bIsOk = false;
                            sMessage = "Having problems recognising the Function, Sub routine or Property name";
                            break;
                    }
                }

                if (bIsOk)
                {
                    ClsDependenciesFunction.generateReport(ref cCodeMapperWrk, sModuleName, sFunctionName, eFuntionType, ePropType, this);

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

        private void FrmDependenciesFunction_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblModule);
                cControlPosition.positionControl(ref cmbModule);

                cControlPosition.positionControl(ref lblFunction);
                cControlPosition.positionControl(ref cmbFunction);

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


        private void cmbModule_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillCmbFunctions();
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
