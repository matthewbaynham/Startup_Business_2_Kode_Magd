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
using System.Text.RegularExpressions;

namespace KodeMagd.Dependencies
{
    public partial class FrmDependenciesVariables : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        const string csListGlobals = "<List global variables in module>";

        public FrmDependenciesVariables()
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

        private void FrmDependenciesVariables_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                cControlPosition.setControl(lblModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblFunction, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbFunction, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblVariable, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbVariable, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                ClsDefaults.FormatControl(ref lblModule);
                ClsDefaults.FormatControl(ref cmbModule);

                ClsDefaults.FormatControl(ref lblModule);
                ClsDefaults.FormatControl(ref cmbModule);

                ClsDefaults.FormatControl(ref lblVariable);
                ClsDefaults.FormatControl(ref cmbVariable);

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref ssStatus);

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModules();
                fillCmbFunctions();
                fillCmbVariables();
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

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                string sModuleName = "";
                string sFunctionName="";
                string sVariableName = "";
                ClsCodeMapper.enumFunctionPropertyType ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;
                bool bIsOk = true;
                bool bGlobalVar = false;
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
                else if (cmbFunction.Text == csListGlobals)
                { bGlobalVar = true; }
                else
                {
                    bGlobalVar = false;
                    string sTempFunctionName = cmbFunction.Text;
                    List<string> lstFunction = sTempFunctionName.Split('|').ToList<string>();

                    switch (lstFunction.Count)
                    {
                        case 2:
                            sFunctionName = lstFunction[0].Trim();
                            break;
                        case 3:
                            sFunctionName = lstFunction[0].Trim();
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

                if (string.IsNullOrWhiteSpace( cmbVariable.Text ))
                {
                    bIsOk = false;
                    sMessage = "You must select a Variable.";
                }
                else
                { sVariableName = cmbVariable.Text; }


                if (bIsOk)
                {
                    if (bGlobalVar)
                    { ClsDepenenciesVariables.reportLocalVariableDependencies(ref cCodeMapperWrk, sModuleName, sVariableName, this); }
                    else
                    { ClsDepenenciesVariables.reportLocalVariableDependencies(ref cCodeMapperWrk, sModuleName, sFunctionName, ePropType, sVariableName, this); }

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

        private void FrmDependenciesVariables_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblModule);
                cControlPosition.positionControl(ref cmbModule);

                cControlPosition.positionControl(ref lblFunction);
                cControlPosition.positionControl(ref cmbFunction);

                cControlPosition.positionControl(ref lblVariable);
                cControlPosition.positionControl(ref cmbVariable);

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

        private void fillCmbFunctions()
        {
            try
            {
                cmbFunction.Items.Clear();
                cmbFunction.Items.Add(csListGlobals);

                string sModuleName = cmbModule.Text;
                cmbFunction.Text = "";
                cmbVariable.Text = "";

                if (string.IsNullOrEmpty(sModuleName))
                { 
                    cmbFunction.Enabled = false;
                    cmbVariable.Enabled = false;
                }
                else
                {
                    cmbFunction.Enabled = true;
                    cmbVariable.Enabled = true;

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

        private void fillCmbVariables()
        {
            try
            {
                cmbVariable.Items.Clear();

                string sModule = "";
                string sFunction = "";
                bool bIsOk = true;
                bool bGlobals = false;

                ClsCodeMapper.enumFunctionPropertyType ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;
                Predicate<ClsCodeMapper.strVariables> prepVar;
                
                prepVar = x => x.sFunctionName.Trim().ToUpper() == sFunction.Trim().ToUpper()
                && x.sModuleName.Trim().ToUpper() == sModule.Trim().ToUpper();

                if (string.IsNullOrWhiteSpace(cmbModule.Text))
                { bIsOk = false; }
                else
                { sModule = cmbModule.Text; }

                if (string.IsNullOrWhiteSpace(cmbFunction.Text))
                { bIsOk = false; }
                else if (cmbFunction.Text == csListGlobals)
                {
                    bGlobals = true;
                }
                else
                {
                    bGlobals = false;
                    string sTemp = cmbFunction.Text;
                    List<string> lstTemp = sTemp.Split('|').ToList<string>();

                    if (lstTemp.Count == 2)
                    {
                        //function or sub
                        sFunction = lstTemp[0];

                        prepVar = x => x.sFunctionName.Trim().ToUpper() == sFunction.Trim().ToUpper()
                        && x.sModuleName.Trim().ToUpper() == sModule.Trim().ToUpper();
                    }
                    else if (lstTemp.Count == 3)
                    {
                        //property (must also filter on get/set/let)
                        sFunction = lstTemp[0].Trim();

                        string sPropertyType = lstTemp[2].Trim();
                        switch (lstTemp[2].Trim().ToUpper())
                        {
                            case "(GET)":
                                ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get;
                                break;
                            case "(LET)":
                                ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Let;
                                break;
                            case "(SET)":
                                ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Set;
                                break;
                            default:
                                ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;
                                break;
                        }
                        prepVar = x => x.sFunctionName.Trim().ToUpper() == sFunction.Trim().ToUpper()
                            && x.ePropType == ePropType;
                    }
                    else
                    {
                        bIsOk = false;
                    }
                }

                if (bIsOk)
                {
                    if (bGlobals)
                    {
                        foreach (ClsCodeMapper.strVariables objVar in cCodeMapperWrk.getLstVariableDetails(sModule).FindAll(prepVar).OrderBy(y => y.sName).ToList<ClsCodeMapper.strVariables>())
                        { cmbVariable.Items.Add(objVar.sName); }
                    }
                    else
                    {
                        foreach (ClsCodeMapper.strVariables objVar in cCodeMapperWrk.getLstVariableDetails(sModule).FindAll(prepVar).OrderBy(y => y.sName).ToList<ClsCodeMapper.strVariables>())
                        { cmbVariable.Items.Add(objVar.sName); }
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

        private void cmbFunction_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillCmbVariables();
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
