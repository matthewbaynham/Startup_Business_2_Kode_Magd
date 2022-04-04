using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using Office = Microsoft.Office.Core;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_Class : ClsInsertCode
    {
        //private const string csSampleCodeModulePrefix = "SampleCode_";
        private const string csPrefixProperty = "prop";

        public struct strParameter
        {
            public string sNameExtermal;
            public string sNameInternal;
            public bool bReadOnly;
            public ClsDataTypes.vbVarType eType;
            public string sDefaultValue;
        }

        List<strParameter> lstParameters = new List<strParameter>();
        private string sClassName;
        private bool bPutSampleCallInOwnNewMod;
        private string sFunctionName = "";
        private string sModuleName = "";
        private string sSampleModuleName = "";

        public string functionName
        {
            get
            {
                try
                {
                    return sFunctionName;
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
        }

        public string moduleName
        {
            get
            {
                try
                {
                    return sModuleName;
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
        }

        public string className
        {
            get
            {
                try
                {
                    return sClassName;
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
            set
            {
                try
                {
                    sClassName = value;
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

        public string SampleModuleName
        {
            get
            {
                try
                { return sSampleModuleName; }
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
            set 
            {
                try
                { sSampleModuleName = value; }
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

        public bool PutSampleCallInOwnNewMod
        {
            get
            {
                try
                {
                    return bPutSampleCallInOwnNewMod;
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
                    return true;
                }
            }
            set
            {
                try
                {
                    bPutSampleCallInOwnNewMod = value;
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

        public void addParameter(string sName, bool bReadOnly, ClsDataTypes.vbVarType eType, string sDefaultValue)
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();
                strParameter objParameter = new strParameter();
                
                string sPrefix = cDataTypes.typePrefix(eType);

                objParameter.sNameExtermal = ClsMiscString.makeValidVarName(sName);
                objParameter.sNameInternal = ClsMiscString.makeValidVarName(sPrefix + " " + sName);
                objParameter.bReadOnly = bReadOnly;
                objParameter.eType = eType;
                objParameter.sDefaultValue = sDefaultValue;

                cDataTypes = null;

                lstParameters.Add(objParameter);
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

        public List<strParameter> parameters
        {
            get
            {
                try
                {
                    return lstParameters;
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

                    return new List<strParameter>();
                }
            }
        }

        public void generateSampleCode() 
        {
            try
            {
                ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                string sWithTemp = "";

                int iIndent = 0;

                VBA.VBComponent vbComp;

                if (bPutSampleCallInOwnNewMod)
                {
                    vbComp = addModule(sSampleModuleName, VBA.vbext_ComponentType.vbext_ct_StdModule);

                    lstCode.Add(cSettings.Indent(iIndent) + "Option Explicit");
                    lstCode.Add(cSettings.Indent(iIndent) + "Option Base 1");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                else
                { vbComp = ClsMisc.ActiveVBComponent(); }


                lstCode.Add(cSettings.Indent(iIndent));
                //if (!cCodeMapper.cursorIsInFunction || bPutSampleCallInOwnNewMod)
                //{
                //    lstCode.Add(cSettings.Indent(iIndent));
                //    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + getNextSampleFunctionName());
                //    if (cSettings.IndentFirstLevel) { iIndent++; }
                //    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                //}
                if (!cCodeMapper.cursorIsInFunction || bPutSampleCallInOwnNewMod)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Class");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim cls as " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cls = new " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));

                if (lstParameters.Count == 0)
                {
                    if (cSettings.UserTips == true)
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + "'The Class was generated without any properties or methods");
                        lstCode.Add(cSettings.Indent(iIndent) + "'But if there were some then call them from here");
                    }
                }
                else
                {
                    if (this.UsingWith)
                    {
                        sWithTemp = "";
                        lstCode.Add(cSettings.Indent(iIndent) + "With cls");
                        iIndent++;
                    }
                    else
                    { sWithTemp = "cls"; }

                    foreach (strParameter objParameter in lstParameters.FindAll(x => x.bReadOnly == false))
                    {
                        string sTemp = cSettings.Indent(iIndent) + sWithTemp + "." + objParameter.sNameExtermal + " = ";

                        switch (cDataTypes.getGeneralType(objParameter.eType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                sTemp += "false";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                sTemp += "Now()";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sTemp += "0";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sTemp += "\"" + Environment.UserName + "\"";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                sTemp += "\"Unknown\"";
                                break;
                            default:
                                sTemp += "???";
                                break;
                        }
                        lstCode.Add(sTemp);
                    }

                    if (lstParameters.FindAll(x => x.bReadOnly == false).Count > 0)
                    { lstCode.Add(cSettings.Indent(iIndent)); }

                    foreach (strParameter objParameter in lstParameters)
                    { lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print " + sWithTemp + "." + objParameter.sNameExtermal); }

                    if (this.UsingWith)
                    {
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End With");
                        sWithTemp = "";
                    }
                    else
                    { sWithTemp = ""; }
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cls = nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction || bPutSampleCallInOwnNewMod)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent));

                this.addCode(ref lstCode, ref vbComp);

                if (!bPutSampleCallInOwnNewMod)
                { 
                    List<string> LstCodeOptions = new List<string>();

                    if (!cCodeMapper.hasOptionExplicit)
                    { LstCodeOptions.Add("Option Explicit"); }

                    if (!cCodeMapper.hasOptionBase)
                    { LstCodeOptions.Add("Option Base " + cSettings.defaultOptionBase); }

                    if (LstCodeOptions.Count > 0)
                    { 
                        LstCodeOptions.Add("");
                        this.addCode(ref LstCodeOptions, enumPosition.ePosBeginningAfterOptions);
                    }
                }

                lstCode = null;
                cSettings = null;
                //cCodeMapper = null;
                cDataTypes = null;
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

        public void generateClass()
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                int iIndent = 0;

                VBA.VBComponent vbComp = addModule(sClassName, VBA.vbext_ComponentType.vbext_ct_ClassModule);

                lstCode.Add(cSettings.Indent(iIndent) + "Option Explicit");
                lstCode.Add(cSettings.Indent(iIndent) + "Option Base 1");
                lstCode.Add(cSettings.Indent(iIndent));

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                foreach (strParameter objParameter in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Private " + objParameter.sNameInternal + " As " + cDataTypes.getName(objParameter.eType)); }

                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strParameter objParameter in lstParameters)
                {
                    if (!objParameter.bReadOnly)
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let " + objParameter.sNameExtermal + "(ByVal " + cDataTypes.typePrefix(objParameter.eType) + "Value As " + cDataTypes.getName(objParameter.eType) + ")");
                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + objParameter.sNameInternal + " = " + cDataTypes.typePrefix(objParameter.eType) + "Value");
                        lstCode.Add(cSettings.Indent(iIndent));
                        addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                        if (cSettings.IndentFirstLevel) { iIndent--; }
                        lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                        lstCode.Add(cSettings.Indent(iIndent));
                    }

                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get " + objParameter.sNameExtermal + "() As " + cDataTypes.getName(objParameter.eType));
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + objParameter.sNameExtermal +  " = " + objParameter.sNameInternal);
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Sub Class_Initialize()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strParameter objParameter in lstParameters)
                {
                    if (!string.IsNullOrEmpty(objParameter.sDefaultValue))
                    {
                        if (cSettings.UserTips == true)
                        {
                            if (cDataTypes.getGeneralType(objParameter.eType) == ClsDataTypes.enumGeneralDateType.eDate)
                            {
                                lstCode.Add(cSettings.Indent(iIndent));
                                lstCode.Add(cSettings.Indent(iIndent) + "'If a date is hardcoded into VBA you have to be extreamly careful of regional date formats");
                                lstCode.Add(cSettings.Indent(iIndent) + "'any code with hardcoded dates requires massive amounts of testing with dates in the ");
                                lstCode.Add(cSettings.Indent(iIndent) + "'first 12 days of the month as well as dates not in the first 12 days of the month.");
                                lstCode.Add(cSettings.Indent(iIndent) + "'Also don't assume that every user has the same region settings, even if there is a company policy.");
                            }
                        }

                        string sTempDefaultValueLine = cSettings.Indent(iIndent) + objParameter.sNameInternal + " = ";
                        
                        switch (cDataTypes.getGeneralType(objParameter.eType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                sTempDefaultValueLine += objParameter.sDefaultValue;
                                bool bTemp;
                                if (!bool.TryParse(objParameter.sDefaultValue, out bTemp))
                                { sTempDefaultValueLine += " 'Please make sure this value is a valid boolean value."; }
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                sTempDefaultValueLine += "#" + objParameter.sDefaultValue + "#";
                                DateTime dteTemp;
                                if (!DateTime.TryParse(objParameter.sDefaultValue, out dteTemp))
                                { sTempDefaultValueLine += " 'Please make sure this hard coded date is actually a valid date value."; }
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sTempDefaultValueLine += objParameter.sDefaultValue;

                                float fTemp;
                                if (!float.TryParse(objParameter.sDefaultValue, out fTemp))
                                { sTempDefaultValueLine += " 'Please make sure the value is a valid number"; }
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sTempDefaultValueLine += "\"" + objParameter.sDefaultValue + "\"";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                sTempDefaultValueLine += "\"" + objParameter.sDefaultValue + "\"";
                                break;
                            default:
                                sTempDefaultValueLine += "\"" + objParameter.sDefaultValue + "\"";
                                break;
                        }

                        lstCode.Add(sTempDefaultValueLine);
                        if (cSettings.UserTips == true)
                        {
                            if (cDataTypes.getGeneralType(objParameter.eType) == ClsDataTypes.enumGeneralDateType.eDate) 
                            {
                                lstCode.Add(cSettings.Indent(iIndent) + "'Be very careful of dates hardcoded in VBA and which date format it is,");
                                lstCode.Add(cSettings.Indent(iIndent) + "'make sure you do a check with a date in the first 12 days of the month");
                                lstCode.Add(cSettings.Indent(iIndent) + "'and a check with a date not in the first 12 days of the month.");
                            }
                        }
                    }
                }

                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Sub Class_Terminate()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));

                addCode(ref lstCode, ref vbComp);
                
                lstCode = null;
                cDataTypes = null;
                cSettings = null;
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

        public bool hasDupicateParameters
        {
            get
            {
                try
                {
                    List<string> lstParameterNames = new List<string>();
                    bool bIsFound = false;

                    foreach (strParameter objParameter in lstParameters)
                    {
                        if (lstParameterNames.Contains(objParameter.sNameExtermal, StringComparer.OrdinalIgnoreCase)) 
                        { bIsFound = true; }

                        lstParameterNames.Add(objParameter.sNameExtermal);
                    }

                    lstParameterNames = null;

                    return bIsFound;
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
                    return true;
                }
            }
        }
    }
}
