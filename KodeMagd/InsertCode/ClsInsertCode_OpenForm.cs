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

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_OpenForm : ClsInsertCode
    {
        private string sFormName = "";
        private bool bNewForm = true;
        private char cDelimiter;
        private string sModuleCallCode = ""; //Where I'm calling the form from
        private string sFunctionCallCode = ""; //Where I'm calling the form from
        private string sInstanceName = "";

        public struct strParameter 
        {
            public string sNamePrivatelyInForm;
            public string sNamePublicOutsideForm;
            public ClsDataTypes.vbVarType eDataType;
            public string sValueGiveToParameter;
            public bool bIsVariable;
        }
        private List<strParameter> lstParameters = new List<strParameter>();

        //private string sModuleCallCode = ""; //Where I'm calling the form from
        //private string sFunctionCallCode = ""; //Where I'm calling the form from

        public string ModuleCallForm
        {
            get
            {
                try
                {
                    return sModuleCallCode;
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

        public string InstanceName
        {
            get
            {
                try
                {
                    return sInstanceName;
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
            set
            {
                try
                {
                    string sTemp = "";

                    if (value == null)
                    { sTemp = ""; }
                    else
                    { sTemp = value; }

                    sInstanceName = ClsMiscString.makeValidVarName(sTemp, "frm");
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

        public string FunctionCallForm
        {
            get
            {
                try
                {
                    return sFunctionCallCode;
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

        public bool isNewForm
        {
            get
            {
                try
                {
                    return bNewForm;
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
            set
            {
                try
                {
                    bNewForm = value;
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

        public string FormName 
        {
            get { 
                try
                {
                    return sFormName;
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
            set 
            { 
                try 
                {
                    sFormName = value; 
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

        public void addParameter(strParameter objParameter) 
        { 
            try
            {
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

        public void fixAmbiguousFieldNames()
        {
            try
            {
                string sPrefix = "";
                bool bIsAmbiguous = true;
                int iCounter = 1;

                while (bIsAmbiguous == true)
                {
                    bIsAmbiguous = false;


                    foreach (strParameter objParameter in lstParameters)
                    {
                        if (lstParameters.Exists(x => x.sNamePublicOutsideForm.Trim().ToLower() == sPrefix + objParameter.sNamePrivatelyInForm.Trim().ToLower()))
                        { bIsAmbiguous = true; }
                    }

                    if (bIsAmbiguous == true)
                    {
                        switch (sPrefix)
                        {
                            case "":
                                sPrefix = "l";
                                break;
                            case "l":
                                sPrefix = "p";
                                break;
                            case "p":
                                sPrefix = "t";
                                break;
                            default:
                                sPrefix = "l" + iCounter.ToString();
                                iCounter++;
                                break;
                        }
                    }
                }

                for (int iIndex = 0; iIndex < lstParameters.Count; iIndex++)
                {
                    strParameter objParameter = lstParameters[iIndex];

                    objParameter.sNamePrivatelyInForm = sPrefix + objParameter.sNamePrivatelyInForm;

                    lstParameters[iIndex] = objParameter;
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

        public void openForm(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = "";

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                sModuleCallCode = ClsMisc.ActiveVBComponent().Name;

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionCallCode = getNextSampleFunctionName(ref cCodeMapper, "_Open_Form");

                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionCallCode);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionCallCode = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sInstanceName + " As " + sFormName);
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                { lstCode.Add(cSettings.Indent(iIndent) + "'UserForm_Initialize is run here so none of the properties are set"); }

                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sInstanceName + " = New " + sFormName);
                lstCode.Add(cSettings.Indent(iIndent));
                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sInstanceName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = sInstanceName; }

                foreach (strParameter objTemp in lstParameters)
                {
                    if (objTemp.bIsVariable == true)
                    {
                        string sTemp = cSettings.Indent(iIndent);
                        sTemp += sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = " + objTemp.sValueGiveToParameter;

                        lstCode.Add(sTemp);
                    }
                    else
                    {
                        string sTemp = "";

                        switch (cDataTypes.getGeneralType(objTemp.eDataType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                if (objTemp.bIsVariable)
                                { lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(" + objTemp.sValueGiveToParameter + ") Then"); }
                                else
                                {
                                    bool bDummy;

                                    if (bool.TryParse(objTemp.sValueGiveToParameter, out bDummy))
                                    { lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(" + objTemp.sValueGiveToParameter + ") Then"); }
                                    else
                                    { 
                                        lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(\"" + objTemp.sValueGiveToParameter + "\") Then");

                                        if (cSettings.UserTips == true)
                                        { lstCode.Add(cSettings.Indent(iIndent) + "'Need to check if this hardcoded value is a boolean value."); }
                                    }
                                }
                                iIndent++;
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = ";

                                sTemp += objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);

                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                if (objTemp.bIsVariable)
                                { lstCode.Add(cSettings.Indent(iIndent) + "If IsDate(" + objTemp.sValueGiveToParameter + ") Then"); }
                                else
                                {
                                    DateTime dDummy;

                                    if (DateTime.TryParse(objTemp.sValueGiveToParameter, out dDummy))
                                    { lstCode.Add(cSettings.Indent(iIndent) + "If IsDate(#" + objTemp.sValueGiveToParameter + "#) Then"); }
                                    else
                                    { 
                                        lstCode.Add(cSettings.Indent(iIndent) + "If IsDate(\"" + objTemp.sValueGiveToParameter + "\") Then");

                                        if (cSettings.UserTips == true)
                                        { lstCode.Add(cSettings.Indent(iIndent) + "'Need to check if this hardcoded value is a date"); }
                                    }
                                }
                                iIndent++;
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = " + objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);
                                if (cSettings.UserTips == true)
                                { lstCode.Add(cSettings.Indent(iIndent) + "'One bug in VBA is all dates hardcoded in VBA have to be in US date format regardless of the machines regional settings, but in general dates formats are a total mess in VBA so make sure you test anything with dates very well"); }

                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                if (objTemp.bIsVariable)
                                { lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(" + objTemp.sValueGiveToParameter + ") Then"); }
                                else
                                {
                                    float fDummy;

                                    if (float.TryParse(objTemp.sValueGiveToParameter, out fDummy))
                                    { lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(" + objTemp.sValueGiveToParameter + ") Then"); }
                                    else
                                    { 
                                        lstCode.Add(cSettings.Indent(iIndent) + "If IsNumeric(\"" + objTemp.sValueGiveToParameter + "\") Then");

                                        if (cSettings.UserTips == true)
                                        { lstCode.Add(cSettings.Indent(iIndent) + "'Need to check if this hardcoded value is a numerical value."); }
                                    }
                                }
                                iIndent++;
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = ";

                                sTemp += objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);

                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = ";

                                sTemp += objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = ";

                                sTemp += objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);
                                break;
                            default:
                                sTemp = sWithTemp + "." + objTemp.sNamePublicOutsideForm + " = ";

                                sTemp += objTemp.sValueGiveToParameter;

                                lstCode.Add(cSettings.Indent(iIndent) + sTemp);
                                break;
                        }
                    }
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                if (cSettings.UserTips == true)
                { lstCode.Add(cSettings.Indent(iIndent) + "'UserForm_Activate is run here and all the properties are set"); }
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Show");
                
                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sInstanceName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                this.addCode(ref lstCode);

                if (lstCodeTop.Count > 0)
                {
                    lstCodeTop.Add("");
                    this.addCode(ref lstCodeTop, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                //cCodeMapper = null;
                lstCode = null;
                lstCodeTop = null;
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

        public void addForm(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeOptions = new List<string>();
                List<string> lstCodeLocals = new List<string>();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                VBA.VBComponent vbForm;

                int iIndent;// = cCodeMapper.cursorCurrentIndentLevel;
                bool bIsOK = true;

                if (bNewForm)
                {
                    iIndent = 0;
                    vbForm = addModule(sFormName, VBA.vbext_ComponentType.vbext_ct_MSForm);
                }
                else
                {
                    vbForm = getModule(sFormName);
                    //cCodeMapper.readCode(vbForm);
                    iIndent = cCodeMapper.cursorCurrentIndentLevel(vbForm.CodeModule.CodePane);
                }

                if (vbForm == null)
                { bIsOK = true; }

                if (bIsOK)
                {
                    if (bNewForm)
                    {
                        lstCodeOptions.Add(cSettings.Indent(iIndent) + "Option Explicit");
                        lstCodeOptions.Add(cSettings.Indent(iIndent) + "Option Base " + cSettings.defaultOptionBase);
                        lstCodeOptions.Add(cSettings.Indent(iIndent));
                    }
                    else
                    {
                        if (!cCodeMapper.hasOptionExplicit)
                        { lstCodeOptions.Add(cSettings.Indent(iIndent) + "Option Explicit"); }
                        if (!cCodeMapper.hasOptionBase)
                        { lstCodeOptions.Add(cSettings.Indent(iIndent) + "Option Base " + cSettings.defaultOptionBase); }
                        lstCodeOptions.Add(cSettings.Indent(iIndent));
                    }

                    addTitleComment(ref lstCode, ref cSettings, iIndent);

                    foreach (strParameter objTemp in lstParameters)
                    { lstCodeLocals.Add(cSettings.Indent(iIndent) + "Private " + objTemp.sNamePrivatelyInForm + " as " + cDataTypes.getName(objTemp.eDataType)); }
                    lstCode.Add(cSettings.Indent(iIndent));

                    if (!bNewForm)
                    {
                        addCode(ref lstCode, ref vbForm, enumPosition.ePosEnd);
                        lstCode.Clear(); 
                    }

                    foreach (strParameter objTemp in lstParameters)
                    {
                        string sLocalName = "";

                        switch (cDataTypes.getGeneralType(objTemp.eDataType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                sLocalName += "bFlag";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                sLocalName += "dDateTime";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sLocalName += "nNumber";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sLocalName += "sText";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                sLocalName += "objTemp";
                                break;
                            default:
                                sLocalName += "objTemp";
                                break;
                        }

                        lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let " + objTemp.sNamePublicOutsideForm + "(ByVal " + sLocalName + " As " + cDataTypes.getName(objTemp.eDataType) + ")");
                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + objTemp.sNamePrivatelyInForm + " = " + sLocalName);
                        lstCode.Add(cSettings.Indent(iIndent));
                        addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                        if (cSettings.IndentFirstLevel) { iIndent--; }
                        lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get " + objTemp.sNamePublicOutsideForm + "() As " + cDataTypes.getName(objTemp.eDataType));
                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + objTemp.sNamePublicOutsideForm + " = " + objTemp.sNamePrivatelyInForm);
                        lstCode.Add(cSettings.Indent(iIndent));
                        addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                        if (cSettings.IndentFirstLevel) { iIndent--; }
                        lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                        lstCode.Add(cSettings.Indent(iIndent));
                    }

                    lstCode.Add(cSettings.Indent(iIndent) + "");
                    //lstCode.Add(cSettings.Indent(iIndent) + "'UserForm_Initialize() is triggered when the command \"Set frm = New " + sFormName + "\" is run.");
                    //lstCode.Add(cSettings.Indent(iIndent) + "'UserForm_Activate() is triggered when the command \"Call frm.Show\" is run.");
                    lstCode.Add(cSettings.Indent(iIndent) + "");

                    if (bNewForm)
                    { addCode(ref lstCode, ref vbForm); }
                    else
                    { addCode(ref lstCode, ref vbForm, enumPosition.ePosEnd); }

                    addCode(ref lstCodeLocals, ref vbForm, enumPosition.ePosBeginningAfterOptions);
                    addCode(ref lstCodeOptions, ref vbForm, enumPosition.ePosBeginning);
                }

                cSettings = null;
                lstCode = null;
                lstCodeOptions = null;
                cDataTypes = null;
                //cCodeMapper = null;
                vbForm = null;
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
