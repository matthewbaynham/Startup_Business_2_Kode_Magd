using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_ExcuteStoredProcedure : ClsInsertCode
    {
        private string sName;
        private string sConnectionString;
        private bool bAsynchronously;
        private string sVbaFnName;

        public struct strParameter 
        {
            public ADODB.Parameter objParameter;
            public int iOrder;
            public string sVbaVariableName;
            public bool bIsVariable;
            public ClsDataTypes.vbVarType eVbaType;
        }

        List<strParameter> lstParameters;

        public string vbaFnName
        {
            get
            {
                try
                {
                    string sResult;
                    
                    if (sVbaFnName == null)
                    { sResult = string.Empty; }
                    else
                    { sResult = sVbaFnName; }

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

        public bool asynchronously
        {
            get
            {
                try
                {
                    return bAsynchronously;
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
                    bAsynchronously = value;
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

        public string connectionString
        {
            get
            {
                try
                {
                    return sConnectionString;
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
                    sConnectionString = value;
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

        public string storedProcedure
        {
            get
            {
                try
                {
                    return sName;
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
                    sName = value;
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

        public ClsInsertCode_ExcuteStoredProcedure() 
        {
            try
            {
                lstParameters = new List<strParameter>();
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

        public void addParameter(strParameter parItem)
        {
            try
            {
                lstParameters.Add(parItem);
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

        public void generateCode(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = "";
                ClsDataTypes cDataTypes = new ClsDataTypes();
                string sTempLine = "";

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));

                    sVbaFnName = getNextSampleFunctionName(ref cCodeMapper, "_Execute_Stored_Procedure");

                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sVbaFnName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sVbaFnName = cCodeMapper.cursorInFunctionName; }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                /*
                 * Dim
                 */
                for (int iIndex = 0; iIndex < lstParameters.Count; iIndex++)
                {
                    strParameter objTemp = lstParameters[iIndex];

                    string sDataType = cDataTypes.getName(cDataTypes.getDataType(lstParameters[iIndex].objParameter.Type));
                    string sPrefix = ClsMiscString.Left(ref sDataType, 1) + "Var";

                    if (objTemp.bIsVariable == false)
                    { objTemp.sVbaVariableName = ClsMiscString.makeValidVarName(objTemp.objParameter.Name, sPrefix); }
                    objTemp.eVbaType = cDataTypes.getDataType(objTemp.objParameter.Type);

                    lstParameters[iIndex] = objTemp;
                }

                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage as String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim cmd as Adodb.Command");

                foreach (strParameter objTemp in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + ClsMiscString.makeValidVarName(objTemp.objParameter.Name, "par") + " As Adodb.Parameter"); }
                foreach (strParameter objTemp in lstParameters.FindAll(x => x.bIsVariable == false))
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + objTemp.sVbaVariableName + " As " + cDataTypes.getName(objTemp.eVbaType)); }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Const csConnectionString as String = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(sConnectionString)));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = True");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));
                
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = New Adodb.Command");

                foreach (strParameter objTemp in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objTemp.objParameter.Name, "par") + " = New Adodb.Parameter"); }
                lstCode.Add(cSettings.Indent(iIndent));
                if (this.UsingWith)
                {
                    sWithTemp = "";
                    lstCode.Add(cSettings.Indent(iIndent) + "With cmd");
                    iIndent++;
                }
                else
                { sWithTemp = "cmd"; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = csConnectionString");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = adCmdStoredProc");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = \"" + sName.Trim() + "\"");

                foreach (strParameter objTemp in lstParameters)
                {
                    if (objTemp.bIsVariable == true)
                    {
                        lstCode.Add(cSettings.Indent(iIndent));

                        switch (cDataTypes.getGeneralType(objTemp.eVbaType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not (CStr(" + objTemp.sVbaVariableName + ") = \"False\" Or CStr(" + objTemp.sVbaVariableName + ") = \"True\") Then");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Please Check that value of '" + objTemp.sVbaVariableName + "' is actually a boolean\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not IsDate(" + objTemp.sVbaVariableName + ") Then");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Please Check that value of '" + objTemp.sVbaVariableName + "' is actually a date\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not IsNumeric(" + objTemp.sVbaVariableName + ") Then");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Please Check that value of '" + objTemp.sVbaVariableName + "' is actually a number\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Len(" + objTemp.sVbaVariableName + ") > " + objTemp.objParameter.Size.ToString() + " Then");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Please Check that value of '" + objTemp.sVbaVariableName + "' is not to long.\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                lstCode.Add(cSettings.Indent(iIndent));
                                break;
                        }
                    }

                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                    iIndent++;

                    sTempLine = cSettings.Indent(iIndent);

                    sTempLine += "set " + ClsMiscString.makeValidVarName(objTemp.objParameter.Name, "par") + " = ";
                    sTempLine += sWithTemp + ".CreateParameter(";
                    sTempLine += "\"" + objTemp.objParameter.Name + "\", ";
                    sTempLine += objTemp.objParameter.Type.ToString() + ", ";
                    sTempLine += objTemp.objParameter.Direction.ToString() + ", ";

                    if (objTemp.bIsVariable)
                    { sTempLine += objTemp.sVbaVariableName + ")"; }
                    else
                    {
                        switch (cDataTypes.getGeneralType(objTemp.eVbaType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                sTempLine += objTemp.objParameter.Value.ToString() + ")";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                try
                                { sTempLine += "#" + objTemp.objParameter.Value.ToString() + "#)"; }
                                catch
                                { sTempLine += "null) 'having problems with the date."; }
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sTempLine += objTemp.objParameter.Value.ToString() + ")";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sTempLine += "\"" + objTemp.objParameter.Value.ToString() + "\")";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                                sTempLine += objTemp.objParameter.Value.ToString() + ")";
                                break;
                        }
                    }

                    lstCode.Add(sTempLine);
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + ClsMiscString.makeValidVarName(objTemp.objParameter.Name, "par") + ")");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                sTempLine = cSettings.Indent(iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                iIndent++;

                sTempLine = cSettings.Indent(iIndent);
                sTempLine += "Call " + sWithTemp + ".Execute(";
                if (bAsynchronously) 
                { sTempLine += "adAsyncExecute"; }
                sTempLine += ")";

                lstCode.Add(sTempLine);


                foreach (strParameter objTemp in lstParameters.FindAll(x => x.bIsVariable == true 
                    && (x.objParameter.Direction == ADODB.ParameterDirectionEnum.adParamInputOutput 
                    || x.objParameter.Direction == ADODB.ParameterDirectionEnum.adParamOutput 
                    || x.objParameter.Direction == ADODB.ParameterDirectionEnum.adParamReturnValue)))
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "If IsNull(" + sWithTemp + ".Parameters(\"" + objTemp.objParameter.Name + "\").Value ) Then");
                    iIndent++;

                    string sTemp = cSettings.Indent(iIndent) + objTemp.sVbaVariableName + " = ";

                    switch(cDataTypes.getGeneralType(objTemp.eVbaType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eBool:
                            sTemp += "false";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate:
                            sTemp += "#00:00:00 30/12/1899#";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            sTemp += "0";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eString:
                            sTemp += "\"\"";
                            break;
                        default:
                            sTemp += "\"\"";
                            break;
                    }
                    lstCode.Add(sTemp);
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "Else");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + objTemp.sVbaVariableName + " = " + sWithTemp + ".Parameters(\"" + objTemp.objParameter.Name + "\").Value");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbCritical, \"Data Issue\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With"); 
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = Nothing");

                foreach (strParameter objTemp in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objTemp.objParameter.Name, "par") + " = Nothing"); }
                lstCode.Add(cSettings.Indent(iIndent));
                
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
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

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
    }
}
