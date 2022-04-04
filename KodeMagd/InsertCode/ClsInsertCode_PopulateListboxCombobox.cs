using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;

namespace KodeMagd.InsertCode
{
    public class ClsInsertCode_PopulateListboxCombobox : ClsInsertCode
    {
        private string sControlName = "";
        private string sValue = "";
        private string sConnectionString = "";
        private string sFieldName = "";
        private ADODB.CommandTypeEnum eCmdType = ADODB.CommandTypeEnum.adCmdUnknown;
        private string sSql = "";
        private string sSheetName = "";
        private string sAddress = "";
        private string sNamedRange = "";
        private string sFunctionName = "";
        private string sModuleName = "";

        public enum enumSource
        {
            eSource_Recordset,
            eSource_OneValue,
            eSource_Array,
            eSource_RangeNamed,
            eSource_RangeByAddress,
            eSource_Unknown
        }

        public struct strParameter
        {
            public string sName;
            public ADODB.DataTypeEnum eDataType;
            public ADODB.ParameterDirectionEnum eDirection;
            public long lSize;
            public string sValue;
        }

        List<strParameter> lstParameters = new List<strParameter>();
        /*
                 *                        Name Optional. A String value that contains the name of the Parameter object.
                 *                        Type Optional. A DataTypeEnum value that specifies the data type of the Parameter object.
                 *                        Direction Optional. A ParameterDirectionEnum value that specifies the type of Parameter object.
                 *                        Size Optional. A Long value that specifies the maximum length for the parameter value in characters or bytes.
                 *                        Value Optional. A Variant that specifies the value for the Parameter object.
         */

        List<string> lstArray = new List<string>();

        private enumSource eSourceType = enumSource.eSource_Unknown;

        public string functionName
        {
            get
            {
                try
                { return sFunctionName; }
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
                { return sModuleName; }
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
                { return lstParameters; }
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
            set
            {
                try
                { lstParameters = value; }
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

        public enumSource sourceType
        {
            get
            {
                try
                { return eSourceType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumSource.eSource_Unknown;
                }
            }
            set
            {
                try
                { eSourceType = value; }
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

        public string address
        {
            get
            {
                try
                { return sAddress; }
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
                { sAddress = value; }
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

        public string fieldName
        {
            get
            {
                try
                { return sFieldName; }
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
                { sFieldName = value; }
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

        public string sheetName
        {
            get
            {
                try
                { return sSheetName; }
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
                { sSheetName = value; }
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

        public string NamedRange
        {
            get
            {
                try
                { return sNamedRange; }
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
                { sNamedRange = value; }
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

        public string sql
        {
            get
            {
                try
                { return sSql; }
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
                { sSql = value; }
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

        public ADODB.CommandTypeEnum cmdType
        {
            get
            {
                try
                { return eCmdType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return ADODB.CommandTypeEnum.adCmdUnknown;
                }
            }
            set
            {
                try
                { eCmdType = value; }
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
                { return sConnectionString; }
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
                { sConnectionString = value; }
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

        public string ControlName {
            get 
            { 
                try 
                { return sControlName; }
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
                { sControlName = value; }
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

        public string Value
        {
            get
            {
                try
                { return sValue; }
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
                { sValue = value; }
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

        public void arrayClear()
        {
            try
            {
                lstArray = new List<string>();
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

        public void arrayAdd(string sItem)
        {
            try
            {
                lstArray.Add(sItem);
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

        public List<string> array 
        {
            get 
            {
                try
                {
                    return lstArray;
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

                    return new List<string>();
                }
            }
        }

        public void generateCode(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                switch (sourceType)
                {
                    case enumSource.eSource_Array:
                        generateArray(ref cCodeMapper);
                        break;
                    case enumSource.eSource_OneValue:
                        generateOneValue(ref cCodeMapper);
                        break;
                    case enumSource.eSource_RangeByAddress:
                        generateRangeAddress(ref cCodeMapper);
                        break;
                    case enumSource.eSource_RangeNamed:
                        generateRangeNamed(ref cCodeMapper);
                        break;
                    case enumSource.eSource_Recordset:
                        generateRecordset(ref cCodeMapper);
                        break;
                    case enumSource.eSource_Unknown:
                        //?????????????????????????????????????
                        break;
                    default:
                        //?????????????????????????????????????
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

        public void generateRecordset(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sWithTemp = "";
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Populate_ListBox_CombBox_Recordset");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim cmd As Adodb.Command");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim rst As Adodb.Recordset");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = New Adodb.Command");
                lstCode.Add(cSettings.Indent(iIndent) + "Set rst = New Adodb.Recordset");
                lstCode.Add(cSettings.Indent(iIndent));
                
                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With cmd");
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "cmd"; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(connectionString)));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = " + cmdType);
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = " + ClsMiscString.addQuotes(sSql));

                foreach(strParameter objParameter in lstParameters)
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    
                    string sVarName = ClsMiscString.makeValidVarName(objParameter.sName, "par");

                    lstCode.Add(cSettings.Indent(iIndent) + "Set " + sVarName + " = " + sWithTemp + ".CreateParameter (\"" + objParameter.sName + "\", " + objParameter.eDataType.ToString() + ", " + objParameter.eDirection.ToString() + ", " + objParameter.lSize.ToString() + ", Value)");
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + sVarName + ")");
                }

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }
                
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "rst.Open cmd, , adOpenForwardOnly, adLockReadOnly");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With Me." + sControlName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "Me." + sControlName; }

                lstCode.Add(cSettings.Indent(iIndent) + "Do While " + sWithTemp + ".ListCount > 0");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".RemoveItem(0)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "do while not rst.Eof");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "if IsError(rst.Fields(\"" + sFieldName + "\").Value) then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & vbcrlf & \"Error in Data\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".AddItem(rst.fields(\"" + sFieldName + "\").Value)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent) + "rst.MoveNext");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set rst = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished loading control\", vbInformation, \"Finished\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbExclamation, \"Warning\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If ");
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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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

        public void generateArray(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sWithTemp = "";
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                int iCounter = 0;

                sModuleName = cCodeMapper.ModuleDetails.sName;
                
                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper,"Populate_ListBox_ComboBox_Array");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim arr(1 to " + lstArray.Count.ToString() +  ") As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim iPos As Long");
                lstCode.Add(cSettings.Indent(iIndent));

                iCounter = 1;
                foreach (string sItem in lstArray)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "arr(" + iCounter.ToString() + ") = \"" + sItem + "\"");
                    iCounter++;
                }

                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With Me." + sControlName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "Me." + sControlName; }


                lstCode.Add(cSettings.Indent(iIndent) + "Do While " + sWithTemp + ".ListCount > 0");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".RemoveItem(0)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "for iPos = lbound(arr) to ubound(arr)");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".AddItem(arr(iPos))");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next iPos");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }
                
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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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

        public void generateOneValue(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sWithTemp = "";
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Populate_Listbox_CombBox_One_Value");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With Me." + sControlName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "Me." + sControlName; }


                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".AddItem(" + sValue + ")");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }

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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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

        public void generateRangeNamed(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sWithTemp = "";
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                int iCounter = 0;

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Populate_Listbox_ComboBox_Named_Range");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim rng As Excel.Range");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With Me." + sControlName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "Me." + sControlName; }

                lstCode.Add(cSettings.Indent(iIndent) + "Do While " + sWithTemp + ".ListCount > 0");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".RemoveItem(0)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "for each rng in ThisWorkbook.Names(\"" + sNamedRange + "\").RefersToRange");
                
                //
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "if IsError(rng.Value) then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & vbcrlf & \"Error in Cell \" + rng.Address");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".AddItem(rng.Value)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next rng");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished loading control\", vbInformation, \"Finished\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbExclamation, \"Warning\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If ");
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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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

        public void generateRangeAddress(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sWithTemp = "";
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                int iCounter = 0;

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Populate_Listbox_ComboBox_Range");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim rng As Excel.Range");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With Me." + sControlName.Trim());
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "Me." + sControlName.Trim(); }


                lstCode.Add(cSettings.Indent(iIndent) + "Do While " + sWithTemp + ".ListCount > 0");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".RemoveItem(0)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "for each rng in ThisWorkbook.Worksheets(\"" + sSheetName + "\").Range(\"" + sAddress + "\")");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "if IsError(rng.Value) then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & vbcrlf & \"Error in Cell \" + rng.Address");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".AddItem(rng.Value)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next rng");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished loading control\", vbInformation, \"Finished\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbExclamation, \"Warning\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If ");
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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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

        public void generateAppend(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Populate_Listbox_ComboBox");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Call Me." + sControlName + ".AddItem(\"" + sValue + "\")");
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

                lstCode = null;
                lstCodeTop = null;
                cSettings = null;
                //cCodeMapper = null;
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
