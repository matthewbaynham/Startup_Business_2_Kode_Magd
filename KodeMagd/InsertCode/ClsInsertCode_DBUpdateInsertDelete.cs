using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using KodeMagd.Misc;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_DBUpdateInsertDelete : ClsInsertCode
    {
        public enum enumType 
        {
            eType_Insert,
            eType_Update,
            eType_Delete,
            eType_Unknown
        }

        public enum enumMethodology 
        {
            eMeth_SQL,
            eMeth_Recordset,
            eMeth_Unknown
        }

        public struct strField
        {
            public string sParameterName; // equals prefix + sName, used for the name of parameters
            public string sName; //Field name relating to table / add prefix "par" and is used as variable name for parameter object
            public string sVariableValue; //value or variable that parameter must equal
            public int iSize;
            public ADODB.DataTypeEnum eDataType;
            public bool bIsVariable;
            public bool bIsSelect;
            public bool bIsConditional;
            public bool bIsAuditCondition;
        }

        private List<strField> LstFields = new List<strField>();
        private string sConnectionString = "";
        private enumMethodology eMethodology = enumMethodology.eMeth_Unknown;
        private enumType eType = enumType.eType_Unknown;
        private string sTableName = "";
        private string sName = "";
        private bool bDoAuditCheck = true;
        private string sFunctionNameCount = "";
        private string sFunctionName = "";
        private string sModuleName = "";

        //private string sClassName;
        //private bool bPutSampleCallInOwnNewMod;

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

        public string functionNameCount
        {
            get
            {
                try
                { return sFunctionNameCount; }
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
                { sFunctionNameCount = value; }
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

        public bool doAuditCheck
        {
            get
            {
                try
                {
                    if (bDoAuditCheck == null)
                    { return true; }
                    else
                    { return bDoAuditCheck; }
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
                    bDoAuditCheck = value;
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

        public string name
        {
            get
            {
                try
                {
                    if (sName == null)
                    { return string.Empty; }
                    else
                    { return sName; }
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

        public string tableName
        {
            get
            {
                try
                {
                    if (sTableName == null)
                    { return string.Empty; }
                    else
                    { return sTableName; }
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
                    sTableName = value;
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
                    if (sConnectionString == null)
                    { return string.Empty; }
                    else
                    { return sConnectionString; }
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

        public void fieldsEmpty() 
        {
            try
            {
                LstFields = new List<strField>();
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

        public List<strField> fields 
        {
            get 
            {
                try
                {
                    return LstFields;
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
                    return new List<strField>();
                }
            }
        }

        public void fieldsAdd(strField objField)
        {
            try
            {
                LstFields.Add(objField);
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

        public void fixAmbiguousFieldNames(ref bool bIsOk, ref string sErrorMessage)
        {
            try
            {
                bIsOk = true;
                sErrorMessage = "";
                string sPrefix = "";
                bool bIsAmbiguous = true;
                int iCounter = 1;

                //if (bIsOk)
                //{
                //    if (LstFields.GroupBy(x => x.sVariableValue).Any(c => c.Count() > 1))
                //    {
                //        bIsOk = false;
                //        sErrorMessage = "Dupicate Variable Names";
                //    }
                //}

                if (bIsOk)
                {
                    if (LstFields.GroupBy(x => x.sName).Any(c => c.Count() > 1))
                    {
                        bIsOk = false;
                        sErrorMessage = "Dupicate Parameter Names";
                    }
                }

                //if (bIsOk)
                //{
                //    if (LstFields.GroupBy(x => x.sParameterName).Any(c => c.Count() > 1))
                //    {
                //        bIsOk = false;
                //        sErrorMessage = "Dupicate Parameter Names";
                //    }
                //}

                if (LstFields.GroupBy(x => x.sParameterName).Any(c => c.Count() > 1))
                {
                    List<string> lstDupeParaName = LstFields.Select(x => x.sParameterName).ToList<string>().GroupBy(x => x).Where(x => x.Count() > 1).Select(x => x.Key).Distinct().ToList<string>();

                    foreach (string sParameterName in lstDupeParaName) 
                    {
                        int iCount = 0;

                        for (int iIndex = 0; iIndex < LstFields.Count; iIndex++)
                        {
                            if (LstFields[iIndex].sParameterName.Trim().ToUpper() == sParameterName.Trim().ToUpper())
                            {
                                strField objFieldTemp = LstFields[iIndex];

                                iCount++;
                                objFieldTemp.sParameterName = sParameterName + iCount.ToString();

                                LstFields[iIndex] = objFieldTemp;
                            }
                        }
                    }
                }

                if (bIsOk)
                {
                    while (bIsAmbiguous == true)
                    {
                        bIsAmbiguous = false;

                        foreach (strField objField in LstFields)
                        {
                            if (objField.sVariableValue != null)
                            {
                                if (LstFields.Exists(x => x.sParameterName.Trim().ToLower() == sPrefix + objField.sVariableValue.Trim().ToLower()))
                                { bIsAmbiguous = true; }
                            }
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

                    for (int iIndex = 0; iIndex < LstFields.Count - 1; iIndex++)
                    {
                        strField objField = LstFields[iIndex];

                        objField.sVariableValue = sPrefix + objField.sVariableValue;

                        LstFields[iIndex] = objField;
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

        public enumType type
        {
            get
            {
                try
                {
                    if (eType == null)
                    { return enumType.eType_Unknown; }
                    else
                    { return eType; }
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
                    return enumType.eType_Unknown;
                }
            }
            set
            {
                try
                {
                    eType = value;
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

        public enumMethodology methodology
        {
            get
            {
                try
                {
                    if (eType == null)
                    { return enumMethodology.eMeth_Unknown; }
                    else
                    { return eMethodology; }
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
                    return enumMethodology.eMeth_Unknown;
                }
            }
            set
            {
                try
                {
                    eMethodology = value;
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


        public void generateCode(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                switch (this.eMethodology)
                {
                    case enumMethodology.eMeth_Recordset:
                        generateCodeRst(ref cCodeMapper);
                        break;
                    case enumMethodology.eMeth_SQL:
                        generateCodeSql(ref cCodeMapper);
                        break;
                    default:
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

        public void generateCodeRst(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeExtraFn = new List<string>();
                string sWithTemp = "";
                string sConnectionName = ClsMiscString.makeValidVarName(sName, "con");
                string sRecordsetName = ClsMiscString.makeValidVarName(sName, "rst");
                string sCmdName = ClsMiscString.makeValidVarName(sName, "cmd");
                int iIndent = 0;

                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();

                List<string> lstCodeOptions = new List<string>();

                sModuleName = cCodeMapper.ModuleDetails.sName;

                List<strField> lstParametersConditional = new List<strField>();
                List<strField> lstParametersSelect = new List<strField>();
                switch (this.eType)
                {
                    case enumType.eType_Delete:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Del");
                            lstParametersConditional.Add(objTemp);
                        }
                        break;
                    case enumType.eType_Insert:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Ins");
                            lstParametersSelect.Add(objTemp);
                        }
                        break;
                    case enumType.eType_Update:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Set");
                            lstParametersSelect.Add(objTemp);
                        }

                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Where");
                            lstParametersConditional.Add(objTemp);
                        }
                        break;
                }

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeOptions.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeOptions.Add("Option Base " + cSettings.defaultOptionBase); }

                /************************************
                 *    start off declaring things    *
                 ***********************************/
                if (!cCodeMapper.cursorIsInFunction)
                {
                    string sSuffix = "";
                    switch (this.type)
                    {
                        case enumType.eType_Delete:
                            sSuffix  = "Delete_DB";
                            break;
                        case enumType.eType_Insert:
                            sSuffix  = "Insert_DB";
                            break;
                        case enumType.eType_Update:
                            sSuffix  = "Update_DB";
                            break;
                        default:
                            sSuffix = "";
                            break;
                    }
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, sSuffix);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sConnectionName + " As ADODB.Connection");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sRecordsetName + " As ADODB.Recordset");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sSql As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sCmdName + " As ADODB.Command");

                foreach (strField objField in lstParametersConditional.Distinct().ToList<strField>())
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + objField.sParameterName + " As ADODB.Parameter"); }
                foreach (strField objField in lstParametersSelect.Distinct().ToList<strField>())
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + objField.sParameterName + " As ADODB.Parameter"); }

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim lCountBefore As Long");
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim lCountAfter As Long");
                }
                string sVarNameRecordAdjusted = "";
                switch (this.type)
                {
                    case enumType.eType_Delete:
                        sVarNameRecordAdjusted = "lRecordsDeleted";
                        break;
                    case enumType.eType_Insert:
                        sVarNameRecordAdjusted = "lRecordsAdded";
                        break;
                    case enumType.eType_Update:
                        sVarNameRecordAdjusted = "lRecordsUpdated";
                        break;
                    default:
                        sVarNameRecordAdjusted = "lRecordsAdjusted";
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sVarNameRecordAdjusted + " As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lSourceLineNo As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsRecordOK As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOK as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sMessage as string");
                lstCode.Add(cSettings.Indent(iIndent));
                /******************
                 *   Initialise   *
                 ******************/
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOK = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sConnectionName + " = New ADODB.Connection");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sRecordsetName + " = New ADODB.Recordset");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = New ADODB.Command");

                foreach (strField objField in lstParametersConditional.Distinct().ToList<strField>())
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = New ADODB.Parameter"); }
                foreach (strField objField in lstParametersSelect.Distinct().ToList<strField>())
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = New ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Const csConnectionString = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(sConnectionString)));
                lstCode.Add(cSettings.Indent(iIndent));

                /*********************
                 *   SQL Statement   *
                 ********************/
                string sSqlLine = "sSql = \"SELECT * FROM [" + tableName + "] WHERE ";

                foreach (strField objField in lstParametersConditional)
                { sSqlLine += "[" + objField.sName + "] = ? AND "; }

                if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 3) == "and")
                { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 3).TrimEnd() + " "; }

                if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 5) == "where")
                { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 5).TrimEnd() + " "; }

                lstCode.Add(sSqlLine);

                lstCode.Add(cSettings.Indent(iIndent));

                /*************************
                 *    Open connection    *
                 ************************/
                lstCode.Add(cSettings.Indent(iIndent) + "'Open connection");
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".ConnectionString = csConnectionString");
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".Open");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'set counter");
                    string sLineCallAuditCheck = cSettings.Indent(iIndent) + "lCountBefore = " + sFunctionNameCount + "(" + sConnectionName;

                    foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true & x.bIsVariable == true).Distinct().ToList<strField>())
                    { sLineCallAuditCheck += ", " + objField.sVariableValue.Trim(); }

                    sLineCallAuditCheck += ")";

                    lstCode.Add(sLineCallAuditCheck);
                }

                lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = 0");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'********************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Open Command   *");
                lstCode.Add(cSettings.Indent(iIndent) + "'********************");
                lstCode.Add(cSettings.Indent(iIndent));
                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sCmdName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = sCmdName; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = " + sConnectionName);
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = adCmdText");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = sSql");

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "lSourceLineNo = 0");
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   ADD START LOOP HERE  *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'In the event of looping though source data put the START of the loop here");
                    lstCode.Add(cSettings.Indent(iIndent) + "'and use lSourceLineNo to find data issues with source files");
                    lstCode.Add(cSettings.Indent(iIndent) + "'If there is a loop make sure that the parameters from the previous");
                    lstCode.Add(cSettings.Indent(iIndent) + "'time around aren't still in the command object.");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent) + "lSourceLineNo = lSourceLineNo + 1");
                lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = True");

                /********************************************************************************
                 *   Add Parameters to the cmd and check data as it's added to the parameters   *
                 ********************************************************************************/
                foreach (strField objField in lstParametersConditional)
                {
                    int iParSize = 0;
                    string sQuoteType = "";
                    lstCode.Add(cSettings.Indent(iIndent));

                    switch (cDataTypes.getGeneralType(objField.eDataType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eBool:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);
                            lstCode.Add(cSettings.Indent(iIndent));
                            sQuoteType = "";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);
                            if (objField.bIsVariable)
                            { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsDate(" + objField.sVariableValue + ") then"); }
                            else
                            {
                                if (cSettings.UserTips == true)
                                {
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Note: be very careful about hardcoding dates in Excel VBA");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'and pay particular attension to the US date format.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Make sure the code is tested for either for the first 12 days of the month.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'as well as testing for the days which are not in the first 12 of each month.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Don't assume the user will have the same regional settings as you.");
                                }
                                lstCode.Add(cSettings.Indent(iIndent) + "If IsDate(#" + objField.sVariableValue + "#) then");
                                sQuoteType = "#";
                            }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed date check on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);
                            lstCode.Add(cSettings.Indent(iIndent) + "If Not IsNumeric(" + objField.sVariableValue + ") then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed numeric check on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            sQuoteType = "";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eString:
                            iParSize = objField.iSize;
                            if (objField.bIsVariable)
                            { lstCode.Add(cSettings.Indent(iIndent) + "If Len(" + objField.sVariableValue + ") > " + objField.iSize.ToString() + " then"); }
                            else
                            {
                                lstCode.Add(cSettings.Indent(iIndent) + "If Len(\"" + objField.sVariableValue + "\") > " + objField.iSize.ToString() + " then");
                                sQuoteType = "\"";
                            }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed because the string was too long, on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        default:
                            sQuoteType = "";
                            break;
                    }
                    lstCode.Add(cSettings.Indent(iIndent));

                    string sTemp = cSettings.Indent(iIndent);

                    sTemp += "Set " + objField.sParameterName + " = ";
                    sTemp += sWithTemp;
                    sTemp += ".CreateParameter(\"\", ";
                    sTemp += objField.eDataType.ToString() + ", ";//Set par = .CreateParameter("", adDouble, adParamInput, 8, 12345)
                    sTemp += "adParamInput, ";
                    sTemp += iParSize.ToString() + ",";
                    sTemp += sQuoteType + objField.sVariableValue + sQuoteType + ")";

                    lstCode.Add(sTemp);
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + objField.sParameterName + ")");
                    if (objField.iSize == 0)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'WARNING: Don't have a parameter of zero size it'll just make the code crash"); }
                }

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'**********************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Open Recordset   *");
                lstCode.Add(cSettings.Indent(iIndent) + "'**********************");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    sWithTemp = "";
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sRecordsetName);
                    iIndent++;
                }
                else
                { sWithTemp = sRecordsetName; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Open " + sCmdName + ", , adOpenDynamic, adLockOptimistic");
                lstCode.Add(cSettings.Indent(iIndent));

                switch (this.type)
                {
                    case enumType.eType_Delete:
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveFirst");
                        lstCode.Add(cSettings.Indent(iIndent) + "Do While Not " + sWithTemp + ".EOF");
                        iIndent++;

                        lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = " + sVarNameRecordAdjusted + " + 1");
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Delete(adAffectCurrent)");
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".UpdateBatch");

                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveNext");
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                        break;
                    case enumType.eType_Insert:
                        lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = " + sVarNameRecordAdjusted + " + 1");
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".AddNew");
                        foreach (strField objField in lstParametersSelect)
                        {
                            string sAssignValue = cSettings.Indent(iIndent);

                            sAssignValue += sWithTemp + ".Fields(\"" + objField.sName + "\").Value = ";

                            if (objField.bIsVariable)
                            { sAssignValue += objField.sVariableValue; }
                            else
                            {
                                switch (cDataTypes.getGeneralType(objField.eDataType))
                                {
                                    case ClsDataTypes.enumGeneralDateType.eBool:
                                        sAssignValue += objField.sVariableValue;
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eDate:
                                        sAssignValue += "#" + objField.sVariableValue + "#";
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eNumber:
                                        sAssignValue += objField.sVariableValue;
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eString:
                                        sAssignValue += "\"" + objField.sVariableValue + "\"";
                                        break;
                                }
                            }
                            lstCode.Add(sAssignValue);
                        }
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Update");

                        break;
                    case enumType.eType_Update:
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveFirst");
                        lstCode.Add(cSettings.Indent(iIndent) + "Do While Not " + sWithTemp + ".EOF");
                        iIndent++;

                        lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = " + sVarNameRecordAdjusted + " + 1");
                        foreach (strField objField in lstParametersSelect)
                        {
                            string sAssignValue = cSettings.Indent(iIndent);

                            sAssignValue += sWithTemp + ".Fields(\"" + objField.sName + "\").Value = ";

                            if (objField.bIsVariable)
                            { sAssignValue += objField.sVariableValue; }
                            else
                            {
                                switch(cDataTypes.getGeneralType(objField.eDataType))
                                {
                                    case ClsDataTypes.enumGeneralDateType.eBool:
                                        sAssignValue += objField.sVariableValue;
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eDate:
                                        sAssignValue += "#" + objField.sVariableValue + "#";
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eNumber:
                                        sAssignValue += objField.sVariableValue;
                                        break;
                                    case ClsDataTypes.enumGeneralDateType.eString:
                                        sAssignValue += "\"" + objField.sVariableValue + "\"";
                                        break;
                                }
                            }
                            lstCode.Add(sAssignValue);
                        }
                        
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Update");
                        lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveNext");
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Close");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }

                lstCode.Add(cSettings.Indent(iIndent));
                if (this.bDoAuditCheck)
                {
                    string sLineCallAuditCheck = cSettings.Indent(iIndent) + "lCountAfter = " + sFunctionNameCount + "(" + sConnectionName;

                    foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true & x.bIsVariable == true).Distinct().ToList<strField>())
                    { sLineCallAuditCheck += ", " + objField.sVariableValue.Trim(); }

                    sLineCallAuditCheck += ")";

                    lstCode.Add(sLineCallAuditCheck);
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent) + "If Not " + sRecordsetName + " Is Nothing Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sRecordsetName + ".State = ADODB.adStateOpen Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + sRecordsetName + ".Close");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If Not " + sConnectionName + " Is Nothing Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sConnectionName + ".State = ADODB.adStateOpen Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".Close");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sRecordsetName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sConnectionName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strField objField in lstParametersConditional)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = Nothing"); }
                foreach (strField objField in lstParametersSelect)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = Nothing"); }

                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sRecordsetName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sConnectionName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'do audit check");
                    switch (this.eType)
                    {
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore + " + sVarNameRecordAdjusted + " = lCountAfter Then");
                            break;
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore - " + sVarNameRecordAdjusted + " = lCountAfter Then");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "'This really doesn't make any sense to check the count of records before and after an undate");
                            lstCode.Add(cSettings.Indent(iIndent) + "'You can prevent the creation of this check with a tickbox on the GUI.");
                            lstCode.Add(cSettings.Indent(iIndent) + "'Unless your updating a field that the Audit count is filtering on then your check will have to");
                            lstCode.Add(cSettings.Indent(iIndent) + "'either Add or subtract " + sVarNameRecordAdjusted + " on one side of this if statement.");
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore = lCountAfter Then");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore = lCountAfter Then");
                            break;
                    }

                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\", vbInformation, \"Finished\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "Else");
                    iIndent++;
                    switch (this.type)
                    {
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records deleted sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records added sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records updated sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records modified sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                    }
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                else
                {
                    switch (this.type)
                    {
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records deleted\", vbInformation, \"Finished\"");
                            break;
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records inserted\", vbInformation, \"Finished\"");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records updated\", vbInformation, \"Finished\"");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records affected\", vbInformation, \"Finished\"");
                            break;
                    }
                }
                lstCode.Add(cSettings.Indent(iIndent));
                
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                lstCode.Add(cSettings.Indent(iIndent));

                if (this.bDoAuditCheck)
                { generateCodeCountFn(ref lstCodeExtraFn, ref cSettings, ref iIndent, ref cDataTypes, sCmdName); }

                this.addCode(ref lstCode, ref vbComp);
                this.addCode(ref lstCodeOptions, ref vbComp, enumPosition.ePosBeginningAfterOptions);
                if (this.bDoAuditCheck)
                { this.addCode(ref lstCodeExtraFn, ref vbComp, enumPosition.ePosEnd); }

                cSettings = null;
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeExtraFn = null;
                lstCodeOptions = null;
                vbComp = null;
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

        public void generateCodeSql(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeExtraFn = new List<string>();
                string sWithTemp = "";
                string sConnectionName = ClsMiscString.makeValidVarName(sName, "con");
                string sCmdName = ClsMiscString.makeValidVarName(sName, "cmd");
                int iIndent = 0;

                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();

                List<string> lstCodeOptions = new List<string>();

                sModuleName = cCodeMapper.ModuleDetails.sName;

                List<strField> lstParameters = new List<strField>();
                switch (this.eType)
                {
                    case enumType.eType_Delete:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Del");
                            lstParameters.Add(objTemp);
                        }
                        break;
                    case enumType.eType_Insert:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Ins");
                            lstParameters.Add(objTemp);
                        }
                        break;
                    case enumType.eType_Update:
                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Set");
                            lstParameters.Add(objTemp);
                        }

                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        {
                            strField objTemp = objField;
                            objTemp.sParameterName = ClsMiscString.makeValidVarName(objField.sName, "par Where");
                            lstParameters.Add(objTemp);
                        }
                        break;
                }

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeOptions.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeOptions.Add("Option Base " + cSettings.defaultOptionBase); }

                /************************************
                 *    start off declaring things    *
                 ***********************************/
                if (!cCodeMapper.cursorIsInFunction)
                {
                    string sSuffix = "";
                    switch (this.type)
                    {
                        case enumType.eType_Delete:
                            sSuffix = "Delete_DB";
                            break;
                        case enumType.eType_Insert:
                            sSuffix = "Insert_DB";
                            break;
                        case enumType.eType_Update:
                            sSuffix = "Update_DB";
                            break;
                        default:
                            sSuffix = "";
                            break;
                    }
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, sSuffix);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sConnectionName + " As ADODB.Connection");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sSql As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sCmdName + " As ADODB.Command");

                foreach (strField objField in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + objField.sParameterName + " As ADODB.Parameter"); }

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim lCountBefore As Long");
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim lCountAfter As Long");
                }
                string sVarNameRecordAdjusted = "";
                switch (this.type)
                {
                    case enumType.eType_Delete:
                        sVarNameRecordAdjusted = "lRecordsDeleted";
                        break;
                    case enumType.eType_Insert:
                        sVarNameRecordAdjusted = "lRecordsAdded";
                        break;
                    case enumType.eType_Update:
                        sVarNameRecordAdjusted = "lRecordsUpdated";
                        break;
                    default:
                        sVarNameRecordAdjusted = "lRecordsAdjusted";
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sVarNameRecordAdjusted + " As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lSourceLineNo As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsRecordOK As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOK as Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sMessage as string");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOK = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sConnectionName + " = New ADODB.Connection");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = New ADODB.Command");

                foreach (strField objField in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = New ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Const csConnectionString = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(sConnectionString)));
                lstCode.Add(cSettings.Indent(iIndent));

                /*************************
                 *    Open connection    *
                 ************************/
                lstCode.Add(cSettings.Indent(iIndent) + "'Open connection");
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".ConnectionString = csConnectionString");
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".Open");
                lstCode.Add(cSettings.Indent(iIndent));
                
                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'set counter");
                    string sLineCallAuditCheck = cSettings.Indent(iIndent) + "lCountBefore = " + sFunctionNameCount + "(" + sConnectionName;

                    foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true & x.bIsVariable == true))
                    { sLineCallAuditCheck += ", " + objField.sVariableValue.Trim(); }

                    sLineCallAuditCheck += ")";
                    
                    lstCode.Add(sLineCallAuditCheck); 
                }

                lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = 0");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'*************************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Insert one Record   *");
                lstCode.Add(cSettings.Indent(iIndent) + "'*************************");

                /*****************************
                 *    declare SQL command    *
                 ****************************/
                string sSqlLine = cSettings.Indent(iIndent);

                switch (eType)
                {
                    case enumType.eType_Delete:
                        sSqlLine += "sSql = \"DELETE FROM [" + tableName + "] WHERE ";

                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        { sSqlLine += "[" + objField.sName + "] = ? AND "; }

                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 3) == "and")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 3).TrimEnd() + " "; }

                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 5) == "where")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 5).TrimEnd() + " "; }

                        lstCode.Add(sSqlLine);
                        break;
                    case enumType.eType_Insert:
                        sSqlLine += "sSql = \"INSERT INTO [" + tableName + "] (";

                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        { sSqlLine += "[" + objField.sName + "], "; }

                        if (ClsMiscString.Right(sSqlLine.TrimEnd(), 1) == ",")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 1) + " "; }

                        sSqlLine += ") VALUES (";

                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        { sSqlLine += " ? , "; }
                
                        if (ClsMiscString.Right(sSqlLine.TrimEnd(), 1) == ",")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 1); }

                        sSqlLine += ");";
                        lstCode.Add(sSqlLine);
                        break;
                    case enumType.eType_Update:
                        sSqlLine += "sSql = \"UPDATE [" + tableName + "] SET ";

                        foreach (strField objField in LstFields.FindAll(x => x.bIsSelect == true))
                        { sSqlLine += "[" + objField.sName + "] = ? , "; }
                        
                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 1) == ",")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 1).TrimEnd() + " "; }

                        sSqlLine += "WHERE ";

                        foreach (strField objField in LstFields.FindAll(x => x.bIsConditional == true))
                        { sSqlLine += "[" + objField.sName + "] = ? AND "; }
                        
                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 1) == ",")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 1).TrimEnd() + " "; }

                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 3) == "and")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 3).TrimEnd() + " "; }

                        if (ClsMiscString.Right(sSqlLine.Trim().ToLower(), 5) == "where")
                        { sSqlLine = ClsMiscString.Left(ref sSqlLine, sSqlLine.TrimEnd().Length - 5).TrimEnd() + " "; }

                        lstCode.Add(sSqlLine);
                        break;
                }

                lstCode.Add(cSettings.Indent(iIndent));

                /*************************
                 *    Open connection    *
                 ************************/
                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sCmdName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = sCmdName; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = " + sConnectionName);
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = adCmdText");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = sSql");

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "lSourceLineNo = 0");
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   ADD START LOOP HERE  *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'In the event of looping though source data put the START of the loop here");
                    lstCode.Add(cSettings.Indent(iIndent) + "'and use lSourceLineNo to find data issues with source files");
                    lstCode.Add(cSettings.Indent(iIndent) + "'If there is a loop make sure that the parameters from the previous");
                    lstCode.Add(cSettings.Indent(iIndent) + "'time around aren't still in the command object.");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent) + "lSourceLineNo = lSourceLineNo + 1");
                lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = True");


                /*
                 * use objField.sParameterName to name the parameter object.
                 */
                foreach (strField objField in lstParameters)
                {
                    int iParSize = 0;
                    string sQuoteType = "";
                    lstCode.Add(cSettings.Indent(iIndent));

                    switch(cDataTypes.getGeneralType(objField.eDataType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eBool:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);
                            lstCode.Add(cSettings.Indent(iIndent));
                            sQuoteType = "";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);
                            if (objField.bIsVariable)
                            {
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not IsDate(" + objField.sVariableValue + ") then");
                            }
                            else
                            {
                                if (cSettings.UserTips == true)
                                {
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Note: be very careful about hardcoding dates in Excel VBA");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'and pay particular attension to the US date format.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Make sure the code is tested for either for the first 12 days of the month.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'as well as testing for the days which are not in the first 12 of each month.");
                                    lstCode.Add(cSettings.Indent(iIndent) + "'Don't assume the user will have the same regional settings as you.");
                                }

                                //lstCode.Add(cSettings.Indent(iIndent) + "If IsDate(#" + objField.sVariableValue + "#) then");
                                DateTime dDummy;
                                if (DateTime.TryParse(objField.sVariableValue, out dDummy))
                                { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsDate(#" + objField.sVariableValue + "#) then"); }
                                else
                                { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsDate(#" + objField.sVariableValue + "#) then 'Note: you choose to assign a hardcoded value that doesn't appear to be a date, this might be worth a look"); }

                                sQuoteType = "#";
                            }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed date check on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            iParSize = ClsDataTypes.getDataTypeSize(objField.eDataType);

                            if (objField.bIsVariable == true)
                            { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsNumeric(" + objField.sVariableValue + ") then"); }
                            else
                            {
                                float fDummy;

                                if (float.TryParse(objField.sVariableValue, out fDummy))
                                { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsNumeric(" + objField.sVariableValue + ") then"); }
                                else
                                { lstCode.Add(cSettings.Indent(iIndent) + "If Not IsNumeric(\"" + objField.sVariableValue + "\") then 'Warning: you have opted to hardcode a value that is not a number"); }
                            }
                            
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed numeric check on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            sQuoteType = "";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eString:
                            iParSize = objField.iSize;
                            if (objField.bIsVariable)
                            { lstCode.Add(cSettings.Indent(iIndent) + "If Len(" + objField.sVariableValue + ") > " + objField.iSize.ToString() + " then"); }
                            else
                            { 
                                lstCode.Add(cSettings.Indent(iIndent) + "If Len(\"" + objField.sVariableValue + "\") > " + objField.iSize.ToString() + " then");
                                sQuoteType = "\"";
                            }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsRecordOK = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sMessage = \"Data Validation Issue: Field " + objField.sName + " failed because the string was too long, on line \" & CStr(lSourceLineNo) & \".\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        default:
                            sQuoteType = "";
                            break;
                    }
                    lstCode.Add(cSettings.Indent(iIndent));

                    string sTemp = cSettings.Indent(iIndent);

                    sTemp += "Set " + objField.sParameterName + " = ";
                    sTemp += sWithTemp;
                    sTemp += ".CreateParameter(\"\", ";
                    sTemp += objField.eDataType.ToString() + ", ";//Set par = .CreateParameter("", adDouble, adParamInput, 8, 12345)
                    sTemp += "adParamInput, ";
                    sTemp += iParSize.ToString() + ",";
                    sTemp += sQuoteType + objField.sVariableValue + sQuoteType + ")";

                    lstCode.Add(sTemp);
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + objField.sParameterName + ")");
                    if (objField.iSize == 0)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'WARNING: Don't have a parameter of zero size it'll just make the code crash"); }
                }
                
                lstCode.Add(cSettings.Indent(iIndent) + "");
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsRecordOK Then");
                iIndent++;

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Execute(adAsyncExecute)");
                    lstCode.Add(cSettings.Indent(iIndent) + sVarNameRecordAdjusted + " = " + sVarNameRecordAdjusted + " + 1");
                }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Execute()"); }

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "'Log record not being loaded");
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sMessage, vbExclamation, \"Data Validation Issue\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'*************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   ADD END LOOP HERE   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'In the event of looping though source data put the END of the loop here"); 
                }
                lstCode.Add(cSettings.Indent(iIndent));
                
                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'count records");
                    string sLineCallAuditCheck = cSettings.Indent(iIndent) + "lCountAfter = " + sFunctionNameCount + "(" + sConnectionName;
                    
                    foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true & x.bIsVariable == true))
                    { sLineCallAuditCheck += ", " + objField.sVariableValue.Trim(); }

                    sLineCallAuditCheck += ")";
                    
                    lstCode.Add(sLineCallAuditCheck); 
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'close everything");
                lstCode.Add(cSettings.Indent(iIndent) + sConnectionName + ".Close");
                lstCode.Add(cSettings.Indent(iIndent));

                //foreach (strField objField in LstFields)
                //{ lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objField.sName, "par") + " = Nothing"); }
                foreach (strField objField in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + objField.sParameterName + " = Nothing"); }
                
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sConnectionName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.bDoAuditCheck)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'do audit check");
                    switch (this.eType)
                    {
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore + " + sVarNameRecordAdjusted + " = lCountAfter Then");
                            break;
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore - " + sVarNameRecordAdjusted + " = lCountAfter Then");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "'This really doesn't make any sense to check the count of records before and after an undate");
                            lstCode.Add(cSettings.Indent(iIndent) + "'You can prevent the creation of this check with a tickbox on the GUI.");
                            lstCode.Add(cSettings.Indent(iIndent) + "'Unless your updating a field that the Audit count is filtering on then your check will have to");
                            lstCode.Add(cSettings.Indent(iIndent) + "'either Add or subtract " + sVarNameRecordAdjusted + " on one side of this if statement.");
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore = lCountAfter Then");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "If lCountBefore = lCountAfter Then");
                            break;
                    }

                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\", vbInformation, \"Finished\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "Else");
                    iIndent++;
                    switch(this.type)
                    {
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records deleted sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records added sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records updated sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Audit Check Failed\" & vbCrLf & \"Not all of the records modified sussessfully made it in to the database.\", vbCritical, \"Audit Check Failed\"");
                            break;
                    }
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                else
                {
                    switch (this.type)
                    {
                        case enumType.eType_Delete:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records deleted\", vbInformation, \"Finished\"");
                            break;
                        case enumType.eType_Insert:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records inserted\", vbInformation, \"Finished\"");
                            break;
                        case enumType.eType_Update:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records updated\", vbInformation, \"Finished\"");
                            break;
                        default:
                            lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Finished\" & vbCrLf & CStr(" + sVarNameRecordAdjusted + ") & \": Records affected\", vbInformation, \"Finished\"");
                            break;
                    }
                }
                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                if (this.bDoAuditCheck)
                { generateCodeCountFn(ref lstCodeExtraFn, ref cSettings, ref iIndent, ref cDataTypes, sCmdName); }

                this.addCode(ref lstCode, ref vbComp);
                this.addCode(ref lstCodeOptions, ref vbComp, enumPosition.ePosBeginningAfterOptions);
                if (this.bDoAuditCheck)
                { this.addCode(ref lstCodeExtraFn, ref vbComp, enumPosition.ePosEnd); }

                cSettings = null;
                //cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeExtraFn = null;
                lstCodeOptions = null;
                vbComp = null;
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

        private void generateCodeCountFn(ref List<string> lstCode, ref ClsSettings cSettings, ref int iIndent, ref ClsDataTypes cDataTypes, string sCmdName)
        {
            try
            {
                string sWithTemp = "";

                lstCode.Add(cSettings.Indent(iIndent));

                string sLineDeclareFunction = cSettings.Indent(iIndent);

                sLineDeclareFunction += "Private Function " + sFunctionNameCount + "(ByRef con As ADODB.Connection";

                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true & x.bIsVariable == true))
                {
                    string sParameterDeclare = ", ";

                    if (cDataTypes.getGeneralType(objField.eDataType) == ClsDataTypes.enumGeneralDateType.eUnknown)
                    { sParameterDeclare += "ByRef "; }
                    else
                    { sParameterDeclare += "ByVal "; }

                    sParameterDeclare += objField.sVariableValue + " As ";

                    switch (cDataTypes.getGeneralType(objField.eDataType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eBool:
                            sParameterDeclare += "Boolean";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate:
                            sParameterDeclare += "Date";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            switch (objField.eDataType)
                            {
                                case ADODB.DataTypeEnum.adBigInt:
                                case ADODB.DataTypeEnum.adInteger:
                                case ADODB.DataTypeEnum.adSmallInt:
                                case ADODB.DataTypeEnum.adTinyInt:
                                case ADODB.DataTypeEnum.adUnsignedBigInt:
                                case ADODB.DataTypeEnum.adUnsignedInt:
                                case ADODB.DataTypeEnum.adUnsignedSmallInt:
                                case ADODB.DataTypeEnum.adUnsignedTinyInt:
                                    sParameterDeclare += "Long";
                                    break;
                                default:
                                    sParameterDeclare += "Double";
                                    break;
                            }
                            break;
                        case ClsDataTypes.enumGeneralDateType.eString:
                            sParameterDeclare += "String";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eUnknown:
                            sParameterDeclare += "Variant";
                            break;
                    }

                    sLineDeclareFunction += sParameterDeclare;
                }

                sLineDeclareFunction += ") As Long";

                //lstCodeExtraFn.Add(cSettings.Indent(iIndent) + "Private Function countRecords(ByRef con As ADODB.Connection) As Long");
                lstCode.Add(sLineDeclareFunction);

                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);

                addTitleComment(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "Dim rst As ADODB.Recordset");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim cmd As ADODB.Command");
                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true))
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + ClsMiscString.makeValidVarName(objField.sName, "par") + " As ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent) + "Dim sSql As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lResult As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "");
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = New ADODB.Command");
                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true))
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objField.sName, "par") + " = New ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent));

                string sSqlAuditLine = cSettings.Indent(iIndent) + "sSql = \"SELECT count(*) as recordCount FROM [" + sTableName + "] WHERE ";

                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true))
                { sSqlAuditLine += " [" + objField.sName + "] = ? AND "; }

                //remove last "AND" on that or if no fields remove "WHERE"
                if (ClsMiscString.Right(sSqlAuditLine.Trim().ToLower(), 3) == "and")
                { sSqlAuditLine = ClsMiscString.Left(ref sSqlAuditLine, sSqlAuditLine.TrimEnd().Length - 3).TrimEnd() + " "; }

                if (ClsMiscString.Right(sSqlAuditLine.Trim().ToLower(), 5) == "where")
                { sSqlAuditLine = ClsMiscString.Left(ref sSqlAuditLine, sSqlAuditLine.TrimEnd().Length - 5).TrimEnd() + " "; }

                sSqlAuditLine += "\"";

                lstCode.Add(sSqlAuditLine);

                //sWithTemp
                /******************
                 *   START WITH   *
                 ****************** 
                 */

                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With cmd");
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "cmd"; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = con");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = adCmdText");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = sSql");
                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true))
                {
                    string sCreateParameter = cSettings.Indent(iIndent);

                    sCreateParameter += "Set " + ClsMiscString.makeValidVarName(objField.sName, "par");
                    sCreateParameter += " = " + sWithTemp + ".CreateParameter(\"\", " + objField.eDataType.ToString() + ", ";
                    sCreateParameter += "adParamInput, " + objField.iSize.ToString() + ", ";

                    if (objField.iSize == 0)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'WARNING: Don't have a parameter of zero size it'll just make the code crash"); }

                    if (objField.bIsVariable)
                    { sCreateParameter += objField.sVariableValue; }
                    else
                    {
                        switch (cDataTypes.getGeneralType(objField.eDataType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                sCreateParameter += objField.sVariableValue;
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                sCreateParameter += "#" + objField.sVariableValue + "#";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sCreateParameter += objField.sVariableValue;
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sCreateParameter += "\"" + objField.sVariableValue + "\"";
                                break;
                        }
                    }

                    sCreateParameter += ")";

                    lstCode.Add(sCreateParameter);

                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + ClsMiscString.makeValidVarName(objField.sName, "par") + ")");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }
                else
                { sWithTemp = ""; }

                /**************** 
                 *   END WITH   *
                 ****************/
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set rst = New ADODB.Recordset");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With rst");
                    iIndent++;
                    sWithTemp = "rst";
                }
                else
                { sWithTemp = sCmdName; }

                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Open cmd, , adOpenForwardOnly, adLockReadOnly");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sWithTemp + ".BOF And " + sWithTemp + ".EOF Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If IsNull(" + sWithTemp + ".Fields(\"recordCount\").Value) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = True");
                lstCode.Add(cSettings.Indent(iIndent) + "lResult = " + sWithTemp + ".Fields(\"recordCount\").Value");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Close");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }
                else
                { sWithTemp = ""; }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set rst = Nothing");
                foreach (strField objField in LstFields.FindAll(x => x.bIsAuditCondition == true))
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objField.sName, "par") + " = Nothing"); }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'return value");
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "countRecords = lResult");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "countRecords = 0");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));

                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Function);
                if (cSettings.IndentFirstLevel) { iIndent--; }

                lstCode.Add(cSettings.Indent(iIndent) + "End Function");
                lstCode.Add(cSettings.Indent(iIndent));
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

        //public string nextCountRecordsFunctionName(ref ClsCodeMapper cCodeMapper, string sPrefix)
        //{
        //    try
        //    {
        //        //string sName = "countRecords";
        //        int iCounter = 1;
        //        string sResult = "";
        //        string sTempName;

        //        if (cCodeMapper.getLstFunctionNames().FindAll(x => x == sPrefix).Count == 0)
        //        { sResult = sPrefix; }
        //        else
        //        {
        //            sTempName = sPrefix;
        //            while (cCodeMapper.getLstFunctionNames().FindAll(x => x == sTempName).Count != 0)
        //            {
        //                iCounter++;
        //                sTempName = sPrefix + iCounter;
        //            }
        //        }

        //        return sResult;
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

        //        return string.Empty;
        //    }
        //}
    }
}
