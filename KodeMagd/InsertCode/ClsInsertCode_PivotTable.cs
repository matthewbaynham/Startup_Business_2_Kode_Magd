using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using ADODB;
using KodeMagd.Misc;

/*
 
old = Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\Test Stuff\Database2.accdb;
ODBC;DBQ=C:\Test Stuff\Database2.accdb;Driver={Microsoft Access Driver (*.mdb, *.accdb)};DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=5;

select * from tblStuff

C:\Test Stuff\Database2.accdb

 */

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_PivotTable : ClsInsertCode
    {
        public enum enumPivotTblColOrientation
        {
            eColOri_Page,
            eColOri_ColumnTitle,
            eColOri_RowTitle,
            eColOri_Aggregate,
            eColOri_NotUsed
        }

        public enum enumSourceType
        {
            eSelectedRange,
            eNamedRange,
            eDatabase,
            eUnknown
        }
        
        public struct strPivotField 
        {
            public string sName;
            public enumPivotTblColOrientation eOrientation;
        }

        private List<strPivotField> lstPivotFields;
        private enumSourceType eSourceType;
        private ADODB.CommandTypeEnum eCmdType;
        private string sConnectionString;
        private string sSql;
        private bool bPromptBeforeDelete;
        private bool bDestinationNewSheet;
        private string sDestinationAddress;
        private string sDestinationSheetName;
        private string sFunctionName = "";
        private string sModuleName = "";
        //private string sConnectionStringWarning = "";

        public string functionName
        {
            get
            {
                try
                {
                    string sTemp = "";

                    if (sFunctionName == null)
                    { sTemp = ""; }
                    else
                    { sTemp = sFunctionName; }

                    return sTemp;
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
                    string sTemp = "";

                    if (sModuleName == null)
                    { sTemp = ""; }
                    else
                    { sTemp = sModuleName; }

                    return sTemp;
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
                    return "";
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
                    return "";
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

        public bool destinationNewSheet
        {
            get
            {
                try
                { return bDestinationNewSheet; }
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
                { bDestinationNewSheet = value; }
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

        public string destinationAddress
        {
            get
            {
                try
                { return sDestinationAddress; }
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
                { sDestinationAddress = value; }
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

        public string destinationSheetName
        {
            get
            {
                try
                { return sDestinationSheetName; }
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
                { sDestinationSheetName = value; }
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

        public ADODB.CommandTypeEnum commandType 
        {
            get 
            {
                try 
                {
                    return eCmdType;
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
                    return ADODB.CommandTypeEnum.adCmdUnknown;
                }
            }
            set 
            {
                try 
                {
                    eCmdType = value;
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
        
        public enumSourceType sourceType 
        {
            get 
            {
                try
                {
                    return eSourceType;
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
                    return enumSourceType.eUnknown;
                }
            }
            set 
            {
                try
                {
                    eSourceType  = value;
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

        public ClsInsertCode_PivotTable() 
        {
            try
            {
                lstPivotFields = new List<strPivotField>();
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

        public static string getNormalName(enumPivotTblColOrientation ePivotTblColOrientation)
        {
            try
            {
                string sResult = "";

                switch (ePivotTblColOrientation) 
                {
                    case enumPivotTblColOrientation.eColOri_Page:
                        sResult = "Page Title";
                        break;
                    case enumPivotTblColOrientation.eColOri_Aggregate:
                        sResult = "Aggregate";
                        break;
                    case enumPivotTblColOrientation.eColOri_ColumnTitle:
                        sResult = "Column Title";
                        break;
                    case enumPivotTblColOrientation.eColOri_NotUsed:
                        sResult = "Not Used";
                        break;
                    case enumPivotTblColOrientation.eColOri_RowTitle:
                        sResult = "Row Title";
                        break;
                    default:
                        sResult = "Unknown";
                        break;
                }

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
                return "";
            }
        }

        public void addPivotField(strPivotField objPivotField) 
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";

                if (bIsOk) 
                {
                    if (lstPivotFields == null)
                    { lstPivotFields = new List<strPivotField>(); }

                    if (lstPivotFields.Any(f => f.sName.Trim().ToUpper() == objPivotField.sName.Trim().ToUpper()))
                    {
                        bIsOk = false;
                        sMessage = "Field " + objPivotField.sName + " is listed twice";
                    }
                }

                if (bIsOk)
                {
                    if (objPivotField.sName != null)
                    { lstPivotFields.Add(objPivotField); }
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

        public void removePivotField(string sName)
        {
            try
            {
                
                bool bIsOk = true;
                string sMessage = "";

                if (bIsOk)
                {
                    if (!lstPivotFields.Any(f => f.sName.Trim().ToUpper() == sName.Trim().ToUpper()))
                    {
                        bIsOk = false;
                        sMessage = "Can't find Field " + sName;
                    }
                }

                if (bIsOk)
                {
                    int iIndex = lstPivotFields.FindIndex(f => f.sName.Trim().ToUpper() == sName.Trim().ToUpper());

                    lstPivotFields.RemoveAt(iIndex);
                }

                if (!bIsOk)
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

        public List<strPivotField> pivotFields 
        {
            get 
            {
                try
                {
                    if (lstPivotFields == null)
                    { return new List<strPivotField>(); }
                    else
                    { return lstPivotFields; }
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
                    return new List<strPivotField>();
                }
            }
        }

        public static ADODB.CommandTypeEnum convertCommandType(string sText) 
        { 
            try
            {
                ADODB.CommandTypeEnum eType = CommandTypeEnum.adCmdUnknown;

                foreach (ADODB.CommandTypeEnum eTemp in Enum.GetValues(typeof(ADODB.CommandTypeEnum)))
                {
                    if (eTemp.ToString() == sText)
                    { eType = eTemp; }

                }

                return eType;
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
                return CommandTypeEnum.adCmdUnknown;
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

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Pivot_Table");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                /*
                 Dim
                 */
                lstCode.Add(cSettings.Indent(iIndent) + "Dim objTable As Excel.PivotTable");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim objCache As Excel.PivotCache");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sht As Excel.Worksheet");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim pivTemp As Excel.PivotTable");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsFound As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk as boolean");

                switch (this.eSourceType)
                {
                    case enumSourceType.eDatabase:
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim wrkConn As Excel.WorkbookConnection");
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim wrkConnTemp As Excel.WorkbookConnection");
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim shtNew As Excel.Worksheet");
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim sTableDestination As String");
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csPivotTableName As String = \"PivotTableFromDB\"");
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csPivotTableDesciption As String = \"Connecting to DB\"");
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csDBConnName As String = \"DB_Comm\"");
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csSqlQuery As String = \"" + sSql + "\"");
                        if (cSettings.UserTips == true)
                        {
                            lstCode.Add(cSettings.Indent(iIndent));
                            lstCode.Add(cSettings.Indent(iIndent) + "'************************");
                            lstCode.Add(cSettings.Indent(iIndent) + "'*   Biggest headacke   *");
                            lstCode.Add(cSettings.Indent(iIndent) + "'************************");
                        }
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csConnectionString As String = \"" + ClsMisc.replaceReturnCharInQuotedTxtWithConst(connectionString) + "\"");
                        if (cSettings.UserTips == true)
                        {
                            lstCode.Add(cSettings.Indent(iIndent) + "'If the code below crashes first thing to check is this connection string");
                            lstCode.Add(cSettings.Indent(iIndent) + "'Just because a connection string works in ADO does not mean it'll work here");
                        }
                        break;
                    case enumSourceType.eNamedRange:
                    case enumSourceType.eSelectedRange:
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim rngDestination As Excel.Range");
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim rngSource As Excel.Range");
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim sTableDestination As String");
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "Const csPivotTableName As String = \"PivotTableFromDB\"");
                        break;
                    default:
                        break;
                }

                foreach (strPivotField objField in lstPivotFields.FindAll(x => x.eOrientation != enumPivotTblColOrientation.eColOri_NotUsed))
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim fld" + ClsMiscString.makeValidVarName(objField.sName) + " As PivotField"); }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = True");
                lstCode.Add(cSettings.Indent(iIndent));

                /*
                 Source
                 */
                switch (this.eSourceType)
                {
                    case enumSourceType.eDatabase:
                        lstCode.Add(cSettings.Indent(iIndent) + "'****************************");
                        lstCode.Add(cSettings.Indent(iIndent) + "'*   Create a connection    *");
                        lstCode.Add(cSettings.Indent(iIndent) + "'****************************");
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "For Each wrkConnTemp In ThisWorkbook.Connections");
                        iIndent++;
                        lstCode.Add(cSettings.Indent(iIndent) + "If wrkConnTemp.Name = csDBConnName Then");
                        iIndent++;
                        if (bPromptBeforeDelete) 
                        {
                            lstCode.Add(cSettings.Indent(iIndent) + "If MsgBox(\"Do you want to delete the old connection to the database?\", vbQuestion+vbYesNo, \"Delete\") = vbYes Then");
                            iIndent++;
                        }
                        lstCode.Add(cSettings.Indent(iIndent) + "wrkConnTemp.Delete");
                        if (bPromptBeforeDelete) 
                        {
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                        }
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End If");
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "Next wrkConnTemp");
                        lstCode.Add(cSettings.Indent(iIndent));
                        if (cSettings.UserTips == true)
                        {
                            lstCode.Add(cSettings.Indent(iIndent) + "'Check the connection string, just because it works for a ADODB object");
                            lstCode.Add(cSettings.Indent(iIndent) + "'is no guarantee it'll work for the Workbook Connection");
                        }
                        if (cSettings.UserTips == true)
                        {
                            lstCode.Add(cSettings.Indent(iIndent) + "'Connection Strings for the Excel.Connections can have differences different from ADODB Connection strings");
                            lstCode.Add(cSettings.Indent(iIndent) + "'Please try adding an item at the beginning to explain what type of string it is.");
                            lstCode.Add(cSettings.Indent(iIndent) + "'For example add \"ODBC;\" to the beginning of a odbc type of connection string");
                            lstCode.Add(cSettings.Indent(iIndent) + "'and it should work as a Excel connection string");
                        }
                        lstCode.Add(cSettings.Indent(iIndent) + "Set wrkConn = ThisWorkbook.Connections.Add(csDBConnName, csPivotTableDesciption, csConnectionString, csSqlQuery, XlConnectionType.xlConnectionTypeODBC)");
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "wrkConn.ODBCConnection.CommandType = xlCmdSql");
                        break;
                    case enumSourceType.eNamedRange:
                        lstCode.Add(cSettings.Indent(iIndent) + "Set rngSource = ThisWorkbook.Names(\"" + sSql + "\").RefersToRange");
                        break;
                    case enumSourceType.eSelectedRange:
                        lstCode.Add(cSettings.Indent(iIndent) + "Set rngSource = ThisWorkbook.Range(\"" + sSql + "\")");
                        break;
                    default:
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'***************************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Create destination    *");
                lstCode.Add(cSettings.Indent(iIndent) + "'***************************");
                lstCode.Add(cSettings.Indent(iIndent));
                if (bDestinationNewSheet)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Set sht = ThisWorkbook.Worksheets.Add()");
                    if (cSettings.UserTips == true)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'might want to add a check to see if sheet exists here"); }
                    lstCode.Add(cSettings.Indent(iIndent) + "sht.Name = \"" + sDestinationSheetName + "\"");
                }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + "Set sht = ThisWorkbook.Worksheets(\"" + sDestinationSheetName + "\")"); }
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                { lstCode.Add(cSettings.Indent(iIndent) + "'address needs to be in R1C1 so if we have A1 this will convert it"); }
                lstCode.Add(cSettings.Indent(iIndent) + "If InStr(sht.Name, \" \") = 0 Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sTableDestination = sht.Name & \"!\" & sht.Range(\"" + sDestinationAddress + "\").Address(ReferenceStyle:=xlR1C1)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sTableDestination = \"'\" & sht.Name & \"'!\" & sht.Range(\"" + sDestinationAddress + "\").Address(ReferenceStyle:=xlR1C1)");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'*****************************************************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Delete any pivot with same name as new pivot    *");
                lstCode.Add(cSettings.Indent(iIndent) + "'*****************************************************");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "For Each pivTemp In sht.PivotTables");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(pivTemp.Name)) = Trim(UCase(csPivotTableName)) Then");
                iIndent++;
                if (bPromptBeforeDelete)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "If MsgBox(\"Do you want to delete the old Pivot table?\", vbQuestion+vbYesNo, \"Delete\") = vbYes Then");
                    iIndent--;
                }
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = True");
                lstCode.Add(cSettings.Indent(iIndent) + "pivTemp.TableRange2.Delete");
                if (bPromptBeforeDelete) 
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next pivTemp");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = False");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "For Each sht In ThisWorkbook.Worksheets");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "For Each pivTemp In sht.PivotTables");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(pivTemp.Name)) = Trim(UCase(csPivotTableName)) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = True");
                if (this.bPromptBeforeDelete)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "If MsgBox(\"Do you want to delete the old pivot table?\", vbQuestion+vbYesNo, \"Delete\") = vbYes Then");
                    iIndent++;
                }
                lstCode.Add(cSettings.Indent(iIndent) + "pivTemp.TableRange2.Delete");
                if (this.bPromptBeforeDelete)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next pivTemp");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next sht");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'*************************");
                lstCode.Add(cSettings.Indent(iIndent) + "'*   Create new pivot    *");
                lstCode.Add(cSettings.Indent(iIndent) + "'*************************");
                lstCode.Add(cSettings.Indent(iIndent));

                switch (this.eSourceType)
                {
                    case enumSourceType.eDatabase:
                        lstCode.Add(cSettings.Indent(iIndent) + "Set objCache = ThisWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlExternal, SourceData:=wrkConn, Version:=xlPivotTableVersion14)");
                        break;
                    case enumSourceType.eNamedRange:
                    case enumSourceType.eSelectedRange:
                        lstCode.Add(cSettings.Indent(iIndent) + "Set objCache = ThisWorkbook.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlExternal, SourceData:=rngSource, Version:=xlPivotTableVersion14)");
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set objTable = objCache.CreatePivotTable(TableDestination:=sTableDestination, TableName:=csPivotTableName, DefaultVersion:=xlPivotTableVersion14)");
                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strPivotField fldTemp in lstPivotFields)
                {
                    string sFieldVariableName = ClsMiscString.makeValidVarName(fldTemp.sName, "fld");

                    switch (fldTemp.eOrientation)
                    {
                        case enumPivotTblColOrientation.eColOri_Page:
                            lstCode.Add(cSettings.Indent(iIndent));
                            if (cSettings.UserTips == true)
                            { lstCode.Add(cSettings.Indent(iIndent) + "'Add a page field for " + fldTemp.sName + " using the variable " + sFieldVariableName); }
                            lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFieldVariableName + " = objTable.PivotFields(\"" + fldTemp.sName + "\")");
                            lstCode.Add(cSettings.Indent(iIndent) + sFieldVariableName + ".Orientation = xlPageField");
                            break;
                        case enumPivotTblColOrientation.eColOri_RowTitle:
                            lstCode.Add(cSettings.Indent(iIndent));
                            if(cSettings.UserTips == true)
                            { lstCode.Add(cSettings.Indent(iIndent) + "'Add a row field for " + fldTemp.sName + " using the variable " + sFieldVariableName); }
                            lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFieldVariableName + " = objTable.PivotFields(\"" + fldTemp.sName + "\")");
                            lstCode.Add(cSettings.Indent(iIndent) + sFieldVariableName + ".Orientation = xlRowField");
                            break;
                        case enumPivotTblColOrientation.eColOri_ColumnTitle:
                            lstCode.Add(cSettings.Indent(iIndent));
                            if (cSettings.UserTips == true)
                            { lstCode.Add(cSettings.Indent(iIndent) + "'Add a column field for " + fldTemp.sName + " using the variable " + sFieldVariableName); }
                            lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFieldVariableName + " = objTable.PivotFields(\"" + fldTemp.sName + "\")");
                            lstCode.Add(cSettings.Indent(iIndent) + sFieldVariableName + ".Orientation = xlColumnField");
                            break;
                        case enumPivotTblColOrientation.eColOri_Aggregate:
                            lstCode.Add(cSettings.Indent(iIndent));
                            if (cSettings.UserTips == true)
                            { lstCode.Add(cSettings.Indent(iIndent) + "'Add the field for the aggregate Function (usually the field that gets totalled)"); }
                            lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFieldVariableName + " = objTable.PivotFields(\"" + fldTemp.sName + "\")");
                            lstCode.Add(cSettings.Indent(iIndent) + sFieldVariableName + ".Orientation = xlDataField");
                            lstCode.Add(cSettings.Indent(iIndent) + sFieldVariableName + ".Function = xlSum");
                            break;
                        case enumPivotTblColOrientation.eColOri_NotUsed:
                            break;
                        default:
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