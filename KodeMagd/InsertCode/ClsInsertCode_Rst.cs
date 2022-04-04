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
using KodeMagd.InsertCode;

/*
    Option Explicit
    Option Base 1

    Public Sub Open_DB()
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim parOne As ADODB.Parameter
    Dim sConnectionString As String
    Dim sSql As String

    Set cmd = New ADODB.Command
    Set rst = New ADODB.Recordset
    Set parOne = New ADODB.Parameter

    sConnectionString = ""
    sSql = "SELECT * FROM <table Name>"

    cmd.ActiveConnection = sConnectionString
    cmd.CommandType = adCmdText
    cmd.CommandText = sSql

    Set parOne = cmd.CreateParameter()
    Call cmd.Parameters.Append(parOne)

    rst.Open cmd, , adOpenForwardOnly, adLockReadOnly

    Do While rst.EOF
    
        rst.MoveNext
    Loop

    rst.Close

    Set cmd = Nothing
    Set rst = Nothing
    Set parOne = Nothing

    End Sub
 
 */

namespace KodeMagd.InsertCode
{
    public class ClsInsertCode_Rst : ClsInsertCode
    {
        private const string csPrefix_Recordset = "rst";
        private const string csPrefix_Command = "cmd";
        private const string csPrefix_Parameter = "par";
        private const string csPrefix_Sql = "sql";

        private ADODB.CursorTypeEnum eRst_CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly;
        private ADODB.LockTypeEnum eRst_LockType = ADODB.LockTypeEnum.adLockReadOnly;
        private ADODB.CommandTypeEnum cmdType = ADODB.CommandTypeEnum.adCmdUnknown;
        private string sFunctionName = "";
        private string sModuleName = "";

        public enum enumDestinationType
        {
            eRstDest_Range,
            eRstDest_ListboxCombo,
            eRstDest_EmptyLoop,
            eRstDest_Unknown
        }

        public enum enumDestinationTypeRangeType 
        {
            eRng_Named,
            eRng_Coordinateds,
            eRng_Unknown
        }

        public struct strParameter
        {
            public string sName;
            public ClsDataTypes.vbVarType eVarType;
            public ADODB.DataTypeEnum eDataType;
            public bool bAssignVariable;
            public int Size;
            public string Value;
        }

        private string sName = "";  //this name will usually be used for most of the valiables (using different prefixes i.e. sName = "Fred" => sRstName = "rstFred", sCmdName = "cmdFred" )
        private string sCmdName = "";
        private string sRstName = "";
        private string sSqlName = "";
        private string sSql = "";
        private string sConnectionString = "";
        private List<strParameter> lstParameters = new List<strParameter>();

        private string sListboxComboboxName = "";
        private string sRstFieldName = "";
        private enumDestinationType eDestinationType = enumDestinationType.eRstDest_EmptyLoop;
        private enumDestinationTypeRangeType eDestinationTypeRangeType = enumDestinationTypeRangeType.eRng_Named;

        private int iDestinationRangeRow = 1;
        private int iDestinationRangeColumn = 1;
        private string sDestinationRangeName = "";
        private string sDestinationRangeShtName = "";
        private string sDestinationRangeWrkName = "";


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
        }

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

        public ADODB.LockTypeEnum LockType
        {
            get
            {
                try
                { return eRst_LockType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return ADODB.LockTypeEnum.adLockReadOnly;
                }
            }
            set
            {
                try
                { eRst_LockType = value; }
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

        public ADODB.CursorTypeEnum CursorType
        {
            get
            {
                try
                { return eRst_CursorType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return ADODB.CursorTypeEnum.adOpenForwardOnly;
                }
            }
            set
            {
                try
                { eRst_CursorType = value; }
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
                { return cmdType; }
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
                    cmdType = value;
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

        public string SQL
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
                {
                    sSql = value;
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

        public string Name
        {
            get
            {
                try
                { return sName; }
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
                    sCmdName = ClsMiscString.makeValidVarName(sName, csPrefix_Command);
                    sRstName = ClsMiscString.makeValidVarName(sName, csPrefix_Recordset);
                    sSqlName = ClsMiscString.makeValidVarName(sName, csPrefix_Sql);
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

        public void AddParameter(string sName, ADODB.DataTypeEnum eDataType, bool bAssignVariable, int iSize, string sValue) 
        { 
            try 
            {
                strParameter sTempPar = new strParameter();

                sTempPar.sName = sName;
                sTempPar.eDataType = eDataType;
                sTempPar.bAssignVariable = bAssignVariable;
                sTempPar.Size = iSize;
                sTempPar.Value = sValue;

                lstParameters.Add(sTempPar);
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

        public string ConnectionString
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

        public string ListboxComboboxName
        {
            get
            {
                try
                { return sListboxComboboxName; }
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
                    sListboxComboboxName = value;
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

        public string RstFieldName
        {
            get
            {
                try
                { return sRstFieldName; }
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
                    sRstFieldName = value;
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

        public enumDestinationTypeRangeType destinationTypeRangeType
        {
            get
            {
                try
                { return eDestinationTypeRangeType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumDestinationTypeRangeType.eRng_Coordinateds;
                }
            }
            set
            {
                try
                {
                    eDestinationTypeRangeType = value;
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

        public enumDestinationType destinationType
        {
            get
            {
                try
                { return eDestinationType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumDestinationType.eRstDest_EmptyLoop;
                }
            }
            set
            {
                try
                {
                    eDestinationType = value;
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

        public int destinationRangeRow
        {
            get
            {
                try
                { return iDestinationRangeRow; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return 1;
                }
            }
            set
            {
                try
                {
                    iDestinationRangeRow = value;
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

        public int destinationRangeColumn
        {
            get
            {
                try
                { return iDestinationRangeColumn; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return 1;
                }
            }
            set
            {
                try
                {
                    iDestinationRangeColumn = value;
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

        public string destinationRangeShtName
        {
            get
            {
                try
                { return sDestinationRangeShtName; }
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
                    sDestinationRangeShtName = value;
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

        public string destinationRangeWrkName
        {
            get
            {
                try
                { return sDestinationRangeWrkName; }
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
                    sDestinationRangeWrkName = value;
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

        public string destinationRangeName
        {
            get
            {
                try
                { return sDestinationRangeName; }
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
                    sDestinationRangeName = value;
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


        public enumDestinationTypeRangeType destinationRangeType
        {
            get
            {
                try
                { return eDestinationTypeRangeType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumDestinationTypeRangeType.eRng_Unknown;
                }
            }
            set
            {
                try
                {
                    eDestinationTypeRangeType = value;
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

        public ClsInsertCode_Rst() 
        {
            try 
            {
                sCmdName = "";
                sRstName = "";
                sName = "";
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

        public void Insert_RstLoop(ref ClsCodeMapper cCodeMapper)
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

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Open_Recordset");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                /*
                 * Dim
                 */
                lstCode.Add(cSettings.Indent(iIndent) + "'Dimension Variables/Objects");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sRstName + " As ADODB.Recordset");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sCmdName + " As ADODB.Command");
            
                foreach (strParameter objParameter in lstParameters.OrderBy(x => x.sName))
                { lstCode.Add(cSettings.Indent(iIndent) + "Dim " + ClsMiscString.makeValidVarName(objParameter.sName, csPrefix_Parameter) + " As ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sSqlName + " As String");

                if (lstParameters.Count > 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOK As Boolean");
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsOK = true");
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\" ");
                }
                lstCode.Add(cSettings.Indent(iIndent));

                /*
                 * Check Parameter Values 
                 */
                foreach (strParameter sTempPar in lstParameters.OrderBy(x => x.sName))
                {
                    switch (sTempPar.eDataType)
                    {
                        case ADODB.DataTypeEnum.adDate:
                        case ADODB.DataTypeEnum.adDBDate:
                        case ADODB.DataTypeEnum.adDBTime:
                        case ADODB.DataTypeEnum.adDBTimeStamp:
                        case ADODB.DataTypeEnum.adFileTime:
                            if (sTempPar.bAssignVariable)
                            { lstCode.Add(cSettings.Indent(iIndent) + "if not isdate(" + sTempPar.Value + ") then"); }
                            else
                            { lstCode.Add(cSettings.Indent(iIndent) + "if not isdate(#" + sTempPar.Value + "#) then"); }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = false");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"Value '" + sTempPar.Value + "' is not a date \" & vbCrLf");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            lstCode.Add(cSettings.Indent(iIndent));
                            break;
                        case ADODB.DataTypeEnum.adBigInt:
                        case ADODB.DataTypeEnum.adCurrency:
                        case ADODB.DataTypeEnum.adDecimal:
                        case ADODB.DataTypeEnum.adInteger:
                        case ADODB.DataTypeEnum.adSingle:
                        case ADODB.DataTypeEnum.adSmallInt:
                        case ADODB.DataTypeEnum.adTinyInt:
                        case ADODB.DataTypeEnum.adUnsignedBigInt:
                        case ADODB.DataTypeEnum.adUnsignedInt:
                        case ADODB.DataTypeEnum.adUnsignedSmallInt:
                        case ADODB.DataTypeEnum.adUnsignedTinyInt:
                        case ADODB.DataTypeEnum.adNumeric:
                        case ADODB.DataTypeEnum.adVarNumeric:
                        case ADODB.DataTypeEnum.adDouble:
                            lstCode.Add(cSettings.Indent(iIndent) + "if not IsNumeric(" + sTempPar.Value + ") then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = false");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"Value '" + sTempPar.Value + "' is not a number \" + vbCrLf");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            lstCode.Add(cSettings.Indent(iIndent));
                            break;
                        case ADODB.DataTypeEnum.adBoolean:
                            break;
                        case ADODB.DataTypeEnum.adLongVarBinary:
                        case ADODB.DataTypeEnum.adBinary:
                            break;
                        case ADODB.DataTypeEnum.adArray:
                        case ADODB.DataTypeEnum.adChapter:
                        case ADODB.DataTypeEnum.adEmpty:
                        case ADODB.DataTypeEnum.adError:
                        case ADODB.DataTypeEnum.adGUID:
                        case ADODB.DataTypeEnum.adIDispatch:
                        case ADODB.DataTypeEnum.adIUnknown:
                        case ADODB.DataTypeEnum.adPropVariant:
                        case ADODB.DataTypeEnum.adUserDefined:
                        case ADODB.DataTypeEnum.adVarBinary:
                        case ADODB.DataTypeEnum.adVariant:
                            break;
                        case ADODB.DataTypeEnum.adBSTR:
                        case ADODB.DataTypeEnum.adChar:
                        case ADODB.DataTypeEnum.adLongVarChar:
                        case ADODB.DataTypeEnum.adLongVarWChar:
                        case ADODB.DataTypeEnum.adVarChar:
                        case ADODB.DataTypeEnum.adVarWChar:
                        case ADODB.DataTypeEnum.adWChar:
                            if (sTempPar.bAssignVariable)
                            { lstCode.Add(cSettings.Indent(iIndent) + "if not len(CStr(" + sTempPar.Value + ")) > " + sTempPar.Size.ToString() + " then"); }
                            else
                            { lstCode.Add(cSettings.Indent(iIndent) + "if not len(CStr(\"" + sTempPar.Value + "\")) > " + sTempPar.Size.ToString() + " then"); }
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = false");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"Long of '" + sTempPar.Value + "' is too long\" ");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            lstCode.Add(cSettings.Indent(iIndent));
                            break;
                        default:
                            break;
                    }
                }

                if (lstParameters.Count > 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "if bIsOk then");
                    iIndent++;
                }

                /*
                 * Initialise Objects
                 */
                if (commandType == ADODB.CommandTypeEnum.adCmdText)
                {
                    //commandType == ADODB.CommandTypeEnum.adCmdText => SQL statement rather than a SP call
                    int iCountQuestionMarksWithSpacing = sSql.ToList().FindAll(x => x.ToString().Contains(" ? ")).Count();
                    int iCountQuestionMarks = sSql.ToList().FindAll(x => x == '?').Count();

                    if (iCountQuestionMarksWithSpacing != iCountQuestionMarks)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'When adding a ? as a parameter in a SQL string, it is best to make sure there is a space before and after the ?"); }

                    if (iCountQuestionMarks == lstParameters.Count)
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + sSqlName + " = " + ClsMiscString.addQuotes(sSql));
                        lstCode.Add(cSettings.Indent(iIndent) + "'It's good to the parameters have already been added to the SQL statement,");
                        lstCode.Add(cSettings.Indent(iIndent) + "'however please make sure that the order of the parameters in the SQL statement");
                        lstCode.Add(cSettings.Indent(iIndent) + "'is the same order as the order in which the VBA adds the parameters to the command object.");
                    }
                    else if (iCountQuestionMarks == 0)
                    {
                        string sSqlFull = sSql;
                        bool bFirstCondition = true;

                        foreach (strParameter objTempPar in lstParameters.OrderBy(x => x.sName))
                        {
                            if (bFirstCondition == true)
                            { sSqlFull += " WHERE ["; }
                            else
                            { sSqlFull += " AND ["; }
                            sSqlFull += objTempPar.sName.Trim() + "] = ? ";
                            bFirstCondition = false;
                        }
                        sSqlFull += ";";

                        lstCode.Add(cSettings.Indent(iIndent) + sSqlName + " = " + ClsMiscString.addQuotes(sSqlFull));
                    }
                    else
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + sSqlName + " = " + ClsMiscString.addQuotes(sSql));
                        lstCode.Add(cSettings.Indent(iIndent) + "'There are " + lstParameters.Count.ToString() + " created in the GUI and would would expect to find the same number of parameter in the SQL.");
                        lstCode.Add(cSettings.Indent(iIndent) + "'However there are " + iCountQuestionMarks.ToString() + " question marks in the SQL statement.");
                    }
                }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + sSqlName + " = " + ClsMiscString.addQuotes(sSql)); }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Initialise an instance of the objects");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = New ADODB.Command");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sRstName + " = New ADODB.Recordset");
                foreach (strParameter objParameter in lstParameters)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objParameter.sName, csPrefix_Parameter) + " = New ADODB.Parameter"); }

                lstCode.Add(cSettings.Indent(iIndent));
                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sCmdName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = sCmdName; }

                /*
                 * Set up connection
                 */
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".ActiveConnection = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(sConnectionString)));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandType = " + cmdType.ToString());
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".CommandText = " + sSqlName);
                lstCode.Add(cSettings.Indent(iIndent));

                foreach (strParameter objTempPar in lstParameters.OrderBy(x => x.sName))
                {
                    string sTempLine;

                    sTempLine = cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objTempPar.sName, csPrefix_Parameter) + " = " + sWithTemp + ".CreateParameter("; 
                    sTempLine += "\"" + objTempPar.sName + "\", "; 
                    sTempLine += objTempPar.eDataType.ToString() + ", "; 
                    sTempLine += "adParamInput, "; 
                    sTempLine += objTempPar.Size.ToString() + ", ";
                    
                    if (objTempPar.bAssignVariable)
                    { sTempLine += objTempPar.Value + ")"; }
                    else
                    { 
                        ClsDataTypes.enumGeneralDateType sGenType = cDataTypes.getGeneralType(objTempPar.eVarType);

                        switch(sGenType)
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                bool bResult;

                                if (bool.TryParse(objTempPar.Value, out bResult))
                                {
                                    if (bResult)
                                    { sTempLine += "true" + ")"; }
                                    else
                                    { sTempLine += "false" + ")"; }
                                
                                }
                                else
                                { sTempLine += objTempPar.Value + ")"; }

                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                sTempLine += "#" + objTempPar.Value + "#" + ")";
                                if (cSettings.UserTips == true)
                                { sTempLine += " 'WARNING: Please double check the date format and be very careful when it's not US date format"; }
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                sTempLine += objTempPar.Value + ")";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eString:
                                sTempLine += "\"" + objTempPar.Value + "\"" + ")";
                                break;
                            case ClsDataTypes.enumGeneralDateType.eUnknown:
                            default:
                                DateTime dteValue;
                                bool bValue;
                                float fValue;

                                if (DateTime.TryParse(objTempPar.Value, out dteValue))
                                { 
                                    sTempLine += "#" + objTempPar.Value + "#" + ")";
                                    if (cSettings.UserTips == true)
                                    { sTempLine += " 'WARNING: Please double check the date format and be very careful when it's not US date format"; }
                                }
                                else if (float.TryParse(objTempPar.Value, out fValue))
                                { sTempLine += objTempPar.Value + ")"; }
                                else if (bool.TryParse(objTempPar.Value, out bValue))
                                { sTempLine += objTempPar.Value + ")"; }
                                else
                                { sTempLine += "\"" + objTempPar.Value + "\"" + ")"; }
                                
                                break;
                        }
                    }

                    lstCode.Add(sTempLine);
                    lstCode.Add(cSettings.Indent(iIndent) + "Call " + sWithTemp + ".Parameters.Append(" + ClsMiscString.makeValidVarName(objTempPar.sName, csPrefix_Parameter) + ")");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }
                else
                { sWithTemp = ""; }

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sRstName);
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = sRstName; }

                /*
                 * Open Recordset
                 */
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Open " + sCmdName + ", , " + eRst_CursorType.ToString() + ", " + eRst_LockType.ToString());
                lstCode.Add(cSettings.Indent(iIndent));

                if (eDestinationType == enumDestinationType.eRstDest_Range)
                {
                    if (this.UsingWith)
                    {
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    }

                    string sTemp = cSettings.Indent(iIndent);

                    sTemp += "Worksheets(\"" + sDestinationRangeWrkName + "\")";

                    //sTemp += "Code not finished yet";
                    switch (eDestinationTypeRangeType)
                    {
                        case enumDestinationTypeRangeType.eRng_Coordinateds:
                            sTemp += ".Worksheets(\"" + sDestinationRangeShtName + "\")";
                            sTemp += ".Cells(" + iDestinationRangeRow.ToString() + "," + iDestinationRangeColumn.ToString() + ")";
                            break;
                        case enumDestinationTypeRangeType.eRng_Named:
                            sTemp += ".Range(\"" + sDestinationRangeName + "\")";
                            break;
                        default:
                            sTemp += "'Need to specify range here and then call ";
                            break;
                    }

                    sTemp += ".CopyFromRecordset " + sRstName;

                    lstCode.Add(sTemp);
                }
                else
                {
                    /*
                     * Loop through Recordset
                     */
                    if (eRst_CursorType == ADODB.CursorTypeEnum.adOpenForwardOnly && cSettings.UserTips == true)
                    { lstCode.Add(cSettings.Indent(iIndent) + "'Note: Recordset is Forward only therefore is already on first record we don't need to call MoveFirst"); }
                    else
                    { lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveFirst"); }
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "If not (" + sWithTemp + ".BOF and " + sWithTemp + ".BOF) Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "Do While Not " + sWithTemp + ".EOF");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));

                    switch (eDestinationType)
                    {
                        case enumDestinationType.eRstDest_EmptyLoop:
                            if (cSettings.UserTips == true)
                            {
                                lstCode.Add(cSettings.Indent(iIndent) + "'   *******************************************************************");
                                lstCode.Add(cSettings.Indent(iIndent) + "'   *        This is where you add code to look at the records        *");
                                lstCode.Add(cSettings.Indent(iIndent) + "'   *******************************************************************");
                                lstCode.Add(cSettings.Indent(iIndent) + "'   Note: The value of a field can be obtained from '" + sWithTemp + ".Fields(\"<Field Name>\").Value'");
                            }
                            break;
                        case enumDestinationType.eRstDest_ListboxCombo:
                            lstCode.Add(cSettings.Indent(iIndent) + "If isNull(" + sWithTemp + ".Fields(\"" + sRstFieldName + "\").Value) Then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "Call Me." + sListboxComboboxName + ".AddItem(\"<Empty>\")");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "Else");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "Call Me." + sListboxComboboxName + ".AddItem(" + sWithTemp + ".Fields(\"" + sRstFieldName + "\").Value)");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                    }

                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent));

                    lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".MoveNext");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                    
                    if (this.UsingWith)
                    {
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    }
                }


                /*
                 * Close Recordset
                 */
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Close Variables and deinitialise releasing resourses");
                lstCode.Add(cSettings.Indent(iIndent) + "If Not " + sRstName + " Is Nothing Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sRstName + ".State = adStateOpen Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + sRstName + ".Close");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));

                /*
                 * deinitialise Objects
                 */
                foreach (strParameter objParameter in lstParameters.OrderBy(x => x.sName))
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ClsMiscString.makeValidVarName(objParameter.sName, csPrefix_Parameter) + " = Nothing"); }

                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sRstName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sCmdName + " = Nothing");

                if (lstParameters.Count > 0)
                {
                    /*
                     * If we had an error display it
                     */
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "Else");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"Please resolve all these data issue's before trying again.\"");
                    if (cSettings.UserTips == true)
                    {
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "'It is recommended that you do not use the word \"Error\" as the title to ");
                        lstCode.Add(cSettings.Indent(iIndent) + "'this messagebox, because when a user see's the word \"Error\" they can ");
                        lstCode.Add(cSettings.Indent(iIndent) + "'panic and act irrationally without actually read the text that is displayed");
                    }
                    lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbCritical, \"Data Issue's\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }

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

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public bool isOk(ref List<string> lstMessages) 
        {
            try
            {
                bool bResult = true;

                switch (this.eDestinationType)
                {
                    case enumDestinationType.eRstDest_Range:
                        if (eDestinationTypeRangeType == enumDestinationTypeRangeType.eRng_Unknown)
                        {
                            bResult = false;
                            lstMessages.Add("Unknown Destination Range Type\n\nYou need to press the button 'Destination Details'\nand select the destination.");
                        }
                        else
                        {
                            switch (eDestinationTypeRangeType)
                            {
                                case enumDestinationTypeRangeType.eRng_Coordinateds:
                                    if (sDestinationRangeShtName == null | sDestinationRangeShtName.Trim() == "")
                                    {
                                        bResult = false;
                                        lstMessages.Add("Destination Sheet Name is blank");
                                    }

                                    if (iDestinationRangeRow == null | iDestinationRangeRow == 0)
                                    {
                                        bResult = false;
                                        lstMessages.Add("Destination Range Row is Blank");
                                    }

                                    if (iDestinationRangeColumn == null | iDestinationRangeColumn == 0)
                                    {
                                        bResult = false;
                                        lstMessages.Add("Destination Range Column is Blank");
                                    }
                                    break;
                                case enumDestinationTypeRangeType.eRng_Named:
                                    if (sDestinationRangeName == null | sDestinationRangeName.Trim() == "")
                                    {
                                        bResult = false;
                                        lstMessages.Add("Destination Named Range is blank");
                                    }
                                    break;
                            }
                        }
                        break;
                    case enumDestinationType.eRstDest_ListboxCombo:
                        if (sListboxComboboxName.Trim() == "")
                        {
                            bResult = false;
                            lstMessages.Add("No Combobox or listbox selected.");
                        }

                        if (sRstFieldName.Trim() == "")
                        {
                            bResult = false;
                            lstMessages.Add("Not source field selected.");
                        }

                        break;
                }

                return bResult;
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
    }
}