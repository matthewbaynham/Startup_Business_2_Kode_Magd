using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;
using KodeMagd.Settings;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_Files : ClsInsertCode
    {
        public enum enumDirection
        {
            eDir_Read,
            eDir_Write,
            eDir_Unknown
        }

        public struct strField_Delimited
        {
            public string sName;
            public int iPosition;
        }

        private enumDirection eDirection = enumDirection.eDir_Unknown;

        //private List<strField_Delimited> lstField_Delimited = new List<strField_Delimited>();
        private List<strFileFormat_FixedColumn> lstFields = new List<strFileFormat_FixedColumn>();
        private bool bHasSpecifiedColumnNames = false;

        public enum enumFileType
        {
            eDelimitedFile,
            eFixedColumnLengthFile,
            eUnknown
        }

        public struct strFileFormat_FixedColumn
        {
            public int iPosStart;
            public int iSize;
            public ClsDataTypes.vbVarType eDataType;
            public string sName;
            //public bool bEnabled;
        }

        public struct strFileFormat
        {
            public char cDelimiter;
            public List<strFileFormat_FixedColumn> lstColumns;
            public enumFileType eFileType;
        }

        private const string sPrefixFieldVariables = "gclCol_";
        //private const string sPrefixArrField = "clCol_";

        private const string csPrefix_Fso = "fso";
        private const string csPrefix_FileName = "sPath";
        private const string csPrefix_TextStream = "ts";
        
        private string sName = "";
        private string sFsoName = "";
        private string sTextStreamName = "";
        private string sFilePathName = ""; //Name of the variable containing the file path
        private string sFilePath = ""; //the actual string value of the file path.
        //private string sVariable = "";
        //private string sText = "";
        private string sFunctionName = "";
        private string sModuleName = "";

        private strFileFormat objFileFormat;

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

        public ClsInsertCode_Files() 
        {
            try
            {
                //lstFields = new List<strFields>();
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

        public strFileFormat fileFormat
        {
            get
            {
                try
                {
                    return objFileFormat;
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

                    return getNullFileFormat();
                }
            }
            set
            {
                try
                {
                    objFileFormat = value;

                    objFileFormat.lstColumns = objFileFormat.lstColumns.OrderBy(x => x.iPosStart).ToList();
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

        public enumDirection direction
        {
            get
            {
                try
                {
                    return eDirection;
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

                    return enumDirection.eDir_Unknown;
                }
            }
            set
            {
                try
                {
                    eDirection = value;
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

        //public void InsertCode_checkFileExists() 
        //{
        //    try
        //    {
        //        ClsCodeMapper cCodeMapper = new ClsCodeMapper();
        //        cCodeMapper.readCode();
        //        ClsSettings cSettings = new ClsSettings();
        //        List<string> lstCode = new List<string>();
        //        List<string> lstCodeTop = new List<string>();
        //        int iIndent = cCodeMapper.cursorCurrentIndentLevel();

        //        sModuleName = cCodeMapper.ModuleDetails.sName;

        //        if (!cCodeMapper.hasOptionExplicit)
        //        { lstCodeTop.Add("Option Explicit"); }

        //        if (!cCodeMapper.hasOptionBase)
        //        { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

        //        lstCode.Add(cSettings.Indent(iIndent));

        //        if (!cCodeMapper.cursorIsInFunction)
        //        {
        //            sFunctionName = getNextSampleFunctionName();
        //            lstCode.Add(cSettings.Indent(iIndent));
        //            lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
        //            if (cSettings.IndentFirstLevel) { iIndent++; }
        //            addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
        //        }
        //        else
        //        { sFunctionName = cCodeMapper.currentFunctionName(); }

        //        addTitleComment(ref lstCode, ref cSettings, iIndent);

        //        lstCode.Add(cSettings.Indent(iIndent) + "'Dimension Variables/Objects");

        //        lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFsoName + " As Scripting.FileSystemObject");
        //        lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFilePathName + " As String");
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = new Scripting.FileSystemObject");
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent) + sFilePathName + " = " + ClsMiscString.addQuotes(sFilePath));
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent) + "If " + sFsoName + ".fileExists(" + sFilePathName + ") then ");
        //        iIndent++;
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent) + "'Write all your code here if the file does exist");
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        iIndent--;
        //        lstCode.Add(cSettings.Indent(iIndent) + "End If");
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        lstCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = nothing");
        //        lstCode.Add(cSettings.Indent(iIndent));
        //        if (!cCodeMapper.cursorIsInFunction)
        //        {
        //            addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
        //            if (cSettings.IndentFirstLevel) { iIndent--; }
        //            lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
        //            lstCode.Add(cSettings.Indent(iIndent));
        //            lstCode.Add(cSettings.Indent(iIndent));
        //        }

        //        this.addCode(ref lstCode);

        //        if (lstCodeTop.Count > 0)
        //        {
        //            lstCodeTop.Add("");
        //            this.addCode(ref lstCodeTop, enumPosition.ePosBeginning);
        //        }

        //        cSettings = null;
        //        cCodeMapper = null;
        //        lstCode = null;
        //        lstCodeTop = null;
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
        //    }
        //}

        private void InsertCode_outputToTextFile(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = string.Empty;

                sModuleName = cCodeMapper.ModuleDetails.sName;

                objFileFormat.lstColumns = objFileFormat.lstColumns.OrderBy(x => x.iPosStart).ToList();

                const string csPrefixFieldVariables = "gclCol_";

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Output_To_Text_File");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "'Dimension Variables/Objects");

                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFsoName + " As Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sTextStreamName + " As Scripting.TextStream");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFilePathName + " As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsEnd As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sLine As String");

                int iOrder = 1;
                foreach (strFileFormat_FixedColumn objField_FixedColumn in lstFields.OrderBy(x => x.iPosStart))
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Const " + ClsMiscString.makeValidVarName(objField_FixedColumn.sName, sPrefixFieldVariables) + " As Long = " + iOrder.ToString());
                    iOrder++;
                }
                lstCode.Add(cSettings.Indent(iIndent));

                switch (objFileFormat.eFileType)
                {
                    case enumFileType.eDelimitedFile:
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim arrLine() As String");
                        break;
                    case enumFileType.eFixedColumnLengthFile:
                        lstCode.Add(cSettings.Indent(iIndent) + "Dim arrLine(0 to " + (objFileFormat.lstColumns.Count - 1).ToString() + ") As String");
                        lstCode.Add(cSettings.Indent(iIndent));

                        int iCounter = 1;
                        foreach (strFileFormat_FixedColumn objTemp in objFileFormat.lstColumns)
                        {
                            lstCode.Add(cSettings.Indent(iIndent) + "const " + ClsMiscString.makeValidVarName(objTemp.sName, sPrefixFieldVariables) + " As Long = " + iCounter.ToString());
                            iCounter++;
                        }
                        break;
                    default:
                        break;
                }
                
                if (objFileFormat.lstColumns.Count != 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent));

                    int iPos = 0;
                    foreach (strFileFormat_FixedColumn objField in objFileFormat.lstColumns.OrderBy(x => x.iPosStart))
                    {
                        lstCodeTop.Add("const " + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + " As long = " + iPos.ToString());
                        iPos++;
                    }
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = new Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sFilePathName + " = " + ClsMiscString.addQuotes(sFilePath));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sFsoName + ".FileExists(" + sFilePathName + ") Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"File already exists\", vbInformation, \"Ooops\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sTextStreamName + " = " + sFsoName + ".OpenTextFile(" + sFilePathName + ", ForWriting, True)");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sTextStreamName);
                    iIndent++;
                    sWithTemp = string.Empty;
                }
                else
                { sWithTemp = sTextStreamName; }

                lstCode.Add(cSettings.Indent(iIndent) + "bIsEnd = false");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Do While Not bIsEnd");
                iIndent++;
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "'************************************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                                                  *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Note: Write some code to populate arrLine with the values you wish to export   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                                                  *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'************************************************************************************");
                }

                foreach (strFileFormat_FixedColumn objField_FixedColumn in lstFields.OrderBy(x => x.iPosStart))
                {
                    lstCode.Add(cSettings.Indent(iIndent) + ClsMiscString.makeValidVarName(objField_FixedColumn.sName, sPrefixFieldVariables) + " = \"\"");
                }
                lstCode.Add(cSettings.Indent(iIndent));

                if (objFileFormat.lstColumns.Count != 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'Error check to check if the data is the correct datatype");
                    foreach (strFileFormat_FixedColumn objField in objFileFormat.lstColumns)
                    {
                        switch (cDataTypes.getGeneralType(objField.eDataType))
                        {
                            case ClsDataTypes.enumGeneralDateType.eBool:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not (Trim(lcase(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) = \"true\" or Trim(lcase(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) = \"false\" or isNumeric(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + ")))) then");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"The Field " + objField.sName + " is not a \"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eDate:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not isDate(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) then'Check is a Boolean");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"The Field " + objField.sName + " is not a Date\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                            case ClsDataTypes.enumGeneralDateType.eNumber:
                                lstCode.Add(cSettings.Indent(iIndent) + "If Not isNumeric(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) then'Check is a Boolean");
                                iIndent++;
                                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"The Field " + objField.sName + " is not a Number\"");
                                iIndent--;
                                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                                break;
                        }
                        lstCode.Add(cSettings.Indent(iIndent));
                    }
                }

                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                iIndent++;

                switch (objFileFormat.eFileType)
                {
                    case enumFileType.eDelimitedFile:
                        if (objFileFormat.cDelimiter.ToString() == "\t")
                        { lstCode.Add(cSettings.Indent(iIndent) + "sLine = join(arrLine, vbTab)"); }
                        else
                        { lstCode.Add(cSettings.Indent(iIndent) + "sLine = join(arrLine, \"" + objFileFormat.cDelimiter.ToString() + "\")"); }
                        break;
                    case enumFileType.eFixedColumnLengthFile:
                        lstCode.Add(cSettings.Indent(iIndent) + "sLine = \"\"");
                        foreach (strFileFormat_FixedColumn objTemp in objFileFormat.lstColumns)
                        { lstCode.Add(cSettings.Indent(iIndent) + "sLine = sLine & Left(arrLine(" + ClsMiscString.makeValidVarName(objTemp.sName, sPrefixFieldVariables) + ") & space(" + objTemp.iSize.ToString() + "), " + objTemp.iSize.ToString() + ")"); }
                        break;
                    default:
                        break;
                }
                
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".WriteLine sLine");

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print \"Log this Error somewhere: \" & sErrorMessage");
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'*******************************************************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                                                                     *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Note: Write code that reports the fact that the format of the data was not the expected format.   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Possibly a Msgbox, could be a log file, or a logged in a table, which ever suits you.             *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                                                                     *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*******************************************************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");

                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'***************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                             *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Note: Write some logic that will change bIsEnd to false   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   when you have exported all the rows you want to export    *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                             *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'***************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent) + "bIsEnd = true");
                lstCode.Add(cSettings.Indent(iIndent));

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Close");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = string.Empty;
                }
                else
                { sWithTemp = string.Empty; }

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sTextStreamName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFsoName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished\", vbInformation, \"Finished\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbCritical, \"Issue's\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
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
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeTop = null;
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
                switch (eDirection)
                {
                    case enumDirection.eDir_Read:
                        InsertCode_inputFromTextFile(ref cCodeMapper);
                        break;
                    case enumDirection.eDir_Write:
                        InsertCode_outputToTextFile(ref cCodeMapper);
                        break;
                    default:
                        MessageBox.Show("Unknown direction.  It should be either read or write.", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private void InsertCode_inputFromTextFile(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = string.Empty;

                sModuleName = cCodeMapper.ModuleDetails.sName;

                objFileFormat.lstColumns = objFileFormat.lstColumns.OrderBy(x => x.iPosStart).ToList();

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Files");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "'Dimension Variables/Objects");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFsoName + " As Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sTextStreamName + " As Scripting.TextStream");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sFilePathName + " As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOk As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");

                lstCode.Add(cSettings.Indent(iIndent) + "Dim arrLine() As String");

                if (cSettings.UserTips == true)
                { lstCode.Add(cSettings.Indent(iIndent) + "'The split function will set this array it's upper and lower bounds, and it'll be base zero regardless of option base"); }

                if (objFileFormat.lstColumns.Count != 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent));

                    int iFieldOrder = 0;
                    foreach (strFileFormat_FixedColumn objTemp in objFileFormat.lstColumns.OrderBy(x => x.iPosStart))
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + "const " + ClsMiscString.makeValidVarName(objTemp.sName, sPrefixFieldVariables) + " As Long = " + iFieldOrder.ToString());
                        iFieldOrder++;
                    }
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = new Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sFilePathName + " = " + ClsMiscString.addQuotes(sFilePath));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If " + sFsoName + ".FileExists(" + sFilePathName + ") Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sTextStreamName + " = " + sFsoName + ".OpenTextFile(" + sFilePathName + ", ForReading, False)");
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With " + sTextStreamName);
                    iIndent++;
                    sWithTemp = string.Empty;
                }
                else
                { sWithTemp = sTextStreamName; }

                lstCode.Add(cSettings.Indent(iIndent) + "Do While Not " + sWithTemp + ".AtEndOfStream");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sLine = " + sWithTemp + ".ReadLine");
                lstCode.Add(cSettings.Indent(iIndent));
                switch (objFileFormat.eFileType) 
                {
                    case enumFileType.eDelimitedFile:
                        if (objFileFormat.cDelimiter == '\t')
                        { lstCode.Add(cSettings.Indent(iIndent) + "arrLine = split(sLine, vbTab)"); }
                        else
                        { lstCode.Add(cSettings.Indent(iIndent) + "arrLine = split(sLine, \"" + objFileFormat.cDelimiter.ToString() + "\")"); }
                        break;
                    case enumFileType.eFixedColumnLengthFile:
                        foreach (strFileFormat_FixedColumn objTemp in objFileFormat.lstColumns)
                        { lstCode.Add(cSettings.Indent(iIndent) + "arrLine(" + ClsMiscString.makeValidVarName(objTemp.sName, sPrefixFieldVariables) + ") = mid(sLine, " + objTemp.iPosStart.ToString() + ", " + objTemp.iSize.ToString() + ")"); }
                        break;
                    default:
                        break;
                }
                lstCode.Add(cSettings.Indent(iIndent));

                if (this.objFileFormat.lstColumns.Count > 0)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'Error check to check if the data is the correct datatype");
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"\"");
                }

                foreach (strFileFormat_FixedColumn objField in this.objFileFormat.lstColumns)
                {
                    switch (cDataTypes.getGeneralType(objField.eDataType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eBool:
                            lstCode.Add(cSettings.Indent(iIndent) + "If Not (Trim(lcase(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) = \"true\" or Trim(lcase(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) = \"false\" or isNumeric(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + ")))) then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"The Field " + objField.sName + " is not a boolean\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate:
                            lstCode.Add(cSettings.Indent(iIndent) + "If Not isDate(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) then'Check is a Date");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"The Field " + objField.sName + " is not a Date\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            lstCode.Add(cSettings.Indent(iIndent) + "If Not isNumeric(Trim(arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + "))) then'Check is a Date");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = sErrorMessage & \"The Field " + objField.sName + " is not a Number\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                            break;
                    }
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'*********************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                           *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Check the line of data is OK here       *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   set bIsOkand sErrorMessage acordingly   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                           *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*********************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                iIndent++;
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Please change the contence of this IF so     *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   that it uses the data in the arrLine array   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'**************************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                foreach (strFileFormat_FixedColumn objField in objFileFormat.lstColumns)
                { lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print \"" + objField.sName + ": \" & arrLine(" + ClsMiscString.makeValidVarName(objField.sName, sPrefixFieldVariables) + ")"); }
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print \"Log this Error somewhere: \" & sErrorMessage");
                lstCode.Add(cSettings.Indent(iIndent));

                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'****************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                              *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Log an error use the text in the variable sErrorMessage.   *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*   Note it's an error in the format of the data.              *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                              *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'****************************************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Close");

                if (this.UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = string.Empty;
                }
                else
                { sWithTemp = string.Empty; }
                
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sTextStreamName + " = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + sFsoName + " = nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished\", vbInformation, \"Finished\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage, vbCritical, \"Issue's\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
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
                cDataTypes = null;
                lstCode = null;
                lstCodeTop = null;
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
                    string sTemp = value;

                    sName = ClsMiscString.makeValidVarName(sTemp);
                    sFsoName = ClsMiscString.makeValidVarName(sTemp, csPrefix_Fso);
                    sFilePathName = ClsMiscString.makeValidVarName(sTemp, csPrefix_FileName);
                    sTextStreamName = ClsMiscString.makeValidVarName(sTemp, csPrefix_TextStream);
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

        public string FullFilePath
        {
            get
            {
                try
                { return sFilePath; }
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
                {sFilePath = value; }
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

        public static strFileFormat getNullFileFormat()
        {
            try
            {
                strFileFormat objTemp;

                objTemp.cDelimiter = ' ';
                objTemp.eFileType = enumFileType.eUnknown;
                objTemp.lstColumns = new List<strFileFormat_FixedColumn>();

                return objTemp;
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

                //OK this is a little daft but nevermind we can't always be perfect
                strFileFormat objTemp;

                objTemp.cDelimiter = ' ';
                objTemp.eFileType = enumFileType.eUnknown;
                objTemp.lstColumns = new List<strFileFormat_FixedColumn>();

                return objTemp;
            }
        }
    }
}
