using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using KodeMagd.Settings;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_TextFileLogClass : ClsInsertCode
    {
        public enum enumAutoPath 
        {
            eAutoPath_Manual,
            eAutoPath_Date
        }

        public struct strVariablesToLog {
            public string sName;
            //public ClsDataTypes.vbVarType eDataType;
            public string sDataType;
            public bool bOptional;
        }

        private enumAutoPath eAutoPathGeneration = enumAutoPath.eAutoPath_Date;
        private List<strVariablesToLog> lstVariablesToLog = new List<strVariablesToLog>();

        private string sPath; //If automatic path is folder if manual path is file
        private char cDelimiter;
        private string sDateFormat_FileName;
        private string sDateFormat_FileContents;
        private string sClassName = "";
        private string sExtension = ".txt";
        private string sFunctionNameCallLog = "";
        private string sModuleNameCallLog = "";

        public List<strVariablesToLog> variablesToLog
        {
            get
            {
                try
                {
                    return lstVariablesToLog;
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

                    return new List<strVariablesToLog>();
                }
            }
        }

        public string functionNameCallLog
        {
            get
            {
                try
                {
                    return sFunctionNameCallLog;
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

        public string moduleNameCallLog
        {
            get
            {
                try
                {
                    return sModuleNameCallLog;
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

        public void addParameter(string sName, string sDataType, bool bOptional) 
        {
            try 
            {
                strVariablesToLog objVar = new strVariablesToLog();
                    
                objVar.sName = sName;
                objVar.sDataType = sDataType;
                objVar.bOptional = bOptional;

                lstVariablesToLog.Add(objVar);
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

        public void addParameter(strVariablesToLog objVar)
        {
            try
            {
                lstVariablesToLog.Add(objVar);
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

        public string DateFormat_FileName
        {
            get
            {
                try
                {
                    return sDateFormat_FileName;
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
                    sDateFormat_FileName = value;
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

        public string DateFormat_FileContents
        {
            get
            {
                try
                {
                    return sDateFormat_FileContents;
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
                    sDateFormat_FileContents = value;
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

        public enumAutoPath PathGeneration
        {
            get 
            {
                try
                {
                    return eAutoPathGeneration;
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

                    return enumAutoPath.eAutoPath_Date;
                }
            }
            set 
            {
                try
                {
                    eAutoPathGeneration = value;
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

        public string Path
        {
            get
            {
                try
                {
                    return sPath;
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
                    string sTemp = value;
                    sTemp = sTemp.Trim();

                    switch (eAutoPathGeneration) 
                    {
                        case enumAutoPath.eAutoPath_Date:
                            //this should be a folder path address and so must end in back slash
                            if (ClsMiscString.Right(ref sTemp, 1) != "\\")
                            { sTemp = sTemp + '\\'; }
                            break;
                        case enumAutoPath.eAutoPath_Manual:
                            //this should be a file path
                            break;
                        default:
                            break;
                    }


                    sPath = sTemp;
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

        public string ClassName
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

                    return string.Empty;
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

        public void CallLog(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sTempLine;

                //this.sFunctionNameCallLog 
                sModuleNameCallLog = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionNameCallLog = getNextSampleFunctionName(ref cCodeMapper, "_Log_to_Text_File_Class");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionNameCallLog);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionNameCallLog = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "'Declaring Class");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim cLog as " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Initialising Class");
                lstCode.Add(cSettings.Indent(iIndent) + "set cLog = new " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'write a long complicated process...");
                lstCode.Add(cSettings.Indent(iIndent) + "'Sometimes calling routine");
                lstCode.Add(cSettings.Indent(iIndent) + "'Note: [ ] square brackets indicate optional parameters");

                sTempLine = "'Call cLog.Log(";
                foreach (strVariablesToLog objVar in lstVariablesToLog)
                {
                    //do the compulsary ones first
                    if (!objVar.bOptional)
                    { sTempLine += "<" + objVar.sName + ">, "; }
                }
                foreach (strVariablesToLog objVar in lstVariablesToLog)
                {
                    //add the optional ones after
                    if (objVar.bOptional)
                    { sTempLine += "[<" + objVar.sName + ">], "; }
                }
                if (ClsMiscString.Right(ref sTempLine, 2) == ", ")
                { sTempLine = ClsMiscString.Left(ref sTempLine, sTempLine.Length - 2); }
                sTempLine += ")";

                lstCode.Add(cSettings.Indent(iIndent) + sTempLine);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "call cLog.CloseLog()");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'At the end of the process");
                lstCode.Add(cSettings.Indent(iIndent) + "If cLog.isUsed Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "msgbox \"Something was logged please look at the log\" & cLog.FullPath");

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox \"Finished. Nothing was Logged\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'deinitialising Class");
                lstCode.Add(cSettings.Indent(iIndent) + "set cLog = Nothing");
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

        public void generateClass(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                int iIndent = 0;
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                string sTempLine;

                VBA.VBComponent vbComp = addModule(sClassName, VBA.vbext_ComponentType.vbext_ct_ClassModule);

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }
                addTitleComment(ref lstCode, ref cSettings, iIndent);

                /*
                 * Dim
                 */
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private fso As Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent) + "Private bIsStarted As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Private lsFullPath As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Private ts As Scripting.TextStream");
                lstCode.Add(cSettings.Indent(iIndent));

                sTempLine = "Private Const lcsDelimiter As String = vbTab";
                if (cDelimiter == '\u0013') //I hope that's decimal
                { sTempLine += "vbtab"; }
                else
                { sTempLine += cDelimiter.ToString(); }

                lstCode.Add(cSettings.Indent(iIndent) + "Private Const lcsDelimiter As String = vbTab");

                switch (eAutoPathGeneration)
                {
                    case enumAutoPath.eAutoPath_Date:
                        lstCode.Add(cSettings.Indent(iIndent) + "Private Const lcsDefaultDirectory As String = \"" + sPath + "\" ");
                        lstCode.Add(cSettings.Indent(iIndent) + "Private Const lcsDefaultExtension As String = \"" + sExtension + "\"");
                        break;
                    case enumAutoPath.eAutoPath_Manual:
                        break;
                    default:
                        break;
                }
                
                lstCode.Add(cSettings.Indent(iIndent));
                /*
                Put these two lines of comments in the GUI not in the code
                lstCode.Add(cSettings.Indent(iIndent) + "'Note: bare in mind what applications the extension will open");
                lstCode.Add(cSettings.Indent(iIndent) + "'You can have a text file with the xls extension and all the log files will open in Excel");
                */
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get FullPath() As String");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "FullPath = lsFullPath");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));

                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get isUsed() As Boolean");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "isUsed = bIsStarted");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Sub Class_Initialize()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set fso = New Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsStarted = False");

                switch (eAutoPathGeneration)
                {
                    case enumAutoPath.eAutoPath_Date:
                        lstCode.Add(cSettings.Indent(iIndent) + "lsFullPath = lcsDefaultDirectory & Format(Now, \"" + sDateFormat_FileName + "\") & lcsDefaultExtension");
                        break;
                    case enumAutoPath.eAutoPath_Manual:
                        lstCode.Add(cSettings.Indent(iIndent) + "lsFullPath = \"" + sPath + "\"");
                        break;
                    default:
                        break;
                }

                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Sub CloseLog()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsStarted Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Not ts Is Nothing Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "ts.Close");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                //lstCode.Add(cSettings.Indent(iIndent) + "bIsStarted = False");
                //lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set ts = Nothing");
                lstCode.Add(cSettings.Indent(iIndent) + "Set fso = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Sub Class_Terminate()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Call CloseLog");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));

                switch (eAutoPathGeneration)
                {
                    case enumAutoPath.eAutoPath_Date:
                        break;
                    case enumAutoPath.eAutoPath_Manual:
                        lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let FullPath(ByVal sFullPath As String)");
                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                        lstCode.Add(cSettings.Indent(iIndent));
                        lstCode.Add(cSettings.Indent(iIndent) + "If bIsStarted Then");
                        iIndent++;
                        lstCode.Add(cSettings.Indent(iIndent) + "MsgBox \"Can't change the Path of the the Log file after it has been opened.\", vbCritical");
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "Else");
                        iIndent++;
                        lstCode.Add(cSettings.Indent(iIndent) + "lsFullPath = sFullPath");
                        //don't open the log file here open it if an error is logged    
                        //lstCode.Add(cSettings.Indent(iIndent) + "Set ts = fso.OpenTextFile(lsFullPath, ForWriting, True)");
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End If");
                        lstCode.Add(cSettings.Indent(iIndent));
                        addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                        if (cSettings.IndentFirstLevel) { iIndent--; }
                        lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                        break;
                    default:
                        break;
                }

                lstCode.Add(cSettings.Indent(iIndent));

                sTempLine = "Public Sub Log(";
                foreach (strVariablesToLog objVar in lstVariablesToLog)
                {
                    //do the compulsary ones first
                    if (!objVar.bOptional)
                    { sTempLine += "ByVal " + objVar.sName + " As " + objVar.sDataType + ", "; }
                }
                foreach (strVariablesToLog objVar in lstVariablesToLog)
                {
                    //add the optional ones after
                    if (objVar.bOptional)
                    { sTempLine += "Optional ByVal " + objVar.sName + " As " + objVar.sDataType + ", "; }
                }
                if (ClsMiscString.Right(ref sTempLine, 2) == ", ")
                { sTempLine = ClsMiscString.Left(ref sTempLine, sTempLine.Length - 2); }
                sTempLine += ")";

                lstCode.Add(cSettings.Indent(iIndent) + sTempLine);
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sLine As String");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Static iLineNo As Long");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If Not bIsStarted Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "'Create the log file");
                lstCode.Add(cSettings.Indent(iIndent) + "Set ts = fso.OpenTextFile(lsFullPath, ForWriting, True)");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'add a title to the log file");
                lstCode.Add(cSettings.Indent(iIndent) + "sLine = \"Line No\"");
                foreach (strVariablesToLog objVar in lstVariablesToLog)
                { lstCode.Add(cSettings.Indent(iIndent) + "sLine = sLine & lcsDelimiter & \"" + objVar.sName + "\""); }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "ts.WriteLine sLine");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsStarted = True");
                lstCode.Add(cSettings.Indent(iIndent) + "iLineNo = iLineNo + 1");
                lstCode.Add(cSettings.Indent(iIndent));

                lstCode.Add(cSettings.Indent(iIndent) + "'output a line to log file");
                lstCode.Add(cSettings.Indent(iIndent) + "sLine = CStr(iLineNo)");
                foreach (strVariablesToLog sVarTemp in lstVariablesToLog)
                {
                    sTempLine = "sLine = sLine & lcsDelimiter & ";
                    switch (cDataTypes.getGeneralType(sVarTemp.sDataType))
                    {
                        case ClsDataTypes.enumGeneralDateType.eString:
                            sTempLine += sVarTemp.sName;
                            break;
                        case ClsDataTypes.enumGeneralDateType.eBool:
                        case ClsDataTypes.enumGeneralDateType.eNumber:
                            sTempLine += "CStr(" + sVarTemp.sName + ")";
                            break;
                        case ClsDataTypes.enumGeneralDateType.eDate: 
                            sTempLine += "Format(" + sVarTemp.sName + ", \"" + sDateFormat_FileContents + "\")";
                            break;
                        default:
                            break;
                    }
                    lstCode.Add(cSettings.Indent(iIndent) + sTempLine);
                }
                lstCode.Add(cSettings.Indent(iIndent) + "");
                lstCode.Add(cSettings.Indent(iIndent) + "ts.WriteLine sLine");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));

                this.addCode(ref lstCode, ref vbComp);

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
    }
}
