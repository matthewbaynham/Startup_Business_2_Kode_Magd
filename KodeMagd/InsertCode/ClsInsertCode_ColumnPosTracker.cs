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

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_ColumnPosTracker : ClsInsertCode
    {
        private string sClassName = "";
        private bool bUsingObjects = false;
        private string sFunctionName = "";
        private string sModuleName = "";

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

        public void setClassNameSuffix(string sName)
        {
            try
            {
                sClassName = ClsMiscString.makeValidVarName(sName, "cls");
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

        public bool UsingObjects
        {
            get
            {
                try
                {
                    return bUsingObjects;
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
                    bUsingObjects = value;
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


        public void CallClass(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();

                sModuleName = cCodeMapper.ModuleName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCodeTop.Add(cSettings.Indent(iIndent));
                lstCodeTop.Add(cSettings.Indent(iIndent) + "Public Const gclError As Long = -1");
                lstCodeTop.Add(cSettings.Indent(iIndent));
                lstCodeTop.Add(cSettings.Indent(iIndent) + "Private sht As Excel.Worksheet");
                lstCodeTop.Add(cSettings.Indent(iIndent));
                
                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Column_Position_Tracker");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim cHeader as " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "set cHeader = new " + sClassName);
                lstCode.Add(cSettings.Indent(iIndent));

                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'**********************************************************");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                        *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*  read through this bit of code and change values here  *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'*                                                        *");
                    lstCode.Add(cSettings.Indent(iIndent) + "'**********************************************************");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                lstCode.Add(cSettings.Indent(iIndent) + "cHeader.headerRow = 1");
                lstCode.Add(cSettings.Indent(iIndent));

                if (bUsingObjects)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "set cHeader.wrkSheet = ThisWorkbook.Worksheets(1)");
                }
                else
                {
                    Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                    lstCode.Add(cSettings.Indent(iIndent) + "cHeader.workbookName = \"" + wrk.Name + "\"");
                    lstCode.Add(cSettings.Indent(iIndent) + "cHeader.sheetName = \"" + wrk.Worksheets[1].Name + "\"");
                }
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'This is intended to be used like sht.Cells(lRow, cHeader.columnPos(\"<Column Name>\")).Value");
                }
                lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print cHeader.columnPos(\"<Column Name>\")");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cHeader = Nothing");
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

        public void generateClass(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = 0;

                VBA.VBComponent vbComp = addModule(sClassName, VBA.vbext_ComponentType.vbext_ct_ClassModule);

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCode.Add(cSettings.Indent(iIndent));
                if (bUsingObjects)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "private sht as Excel.Worksheet");
                }
                else
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Private sWorkbookName As String");
                    lstCode.Add(cSettings.Indent(iIndent) + "Private sSheetName As String");
                }
                lstCode.Add(cSettings.Indent(iIndent) + "Private lRowHeader As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Private bContainsDuplicates As Boolean");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Type typColumn");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sName As String");
                lstCode.Add(cSettings.Indent(iIndent) + "lPosition As Long");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End Type");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private arrColumns() As typColumn");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get headerRow() As Long");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "headerRow = lRowHeader");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let headerRow(lRow As Long)");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "lRowHeader = lRow");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));
                if (bUsingObjects)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get wrkSheet() As Excel.Worksheet");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Set wrkSheet = sht");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Call scanHeader");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                }
                else
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get workbookName() As String");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "workbookName = sWorkbookName");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Call scanHeader");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let workbookName(sName As String)");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent) + "");
                    lstCode.Add(cSettings.Indent(iIndent) + "sWorkbookName = sName");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Call scanHeader");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get sheetName() As String");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "sheetName = sSheetName");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Call scanHeader");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Let sheetName(sName As String)");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "sSheetName = sName");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Call scanHeader");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Private Sub scanHeader()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsReady As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lColumn As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sHeader As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lPosition As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim larrMax As Long");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsReady = True");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If lRowHeader = 0 Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsReady = False");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                if (bUsingObjects)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "If sht is Nothing Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsReady = False");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                else
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "If sWorkbookName = \"\" Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsReady = False");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "If sSheetName = \"\" Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsReady = False");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsReady Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "lColumn = 1");
                lstCode.Add(cSettings.Indent(iIndent) + "larrMax = 0");
                lstCode.Add(cSettings.Indent(iIndent) + "ReDim arrColumns(1 To 1)");
                lstCode.Add(cSettings.Indent(iIndent) + "");
                if (bUsingObjects)
                { lstCode.Add(cSettings.Indent(iIndent) + "With sht"); }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + "With Workbooks(sWorkbookName).Worksheets(sSheetName)"); }
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Do While Not .Cells(lRowHeader, lColumn).Value = \"\"");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sHeader = .Cells(lRowHeader, lColumn).Value");
                lstCode.Add(cSettings.Indent(iIndent) + "lPosition = lColumn");
                lstCode.Add(cSettings.Indent(iIndent) + "sHeader = Trim(CStr(.Cells(lRowHeader, lColumn).Value))");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If isInArray(sHeader) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bContainsDuplicates = True");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "larrMax = larrMax + 1");
                lstCode.Add(cSettings.Indent(iIndent) + "ReDim Preserve arrColumns(1 To larrMax) As typColumn");
                lstCode.Add(cSettings.Indent(iIndent) + "arrColumns(larrMax).sName = sHeader");
                lstCode.Add(cSettings.Indent(iIndent) + "arrColumns(larrMax).lPosition = lColumn");
                lstCode.Add(cSettings.Indent(iIndent) + "");
                lstCode.Add(cSettings.Indent(iIndent) + "lColumn = lColumn + 1");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End With");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get columnPos(ByVal sName As String) As Long");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lCounter As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsFound As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lPos As Long");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = False");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "lCounter = LBound(arrColumns)");
                lstCode.Add(cSettings.Indent(iIndent) + "Do While lCounter <= UBound(arrColumns) And Not bIsFound");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(arrColumns(lCounter).sName)) = Trim(UCase(sName)) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = True");
                lstCode.Add(cSettings.Indent(iIndent) + "lPos = lCounter");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent) + "lCounter = lCounter + 1");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsFound Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "columnPos = lPos");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "columnPos = gclError");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get containsDuplicates(ByVal sName As String) As Long");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "containsDuplicates = bContainsDuplicates");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get isInArray(ByVal sName As String) As Boolean");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent) + "Dim lCounter As Long");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsFound As Boolean");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = False");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "lCounter = LBound(arrColumns)");
                lstCode.Add(cSettings.Indent(iIndent) + "Do While lCounter <= UBound(arrColumns) And Not bIsFound");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(arrColumns(lCounter).sName)) = Trim(UCase(sName)) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsFound = True");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent) + "lCounter = lCounter + 1");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "isInArray = bIsFound");
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Property");
                lstCode.Add(cSettings.Indent(iIndent));

                this.addCode(ref lstCode, ref vbComp);

                if (lstCodeTop.Count > 0)
                {
                    lstCodeTop.Add("");
                    this.addCode(ref lstCodeTop, ref vbComp, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                cCodeMapper = null;
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
