using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using System.Diagnostics;
using KodeMagd.Rename;
using KodeMagd.Format;
using KodeMagd.InsertCode;
using Microsoft.Office.Interop.Excel;

namespace KodeMagd
{
    public class ClsCodeMapper
    {
        private bool bIsRead = false;
        private bool bIsOptionExplicit = false;
        private bool bIsOptionBase = false;
        private const string csChar_SingleQuote = "'";
        private const string csChar_DoubleQuote = "\"";
        private List<strVariables> lstVariablesMod = new List<strVariables>();
        private List<int> lstLineNoReferenced = new List<int>();

        public enum enumVarDimType
        {
            eVarDim_OneSpace,
            eVarDim_InLine,
            eVarDim_Nothing
        }
        
        public enum enumFunctionType
        {
            eFnType_Error,
            eFnType_Function,
            eFnType_Sub,
            eFnType_Property,
            eFnType_None
        }

        public enum enumFunctionPropertyType
        {
            ePropType_Let,
            ePropType_Get,
            ePropType_Set,
            ePropType_NA
        }

        public enum enumLineType 
        {
            eLineType_Options, //Option Explicit, Option Base ...
            eLineType_FunctionName,
            eLineType_DllFunctionDeclare,
            eLineType_With,
            eLineType_EndWith,
            eLineType_If,
            eLineType_ElseIF,
            eLineType_Else,
            eLineType_EndIf,
            eLineType_BeginLoop,
            eLineType_EndLoop,
            eLineType_EndFunction,
            eLineType_ContinuedFromAbove,
            eLineType_Empty,
            eLineType_Comment,
            eLineType_AssignValue,
            eLineType_Call,
            eLineType_Dim,
            eLineType_Initialise,
            eLineType_DeInitialise,
            eLineType_Goto,
            eLineType_OnError,
            eLineType_ErrorHandler,
            eLineType_ExitFn,
            eLineType_ExitLoop,
            eLineType_ExitIf,
            eLineType_Output,
            eLineType_Input,
            eLineType_Unknown
        }

        private struct strLineSimple 
        {
            public int iNo;
            public string sText;
            public string sLabel;
        }

        public struct strLine
        {
            public int iIndex;
            public int iOrder; //An incrementing number related to the order the line is in
            public string sText_Orig;
            public string sText_NoComment;
            public string sText_Comment;
            public string sModuleName;
            public string sFunctionName;
            public string sLabel; //Label is at the side of the text and can be used by Goto statements
            public string sLineNo; //written at the side or the text in the line
            public List<enumLineType> lstLineType;
            public enumFunctionType eFunctionType;
            public enumFunctionPropertyType ePropertyType;
            public int iIndentSize;
            public bool bValidateable;
            public int iOriginalLineNo; //refers to where in the module the line was if multiple lines (using :) are on the same line then this property will have duplicates
        }
        private List<strLine> lstLines = new List<strLine>();

        public enum enumScopeVar
        {
            eScope_Function,
            eScope_Module,
            eScope_Global
        }

        public enum enumScopeFn
        {
            eScopeFn_Public,
            eScopeFn_Private,
            eScopeFn_Friend
        }

        public enum enumParamType
        {
            eParTyp_Unknown,
            eParTyp_ByRef,
            eParTyp_ByVal,
            eParTyp_NA /*if it's a local variable*/
        }

        public struct strVariables 
        {
            public string sName;
            public enumScopeVar eScope;
            public ClsDataTypes.vbVarType eType;
            public bool bIsParameter;
            public enumParamType eParaType;
            public enumFunctionType eFunctionType;
            public enumFunctionPropertyType ePropType;
            public string sDatatype;
            public string sFunctionName;
            public string sModuleName;
            public bool bIsConstant;
        }

        public struct strFunctionIdentity 
        {
            public string sName;
            public enumFunctionType eFunctionType;
            public enumFunctionPropertyType ePropertyType;
            public string sModuleName;
        }

        public struct strFunctions
        {
            public string sName;
            public enumFunctionType eFunctionType;
            public enumFunctionPropertyType ePropertyType;
            public enumScopeFn eScope;
            public bool bIsStatic;
            public List<strVariables> lstVariablesFn;
            public int iLineNoStart;
            public int iLineNoEnd;
            public string sModuleName;
            public bool bHasErrorHandler;
        }
        private List<strFunctions> lstFunctions = new List<strFunctions>();

        public struct strModuleDetails
        {
            public string sName;
            public VBA.vbext_ComponentType eType;
        }
        private string sModuleName;
        private VBA.VBComponent vbComponent;

        private strModuleDetails objModuleDetails;

        public bool isRead
        {
            get
            {
                try
                {
                    return bIsRead;
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

        public bool hasOptionExplicit
        {
            get
            {
                try
                {
                    return bIsOptionExplicit;
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

        public string ModuleName 
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

                    return "";
                }
            }
        }

        public strModuleDetails ModuleDetails
        {
            get
            {
                try
                {
                    //strModuleDetails objTemp = new strModuleDetails();

                    //objTemp.sName = sModuleName;
                    //objTemp.eType = ;

                    return objModuleDetails;
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

                    return new strModuleDetails();
                }
            }
        }

        public bool hasOptionBase
        {
            get
            {
                try
                {
                    return bIsOptionBase;
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

        public List<strLine> lines
        {
            get 
            {
                try
                {
                    return lstLines;
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

                    return new List<strLine>();
                }
            }
        }

        public void readCode()
        {
            try
            {
                VBA.VBComponent vbcomp = ClsMisc.ActiveVBComponent();

                readCode(vbcomp);
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

        public ClsCodeMapper()
        {
            try
            {
                bIsRead = false;
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

        ~ClsCodeMapper()
        {
            try
            {
                lstFunctions = null;
                lstLineNoReferenced = null;
                lstLines = null;
                lstVariablesMod = null;
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

        public void readCode(VBA.VBComponent vbComp)
        {
            /*
             *  scan through a function or sub or property 
             *  recognise keywords
             *  recognise variables
             *  recognise comments
             *  recognise dim variables
             *  recognise strings and hardcoded values
             *  recognise objects that are not keywords (e.g. Recordset)
             *  take note of the keywords that change the flow (e.g. if then end if, do while until loop, exit, goto)
             *  build up a data structure to reperesent the code
             */
            try
            {
                bIsRead = true;
                Excel.Application app = ClsMisc.ActiveApplication();

                XlMousePointer objOrigCursor = app.Cursor;
                app.Cursor = XlMousePointer.xlWait;

                vbComponent = vbComp;
                VBA.CodeModule objCode = vbComp.CodeModule;
                Queue<strLineSimple> qLinesSource = stripMultiLines(objCode); //if a line is on multiple lines with underscores then put it on one line
                int iPreviousFunctionLine = objCode.CountOfLines;
                int iLine = 1;
                strFunctions objCurrentFunction = new strFunctions();
                string sCurrentFunctionName = "";
                enumFunctionPropertyType eCurrentPropertyType = enumFunctionPropertyType.ePropType_NA;
                enumFunctionType eCurrentFunctionType = enumFunctionType.eFnType_Error;
                lstVariablesMod = new List<strVariables>();

                sModuleName = vbComp.Name;
                objModuleDetails.sName = vbComp.Name;
                objModuleDetails.eType = vbComp.Type;

                foreach (strLineSimple strTempLine in qLinesSource) 
                {
                    string sLine = strTempLine.sText;
                    
                    //Check Comments
                    strLine objLine = new strLine();
                    strFunctions objFunction;

                    objLine.sModuleName = sModuleName;
                    objLine.iIndex = 0;
                    objLine.iOrder = iLine;
                    objLine.iOriginalLineNo = strTempLine.iNo;
                    objLine.sText_Orig = sLine;
                    objLine.sLineNo = ClsMisc.getLineNumbers(sLine);
                    string sLineTemp = ClsMisc.stripLineNumbers(sLine);
                    objLine.sLabel = strTempLine.sLabel;
                    //objLine.sLabel = ClsMisc.getLineLabels(sLineTemp);
                    sLineTemp = ClsMisc.stripLineLabels(ref objLine, ref sLineTemp);
                    objLine.sText_Comment = return_Comment(sLineTemp);
                    objLine.sText_NoComment = remove_Comment(sLineTemp);
                    objLine.lstLineType = new List<enumLineType>();

                    objCurrentFunction = functionName(ref objLine, objCurrentFunction, vbComp.Name); //get the list of functions populated inn here
                    sCurrentFunctionName = objCurrentFunction.sName;
                    eCurrentPropertyType = objCurrentFunction.ePropertyType;
                    eCurrentFunctionType = objCurrentFunction.eFunctionType;

                    if (sCurrentFunctionName == null)
                    { objLine.sFunctionName = ""; }
                    else
                    { objLine.sFunctionName = sCurrentFunctionName; }

                    objLine.eFunctionType = eCurrentFunctionType;
                    objLine.ePropertyType = eCurrentPropertyType;

                    //check Option Explicit
                    check_OptionExplicit(ref objLine);
                    check_OptionBase(ref objLine);

                    //Get all the variables - local to functions
                    objFunction = getFunction(ref objLine, vbComp.Name);
                    objFunction.sModuleName = vbComp.Name;
                    objLine.eFunctionType = objFunction.eFunctionType;
 
                    getLineType(ref objLine);
                    logGotoNumbers(ref objLine);

                    /* 
                     * bug here if it couldn't find a function then the the object is not set 
                     * and when we add variables to a object of a function that is not set it 
                     * all goes wrong.
                     */

                    find_Variables_SplitLine(ref objLine, ref objFunction);

                    //get all the variables - local to module
                    //get all the variables - global

                    lstLines.Add(objLine);
                    iLine++;
                }

                getLineTypeSecondScan(ref lstLines);

                findEndOfFunctions();
                functionsWithOnErrorGoto();

                fixIndex();

                foreach (strFunctions objFunction in lstFunctions)
                { markErrorHandlers(objFunction); }

                fixIndex();

                app.Cursor = objOrigCursor;
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

        public void fixIndex()
        {
            try
            {
                for(int iIndex = 0; iIndex < lstLines.Count; iIndex++)
                {
                    strLine objLine = lstLines[iIndex];

                    objLine.iIndex = iIndex;

                    lstLines[iIndex] = objLine;
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

        private Queue<strLineSimple> stripMultiLines(VBA.CodeModule objCode) 
        {
            try
            {
                Queue<strLineSimple> qResult = new Queue<strLineSimple>();
                string sLineFull = "";
                bool bPreviousLineContinues = false;
                string sLabel = "";

                for (int iLine = 0; iLine <= objCode.CountOfLines; iLine++)
                {
                    string sLine = objCode.get_Lines(iLine + 1, 1);

                    /*
                     strip labels off here.
                     
                     */

                    if (bPreviousLineContinues)
                    { sLineFull += " " + sLine; }
                    else
                    { sLineFull = sLine; }

                    if (sLineFull.Trim().EndsWith(" _"))
                    {
                        sLineFull = ClsMiscString.Left(ref sLineFull, sLineFull.Length - 2);
                        bPreviousLineContinues = true;
                    }
                    else
                    {
                        sLabel = ClsMisc.getLineLabels(sLine);

                        if (sLabel.Trim() != "")
                        {
                            //remove label here
                            if (sLine.Trim().ToUpper().StartsWith(sLabel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                int iPos = sLine.ToUpper().IndexOf(sLabel.ToUpper());

                                sLineFull = ClsMiscString.Right(sLine, sLine.Length - sLabel.Length - iPos);
                            }
                        }

                        bPreviousLineContinues = false;

                        /*Not Tested yet*/

                        List<string> lstTemp = ClsMisc.splitButNotInQuotes(ref sLineFull, ':');

                        if (lstTemp.Count == 0)
                        {
                            strLineSimple sLineSimple;
                            sLineSimple.sText = "";
                            sLineSimple.iNo = iLine + 1;
                            sLineSimple.sLabel = sLabel;

                            qResult.Enqueue(sLineSimple);

                            sLabel = "";
                        }
                        else
                        {
                            foreach (string sTemp in lstTemp)
                            {
                                strLineSimple sLineSimple;
                                sLineSimple.sText = sTemp;
                                sLineSimple.iNo = iLine + 1;
                                sLineSimple.sLabel = sLabel;

                                qResult.Enqueue(sLineSimple);

                                sLabel = "";
                            }
                        }

                        sLineFull = "";
                    }
                }

                return qResult;
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
                return null;
            }
        }

        private string return_Comment(string sLine)
        {
            try 
            {
                /*
                 ************************************************************
                 *                                                          *
                 *   BUG: comments can be from the REM statement as well.   *
                 *                                                          *
                 ************************************************************
                 */
                //use containsSingeQuoteOrREM

                string sTemp = sLine;
                bool bContainsQuotationMark;
                string sResult;
                bool bIsFinished;

                bIsFinished = false;
                sResult = "";

                bContainsQuotationMark = sTemp.Contains(csChar_SingleQuote);

                if (!bContainsQuotationMark)
                {
                    sResult = "";
                    bIsFinished = true;
                }

                if (!bIsFinished)
                {
                    if (sTemp.Contains(csChar_DoubleQuote))
                    {
                        bool bIsInString = false;
                        int iPos = 0;
                        bool bCommentFound = false;
                        bool bEndOfSearch = false;

                        while (!bEndOfSearch & !bCommentFound)
                        {
                            if (sTemp.ToString().Contains(csChar_DoubleQuote))
                            {
                                iPos = sTemp.ToString().IndexOf(csChar_DoubleQuote);
                                sTemp = sTemp.Substring(iPos + 1);
                                bIsInString = !bIsInString;
                            }
                            else
                            { bEndOfSearch = true; }

                            if (!bIsInString & sTemp.ToString().Contains(csChar_SingleQuote))
                            {
                                if (sTemp.ToString().Contains(csChar_DoubleQuote))
                                {
                                    if (sTemp.ToString().IndexOf(csChar_DoubleQuote) > sTemp.ToString().IndexOf(csChar_SingleQuote))
                                    { bCommentFound = true; }
                                }
                                else 
                                { bCommentFound = true; }

                                if (bCommentFound)
                                {
                                    iPos = sTemp.ToString().IndexOf(csChar_SingleQuote);
                                    sTemp = sTemp.Substring(iPos);
                                    sResult = sLine.Substring(sLine.Length - sTemp.Length);
                                    bCommentFound = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        int iPosCommentStart = sTemp.IndexOf(csChar_SingleQuote);

                        sResult = sLine.ToString().Substring(iPosCommentStart);
                        bIsFinished = true;
                    }
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
                return string.Empty;
            }
        }

        private string remove_Comment(string sLine)
        {
            try 
            {
                /*
                 ************************************************************
                 *                                                          *
                 *   BUG: comments can be from the REM statement as well.   *
                 *                                                          *
                 ************************************************************
                 */

                string sTemp = sLine;
                bool bContainsQuotationMark;
                string sResult;
                bool bIsFinished;

                bIsFinished = false;
                sResult = "";

                bContainsQuotationMark = sTemp.Contains(csChar_SingleQuote);

                if (!bContainsQuotationMark)
                {
                    sResult = sLine;
                    bIsFinished = true;
                }
                else
                {
                    int iPosStartComment = ClsMisc.charFirstPosNotInQuotes(ref sTemp, '\'');

                    if (iPosStartComment == -1)
                    {
                        sResult = sLine;
                        bIsFinished = true;
                    }
                    else
                    {
                        sResult = sLine.Substring(0, iPosStartComment);
                        bIsFinished = true;
                    }
                }

                /*
                if (!bIsFinished)
                {
                    if (sTemp.Contains(csChar_DoubleQuote))
                    {
                        bool bIsInString = false;
                        int iPos = 0;
                        bool bCommentFound = false;
                        bool bEndOfSearch = false;

                        while (!bEndOfSearch & !bCommentFound)
                        {
                            if (sTemp.ToString().Contains(csChar_DoubleQuote))
                            {
                                iPos = sTemp.ToString().IndexOf(csChar_DoubleQuote);
                                sTemp = sTemp.Substring(iPos + 1);
                                bIsInString = !bIsInString;
                            }
                            else
                            { bEndOfSearch = true; }

                            if (!bIsInString & sTemp.ToString().Contains(csChar_SingleQuote))
                            {
                                if (sTemp.ToString().Contains(csChar_DoubleQuote))
                                {
                                    if (sTemp.ToString().IndexOf(csChar_DoubleQuote) > sTemp.ToString().IndexOf(csChar_SingleQuote))
                                    { bCommentFound = true; }
                                }
                                else
                                { bCommentFound = true; }

                                if (bCommentFound)
                                {
                                    iPos = sTemp.ToString().IndexOf(csChar_SingleQuote);
                                    sTemp = sTemp.Substring(iPos);
                                    sResult = sLine.Substring(0, sLine.Length - sTemp.Length);
                                    bCommentFound = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        int iPosCommentStart = sTemp.IndexOf(csChar_SingleQuote);

                        sResult = sLine.ToString().Substring(0, iPosCommentStart);
                        bIsFinished = true;
                    }
                }
                */
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

        private strFunctions functionName(ref strLine objLine, strFunctions objOldFunctions, string sModuleName)
        {
            try
            {
                string sLine = objLine.sText_Orig;
                strFunctions objResult = new strFunctions();

                strFunctions objFnTemp = functionDetails(ref objLine, sModuleName);

                if (objFnTemp.eFunctionType == enumFunctionType.eFnType_Error || objFnTemp.eFunctionType == enumFunctionType.eFnType_None)
                { objResult = objOldFunctions; }
                else
                { objResult = objFnTemp; }

                return objResult;
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
                return new strFunctions();
            }
        }

        private strFunctions functionDetails(ref strLine objLine, string sModuleName)
        {
            try
            {
                string sLine = objLine.sText_NoComment;
                strFunctions objFnTemp = new strFunctions();
                string sResult = "";
                string sTemp;
                bool bIsFunction;

                sTemp = sLine.Trim();
                if (sTemp.Trim().Length > 10)
                {
                    objFnTemp.eScope = enumScopeFn.eScopeFn_Public;
                    if (sTemp.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Substring(6).Trim();
                        objFnTemp.eScope = enumScopeFn.eScopeFn_Public;
                    }
                    else if (sTemp.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Substring(7).Trim();
                        objFnTemp.eScope = enumScopeFn.eScopeFn_Private;
                    }
                    else if (sTemp.Trim().ToUpper().StartsWith("FRIEND ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Substring(6).Trim();
                        objFnTemp.eScope = enumScopeFn.eScopeFn_Friend;
                    }
                }

                bIsFunction = false;
                sTemp = sTemp.Trim();

                objFnTemp.bIsStatic = false;
                if (sTemp.Trim().Length > 10)
                {
                    if (sTemp.Trim().ToUpper().StartsWith("STATIC ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = ClsMiscString.Right(sTemp.Trim(), sTemp.Length - 8).Trim();
                        objFnTemp.bIsStatic = true;
                    }
                }

                sTemp = sTemp.Trim();

                if (sTemp.Trim().Length > 6)
                {
                    objFnTemp.eFunctionType = enumFunctionType.eFnType_None;
                    objFnTemp.ePropertyType = enumFunctionPropertyType.ePropType_NA;
                    if (sTemp.Trim().ToUpper().StartsWith("FUNCTION ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Trim().Substring(8).Trim();
                        objFnTemp.eFunctionType = enumFunctionType.eFnType_Function;
                        bIsFunction = true;
                    }
                    else if (sTemp.Trim().ToUpper().StartsWith("SUB ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Trim().Substring(3).Trim();
                        objFnTemp.eFunctionType = enumFunctionType.eFnType_Sub;
                        bIsFunction = true;
                    }
                    else if (sTemp.Trim().ToUpper().StartsWith("PROPERTY ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sTemp = sTemp.Trim().Substring(8).Trim();
                        objFnTemp.eFunctionType = enumFunctionType.eFnType_Property;
                        if (sTemp.Trim().ToUpper().StartsWith("LET ", StringComparison.CurrentCultureIgnoreCase))
                        { 
                            objFnTemp.ePropertyType = enumFunctionPropertyType.ePropType_Let;
                            sTemp = sTemp.Trim().Substring(3).Trim();
                        }
                        else if (sTemp.Trim().ToUpper().StartsWith("GET ", StringComparison.CurrentCultureIgnoreCase))
                        { 
                            objFnTemp.ePropertyType = enumFunctionPropertyType.ePropType_Get;
                            sTemp = sTemp.Trim().Substring(3).Trim();
                        }
                        else if (sTemp.Trim().ToUpper().StartsWith("SET ", StringComparison.CurrentCultureIgnoreCase))
                        { 
                            objFnTemp.ePropertyType = enumFunctionPropertyType.ePropType_Set;
                            sTemp = sTemp.Trim().Substring(3).Trim();
                        }

                        bIsFunction = true;
                    }
                }

                if (bIsFunction)
                {
                    objLine.lstLineType.Add(enumLineType.eLineType_FunctionName);
                    int iPosBracket = sTemp.Trim().LastIndexOf("(");

                    sResult = ClsMiscString.Left(sTemp.Trim(), iPosBracket);
                    objFnTemp.sName = sResult;
                    objFnTemp.iLineNoStart = objLine.iOriginalLineNo;
                    objFnTemp.lstVariablesFn = getParamters(sTemp, objFnTemp.sName);
                    objFnTemp.bHasErrorHandler = false;
                    objFnTemp.sModuleName = sModuleName;
                    lstFunctions.Add(objFnTemp);
                }

                return objFnTemp;
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

                strFunctions objFnTemp = new strFunctions();

                objFnTemp.bIsStatic = false;
                objFnTemp.eFunctionType = enumFunctionType.eFnType_Error;
                objFnTemp.eScope = enumScopeFn.eScopeFn_Public;
                objFnTemp.lstVariablesFn = new List<strVariables>();
                objFnTemp.sName = "";

                return objFnTemp;
            }
        }

        private List<strVariables> getParamters(string sLine, string sFunctionName) 
        { 
            try 
            {
                List<strVariables> lstResult = new List<strVariables>();

                string sParameters;
                int iPosStart;
                int iPosEnd;

                enumFunctionPropertyType ePropType;
                enumFunctionType eFunctionType;

                iPosStart = sLine.IndexOf('(') + 1;
                iPosEnd = sLine.LastIndexOf(')');

                string sFunctionDeclaration = ClsMiscString.Left(sLine, iPosStart);

                if (sFunctionDeclaration.Trim().ToUpper().Contains(" SUB ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("SUB "))
                {
                    ePropType = enumFunctionPropertyType.ePropType_NA;
                    eFunctionType = enumFunctionType.eFnType_Sub;
                }
                else if (sFunctionDeclaration.Trim().ToUpper().Contains(" FUNCTION ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("FUNCTION "))
                {
                    ePropType = enumFunctionPropertyType.ePropType_NA;
                    eFunctionType = enumFunctionType.eFnType_Function;
                }
                else if (sFunctionDeclaration.Trim().ToUpper().Contains(" PROPERTY ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("PROPERTY "))
                {
                    ePropType = enumFunctionPropertyType.ePropType_NA;
                    eFunctionType = enumFunctionType.eFnType_Property;

                    if (sFunctionDeclaration.Trim().ToUpper().Contains(" PROPERTY LET ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("PROPERTY LET "))
                    { ePropType = enumFunctionPropertyType.ePropType_Let; }
                    else if (sFunctionDeclaration.Trim().ToUpper().Contains(" PROPERTY GET ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("PROPERTY GET "))
                    { ePropType = enumFunctionPropertyType.ePropType_Get; }
                    else if (sFunctionDeclaration.Trim().ToUpper().Contains(" PROPERTY SET ") || sFunctionDeclaration.Trim().ToUpper().StartsWith("PROPERTY SET "))
                    { ePropType = enumFunctionPropertyType.ePropType_Set; }
                }
                else
                {
                    ePropType = enumFunctionPropertyType.ePropType_NA;
                    eFunctionType = enumFunctionType.eFnType_Error;
                }


                sParameters = sLine.Substring(iPosStart, iPosEnd - iPosStart);

                List<string> lstParameterStrings = splitParameters(sParameters);

                foreach (string sParameter in lstParameterStrings) 
                { 
                    if (sParameter.Trim() != "")
                    {
                        strVariables sVar = new strVariables();

                        sVar.ePropType = ePropType;
                        sVar.eScope = enumScopeVar.eScope_Function;
                        sVar.bIsParameter = true;

                        //1) get the name

                        int iPosNameEnd;
                        if (sParameter.ToUpper().Contains(" AS "))
                        { iPosNameEnd = sParameter.ToUpper().IndexOf(" AS "); }
                        else
                        { iPosNameEnd = sParameter.Length; }
                        sVar.sName = ClsMiscString.Left(sParameter, iPosNameEnd);

                        sVar.sName = sVar.sName.Trim();
                        if (ClsMiscString.Left(ref sVar.sName, 6).ToUpper() == "BYVAL " || ClsMiscString.Left(ref sVar.sName, 6).ToUpper() == "BYREF ") 
                        {
                            sVar.sName = ClsMiscString.Right(ref sVar.sName, sVar.sName.Length - 6);
                            sVar.sName = sVar.sName.Trim();
                        }

                        //2) Get type
                        int iPosAs = sParameter.ToUpper().IndexOf(" AS ");

                        sVar.sDatatype = ClsMiscString.Right(sParameter, sParameter.Length - iPosAs - 4).Trim();

                        sVar.eType = ClsMisc.getVBA_VarType(sVar.sDatatype);

                        //3) Get ByRef or ByVal (note: if not speciaffied then look at the datatype)
                        if (ClsMiscString.Left(sParameter.Trim().ToUpper(), 6) == "BYVAL ")
                        { sVar.eParaType = enumParamType.eParTyp_ByVal; }
                        else if (ClsMiscString.Left(sParameter.Trim().ToUpper(), 6) == "BYREF ")
                        { sVar.eParaType = enumParamType.eParTyp_ByRef; }
                        else
                        {
                            switch (sVar.eType)
                            { 
                                case ClsDataTypes.vbVarType.vbArray:
                                case ClsDataTypes.vbVarType.vbDataObject:
                                case ClsDataTypes.vbVarType.vbObject:
                                    sVar.eParaType = enumParamType.eParTyp_ByRef;
                                    break;
                                case ClsDataTypes.vbVarType.vbUnknown:
                                case ClsDataTypes.vbVarType.vbUserDefinedType:
                                    sVar.eParaType = enumParamType.eParTyp_Unknown;
                                    break;
                                default:
                                    sVar.eParaType = enumParamType.eParTyp_ByVal;
                                    break;
                            }
                        }

                        sVar.sModuleName = sModuleName;
                        sVar.sFunctionName = sFunctionName;
                        sVar.ePropType = ePropType;
                        sVar.eFunctionType = eFunctionType;

                        lstVariablesMod.Add(sVar);
                    }
                }

                return lstResult;
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

                return null;
            }
        }

        private List<string> splitParameters(string sParameters) 
        { 
            try {
                List<string> lstResult = new List<string>();
                string sCurrentParameter = "";
                int iBracketRunningTotal = 0;

                foreach (string sTemp in sParameters.Split(',')) 
                {
                    //int iBracketCountOpen = ClsMiscString.stringCountChar(sTemp, '(');
                    //int iBracketCountClose = ClsMiscString.stringCountChar(sTemp, ')');
                    int iBracketCountOpen = sTemp.Count(x => x == '(');
                    int iBracketCountClose = sTemp.Count(x => x == ')');

                    iBracketRunningTotal += iBracketCountOpen;
                    iBracketRunningTotal -= iBracketCountClose;

                    sCurrentParameter += sTemp;

                    if (iBracketRunningTotal == 0)
                    {
                        lstResult.Add(sCurrentParameter);
                        sCurrentParameter = "";
                    }
                }

                if (iBracketRunningTotal != 0)
                { lstResult.Add(sCurrentParameter); }

                return lstResult;
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

                return null;
            }
        }

        private void check_OptionExplicit(ref strLine objLine)
        {
            try
            {
                const char csDelimiter = ':';
                string sLine = objLine.sText_Orig;

                string[] lstLines = sLine.Split(csDelimiter);

                foreach (string sTemp in lstLines) {
                    if (sTemp.Trim().ToUpper() == "OPTION EXPLICIT")
                    {
                        bIsOptionExplicit = true;
                        objLine.lstLineType.Add(enumLineType.eLineType_Options);
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

        private void check_OptionBase(ref strLine objLine)
        {
            try
            {
                const char csDelimiter = ':';
                string sLine = objLine.sText_Orig;

                string[] lstLines = sLine.Split(csDelimiter);

                foreach (string sTemp in lstLines)
                {
                    if (sTemp.Trim().ToUpper() == "OPTION BASE 1" || sTemp.Trim().ToUpper() == "OPTION BASE 0")
                    {
                        bIsOptionBase = true;
                        objLine.lstLineType.Add(enumLineType.eLineType_Options);
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

        private void find_Variables(ref strFunctions objFunction, string sOneDimStatement)
        {
            try
            {
                strVariables sVar = new strVariables();

                sVar.eScope = enumScopeVar.eScope_Function;
                sVar.eParaType = enumParamType.eParTyp_NA;

                if (sOneDimStatement.Trim().ToUpper().StartsWith("DIM ", StringComparison.CurrentCultureIgnoreCase))
                { 
                    sOneDimStatement = sOneDimStatement.Substring(4);
                    if (string.IsNullOrEmpty(objFunction.sName))
                    { sVar.eScope = enumScopeVar.eScope_Module; }
                    else
                    { sVar.eScope = enumScopeVar.eScope_Function; }
                }

                if (sOneDimStatement.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase))
                {
                    sOneDimStatement = sOneDimStatement.Trim().Substring(7);
                    sVar.eScope = enumScopeVar.eScope_Global;
                }

                if (sOneDimStatement.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                {
                    sOneDimStatement = sOneDimStatement.Trim().Substring(8);
                    sVar.eScope = enumScopeVar.eScope_Module;
                }

                if (sOneDimStatement.Trim().ToUpper().StartsWith("CONST ", StringComparison.CurrentCultureIgnoreCase))
                {
                    sOneDimStatement = sOneDimStatement.Trim().Substring(6);
                    sVar.bIsConstant = true;
                }
                else
                { sVar.bIsConstant = false; }

                if (sOneDimStatement.ToUpper().Contains(" AS "))
                {
                    int iPos = sOneDimStatement.ToUpper().IndexOf(" AS ");

                    sVar.sName = ClsMiscString.Left(ref sOneDimStatement, iPos).Trim();

                    string sDataTypeText = ClsMiscString.Right(ref sOneDimStatement, sOneDimStatement.Length - iPos - 4).Trim(); //wrong

                    if (sDataTypeText.Trim().ToUpper().StartsWith("NEW ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        sDataTypeText = ClsMiscString.Right(ref sDataTypeText, sDataTypeText.Length - 4).Trim();
                        sVar.sDatatype = sDataTypeText;
                    }
                    else if (sDataTypeText.Contains('='))
                    {
                        //is constant
                        int iPosEquals = sDataTypeText.IndexOf('=');

                        sVar.sDatatype = ClsMiscString.Left(ref sDataTypeText, iPosEquals - 1).Trim();
                    }
                    else
                    { sVar.sDatatype = sDataTypeText; }

                    sVar.eType = ClsMisc.getVBA_VarType(sVar.sDatatype);
                }
                else
                {
                    sVar.sName = sOneDimStatement.Trim();
                    sVar.sDatatype = "Variant";
                    sVar.eType = ClsDataTypes.vbVarType.vbVariant;
                }

                if (objFunction.sName == null)
                { 
                    sVar.sFunctionName = string.Empty;
                    sVar.eFunctionType = enumFunctionType.eFnType_None;
                    sVar.ePropType = enumFunctionPropertyType.ePropType_NA;
                }
                else
                {
                    sVar.sFunctionName = objFunction.sName;
                    sVar.eFunctionType = objFunction.eFunctionType;
                    sVar.ePropType = objFunction.ePropertyType;
                }

                if (objModuleDetails.sName == null)
                { sVar.sModuleName = string.Empty; }
                else
                { sVar.sModuleName = objModuleDetails.sName; }

                if (string.IsNullOrEmpty(objFunction.sName))
                { this.lstVariablesMod.Add(sVar); }
                else
                { objFunction.lstVariablesFn.Add(sVar); }
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

        private List<string> splitDim(string sFullLine)
        {
            try {
                List<string> lstResult = new List<string>();

                foreach (string sLine in sFullLine.Split(':'))
                {
                    //check it's a dim line first
                    string sCompare = sLine.Trim().ToLower();

                    if (sCompare.Trim().ToUpper().StartsWith("DIM ", StringComparison.CurrentCultureIgnoreCase)
                        || ((!sCompare.Contains(" sub ") && !sCompare.Contains(" function ") && !sCompare.Contains(" property "))
                            && (sCompare.Trim().ToUpper().StartsWith("CONST ", StringComparison.CurrentCultureIgnoreCase)
                                || sCompare.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase)
                                || sCompare.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))))
                    {
                        string sComplicatedElement = "";
                        int iBracketDepth = 0;

                        foreach (string sTemp in sLine.Split(','))
                        {
                            //int iCountOpenBrackets = ClsMiscString.stringCountChar(sTemp, '(');
                            //int iCountCloseBrackets = ClsMiscString.stringCountChar(sTemp, ')');
                            int iCountOpenBrackets = sTemp.Count(x => x == '(');
                            int iCountCloseBrackets = sTemp.Count(x => x == ')');

                            iBracketDepth += iCountOpenBrackets - iCountCloseBrackets;

                            if (iBracketDepth == 0)
                            {
                                if (string.IsNullOrEmpty(sComplicatedElement))
                                { lstResult.Add(sTemp.Trim()); }
                                else
                                { lstResult.Add(sComplicatedElement.Trim() + ", " + sTemp.Trim()); }
                                sComplicatedElement = "";
                            }
                            else
                            {
                                if (string.IsNullOrEmpty(sComplicatedElement))
                                { sComplicatedElement = sTemp.Trim(); }
                                else
                                { sComplicatedElement += ", " + sTemp.Trim(); }
                            }
                        }
                    }
                }
                return lstResult;
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
                return null;
            }
        }

        private void find_Variables_SplitLine(ref strLine objLine, ref strFunctions objFunction)
        {
            try 
            {
                List<string> lstLines = splitDim(objLine.sText_NoComment);

                foreach (string sTemp in lstLines)
                { find_Variables(ref objFunction, sTemp); }
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

        private strFunctions getFunction(ref strLine objLine, string sModuleName)
        {
            try
            {
                string sFunctionName = objLine.sFunctionName;
                enumFunctionPropertyType ePropType = objLine.ePropertyType;

                Predicate<strFunctions> predFunction;
                switch (objLine.eFunctionType)
                {
                    case enumFunctionType.eFnType_Property:
                        predFunction = Fn => Fn.sName == sFunctionName && Fn.ePropertyType == ePropType;
                        break;
                    case enumFunctionType.eFnType_Function:
                        predFunction = Fn => Fn.sName == sFunctionName;
                        break;
                    case enumFunctionType.eFnType_Sub:
                        predFunction = Fn => Fn.sName == sFunctionName;
                        break;
                    default:
                        predFunction = Fn => Fn.sName == sFunctionName;
                        break;
                 }

                strFunctions objFunction = lstFunctions.Find(predFunction);

                if (objFunction.Equals(null))
                { strFunctions objNewFn = functionDetails(ref objLine, sModuleName); }

                return objFunction;
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

                strFunctions objFunction;
                objFunction.eFunctionType = enumFunctionType.eFnType_Error;
                objFunction.ePropertyType = enumFunctionPropertyType.ePropType_NA;
                objFunction.bIsStatic = false;
                objFunction.eScope = enumScopeFn.eScopeFn_Public;
                objFunction.sName = "";
                objFunction.lstVariablesFn = null;
                objFunction.iLineNoStart = 0;
                objFunction.iLineNoEnd = 0;
                objFunction.sModuleName = "";
                objFunction.bHasErrorHandler = false;

                return objFunction;
            }
        }

        public void Indenting()
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();
                int iLevel = 0;
                int iCounter = 0;

                for (iCounter = 0; iCounter < lstLines.Count; iCounter++ )
                {
                    strLine objTemp = lstLines[iCounter];
                    int iIndent;

                    //If beginning of module or beginning of function then iLevel = 0
                    //if (objTemp.lstLineType.Contains(enumLineType.eLineType_FunctionName) || objTemp.lstLineType.Contains(enumLineType.eLineType_EndFunction))
                    //{ iLevel = 0; }

                    //if exiting a "if" or a loop then iLevel-- 
                    iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndWith);
                    iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndIf);
                    iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndLoop);

                    //if (cSettings.IndentFirstLevel)
                    //{ iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndFunction); }

                    if (cSettings.IndentFirstLevel)
                    { iIndent = (iLevel + 1) * cSettings.IndentSize; }
                    else
                    { iIndent = iLevel * cSettings.IndentSize; }

                    //else
                    iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_Else);
                    iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_ElseIF);
                    //iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_FunctionName);
                    //iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_EndFunction);


                    //if (cSettings.IndentFirstLevel) 
                    //{ iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_FunctionName); }

                    if (iIndent < 0)
                    { iIndent = 0; }

                    //if entering a "if" or a loop then iLevel++
                    iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_With);
                    iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_If);
                    iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_BeginLoop);

                    //if (countIF(objTemp.lstLineType, enumLineType.eLineType_FunctionName) > 0 | countIF(objTemp.lstLineType, enumLineType.eLineType_EndFunction) > 0)
                    if (objTemp.lstLineType.Contains(enumLineType.eLineType_FunctionName)
                        || objTemp.lstLineType.Contains(enumLineType.eLineType_EndFunction)
                        || objTemp.lstLineType.Contains(enumLineType.eLineType_Options))
                    { 
                        iIndent = 0;
                        iLevel = 0;
                    }

                    //cLog.LOG(objTemp.sText_Orig, "Indent: " + iIndent.ToString(), objTemp.lstLineType.Count.ToString());

                    objTemp.sText_Orig = objTemp.sText_Orig.Trim().PadLeft(iIndent + objTemp.sText_Orig.Trim().Length, ' ');
                    objTemp.sText_NoComment = objTemp.sText_NoComment.Trim().PadLeft(iIndent + objTemp.sText_NoComment.Trim().Length, ' ');

                    lstLines[iCounter] = objTemp;
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
/*
        public void Indenting()
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();
                int iLevel = 0;
                int iCounter = 0;

                for (iCounter = 0; iCounter < lstLines.Count; iCounter++)
                {
                    strLine objTemp = lstLines[iCounter];
                    int iIndent;

                    //If beginning of module or beginning of function then iLevel = 0
                    if (objTemp.lstLineType.Contains(enumLineType.eLineType_FunctionName) || objTemp.lstLineType.Contains(enumLineType.eLineType_EndFunction))
                    { iLevel = 0; }


                    //if exiting a "if" or a loop then iLevel-- 
                    iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndIf);
                    iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndLoop);

                    if (cSettings.IndentFirstLevel)
                    { iLevel -= countIF(objTemp.lstLineType, enumLineType.eLineType_EndFunction); }

                    if (cSettings.IndentFirstLevel)
                    { iIndent = (iLevel + 1) * cSettings.IndentSize; }
                    else
                    { iIndent = iLevel * cSettings.IndentSize; }

                    //else
                    iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_Else);
                    iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_FunctionName);
                    iIndent -= cSettings.IndentSize * countIF(objTemp.lstLineType, enumLineType.eLineType_EndFunction);


                    if (cSettings.IndentFirstLevel)
                    { iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_FunctionName); }

                    if (iIndent < 0)
                    { iIndent = 0; }

                    //if entering a "if" or a loop then iLevel++
                    iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_If);
                    iLevel += countIF(objTemp.lstLineType, enumLineType.eLineType_BeginLoop);

                    //cLog.LOG(objTemp.sText_Orig, "Indent: " + iIndent.ToString(), objTemp.lstLineType.Count.ToString());

                    objTemp.sText_Orig = objTemp.sText_Orig.Trim().PadLeft(iIndent + objTemp.sText_Orig.Trim().Length, ' ');
                    objTemp.sText_NoComment = objTemp.sText_NoComment.Trim().PadLeft(iIndent + objTemp.sText_NoComment.Trim().Length, ' ');

                    lstLines[iCounter] = objTemp;
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
*/
        public static int countIF(List<enumLineType> lstItems, enumLineType eItem) 
        {
            try
            {
                int iResult = 0;

                foreach (enumLineType eTemp in lstItems) 
                {
                    if (eTemp == eItem)
                    { iResult++; }
                }

                return iResult;
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
                
                return 0;
            }
        }

        public void addLine(int iLineIndex, ref strLine objLine)
        {
            try
            {
                lstLines.Insert(iLineIndex, objLine);
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

        public void ImplementChanges() 
        { 
            try 
            {
                ImplementChanges(vbComponent); 
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


        public void ImplementChanges(VBA.VBComponent vbComp) 
        { 
            try
            {
                VBA.CodeModule objCode = vbComp.CodeModule;

                if (objCode.CountOfLines > 0)
                { objCode.DeleteLines(1, objCode.CountOfLines); }

                int iCounter = 1;

                int iMaxCharLineNumber = maxCharLineNumber();

                foreach (strLine objLine in lstLines)
                {
                    string sLine = "";

                    if (objLine.sLineNo.Trim() != "")
                    {
                        if (objLine.sText_NoComment.StartsWith(" "))
                        { sLine += objLine.sLineNo.Trim(); }
                        else
                        { sLine += objLine.sLineNo.Trim().PadRight(iMaxCharLineNumber + 1, ' '); }
                    }

                    if (objLine.sLabel.Trim() != "")
                    { sLine += objLine.sLabel.Trim() + ": "; }

                    //if (objLine.sText_NoComment.Trim() != "")
                    //{ sLine += objLine.sText_NoComment + " "; }
                    sLine += objLine.sText_NoComment;

                    if (objLine.sText_Comment.Trim() != "")
                    { sLine += objLine.sText_Comment; }

                    objCode.InsertLines(iCounter, sLine);
                    iCounter++;
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

        private int maxCharLineNumber() 
        {
            try
            {
                int iMaxNo = 0;

                foreach (strLine objLine in lstLines.FindAll(x => x.iOriginalLineNo > 0))
                {
                    if (iMaxNo < objLine.iOriginalLineNo)
                    { iMaxNo = objLine.iOriginalLineNo; }
                }

                string sMaxNo = iMaxNo.ToString();

                return sMaxNo.Length;
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

                return 0;
            }
        }

        public static void getLineType(ref strLine objLine) 
        {
            try 
            {        
                objLine.lstLineType.Clear();

                if (objLine.sText_NoComment.Contains(':'))
                {
                    foreach (string sLine in objLine.sText_NoComment.Split(':'))
                    {
                        enumLineType eLineType = getLineType(sLine);

                        objLine.lstLineType.Add(eLineType);
                    }
                }
                else
                {
                    enumLineType eLineType = getLineType(objLine.sText_NoComment);
                    
                    objLine.lstLineType.Add(eLineType);
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

        private void logGotoNumbers(ref strLine objLine)
        {
            try
            {
                if (countIF(objLine.lstLineType, enumLineType.eLineType_Goto) > 0)
                {
                    string sNumber = ClsMiscString.Right(objLine.sText_NoComment.Trim(), objLine.sText_NoComment.Trim().Length - "GOTO ".Length);
                    int iNumber = 0;

                    if (int.TryParse(sNumber, out iNumber))
                    {
                        lstLineNoReferenced.Add(iNumber);
                        lstLineNoReferenced = lstLineNoReferenced.Distinct().ToList();
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

        public static enumLineType getLineType(string sLine)
        {
            try
            {
                bool bIsUnknown = true;
                enumLineType eResult = enumLineType.eLineType_Unknown;

                if (bIsUnknown)
                {
                    if (string.IsNullOrEmpty(sLine.Trim()))
                    {
                        eResult = enumLineType.eLineType_Empty;
                        bIsUnknown = false;
                    }
                }

                //Option
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("OPTION ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        eResult = enumLineType.eLineType_Options;
                        bIsUnknown = false;
                    }
                }

                //Dim
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("DIM ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        eResult = enumLineType.eLineType_Dim;
                        bIsUnknown = false;
                    }
                }

                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("REDIM ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        eResult = enumLineType.eLineType_Dim;
                        bIsUnknown = false;
                    }
                }

                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("CONST ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        eResult = enumLineType.eLineType_Dim;
                        bIsUnknown = false;
                    }
                }

                //With
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("WITH "))
                    {
                        eResult = enumLineType.eLineType_With;
                        bIsUnknown = false;
                    }
                }

                //if then
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("IF ") & sLine.Trim().ToUpper().EndsWith(" THEN"))
                    {
                        eResult = enumLineType.eLineType_If;
                        bIsUnknown = false;
                    }
                }

                //elseif
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("ELSEIF ") & sLine.Trim().ToUpper().EndsWith(" THEN"))
                    {
                        eResult = enumLineType.eLineType_ElseIF;
                        bIsUnknown = false;
                    }
                }

                //else
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "ELSE")
                    {
                        eResult = enumLineType.eLineType_Else;
                        bIsUnknown = false;
                    }
                }

                //end if
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "END IF")
                    {
                        eResult = enumLineType.eLineType_EndIf;
                        bIsUnknown = false;
                    }
                }

                //end with
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "END WITH")
                    {
                        eResult = enumLineType.eLineType_EndWith;
                        bIsUnknown = false;
                    }
                }

                //#if then
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("#IF ") & sLine.Trim().ToUpper().EndsWith(" THEN"))
                    {
                        eResult = enumLineType.eLineType_If;
                        bIsUnknown = false;
                    }
                }

                //#elseif
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("#ELSEIF ") & sLine.Trim().ToUpper().EndsWith(" THEN"))
                    {
                        eResult = enumLineType.eLineType_Else;
                        bIsUnknown = false;
                    }
                }

                //#else
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "#ELSE")
                    {
                        eResult = enumLineType.eLineType_Else;
                        bIsUnknown = false;
                    }
                }

                //#end if
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "#END IF")
                    {
                        eResult = enumLineType.eLineType_EndIf;
                        bIsUnknown = false;
                    }
                }

                /*
                 ********************* 
                 *                   * 
                 *   begin of loop   *
                 *                   * 
                 ********************* 
                 */
                //Do
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "DO")
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }
                //Do While
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("DO WHILE "))
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }
                //While
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("WHILE "))
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }
                //Do Until
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("DO UNTIL "))
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }
                //For each
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("FOR EACH ") & sLine.Trim().ToUpper().Contains(" IN "))
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }
                //For 
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("FOR ") & sLine.Trim().ToUpper().Contains(" = ") & sLine.Trim().ToUpper().Contains(" TO "))
                    {
                        eResult = enumLineType.eLineType_BeginLoop;
                        bIsUnknown = false;
                    }
                }

                /*
                 ********************* 
                 *                   * 
                 *   end of loop     *
                 *                   * 
                 ********************* 
                 */
                //Loop
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "LOOP")
                    {
                        eResult = enumLineType.eLineType_EndLoop;
                        bIsUnknown = false;
                    }
                }

                //Loop While
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("LOOP WHILE "))
                    {
                        eResult = enumLineType.eLineType_EndLoop;
                        bIsUnknown = false;
                    }
                }

                //Wend - end while loop
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("Wend"))
                    {
                        eResult = enumLineType.eLineType_EndLoop;
                        bIsUnknown = false;
                    }
                }

                //Loop Until
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("LOOP UNTIL "))
                    {
                        eResult = enumLineType.eLineType_EndLoop;
                        bIsUnknown = false;
                    }
                }

                //Next
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("NEXT ") || sLine.Trim().ToUpper() == "NEXT")
                    {
                        eResult = enumLineType.eLineType_EndLoop;
                        bIsUnknown = false;
                    }
                }

                //error
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("ON ERROR "))
                    {
                        eResult = enumLineType.eLineType_OnError;
                        bIsUnknown = false;
                    }
                }

                //go to
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("GOTO "))
                    {
                        eResult = enumLineType.eLineType_Goto;
                        bIsUnknown = false;
                    }
                }

                //end function
                if (bIsUnknown)
                {
                    if ((sLine.Contains('(') & sLine.Contains(')'))
                        &&
                        (sLine.Trim().ToUpper().StartsWith("FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("FRIEND FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC STATIC FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE STATIC FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("FRIEND STATIC FUNCTION ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("FRIEND SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC STATIC SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE STATIC SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("FRIEND STATIC SUB ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PROPERTY ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC PROPERTY ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE PROPERTY ", StringComparison.CurrentCultureIgnoreCase)
                        || sLine.Trim().ToUpper().StartsWith("FRIEND PROPERTY ", StringComparison.CurrentCultureIgnoreCase))
                        && !((sLine.Trim().ToUpper().StartsWith("DECLARE ", StringComparison.CurrentCultureIgnoreCase) 
                                || sLine.ToUpper().Contains(" DECLARE "))
                                && sLine.ToUpper().Contains(" LIB "))
                        )
                    {
                        eResult = enumLineType.eLineType_FunctionName;
                        bIsUnknown = false;
                    }
                }

                //exit function
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "EXIT FUNCTION")
                    {
                        eResult = enumLineType.eLineType_ExitFn;
                        bIsUnknown = false;
                    }
                }

                //exit sub
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "EXIT SUB")
                    {
                        eResult = enumLineType.eLineType_ExitFn;
                        bIsUnknown = false;
                    }
                }

                //exit property
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "EXIT PROPERTY")
                    {
                        eResult = enumLineType.eLineType_ExitFn;
                        bIsUnknown = false;
                    }
                }

                //end function
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "END FUNCTION")
                    {
                        eResult = enumLineType.eLineType_EndFunction;
                        bIsUnknown = false;
                    }
                }

                //end sub
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "END SUB")
                    {
                        eResult = enumLineType.eLineType_EndFunction;
                        bIsUnknown = false;
                    }
                }

                //end property
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper() == "END PROPERTY")
                    {
                        eResult = enumLineType.eLineType_EndFunction;
                        bIsUnknown = false;
                    }
                }

                //MsgBox
                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("MSGBOX"))
                    {
                        eResult = enumLineType.eLineType_Output;
                        bIsUnknown = false;
                    }
                }

                if (bIsUnknown)
                {
                    if ((sLine.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                        && !(sLine.ToUpper().Contains(" SUB ") || sLine.ToUpper().Contains(" FUNCTION "))
                        && (sLine.ToUpper().Contains(" DECLARE ") && sLine.ToUpper().Contains(" LIB ")))
                    {
                        eResult = enumLineType.eLineType_DllFunctionDeclare;
                        bIsUnknown = false;
                    }
                    else if (!(sLine.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                        && !(sLine.ToUpper().Contains(" SUB ") || sLine.ToUpper().Contains(" FUNCTION "))
                        && (sLine.ToUpper().StartsWith("DECLARE ", StringComparison.CurrentCultureIgnoreCase) && sLine.ToUpper().Contains(" LIB ")))
                    {
                        eResult = enumLineType.eLineType_DllFunctionDeclare;
                        bIsUnknown = false;
                    }
                }

                if (bIsUnknown)
                {
                    if ((sLine.Trim().ToUpper().StartsWith("FRIEND ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                        && !(sLine.ToUpper().Contains(" SUB ") || sLine.ToUpper().Contains(" FUNCTION ") || sLine.ToUpper().Contains(" PROPERTY "))
                        && !(sLine.ToUpper().Contains(" DECLARE ") || sLine.ToUpper().Contains(" LIB ")))
                    {
                        eResult = enumLineType.eLineType_Dim;
                        bIsUnknown = false;
                    }
                   /*
                    else if (!(sLine.Trim().ToUpper().StartsWith("FRIEND ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase) || sLine.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                        && !(sLine.ToUpper().StartsWith("SUB ") || sLine.ToUpper().StartsWith("FUNCTION ") || sLine.ToUpper().StartsWith("PROPERTY "))
                        && !(sLine.ToUpper().Contains(" DECLARE ") || sLine.ToUpper().Contains(" LIB ")))
                    {
                        eResult = enumLineType.eLineType_Dim;
                        bIsUnknown = false;
                    }
                    */
                }

                //end property
                if (bIsUnknown)
                {
                    if (sLine.Contains('='))
                    {
                        int iPos = sLine.IndexOf('=');

                        if (iPos > 1)
                        {
                            char cBefore = sLine[iPos - 1];

                            if (cBefore != ':')
                            {

                            //string sPrefix = ClsMiscString.Left(ref sLine, iPos - 1).Trim();

                            //if (ClsMiscString.isValidVariableName(sPrefix))
                            //{
                                eResult = enumLineType.eLineType_AssignValue;
                                bIsUnknown = false;
                            }
                        }

                        if (sLine.Length - 1 > iPos)
                        {
                            string sSuffix = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 1).Trim();

                            if (sSuffix.Trim().ToUpper().StartsWith("INPUTBOX"))
                            {
                                eResult = enumLineType.eLineType_Input;
                                bIsUnknown = false;
                            }

                        }
                    }
                }

                if (bIsUnknown)
                {
                    if (sLine.Trim().ToUpper().StartsWith("CALL "))
                    {
                        eResult = enumLineType.eLineType_Call;
                        bIsUnknown = false;
                    }
                }

                return eResult;
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

                return enumLineType.eLineType_Unknown;
            }
        }

        public List<string> variableNames()
        {
            try 
            {
                int iLineNo;
                int iStartLine;
                int iStartColumn;
                int iEndLine;
                int iEndColumn;

                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();

                vbComp.CodeModule.CodePane.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                iLineNo = iStartLine;

                List<string> lstResult = variableNames(iLineNo);

                return lstResult;
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

                return null;
            }
        }

        public List<string> variableNames(int iLineNo) 
        {
            try 
            {
                List<string> lstResult = new List<string>();

                /*
                 * 1) find the function using the line number
                 * 2) loop through the variables in that function and the global variables for the module.
                 * 3) if the module is a form then add the controls that have a "value" property.
                 */
                strFunctions objFunction = new strFunctions();
                string sFunctionName = "";
                bool bIsFound_LineNo = false;
                bool bIsFound_Function = false;

                foreach (strLine objLine in lstLines) 
                {
                    if (objLine.iOriginalLineNo == iLineNo)
                    {
                        sFunctionName = objLine.sFunctionName;
                        bIsFound_LineNo = true;
                    }
                }

                if (bIsFound_LineNo) 
                {
                    foreach (strFunctions objFunctionTemp in lstFunctions) 
                    {
                        if (objFunctionTemp.sName.Trim().ToUpper() == sFunctionName.Trim().ToUpper()) 
                        {
                            objFunction = objFunctionTemp;
                            bIsFound_Function = true;
                        }
                    }
                }

                if (bIsFound_Function) 
                { 
                    foreach (strVariables objVar in objFunction.lstVariablesFn)
                    { lstResult.Add(objVar.sName); }
                }

                lstResult = lstResult.Distinct().ToList(); //not sure might work

                lstResult.Sort();

                //lstResult.Add(csChar_DefaultParameterOther);

                return lstResult;
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

                return null;
            }
        }

        public List<strVariables> lstVariables() 
        {
            try
            {
                List<strVariables> lstResult = new List<strVariables>();
                foreach (strVariables objGlobalVarInMod in lstVariablesMod)
                { lstResult.Add(objGlobalVarInMod); }

                foreach (strFunctions objFunction in lstFunctions)
                {
                    foreach (strVariables objVariable in objFunction.lstVariablesFn) 
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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

                return null;
            }
        }

        public List<strVariables> lstVariablesInCurrentScope()
        {
            try
            {
                List<strVariables> lstResult = new List<strVariables>();
                strFunctions objCurrentFunction = this.currentFunction;

                foreach (strVariables objGlobalVarInMod in lstVariablesMod)
                { lstResult.Add(objGlobalVarInMod); }

                Predicate<strFunctions> predFunction;

                if (objCurrentFunction.eFunctionType == enumFunctionType.eFnType_Property)
                { predFunction = x => x.sName.Trim().ToUpper() == objCurrentFunction.sName.Trim().ToUpper() && x.eFunctionType == objCurrentFunction.eFunctionType && x.ePropertyType == objCurrentFunction.ePropertyType; }
                else
                { predFunction = x => x.sName.Trim().ToUpper() == objCurrentFunction.sName.Trim().ToUpper() && x.eFunctionType == objCurrentFunction.eFunctionType; }

                foreach (strFunctions objFunction in lstFunctions.FindAll(predFunction))
                {
                    foreach (strVariables objVariable in objFunction.lstVariablesFn)
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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

                return null;
            }
        }


        public List<strVariables> lstVariables(string sName)
        {
            try
            {
                List<strVariables> lstResult = new List<strVariables>();
                foreach (strVariables objGlobalVarInMod in lstVariablesMod.FindAll(x => x.sName.Trim().ToUpper() == sName.Trim().ToUpper()))
                { lstResult.Add(objGlobalVarInMod); }

                foreach (strFunctions objFunction in lstFunctions)
                {
                    foreach (strVariables objVariable in objFunction.lstVariablesFn.FindAll(x => x.sName.Trim().ToUpper() == sName.Trim().ToUpper()))
                    { lstResult.Add(objVariable); }
                }

                lstResult = lstResult.OrderBy(x => x.sName).ToList();

                return lstResult;
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

                return null;
            }
        }
        //public List<VBA.Forms.Control> getLstControls(VBA.VBComponent vbComp)
        //{ 
        //    try 
        //    {
        //        List<VBA.Forms.Control> lstResult = new List<VBA.Forms.Control>();
                
        //        if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
        //        {
        //            foreach (VBA.Forms.Control ctrl in vbComp.Designer.Controls) 
        //            {
        //                //if (ctrl.GetType() ) { }
        //            }
        //        }

        //        Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

        //        //VBA.Forms.
        //        //wrk.VBProject.VBComponents.

        //        return lstResult;
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

        //        return null;
        //    }
        //}

        public List<string> getLstFunctionNames()
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (strFunctions objFunction in lstFunctions)
                { lstResult.Add(objFunction.sName); }

                return lstResult;
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

        public List<strFunctions> getLstFunctions()
        {
            try
            {
                return lstFunctions;
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

                return new List<strFunctions>();
            }
        }

        public List<string> getLstFunctionSampleCalls()
        {
            try
            {
                List<string> lstResult = new List<string>();
                ClsDataTypes cDataTypes = new ClsDataTypes();

                foreach (strFunctions objFunction in lstFunctions)
                {
                    string sSampleCall = "Call " + sModuleName + objFunction.sName + "(";
                    
                    foreach (strVariables objVar in objFunction.lstVariablesFn)
                    {
                        if (objVar.bIsParameter) 
                        {
                            sSampleCall += objVar.sName + ":=";
                            switch (cDataTypes.getGeneralType(objVar.eType))
                            {
                                case ClsDataTypes.enumGeneralDateType.eBool:
                                    sSampleCall += objVar.sName + "true";
                                    break;
                                case ClsDataTypes.enumGeneralDateType.eDate:
                                    sSampleCall += objVar.sName + "Now()";
                                    break;
                                case ClsDataTypes.enumGeneralDateType.eNumber:
                                    sSampleCall += objVar.sName + "0";
                                    break;
                                case ClsDataTypes.enumGeneralDateType.eString:
                                    sSampleCall += objVar.sName + "\"\"";
                                    break;
                                default:
                                    break;

                            }

                        }
                    }

                    lstResult.Add(objFunction.sName); }

                return lstResult;
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

                return null;
            }
        }

        public bool cursorIsInFunction
        {
            get
            {
                try
                {
                    bool bResult = false;

                    VBA.CodePane cp = ClsMisc.ActiveVBCodePane();

                    int iStartLine;
                    int iStartColumn;
                    int iEndLine;
                    int iEndColumn;

                    cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                    foreach (strFunctions objFunction in lstFunctions)
                    {
                        if (objFunction.iLineNoStart <= iStartLine && objFunction.iLineNoEnd >= iStartLine)
                        { bResult = true; }
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

        public strFunctions currentFunction
        {
            get
            {
                try
                {
                    strFunctions objResult = new strFunctions();

                    VBA.CodePane cp = ClsMisc.ActiveVBCodePane();

                    int iStartLine;
                    int iStartColumn;
                    int iEndLine;
                    int iEndColumn;

                    cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                    foreach (strFunctions objFunction in lstFunctions)
                    {
                        if (objFunction.iLineNoStart <= iStartLine && objFunction.iLineNoEnd >= iStartLine)
                        { objResult = objFunction; }
                    }

                    return objResult;
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

                    return new strFunctions();
                }
            }
        }

        public string cursorInFunctionName 
        { 
            get 
            {
                try
                {
                    string sResult = "";

                    VBA.CodePane cp = ClsMisc.ActiveVBCodePane();

                    int iStartLine;
                    int iStartColumn;
                    int iEndLine;
                    int iEndColumn;

                    cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                    foreach (strFunctions objFunction in lstFunctions) 
                    {
                        if (objFunction.iLineNoStart <= iStartLine & objFunction.iLineNoEnd > iStartLine)
                        { sResult = objFunction.sName; }
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

                    return string.Empty;
                }
            }
        }

        public int cursorCurrentIndentLevel(VBA.CodePane cp)
        {
            try
            {
                int iResult;
                if (cp.CodeModule.CountOfLines == 0)
                { iResult = 0; }
                else
                {
                    ClsSettings cSettings = new ClsSettings();

                    int iStartLine;
                    int iStartColumn;
                    int iEndLine;
                    int iEndColumn;
                    iResult = 0;

                    cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                    int iLineIndex = lstLines.FindIndex(x => x.iOriginalLineNo == iStartLine);
                    int iLineIndexAbove = iLineIndex;
                    int iLineIndexBelow = iLineIndex;

                    //find a line above which is not empty

                    if (iLineIndexAbove > 0)
                    {
                        bool bFinished = false;
                        while (!bFinished)
                        {
                            if (iLineIndexAbove <= 0)
                            {
                                bFinished = true;
                                iLineIndexAbove++;
                            }
                            else
                            {
                                if (lstLines[iLineIndexAbove].sText_Orig.Trim() == "")
                                { iLineIndexAbove--; }
                                else
                                { bFinished = true; }
                            }
                        }
                    }

                    //find a line below which is not empty
                    if (iLineIndexBelow > 0)
                    {
                        bool bFinished = false;
                        while (!bFinished)
                        {
                            if (iLineIndexBelow >= lstLines.Count)
                            {
                                bFinished = true;
                                iLineIndexBelow--;
                            }
                            else
                            {
                                if (lstLines[iLineIndexBelow].sText_Orig.Trim() == "")
                                { iLineIndexBelow++; }
                                else
                                { bFinished = true; }
                            }
                        }
                    }

                    //get the max indent of the above and below line.
                    int iLineIndentAbove = ClsMiscString.indentSize(lstLines[iLineIndexAbove].sText_Orig);
                    int iLineIndentBelow = ClsMiscString.indentSize(lstLines[iLineIndexBelow].sText_Orig);
                    int iLineIndent = Math.Max(iLineIndentAbove, iLineIndentBelow);

                    iResult = iLineIndent / cSettings.IndentSize;
                }

                return iResult;
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

                return 0;
            }
        }

        public void updateLine(strLine objLine)
        {
            try
            {
                int iLineNo = objLine.iIndex;
                lstLines[iLineNo] = objLine; 
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

        public void updateLine(int iLineNo, strLine objLine)
        {
            try
            {
                lstLines[iLineNo] = objLine;
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

        public int getLineNumber(strLine objLine) 
        {
            try
            {
                return lstLines.IndexOf(objLine);
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

                return -1;
            }
        }

        public int indentLevel(int iLineIndex)
        {
            try
            {
                int iResult = 0;

                if (0 < iLineIndex && iLineIndex < lstLines.Count)
                { iResult = lstLines[iLineIndex].sText_NoComment.Length - lstLines[iLineIndex].sText_NoComment.TrimStart().Length; }
                else
                { iResult = 0; }

                return iResult;
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

                return 0;
            }
        }
        
        
        public int cursorCurrentIndentLevel()
        {
            try
            {
                VBA.CodePane cp = ClsMisc.ActiveVBCodePane();
                ClsSettings cSettings = new ClsSettings();

                int iStartLine;
                int iStartColumn;
                int iEndLine;
                int iEndColumn;
                int iResult = 0;

                cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                int iLineIndex = lstLines.FindIndex(x => x.iOriginalLineNo == iStartLine);
                int iLineIndexAbove = iLineIndex;
                int iLineIndexBelow = iLineIndex;

                //find a line above which is not empty
                if (iLineIndexAbove > 0)
                {
                    bool bFinished = false;
                    while (!bFinished)
                    {
                        if (iLineIndexAbove <= 0)
                        { 
                            bFinished = true;
                            iLineIndexAbove++;
                        }
                        else
                        {
                            if (lstLines[iLineIndexAbove].sText_Orig.Trim() == "")
                            { iLineIndexAbove--; }
                            else
                            { bFinished = true; }
                        }
                    }
                }

                //find a line below which is not empty
                if (iLineIndexBelow > 0)
                {
                    bool bFinished = false;
                    while (!bFinished)
                    {
                        if (iLineIndexBelow >= lstLines.Count)
                        { 
                            bFinished = true;
                            iLineIndexBelow--;
                        }
                        else
                        {
                            if (lstLines[iLineIndexBelow].sText_Orig.Trim() == "")
                            { iLineIndexBelow++; }
                            else
                            { bFinished = true; }
                        }
                    }
                }

                //get the max indent of the above and below line.
                if (lstLines.Count == 0)
                { iResult = 0; }
                else
                {
                    int iLineIndentAbove = ClsMiscString.indentSize(lstLines[iLineIndexAbove].sText_Orig);
                    int iLineIndentBelow = ClsMiscString.indentSize(lstLines[iLineIndexBelow].sText_Orig);
                    int iLineIndent = Math.Max(iLineIndentAbove, iLineIndentBelow);

                    iResult = iLineIndent / cSettings.IndentSize;
                }

                return iResult;
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

                return 0;
            }
        }
        /*
        public int cursorCurrentIndentLevel(string sCodePaneName)
        {
            try
            {
                VBA.CodePane cp = ClsMisc.VBCodePane(sCodePaneName);
                ClsSettings cSettings = new ClsSettings();

                int iStartLine;
                int iStartColumn;
                int iEndLine;
                int iEndColumn;
                int iResult = 0;

                cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                int iLineIndex = lstLines.FindIndex(x => x.iOriginalLineNo == iStartLine);
                int iLineIndexAbove = iLineIndex;
                int iLineIndexBelow = iLineIndex;

                //find a line above which is not empty
                if (iLineIndexAbove > 0)
                {
                    bool bFinished = false;
                    while (!bFinished)
                    {
                        if (iLineIndexAbove <= 0)
                        {
                            bFinished = true;
                            iLineIndexAbove++;
                        }
                        else
                        {
                            if (lstLines[iLineIndexAbove].sText_Orig.Trim() == "")
                            { iLineIndexAbove--; }
                            else
                            { bFinished = true; }
                        }
                    }
                }

                //find a line below which is not empty
                if (iLineIndexBelow > 0)
                {
                    bool bFinished = false;
                    while (!bFinished)
                    {
                        if (iLineIndexBelow >= lstLines.Count)
                        {
                            bFinished = true;
                            iLineIndexBelow--;
                        }
                        else
                        {
                            if (lstLines[iLineIndexBelow].sText_Orig.Trim() == "")
                            { iLineIndexBelow++; }
                            else
                            { bFinished = true; }
                        }
                    }
                }

                //get the max indent of the above and below line.
                if (lstLines.Count == 0)
                { iResult = 0; }
                else
                {
                    int iLineIndentAbove = ClsMiscString.indentSize(lstLines[iLineIndexAbove].sText_Orig);
                    int iLineIndentBelow = ClsMiscString.indentSize(lstLines[iLineIndexBelow].sText_Orig);
                    int iLineIndent = Math.Max(iLineIndentAbove, iLineIndentBelow);

                    iResult = iLineIndent / cSettings.IndentSize;
                }

                return iResult;
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

                return 0;
            }
        }
        */
        private void findEndOfFunctions() 
        {
            try
            {
                /*
                 loop through the lines and for each line find a the function that corisponds then check the end line number is correct.
                 End line number must have "End Sub", "End Function" or "End Property"
                 */

                foreach (strLine objLine in lstLines) 
                {
                    if (countIF(objLine.lstLineType, enumLineType.eLineType_EndFunction) > 0)
                    {
                        if (lstFunctions.Exists(x => x.sName == objLine.sFunctionName && x.ePropertyType == objLine.ePropertyType))
                        {
                            int iFunctionIndex = lstFunctions.FindIndex(x => x.sName == objLine.sFunctionName && x.ePropertyType == objLine.ePropertyType);

                            strFunctions objFunction = lstFunctions[iFunctionIndex];

                            if (objFunction.iLineNoEnd == 0)
                            {
                                objFunction.iLineNoEnd = objLine.iOriginalLineNo;
                                lstFunctions[iFunctionIndex] = objFunction;
                            }
                            else if (objFunction.iLineNoEnd < objLine.iOriginalLineNo)
                            {
                                objFunction.iLineNoEnd = objLine.iOriginalLineNo;
                                lstFunctions[iFunctionIndex] = objFunction;
                            }
                        }
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

        private bool containsSingeQuoteOrREM(string sLine)
        {
            try
            {
                bool bResult;

                //Comment could begin with either single quote or REM

                /*
                 WRONG: REM only works at the beginning of the line and don't check char after it.
                 */

                if (sLine.Contains(csChar_SingleQuote))
                { bResult = true; }
                else
                {
                    if (sLine.ToUpper().Contains("REM"))
                    {
                        //Could be a longer word
                        bool bPreseededByChar;
                        //bool bFollowedByChar;

                        int iPosRem = sLine.ToUpper().IndexOf("REM");

                        if (iPosRem == 0)
                        { bPreseededByChar = false; }
                        else
                        {
                            //char cPressedingChar = sLine.Substring(iPosRem-1,1).ToUpper().ToCharArray()[0];

                            if (char.IsLetterOrDigit(sLine, iPosRem - 1))
                            { bPreseededByChar = true; }
                            else
                            { bPreseededByChar = false; }
                        }

                        if (bPreseededByChar)
                        { bResult = false; }
                        else
                        { bResult = true; }
                    }
                    else
                    { bResult = false; }
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

        public void dumpIntoExcelSheet(Excel.Workbook wrk)
        {
            try
            {
                /*****
                 Lines
                 *****/
                Excel.Worksheet shtLines = wrk.Worksheets.Add(Type.Missing, Type.Missing);

                shtLines.Name = ClsMisc.newWorksheetName(ref wrk, "Lines");

                int iRowLines = 1;
                const int iColLineFunctionName = 1;
                const int iColLineOrigLineNo = 2;
                const int iColLineOrigText = 3;
                const int iColLineComment = 4;
                const int iColLineNoComment = 5;

                shtLines.Cells[iRowLines, iColLineFunctionName].Value = "sFunctionName";
                shtLines.Cells[iRowLines, iColLineOrigLineNo].Value = "iOriginalLineNo";
                shtLines.Cells[iRowLines, iColLineOrigText].Value = "sText_Orig";
                shtLines.Cells[iRowLines, iColLineComment].Value = "sText_Comment";
                shtLines.Cells[iRowLines, iColLineNoComment].Value = "sText_NoComment";

                foreach (strLine objLine in lstLines)
                {
                    iRowLines++;

                    shtLines.Cells[iRowLines, iColLineFunctionName].Value = objLine.sFunctionName;
                    shtLines.Cells[iRowLines, iColLineOrigLineNo].Value = objLine.iOriginalLineNo;
                    shtLines.Cells[iRowLines, iColLineOrigText].Value = objLine.sText_Orig;
                    shtLines.Cells[iRowLines, iColLineComment].Value = objLine.sText_Comment;
                    shtLines.Cells[iRowLines, iColLineNoComment].Value = objLine.sText_NoComment;
                }

                /*********
                 Functions
                 *********/
                Excel.Worksheet shtFunctions = wrk.Worksheets.Add(Type.Missing, Type.Missing);

                shtFunctions.Name = ClsMisc.newWorksheetName(ref wrk, "Functions");
                
                int iRowFunction = 1;
                const int iColFnFunctionName = 1;
                const int iColFnType = 2;
                const int iColFnVarCount = 3;
                const int iColFnLineNoStart = 4;
                const int iColFnLineNoEnd = 5;

                shtFunctions.Cells[iRowFunction, iColFnFunctionName].Value = "sFunctionName";
                shtFunctions.Cells[iRowFunction, iColFnType].Value = "eFunctionType";
                shtFunctions.Cells[iRowFunction, iColFnVarCount].Value = "lstVariablesFn.Count()";
                shtFunctions.Cells[iRowFunction, iColFnLineNoStart].Value = "iLineNoStart";
                shtFunctions.Cells[iRowFunction, iColFnLineNoEnd].Value = "iLineNoEnd";

                foreach (strFunctions objFunctions in lstFunctions)
                {
                    iRowFunction++;

                    shtFunctions.Cells[iRowFunction, iColFnFunctionName].Value = objFunctions.sName;
                    shtFunctions.Cells[iRowFunction, iColFnType].Value = objFunctions.eFunctionType.ToString();
                    shtFunctions.Cells[iRowFunction, iColFnVarCount].Value = objFunctions.lstVariablesFn.Count();
                    shtFunctions.Cells[iRowFunction, iColFnLineNoStart].Value = objFunctions.iLineNoStart;
                    shtFunctions.Cells[iRowFunction, iColFnLineNoEnd].Value = objFunctions.iLineNoEnd;
                }

                /*****
                 Variable
                 *****/
                Excel.Worksheet shtVar = wrk.Worksheets.Add(Type.Missing, Type.Missing);

                shtVar.Name = ClsMisc.newWorksheetName(ref wrk, "Variable");
                
                int iRowVar = 1;
                const int iColVarFunction = 1;
                const int iColVarName = 2;
                const int iColVarParameter = 3;
                const int iColVarScope = 4;
                const int iColVarType = 5;
                const int iColVarDataType = 6;
                const int iColVarParaType = 7;

                shtVar.Cells[iRowVar, iColVarFunction].Value = "Function";
                shtVar.Cells[iRowVar, iColVarName].Value = "sName";
                shtVar.Cells[iRowVar, iColVarParameter].Value = "bIsParameter";
                shtVar.Cells[iRowVar, iColVarScope].Value = "eScope";
                shtVar.Cells[iRowVar, iColVarType].Value = "eType";
                shtVar.Cells[iRowVar, iColVarDataType].Value = "sDatatype";
                shtVar.Cells[iRowVar, iColVarParaType].Value = "eParaType";

                foreach (strVariables objVar in lstVariablesMod)
                {
                    iRowVar++;

                    shtVar.Cells[iRowVar, iColVarFunction].Value = "<No Function>";
                    shtVar.Cells[iRowVar, iColVarName].Value = objVar.sName;
                    shtVar.Cells[iRowVar, iColVarParameter].Value = objVar.bIsParameter;
                    shtVar.Cells[iRowVar, iColVarScope].Value = objVar.eScope.ToString();
                    shtVar.Cells[iRowVar, iColVarType].Value = objVar.eType.ToString();
                    shtVar.Cells[iRowVar, iColVarDataType].Value = objVar.sDatatype.ToString();
                    shtVar.Cells[iRowVar, iColVarParaType].Value = objVar.eParaType.ToString();
                }

                foreach (strFunctions objFunctions in lstFunctions)
                {
                    
                    foreach(strVariables objVar in objFunctions.lstVariablesFn)
                    {
                        iRowVar++;

                        shtVar.Cells[iRowVar, iColVarFunction].Value = objFunctions.sName;
                        shtVar.Cells[iRowVar, iColVarName].Value = objVar.sName;
                        shtVar.Cells[iRowVar, iColVarParameter].Value = objVar.bIsParameter;
                        shtVar.Cells[iRowVar, iColVarScope].Value = objVar.eScope.ToString();
                        shtVar.Cells[iRowVar, iColVarType].Value = objVar.eType.ToString();
                        shtVar.Cells[iRowVar, iColVarDataType].Value = objVar.sDatatype.ToString();
                        shtVar.Cells[iRowVar, iColVarParaType].Value = objVar.eParaType.ToString();
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

        public bool variableNameExistsInModule(string sNewName, bool bOnlyGlobal) 
        {
            try
            {
                /*
                 Checks if global variables or function names exist in module
                 */
                bool bIsFound = false;

                foreach (strFunctions objFunction in lstFunctions)
                {
                    if (bOnlyGlobal)
                    {
                        if (objFunction.eScope == enumScopeFn.eScopeFn_Public & objFunction.sName.Trim().ToLower() == sNewName.Trim().ToLower())
                        { bIsFound = true; }
                    }
                    else
                    {
                        //function name is in confict
                        if (objFunction.sName.Trim().ToLower() == sNewName.Trim().ToLower())
                        { bIsFound = true; }

                        //variable in function is in conflict
                        foreach (strVariables objVarInFn in objFunction.lstVariablesFn) 
                        {
                            if (objVarInFn.sName.Trim().ToLower() == sNewName.Trim().ToLower())
                            { bIsFound = true; }
                        }
                    }
                }

                if (!bIsFound)
                {
                    //variable that are global through out module are in confict
                    foreach (strVariables objVar in this.lstVariablesMod)
                    {
                        if (objVar.sName.Trim().ToLower() == sNewName.Trim().ToLower())
                        { bIsFound = true; }
                    }
                }

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

                return false;
            }
        }

        public void renameModuleInList(string sNewName) 
        {
            try
            {
                sModuleName = sNewName;
                objModuleDetails.sName = sNewName;
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

        public void renameFunctionInList(string sNewName, string sOldName)
        {
            try
            {
                List<int> lstIndexes = new List<int>();

                foreach(strFunctions objFunction in lstFunctions.FindAll(x => x.sName.Trim().ToLower() == sOldName.Trim().ToLower()))
                { 
                    int iIndex = lstFunctions.FindIndex(x => x.sName == objFunction.sName && x.ePropertyType == objFunction.ePropertyType);
                    lstIndexes.Add(iIndex); 
                }

                foreach (int iIndex in lstIndexes)
                {
                    strFunctions objFunction = lstFunctions[iIndex];
                    objFunction.sName = sNewName;
                    lstFunctions[iIndex] = objFunction;
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

        public void renameVariableInFunctionList(string sVariableNameOld, string sVariableNameNew, string sFunctionName)
        {
            try
            {
                /*
                 check if variable name exists locally in a function
                 */
                bool bIsFound = false;
                int iFnIndex = lstFunctions.FindIndex(x => x.sName.Trim().ToLower() == sFunctionName.Trim().ToLower());

                if (iFnIndex == -1)
                { bIsFound = false; }
                else
                {
                    int iVarIndex = lstFunctions[iFnIndex].lstVariablesFn.FindIndex(x => x.sName.Trim().ToLower() == sVariableNameOld.Trim().ToLower());

                    strVariables objVar = lstFunctions[iFnIndex].lstVariablesFn[iVarIndex];
                    objVar.sName = sVariableNameNew;
                    lstFunctions[iFnIndex].lstVariablesFn[iVarIndex] = objVar;
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

        public bool variableNameExistsInFunction(string sVariableName, string sFunctionName) 
        {
            try
            {
                /*
                 check if variable name exists locally in a function
                 */
                bool bIsFound = false;
                int iFnIndex = lstFunctions.FindIndex(x => x.sName.Trim().ToLower() == sFunctionName.Trim().ToLower());

                if (iFnIndex == -1)
                { bIsFound = false; }
                else
                {
                    if (lstFunctions[iFnIndex].lstVariablesFn.Exists(x => x.sName.Trim().ToLower() == sVariableName.Trim().ToLower()))
                    { bIsFound = true; }
                }

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
                
                return false;
            }
        }

        public List<ClsCodeMapper.strLine> findModuleReferences(string sModuleName)
        {
            try
            {
                List<ClsCodeMapper.strLine> lstResults = new List<ClsCodeMapper.strLine>();

                foreach (ClsCodeMapper.strLine objLine in lstLines.FindAll(x => x.sText_NoComment.ToUpper().Trim().Contains(sModuleName.ToUpper().Trim())))
                {
                    bool bIsIncluded = false;

                    if (objLine.sText_NoComment.Contains('"'))
                    {
                        int iPosDoubleQuote = objLine.sText_NoComment.IndexOf('"');
                        int iPosModName = objLine.sText_NoComment.IndexOf(sModuleName);

                        if (iPosModName < iPosDoubleQuote)
                        { bIsIncluded = true; }
                        else
                        {
                            int iNoOfQuotesToLeft = ClsMiscString.Left(objLine.sText_NoComment, iPosModName).Count(x => x == '"');

                            if (iNoOfQuotesToLeft % 2 == 0)
                            { bIsIncluded = true; }
                            else
                            { bIsIncluded = false; }
                        }
                    }
                    else
                    { bIsIncluded = true; }

                    /*
                     * To do:
                     * If searching for a class then we will find the "Dim x as Class1" line 
                     * but then we will want to also catch all the lines with x in then after this
                     */


                    if (bIsIncluded)
                    { 
                        lstResults.Add(objLine);

                    /*
                        //if it's the decloration of a variable
                        if (objLine.lstLineType.Contains(enumLineType.eLineType_Dim))
                        {
                            string sVarName = "";
                            
                            if (objLine.sFunctionName == "")
                            {
                                //local to a module

                                //global
                            }
                            else
                            {
                                //local to a function
                            }
                        }
                    */
                    }
                }



                return lstResults;
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

                return new List<ClsCodeMapper.strLine>();
            }
        }


        public void RenameVariable(string sNewName, string sOldName)
        {
            try
            {
                for (int iIndex = 0; iIndex < lstLines.Count; iIndex++)
                {
                    strLine objTemp = lstLines[iIndex];

                    string sTemp = objTemp.sText_NoComment;

                    ClsMiscRename.RenameVariable(sNewName, sOldName, ref sTemp);

                    objTemp.sText_NoComment = sTemp;

                    lstLines[iIndex] = objTemp;
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

        public void RenameVariable(string sNewName, string sOldName, string sFunctionName)
        {
            try
            {
                for (int iIndex = 0; iIndex < lstLines.Count; iIndex++)
                {
                    if (ClsMiscString.ingoreNull(lstLines[iIndex].sFunctionName).Trim().ToLower() == ClsMiscString.ingoreNull(sFunctionName).Trim().ToLower())
                    {
                        strLine objTemp = lstLines[iIndex];

                        string sTemp = objTemp.sText_NoComment;

                        ClsMiscRename.RenameVariable(sNewName, sOldName, ref sTemp);

                        objTemp.sText_NoComment = sTemp;

                        lstLines[iIndex] = objTemp;
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
        
        public strFunctions getFunction(string sName)
        {
            try
            {
                strFunctions objFunction = lstFunctions.Find(x => x.sName.Trim().ToLower() == sName.Trim().ToLower());

                return objFunction;
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

                return new strFunctions();
            }
        }

        public void RenameFunction(string sNewName, string sOldName)
        {
            try
            {
                RenameVariable(sNewName, sOldName);

                //if (lstFunctions.Exists(x => x.sName.Trim().ToLower() == sOldName.Trim().ToLower()))
                //{
                //    strFunctions objFn = lstFunctions.Find(x => x.sName.Trim().ToLower() == sOldName.Trim().ToLower());
                //    if (objFn.eScope == enumScopeFn.eScopeFn_Private)
                //    {
                //    }
                //}
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
        /*
        public void RenameModule()
        {
            try
            {


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
        */
        public static string convertToText(enumScopeFn eScopeFn)
        {
            try
            {
                string sResult;

                switch (eScopeFn)
                {
                    case enumScopeFn.eScopeFn_Friend:
                        sResult = "Friend";
                        break;
                    case enumScopeFn.eScopeFn_Private:
                        sResult = "Private";
                        break;
                    case enumScopeFn.eScopeFn_Public:
                        sResult = "Public";
                        break;
                    default:
                        sResult = "";
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
                return string.Empty;
            }
        }

        public static string convertToText(enumFunctionType eFunctionType)
        {
            try
            {
                string sResult;

                switch (eFunctionType)
                {
                    case enumFunctionType.eFnType_Error:
                        sResult = "Error";
                        break;
                    case enumFunctionType.eFnType_Function:
                        sResult = "Function";
                        break;
                    case enumFunctionType.eFnType_Sub:
                        sResult = "Sub";
                        break;
                    case enumFunctionType.eFnType_Property:
                        sResult = "Property";
                        break;
                    case enumFunctionType.eFnType_None:
                        sResult = "None";
                        break;
                    default:
                        sResult = "";
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
                return string.Empty;
            }
        }

        public static string convertToText(enumFunctionPropertyType ePropertyType)
        {
            try
            {
                string sResult;

                switch (ePropertyType)
                {
                    case enumFunctionPropertyType.ePropType_Let:
                        sResult = "Let";
                        break;
                    case enumFunctionPropertyType.ePropType_Get:
                        sResult = "Get";
                        break;
                    case enumFunctionPropertyType.ePropType_Set:
                        sResult = "Set";
                        break;
                    case enumFunctionPropertyType.ePropType_NA:
                        sResult = "NA";
                        break;
                    default:
                        sResult = "";
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
                return string.Empty;
            }
        }

        public void RenameModule(string sName)
        {
            try
            {
                vbComponent.Name = sName;
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

        public bool existsGoToWithLineNo(string sFunctionName)
        {
            bool bResult = false;

            try
            {
                bool bIsFound = false;

                foreach (strLine objLine in lstLines.FindAll(x => x.sFunctionName.ToLower().Trim() == sFunctionName.ToLower().Trim() && x.sText_NoComment.ToLower().Contains("goto")))
                {
                    //Loop through all the lines of code with the text "goto" and 
                    
                    string sLine = objLine.sText_NoComment;

                    for (int iPos = 0; iPos < sLine.Length - 5; iPos++) //note -5 refers to -1 converting base 0 pos and base 1 lengh and -4 for the goto text
                    {
                        if(sLine.ToLower().Substring(iPos, 4) == "goto")
                        {
                            string sOnRight = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 4).Trim();

                            if (sOnRight.Length > 0)
                            {
                                if (Regex.IsMatch(ClsMiscString.Left(ref sOnRight, 1), "[0-9]"))
                                { bIsFound = true; }
                            }
                        }
                    }
                }

                bResult = bIsFound;

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

                return bResult;
            }
        }

        public void removeLineNo()
        {
            try
            {
                //foreach (strLine objLine in lstLines)
                for (int iIndex = 0; iIndex < lstLines.Count; iIndex++)
                {
                    //int iPos = lstLines.IndexOf(objLine);
                    strLine objTemp = lstLines[iIndex];

                    int iLineNo = 0;

                    if (int.TryParse(objTemp.sLineNo.Trim(), out iLineNo))
                    {
                        if (!lstLineNoReferenced.Contains(iLineNo))
                        { objTemp.sLineNo = ""; }
                    }
                    else
                    { objTemp.sLineNo = ""; }

                    lstLines[iIndex] = objTemp;
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

        public void removeLineNo(string sFunctionName)
        {
            try
            {
                foreach (strLine objLine in lstLines.FindAll(x => x.sFunctionName.Trim().ToLower() == sFunctionName.Trim().ToLower() || sFunctionName.Trim().ToLower() == ClsDefaults.textAll.Trim().ToLower()))
                {
                    int iPos = lstLines.IndexOf(objLine);
                    strLine objTemp = objLine;

                    int iLineNo = 0;

                    if (int.TryParse(objTemp.sLineNo.Trim(), out iLineNo))
                    {
                        if (!lstLineNoReferenced.Contains(iLineNo))
                        { objTemp.sLineNo = ""; }
                    }
                    else
                    { objTemp.sLineNo = ""; }

                    lstLines[iPos] = objTemp;
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

        public void splitLines()
        {
            try
            {
                splitLines(lstLines);
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

        public void splitLines(string sFunctionName)
        {
            try
            {
                splitLines(lstLines.FindAll(x => x.sFunctionName.Trim().ToLower() == sFunctionName.Trim().ToLower()));
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

        private void splitLines(List<strLine> lst) 
        {
            try
            {
                splitLinesColon(lst);
                splitLinesDim(lst);
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

        private void splitLinesDim(List<strLine> lst) 
        {
            try
            {
                List<int> lstIndexes = new List<int>();
                
                foreach (strLine objLine in lst)
                { lstIndexes.Add(lstLines.IndexOf(objLine)); }

                lstIndexes.Sort();

                for (int iIndex = lstIndexes.Count - 1; iIndex > 0; iIndex--)
                {
                    strLine objLine = lstLines[lstIndexes[iIndex]];

                    string sLine = objLine.sText_NoComment;
                    int iIndent = sLine.Length - sLine.TrimStart().Length;

                    //if Dim statement
                    bool bIsVariableDeclaration = false;

                    //note single & means "contains" is fast but extra might get through but second half of "if" will get it absolutely right
                    if (sLine.Contains(',') & ClsMiscString.containsOutsideQuotes(sLine, ","))
                    {
                        if (sLine.Trim().ToUpper().StartsWith("Dim ", StringComparison.CurrentCultureIgnoreCase))
                        { bIsVariableDeclaration = true; }
                        else if (sLine.Trim().ToUpper().StartsWith("Const ", StringComparison.CurrentCultureIgnoreCase))
                        { bIsVariableDeclaration = true; }
                        else if (!(sLine.Contains(" Sub ") || sLine.Contains(" Function ") || sLine.Contains(" Property ")))
                        {
                            if ((sLine.Trim().ToUpper().StartsWith("Public ", StringComparison.CurrentCultureIgnoreCase) | sLine.Trim().ToUpper().StartsWith("Private ", StringComparison.CurrentCultureIgnoreCase)))
                            { bIsVariableDeclaration = true; }
                        }
                    }

                    if (bIsVariableDeclaration)
                    {
                        bool bIsCommaInBrackets = false;
                        bool bIsFinished = false;

                        int iBracketCurved = 0;
                        int iBracketSquare = 0;
                        int iBracketSquigly = 0;
                        int iPos = 0;

                        while (!bIsFinished)
                        {
                            char cCurrChar = sLine[iPos];

                            if (cCurrChar == '"')
                            { bIsCommaInBrackets = !bIsCommaInBrackets; }

                            if (!bIsCommaInBrackets)
                            {
                                switch (cCurrChar)
                                {
                                    case '(':
                                        iBracketCurved++;
                                        break;
                                    case ')':
                                        iBracketCurved--;
                                        break;
                                    case '{':
                                        iBracketSquigly++;
                                        break;
                                    case '}':
                                        iBracketSquigly--;
                                        break;
                                    case '[':
                                        iBracketSquare++;
                                        break;
                                    case ']':
                                        iBracketSquare--;
                                        break;
                                }

                                if (cCurrChar == ',' & iBracketCurved == 0 & iBracketSquare == 0 & iBracketSquigly == 0)
                                {
                                    //cut here
                                    string sPrefix = "";
                                    if (sLine.Trim().ToUpper().StartsWith("DIM ", StringComparison.CurrentCultureIgnoreCase))
                                    { sPrefix = "Dim"; }

                                    if (sLine.Trim().ToUpper().StartsWith("CONST ", StringComparison.CurrentCultureIgnoreCase))
                                    { sPrefix = "Const"; }

                                    if (sLine.Trim().ToUpper().StartsWith("PRIVATE ", StringComparison.CurrentCultureIgnoreCase))
                                    { sPrefix = "private"; }

                                    if (sLine.Trim().ToUpper().StartsWith("PUBLIC ", StringComparison.CurrentCultureIgnoreCase))
                                    { sPrefix = "public"; }

                                    strLine objLine1 = objLine;
                                    strLine objLine2 = objLine;

                                    objLine1.sText_NoComment = ClsMiscString.Left(ref sLine, iPos).Trim();
                                    objLine1.sText_NoComment = objLine1.sText_NoComment.PadLeft(iIndent + objLine1.sText_NoComment.Length, ' ');
                                    
                                    objLine2.sText_NoComment = sPrefix + " " + ClsMiscString.Right(ref sLine, sLine.Length - iPos - 1);
                                    objLine2.sText_NoComment = objLine2.sText_NoComment.PadLeft(iIndent + objLine2.sText_NoComment.Length, ' ');

                                    //lstLines[iIndex] = objLine1;
                                    //lstLines.Insert(iIndex, objLine2);

                                    lstLines.Insert(iIndex, objLine1);
                                    iIndex++;
                                    lstLines[iIndex] = objLine2;

                                    objLine = objLine2;

                                    sLine = objLine2.sText_NoComment;
                                    iPos = 0;
                                }

                                if (iPos >= sLine.Length - 1)
                                { bIsFinished = true; }
                                else
                                { iPos++; }
                            }
                        }
                    }

                    objLine.sText_NoComment = sLine;
                    
                    lstLines[lstIndexes[iIndex]] = objLine;
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

        private void splitLinesColon(List<strLine> lst) 
        {
            try
            {
                List<int> lstIndexes = new List<int>();
                
                foreach (strLine objLine in lst)
                { lstIndexes.Add(lstLines.IndexOf(objLine)); }

                lstIndexes.Sort();

                for (int iIndex = lstIndexes.Count - 1; iIndex > 0; iIndex--)
                {
                    strLine objLine = lstLines[lstIndexes[iIndex]];

                    string sLine = objLine.sText_NoComment;

                    //if ; is joining two statements
                    if (sLine.Contains(':') & ClsMiscString.containsOutsideQuotes(sLine, ":"))
                    {
                        int iPos = sLine.Length - sLine.TrimStart().Length;
                        bool bIsFinished = false;
                        bool bIsInQuotes = false;

                        while(!bIsFinished)
                        {
                            char cCurr = sLine[iPos];
                            char cNext = ' ';

                            if (sLine.Length > iPos + 1)
                            { cNext = sLine[iPos + 1]; }

                            if (cCurr == '"')
                            { bIsInQuotes=!bIsInQuotes; }

                            if (!bIsInQuotes)
                            {
                                if (cCurr == ':' && cNext != '=')
                                {
                                    strLine objLine1 = objLine;
                                    strLine objLine2 = objLine;

                                    objLine1.sText_NoComment = ClsMiscString.Left(ref sLine, iPos);
                                    objLine2.sText_NoComment = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 1);

                                    lstLines[iIndex] = objLine1;
                                    lstLines.Insert(iIndex, objLine2);

                                    objLine = objLine2;

                                    sLine = objLine2.sText_NoComment;

                                    iPos = 0;
                                }
                            }

                            if (iPos < sLine.Length - 1)
                            { iPos++; }
                            else
                            { bIsFinished = true; }
                        }
                    }

                    objLine.sText_NoComment = sLine;
                    
                    lstLines[lstIndexes[iIndex]] = objLine;
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

        public void setLineLength()
        {
            try
            {
                setLineLength(this.vbComponent.CodeModule.CodePane);
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


        public void setLineLength(VBA.CodePane objCodePane)
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();
                bool bIsFinished = false;
                bool bIsDoubleQuotes = false;
                int iNoOfCuts = 0;
                int iLinePos = 0;
                List<string> lstNewLines = new List<string>();

                for (int iIndex = lstLines.Count - 1; iIndex >= 0; iIndex--)
                {
                    strLine objLine = lstLines[iIndex];
                    lstNewLines = new List<string>();

                    string sLineNew = objLine.sText_NoComment;

                    ClsMiscString.RemoveDoubleSpaces(ref sLineNew);

                    if (ClsMiscString.Left(sLineNew.TrimStart(), 1) == "'" || ClsMiscString.Left(sLineNew.ToUpper().TrimStart(), 3) == "REM")
                    {
                        lstNewLines.Add(sLineNew);
                        iLinePos++;
                    }
                    else
                    {
                        int iCharCutOff = cSettings.InsertCode_Format_CharCutOffPoint;
                        int iNextIndenting = 0;
                        int iPosCut = 0;
                        ClsInsertCode.findWhereToCutLine(ref cSettings, sLineNew, iCharCutOff, ref iPosCut, ref iNextIndenting, objLine.lstLineType);

                        string sLineSoFar = ClsMiscString.Left(ref sLineNew, iPosCut);
                        string sLineRemainer = ClsMiscString.Right(ref sLineNew, sLineNew.Length - iPosCut);
                        /*
                         * new indent is for new lines it's the previous open bracket from the cut.
                         * however if there is no previous open bracket then excluding the indent 
                         * it's the first white space that is not inclosed in double quotes
                         */
                        if (iPosCut == 0 || iPosCut >= sLineNew.Length - 1)
                        {
                            if (sLineNew.Trim() != "_")
                            {
                                lstNewLines.Add(sLineNew);
                                iLinePos++;
                            }
                        }
                        else
                        {
                            bool bIsLineFunctionDeclare = false;
                            bool bIsLineEquals = false;

                            if (sLineNew.Trim().ToUpper().StartsWith("PUBLIC SUB ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PRIVATE SUB ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PUBLIC FUNCTION ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PRIVATE FUNCTION ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PUBLIC PROPERTY ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PRIVATE PROPERTY ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("SUB ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("FUNCTION ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }
                            if (sLineNew.Trim().ToUpper().StartsWith("PROPERTY ", StringComparison.CurrentCultureIgnoreCase))
                            { bIsLineFunctionDeclare = true; }

                            if (sLineNew.Contains('='))
                            { bIsLineEquals = true; }

                            iNoOfCuts++; //First line

                            string sLineOne = ClsMiscString.Left(ref sLineNew, iPosCut);
                            bIsFinished = false;

                            sLineOne += " _";

                            lstNewLines.Add(sLineOne);
                            iLinePos++;

                            while (!bIsFinished)
                            {
                                int iPosNextCut = 0;

                                if (iNoOfCuts < 22)
                                { ClsInsertCode.findNextCut(ref cSettings, sLineNew, iPosCut, ref iPosNextCut, ref iNextIndenting, bIsLineFunctionDeclare, bIsLineEquals); }
                                else
                                { iPosNextCut = 0; }

                                string sLineNext;

                                if (iPosNextCut == 0)
                                {
                                    sLineNext = ClsMiscString.Right(ref sLineNew, sLineNew.Length - iPosCut).TrimStart();
                                    sLineNext = sLineNext.PadLeft(iNextIndenting + sLineNext.Length);
                                    bIsFinished = true;
                                }
                                else
                                {
                                    iNoOfCuts++;
                                    sLineNext = sLineNew.Substring(iPosCut, iPosNextCut - iPosCut).TrimStart() + " _ ";
                                    sLineNext = sLineNext.PadLeft(iNextIndenting + sLineNext.Length);
                                }

                                if (sLineNext.Trim() != "_")
                                {
                                    lstNewLines.Add(sLineNext);
                                    iLinePos++;
                                }

                                if (iPosNextCut >= sLineNew.Length - 1)
                                { bIsFinished = true; }

                                iPosCut = iPosNextCut;
                            }
                        }
                    }

                    if (iNoOfCuts > 0)
                    {
                        //add new ones
                        int iPos = lstLines.IndexOf(objLine);

                        foreach (string sTemp in lstNewLines)
                        {
                            strLine objLineNew = objLine;

                            objLineNew.sText_NoComment = sTemp;

                            if (ClsMiscString.Right(sTemp.TrimEnd(), 2) == " _")
                            { objLineNew.sText_Comment = ""; }

                            this.lstLines.Insert(iPos, objLineNew);
                            iPos++;
                        }

                        //remove old
                        lstLines.RemoveAt(lstLines.IndexOf(objLine));
                    }
                }

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

        public void alignVariableDim() 
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                enumVarDimType eVarDimType = cSettings.FormatVarDimType;
                
                switch (eVarDimType)
                {
                    case enumVarDimType.eVarDim_InLine:
                        List<string> lstFunctionNames = new List<string>();

                        foreach (strFunctions objFunctions in lstFunctions)
                        { lstFunctionNames.Add(objFunctions.sName); }

                        lstFunctionNames.Add(string.Empty);//Used for global variables

                        foreach (string sFunctionName in lstFunctionNames)
                        {
                            List<strVariables> lstVar = this.lstVariables();
                            lstVar = lstVar.FindAll(x => x.sFunctionName.Trim().ToLower() == sFunctionName.Trim().ToLower() & x.bIsParameter == false);

                            if (lstVar.Count > 0)
                            {
                                int iMaxVariableNameLength = lstVar.Max(x => x.sName.Trim().Length);
                                List<int> lstIndexes = new List<int>();

                                foreach (strLine objLine in lstLines.FindAll(x => x.sFunctionName.Trim().ToLower() == sFunctionName.Trim().ToLower() && countIF(x.lstLineType, enumLineType.eLineType_Dim) > 0))
                                { lstIndexes.Add(lstLines.IndexOf(objLine));}

                                int iMaxPrefixCharCount = 0;

                                foreach (int iIndex in lstIndexes)
                                {
                                    strLine objLine = lstLines[iIndex];

                                    string sLine = objLine.sText_NoComment;
                                
                                    int iPrefixCharCount = sLine.Length - sLine.TrimStart().Length;

                                    //this needs to move to calculate the max per function.
                                    if (sLine.Trim().ToUpper().StartsWith("dim ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "dim ".Length; }
                                    else if (sLine.Trim().ToUpper().StartsWith("public const ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "public const ".Length; }
                                    else if (sLine.Trim().ToUpper().StartsWith("public ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "public ".Length; }
                                    else if (sLine.Trim().ToUpper().StartsWith("private const ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "private const ".Length; }
                                    else if (sLine.Trim().ToUpper().StartsWith("private ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "private ".Length; }
                                    else if (sLine.Trim().ToUpper().StartsWith("const ", StringComparison.CurrentCultureIgnoreCase))
                                    { iPrefixCharCount += "const ".Length; }

                                    if (iMaxPrefixCharCount < iPrefixCharCount)
                                    { iMaxPrefixCharCount = iPrefixCharCount; }
                                }


                                foreach (int iIndex in lstIndexes)
                                {
                                    strLine objLine = lstLines[iIndex];

                                    string sLine = objLine.sText_NoComment;

                                    /*
                                    int iPrefixCharCount = sLine.Length - sLine.TrimStart().Length;

                                    //this needs to move to calculate the max per function.
                                    if (ClsMiscString.stringStartsWith(sLine, "dim "))
                                    { iPrefixCharCount += "dim ".Length; }
                                    else if (ClsMiscString.stringStartsWith(sLine, "public const "))
                                    { iPrefixCharCount += "public const ".Length; }
                                    else if (ClsMiscString.stringStartsWith(sLine, "public "))
                                    { iPrefixCharCount += "public ".Length; }
                                    else if (ClsMiscString.stringStartsWith(sLine, "private const "))
                                    { iPrefixCharCount += "private const ".Length; }
                                    else if (ClsMiscString.stringStartsWith(sLine, "private "))
                                    { iPrefixCharCount += "private ".Length; }
                                    else if (ClsMiscString.stringStartsWith(sLine, "const "))
                                    { iPrefixCharCount += "const ".Length; }
                                    */

                                    for (int iPos = sLine.Length - 4; iPos > 0; iPos--)
                                    {
                                        if (sLine.Substring(iPos, 4).ToLower() == " as ")
                                        {
                                            if (iPos > iMaxVariableNameLength + iMaxPrefixCharCount + 1) //5=>"Dim ".length = 4 before variable and 1 after variable
                                            {
                                                if (sLine.Substring(iPos - 1, 5).ToLower() == "  as ")
                                                {
                                                    string sBefore = ClsMiscString.Left(ref sLine, iPos - 1);
                                                    string sAfter = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 3);

                                                    sLine = sBefore + " As " + sAfter;
                                                }
                                                //remove spaces one at a time checking along the way
                                            }

                                            if (iPos < iMaxVariableNameLength + iMaxPrefixCharCount + 1) //5=>"Dim ".length = 4 before variable and 1 after variable
                                            {
                                                //add spaces to pad out
                                                string sBefore = ClsMiscString.Left(ref sLine, iPos).PadRight(iMaxVariableNameLength + iMaxPrefixCharCount + 1);
                                                string sAfter = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 3);

                                                sLine = sBefore + " As " + sAfter;
                                            }
                                        }
                                    }

                                    objLine.sText_NoComment = sLine;
                                    lstLines[iIndex] = objLine;
                                }
                            }

                        }
                        break;
                    case enumVarDimType.eVarDim_OneSpace:
                        for(int iIndex = 0; iIndex < lstLines.Count;iIndex++)
                        {
                            strLine objLine = lstLines[iIndex];

                            if (countIF(objLine.lstLineType, enumLineType.eLineType_Dim) > 0)
                            {
                                string sLine = objLine.sText_NoComment;
                                //int iPos = 0;
                                bool bIsInQuotes = false;

                                for (int iPos = sLine.Length - 1; iPos > 0; iPos--)
                                {
                                    char cCurrChar = sLine[iPos];

                                    if (cCurrChar == '"')
                                    { bIsInQuotes = !bIsInQuotes; }

                                    if (!bIsInQuotes)
                                    {
                                        if (iPos < sLine.Length - 4)
                                        {
                                            if (sLine.Substring(iPos, 5).ToLower() == "  as ")
                                            {
                                                string sBefore = ClsMiscString.Left(ref sLine, iPos);
                                                string sAfter = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 4);

                                                sLine = sBefore + " As " + sAfter.TrimStart();
                                            }
                                        }
                                    }
                                }

                                objLine.sText_NoComment = sLine;
                            }

                            lstLines[iIndex] = objLine;
                        }
                        break;
                    case enumVarDimType.eVarDim_Nothing:
                        break;
                }

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

        public string currentFunctionName()
        {
            try
            {
                VBA.CodePane objCodePane = this.vbComponent.CodeModule.CodePane;

                int iStartLine = 0;
                int iStartColumn = 0;
                int iEndLine = 0;
                int iEndColumn = 0;

                objCodePane.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                string sCodeLine = objCodePane.CodeModule.get_Lines(iStartLine, 1);
                bool bIsFound = false;
                string sFunctionName = "";

                foreach (strLine objLine in this.lstLines.FindAll(x => x.iOriginalLineNo == iStartLine && x.sFunctionName.Trim() != ""))
                {
                    sFunctionName = objLine.sFunctionName;
                    bIsFound = true;
                }

                if (bIsFound)
                { return sFunctionName; }
                else
                { return string.Empty; }
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

                return string.Empty;
            }
        }

        private void functionsWithOnErrorGoto() 
        {
            try
            {
                List<string> lstFunctionsWithOnErrorGoto = new List<string>();
                List<strFunctions> lstChanges = new List<strFunctions>();

                foreach(strLine objLine in lstLines.FindAll(x => x.lstLineType.Contains(enumLineType.eLineType_OnError)))
                { lstFunctionsWithOnErrorGoto.Add(objLine.sFunctionName.Trim()); }

                foreach (strFunctions objFunction in lstFunctions)
                {
                    strFunctions objTemp = objFunction;
                    bool bValue;
                    if (lstFunctionsWithOnErrorGoto.Exists(x => x.Trim().ToLower() == objFunction.sName.Trim().ToLower() ))
                    { bValue = true; }
                    else
                    { bValue = false; }

                    if (objTemp.bHasErrorHandler != bValue)
                    {
                        objTemp.bHasErrorHandler = bValue;
                        lstChanges.Add(objTemp);
                    }
                }

                foreach (strFunctions objFunction in lstChanges)
                {
                    int iFnIndex = lstFunctions.FindIndex(x => x.sName.Trim().ToLower() == objFunction.sName.Trim().ToLower()  && x.sModuleName.Trim().ToLower() == objFunction.sModuleName.Trim().ToLower() && x.ePropertyType == objFunction.ePropertyType);

                    lstFunctions[iFnIndex] = objFunction;
                }

                lstChanges = null;
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

        private void markErrorHandlers(strFunctions objFunction)
        {
            try
            {
                bool bIsOk = true;

                List<strLine> lstFn = lstLines.FindAll(x => x.sFunctionName.Trim().ToLower() == objFunction.sName.Trim().ToLower() && x.ePropertyType == objFunction.ePropertyType);
                List<strLine> lstOnError = lstFn.FindAll(x => x.lstLineType.Contains(enumLineType.eLineType_OnError));

                foreach (strLine objOnError in lstOnError)
                {
                    bIsOk = true;
                    string sOnErrorSuffix = "";
                    int iLineNo = -1;
                    int iErrorHandlerIndex = -1;
                    int iEndSubIndex = -1;

                    string sLine = objOnError.sText_NoComment.Trim();
                    int iOnErrorPos = sLine.ToUpper().IndexOf("ON ERROR GOTO");

                    if (iOnErrorPos == -1)
                    { bIsOk = false; }

                    if (bIsOk == true)
                    {
                        sOnErrorSuffix = ClsMiscString.Right(ref sLine, sLine.Length - iOnErrorPos - 13).Trim();

                        if (sOnErrorSuffix.Contains(' '))
                        { bIsOk = false; }
                    }

                    if (bIsOk)
                    {
                        if (int.TryParse(sOnErrorSuffix, out iLineNo))
                        {
                            //goto line no
                            iErrorHandlerIndex = lstFn.FindIndex(x => x.sLineNo.Trim().Contains(sOnErrorSuffix.Trim()));
                        }
                        else
                        {
                            //goto label
                            iErrorHandlerIndex = lstFn.FindIndex(x => x.sLabel.ToLower().Trim().Contains(sOnErrorSuffix.ToLower().Trim()));
                        }

                        if (iErrorHandlerIndex != -1)
                        { iEndSubIndex = lstFn.FindIndex(iErrorHandlerIndex, x => x.lstLineType.Contains(enumLineType.eLineType_EndFunction)); }

                        if (iEndSubIndex == -1 || iErrorHandlerIndex == -1 || iEndSubIndex == -1)
                        { bIsOk = false; }
                    }

                    if (bIsOk)
                    {
                        //mark everything from the label or line number down to the end of the function
                        foreach (strLine objErrorHandler in lstFn.GetRange(iErrorHandlerIndex, iEndSubIndex - iErrorHandlerIndex))
                        {
                            int iTempIndex = lstLines.IndexOf(objErrorHandler);

                            strLine objTemp = lstLines[iTempIndex];
                            objTemp.lstLineType.Add(enumLineType.eLineType_ErrorHandler);
                            lstLines[iTempIndex] = objTemp;
                        }

                        //check upwards if the next line above is Exit Function
                        bool bCheckExitFn = true;
                        int iCheckPos = lstFn[iErrorHandlerIndex].iIndex - 1;
                        while (bCheckExitFn)
                        {
                            if (iCheckPos > 0)
                            {
                                strLine objExit = lstLines[iCheckPos];

                                if (objExit.sText_Orig.Trim() == "" || objExit.sText_NoComment.Trim() == "")
                                {
                                    strLine objTemp = lstLines[iCheckPos];
                                    objTemp.lstLineType.Add(enumLineType.eLineType_ErrorHandler);
                                    lstLines[iCheckPos] = objTemp;
                                    bCheckExitFn = false;
                                }
                                else if (objExit.lstLineType.Contains(enumLineType.eLineType_ExitFn))
                                {
                                    strLine objTemp = lstLines[iCheckPos];
                                    objTemp.lstLineType.Add(enumLineType.eLineType_ErrorHandler);
                                    lstLines[iCheckPos] = objTemp;
                                    bCheckExitFn = false;
                                }
                                else
                                { iCheckPos--; }
                            }
                            else
                            { bCheckExitFn = false; }
                        }
                    }
                }
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




        //public void addErrorHandlerToFunction(string sFunctionName, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        //{
        //    try
        //    {
        //        ClsSettings cSettings = new ClsSettings();
        //        bool bIsOk = true;
        //        string sErrorMesage;
        //        int iIndent = 0;
        //        List<strLine> lstFn = lstLines.FindAll(x => x.sFunctionName.Trim().ToLower() == sFunctionName.Trim().ToLower());

        //        strLine objStart = lstFn.Find(x => x.lstLineType.Contains(enumLineType.eLineType_FunctionName));
        //        strLine objEnd = lstFn.Find(x => x.lstLineType.Contains(enumLineType.eLineType_EndFunction));

        //        int iStart;
        //        int iEnd;

        //        if (!int.TryParse(objStart.sLineNo, out iStart))
        //        { 
        //            bIsOk = false;
        //            sErrorMesage = "Line No (Function declaration) is not recognised as a number";
        //        }

        //        if(!int.TryParse(objEnd.sLineNo, out iEnd))
        //        { 
        //            bIsOk = false;
        //            sErrorMesage = "Line No (End ...) is not recognised as a number";
        //        }

        //        if (lstFn.Exists(x => x.lstLineType.Contains(enumLineType.eLineType_OnError)))
        //        {
                    
        //        }

        //        //int i = 0;

        //        List<string> lstCall = new List<string>();
        //        List<string> lstBody = new List<string>();

        //        ClsInsertCode.addErrorHandlerCall(ref lstCall, ref cSettings, iIndent);
        //        ClsInsertCode.addLines(ref lstCall, ref vbComponent.CodeModule.CodePane, ref iIndent);  
        //        ClsInsertCode.addErrorHandlerBody(ref lstBody, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.);

        //        /*
        //         * 1) What if the function already has a on error goto.
        //         * a) There is only one on error goto and it is at the top of the function
        //         * b) There is only one on error goto and it is NOT at the top of the function
        //         * c) There are many on error goto's.
        //         * 
        //         * If (a) it's OK to replace 
        //         * if (b) or (c) the user has to replace and it can't be done by the in an automated way because it is part of the flow of the code.
        //         * 
        //         * 2) If (a) check the flow of the code doesn't do something silly for example more goto's inside the error handler
        //         * 
        //         * 
        //         * 
        //         */





        //        cSettings = null;
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
        private void getLineTypeSecondScan(ref List<strLine> lstLines)
        {
            try
            {
                List<strLine> lstLineCalls = new List<strLine>();

                foreach (strFunctions objFunction in lstFunctions)
                {
                    foreach (strLine objLine in lstLines.FindAll(x => x.sText_NoComment.ToUpper().Contains(objFunction.sName.ToUpper()) && !(x.lstLineType.Contains(enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(enumLineType.eLineType_Call))))
                    {
                        string sLine = objLine.sText_NoComment.ToUpper();
                        string sFunctionName = objFunction.sName.ToUpper();
                        bool bIsFound = false;

                        int iPos = sLine.IndexOf(sFunctionName);

                        while (iPos != -1)
                        {
                            bool bCheckBeforeOK = false;
                            bool bCheckAfterOK = false;
                            bool bIsInQuotes = false;

                            if (iPos > 0)
                            {
                                //if count of double quotes before is odd => in string
                                string sPrefix = sLine.Substring(0, iPos - 1);
                                if (sPrefix.Count(x => x == '"') % 2 != 0)
                                { bIsInQuotes = true; }
                                else
                                { bIsInQuotes = false; }
                            }

                            if (!bIsInQuotes)
                            {
                                if (iPos == 0)
                                { bCheckBeforeOK = true; }
                                else
                                {
                                    //Check char before function name
                                    char cBefore = sLine[iPos - 1];

                                    switch (cBefore)
                                    {
                                        case '.':
                                        case ' ':
                                        case '(':
                                        case ')':
                                        case '{':
                                        case '}':
                                        case '[':
                                        case ']':
                                            bCheckBeforeOK = true;
                                            break;
                                        default:
                                            bCheckBeforeOK = false;
                                            break;
                                    }
                                }

                                if (iPos == sLine.Length - sFunctionName.Length)
                                { bCheckAfterOK = true; }
                                else
                                {
                                    //Check char after function name
                                    char cAfter = sLine[iPos + sFunctionName.Length];

                                    switch (cAfter)
                                    {
                                        case '.':
                                        case ' ':
                                        case '(':
                                        case ')':
                                        case '{':
                                        case '}':
                                        case '[':
                                        case ']':
                                            bCheckAfterOK = true;
                                            break;
                                        default:
                                            bCheckAfterOK = false;
                                            break;
                                    }
                                }
                            }

                            if (bCheckBeforeOK == true && bCheckAfterOK == true)
                            { bIsFound = true; }

                            iPos = sLine.IndexOf(sFunctionName, iPos + 1);
                        }

                        if (bIsFound)
                        {
                            lstLineCalls.Add(objLine);

                            //objLine.lstLineType.Add(enumLineType.eLineType_Call);
                        }
                    }
                }

                foreach (strLine objLine in lstLineCalls)
                {
                    int iIndex = lstLines.FindIndex(x => x.sText_NoComment == objLine.sText_NoComment);

                    while (iIndex != -1)
                    {
                        if (!lstLines[iIndex].lstLineType.Contains(enumLineType.eLineType_Call))
                        {
                            strLine objLine2 = lstLines[iIndex];
                            objLine2.lstLineType.Add(enumLineType.eLineType_Call);
                            lstLines[iIndex] = objLine2;
                        }
                        
                        iIndex = lstLines.FindIndex(iIndex + 1, x => x.sText_NoComment == objLine.sText_NoComment);
                    }

                
                }

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
