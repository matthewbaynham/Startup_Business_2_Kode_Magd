using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_ErrorHandler : ClsInsertCode
    {
        private enum eDetails 
        {
            eCommentOutErrorHandler,
            eAddNewErrorHandler,
            eNoPreviousErrorHandler,
            eIssueNeedsReporting
        }

        private struct strLog
        {
            public string sModule;
            public string sFunction;
            public ClsCodeMapper.enumFunctionType eFunctionType;
            public ClsCodeMapper.enumFunctionPropertyType ePropertyType;
            public List<eDetails> lstDetails;
            public string sIssue;
        }

        List<strLog> lstLog = new List<strLog>();

        public void replaceErrorRoutines_IgnoreOldAddNewRegardless(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, List<string> lstFunctionNames, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        {
            try
            {
                ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(sModuleName);
                if (!cCodeMapper.isRead)
                { cCodeMapper.readCode(); }
                cCodeMapper.fixIndex();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> LstCodeTop = new List<string>();

                List<ClsCodeMapper.strLine> lstToBeCommentedOut = new List<ClsCodeMapper.strLine>();

                Predicate<ClsCodeMapper.strLine> predErrorHandler;
                Predicate<ClsCodeMapper.strLine> predFunctionNameAndEnd;

                //set search criteria for which lines we want to look at
                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                {
                    predErrorHandler = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError);
                    predFunctionNameAndEnd = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction);
                }
                else
                {
                    predErrorHandler = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError));
                    predFunctionNameAndEnd = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction));
                }

                if (cSettings.UserTips == true)
                {
                    //loop through the lines collect any that need to be commented out
                    List<ClsCodeMapper.strLine> lstErrorLine = cCodeMapper.lines.FindAll(predErrorHandler);
                    
                    foreach (ClsCodeMapper.strLine objLine in lstErrorLine.OrderByDescending(x => x.iIndex))
                    {
                        if (lstErrorLine.FindAll(y => y.iIndex == (objLine.iIndex + 1)).Count == 0)
                        {
                            //no line below
                            ClsCodeMapper.strLine objNewLine = objLine;

                            int iNewIndex = objLine.iIndex + 1;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "'The previous error handler NOT been removed, it is strongly recommended that you remove it.";
                            objNewLine.sText_NoComment = "";
                            objNewLine.sText_Orig = "";
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_Comment);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);
                        }

                        ClsCodeMapper.strLine objTemp = objLine;

                        if (objTemp.sLabel.Trim() == "")
                        { objTemp.sText_NoComment = objTemp.sText_NoComment + "'Old Code recommend removing"; }
                        else
                        {
                            objTemp.sText_NoComment = objTemp.sLabel + ": " + objTemp.sText_NoComment + "'Old Code recommend removing";
                            objTemp.sLabel = "";
                        }

                        cCodeMapper.updateLine(objTemp);
                        
                        strLog objLog = new strLog();

                        objLog.sModule = sModuleName;
                        objLog.sFunction = objLine.sFunctionName;
                        objLog.eFunctionType = objLine.eFunctionType;
                        objLog.ePropertyType = objLine.ePropertyType;
                        objLog.lstDetails = new List<eDetails>();
                        objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                        lstLog.Add(objLog);

                        //lstToBeCommentedOut.Add(objTemp);
                    }

                    ////loop though the list of lines that need to be commented out
                    //foreach (ClsCodeMapper.strLine objLine in lstToBeCommentedOut.OrderByDescending(x => x.iIndex))
                    //{ cCodeMapper.updateLine(objLine); }
                }

                //although it shouldn't be broken but just to be sure.
                cCodeMapper.fixIndex();

                //find beginning and end of functions, note with properties there will be multiple functions.
                foreach (ClsCodeMapper.strLine objLine in cCodeMapper.lines.FindAll(predFunctionNameAndEnd).OrderByDescending(x => x.iIndex))
                {

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                    {
                        List<string> lstNewCall = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCallNoCheck(ref lstNewCall, ref cSettings, iIndent);

                        foreach (string sNewLine in lstNewCall)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_OnError);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewCall, iNewIndex);
                    }

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                    {
                        List<string> lstNewHandler = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex - 1;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerBodyNoCheck(ref lstNewHandler, ref cSettings, iIndent, objLine.eFunctionType);

                        foreach (string sNewLine in lstNewHandler)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_ErrorHandler);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewHandler, iNewIndex);
                    }
                }
                //Add calls and add handlers

                cCodeMapper.ImplementChanges();

                cSettings = null;
                //cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                LstCodeTop = null;
                lstToBeCommentedOut = null;
                predErrorHandler = null;
                predFunctionNameAndEnd = null;
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

        public void replaceErrorRoutines_OneOrManyThenReplace(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, List<string> lstFunctionNames, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        {
            try
            {
                ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(sModuleName);
                if (!cCodeMapper.isRead)
                { cCodeMapper.readCode(); }
                cCodeMapper.fixIndex();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();

                List<ClsCodeMapper.strLine> lstToBeCommentedOut = new List<ClsCodeMapper.strLine>();

                Predicate<ClsCodeMapper.strLine> predErrorHandler;
                Predicate<ClsCodeMapper.strLine> predFunctionNameAndEnd;

                //set search criteria for which lines we want to look at
                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                { 
                    predErrorHandler = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError); 
                    predFunctionNameAndEnd= x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction); 
                }
                else
                { 
                    predErrorHandler = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError));
                    predFunctionNameAndEnd = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction));
                }

                //loop through the lines collect any that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in cCodeMapper.lines.FindAll(predErrorHandler))
                {
                    ClsCodeMapper.strLine objTemp = objLine;

                    if (objTemp.sLabel.Trim() == "")
                    { objTemp.sText_NoComment = "'" + objTemp.sText_NoComment; }
                    else
                    {
                        objTemp.sText_NoComment = "'" + objTemp.sLabel + ": " + objTemp.sText_NoComment;
                        objTemp.sLabel = "";
                    }

                    lstToBeCommentedOut.Add(objTemp);
                }

                //loop though the list of lines that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in lstToBeCommentedOut.OrderByDescending(x => x.iIndex))
                { 
                    cCodeMapper.updateLine(objLine);

                    strLog objLog = new strLog();

                    objLog.sModule = sModuleName;
                    objLog.sFunction = objLine.sFunctionName;
                    objLog.eFunctionType = objLine.eFunctionType;
                    objLog.ePropertyType = objLine.ePropertyType;
                    objLog.lstDetails = new List<eDetails>();
                    objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                    lstLog.Add(objLog);
                }
                
                //although it shouldn't be broken but just to be sure.
                cCodeMapper.fixIndex();

                //find beginning and end of functions, note with properties there will be multiple functions.
                foreach (ClsCodeMapper.strLine objLine in cCodeMapper.lines.FindAll(predFunctionNameAndEnd).OrderByDescending(x => x.iIndex))
                {
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                    {
                        List<string> lstNewCall = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCallNoCheck(ref lstNewCall, ref cSettings, iIndent);

                        foreach (string sNewLine in lstNewCall)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_OnError);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewCall, iNewIndex);
                    }

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                    {
                        List<string> lstNewHandler = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex - 1;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerBodyNoCheck(ref lstNewHandler, ref cSettings, iIndent, objLine.eFunctionType);

                        foreach (string sNewLine in lstNewHandler)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_ErrorHandler);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewHandler, iNewIndex);
                    }
                }
                //Add calls and add handlers

                cCodeMapper.ImplementChanges();

                cSettings = null;
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeTop = null;
                lstToBeCommentedOut = null;
                predErrorHandler = null;
                predFunctionNameAndEnd = null;
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

        public void replaceErrorRoutines_DoNothingIfErrorHandlerExists(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, List<string> lstFunctionNames, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        {
            try
            {
                ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(sModuleName);
                if (!cCodeMapper.isRead)
                { cCodeMapper.readCode(); }
                cCodeMapper.fixIndex();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> LstCodeTop = new List<string>();

                List<ClsCodeMapper.strLine> lstLinesOfErrorCode = new List<ClsCodeMapper.strLine>();

                Predicate<ClsCodeMapper.strLine> predErrorHandler;
                Predicate<ClsCodeMapper.strLine> predFunctionNameAndEnd;

                //set search criteria for which lines we want to look at
                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                {
                    predErrorHandler = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError);
                    //predFunctionNameAndEnd = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction);
                }
                else
                {
                    predErrorHandler = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError));
                    //predFunctionNameAndEnd = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction));
                }

                /*
                 * Get list of all functions that should have error handlers added
                 * Get list of all functions which already have error handlers
                 * 
                 * Note: if function has call but not handler or if function has function but not call then, ignore but report.
                 * 
                 * 
                 * 
                 */
                List<strLog> lstFnWithErrorCode = new List<strLog>();
                List<strLog> lstFnWithErrorGotoNotHandler = new List<strLog>();
                List<strLog> lstFnWithErrorHandlerNotGoto = new List<strLog>();

                //loop through the lines collect any that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in cCodeMapper.lines.FindAll(predErrorHandler))
                {
                    ClsCodeMapper.strLine objTemp = objLine;

                    if (objTemp.sLabel.Trim() == "")
                    { objTemp.sText_NoComment = "'" + objTemp.sText_NoComment; }
                    else
                    {
                        objTemp.sText_NoComment = "'" + objTemp.sLabel + ": " + objTemp.sText_NoComment;
                        objTemp.sLabel = "";
                    }

                    lstLinesOfErrorCode.Add(objTemp);

                    strLog objFn = new strLog();

                    objFn.sModule = cCodeMapper.ModuleDetails.sName;
                    objFn.sFunction = objLine.sFunctionName;
                    objFn.eFunctionType = objLine.eFunctionType;
                    objFn.ePropertyType = objLine.ePropertyType;
                    objFn.lstDetails = new List<eDetails>();
                    objFn.sIssue = "";

                    lstFnWithErrorCode.Add(objFn);

                    objFn.lstDetails.Add(eDetails.eIssueNeedsReporting);

                    //add any function with on error
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError))
                    {
                        objFn.sIssue = "Contains On Error but no Error Handler.";
                        lstFnWithErrorGotoNotHandler.Add(objFn);
                    
                    }

                    //add any function with error hander
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler))
                    {
                        objFn.sIssue = "Contains Error Handler but no On Error.";
                        lstFnWithErrorHandlerNotGoto.Add(objFn);
                    }
                }

                lstFnWithErrorCode = lstFnWithErrorCode.Distinct().ToList();
                lstFnWithErrorGotoNotHandler = lstFnWithErrorGotoNotHandler.Distinct().ToList();
                lstFnWithErrorHandlerNotGoto = lstFnWithErrorHandlerNotGoto.Distinct().ToList();;

                //remove any function with on error 
                for (int iCounter = lstFnWithErrorGotoNotHandler.Count - 1; iCounter >= 0; iCounter--)
                {
                    if (iCounter < lstFnWithErrorGotoNotHandler.Count)
                    {
                        strLog objItem = lstFnWithErrorGotoNotHandler[iCounter];

                        if (lstFnWithErrorHandlerNotGoto.Exists(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                && x.ePropertyType == objItem.ePropertyType
                                                                && x.eFunctionType == objItem.eFunctionType))
                        {
                            lstFnWithErrorHandlerNotGoto.RemoveAll(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                    && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                    && x.ePropertyType == objItem.ePropertyType
                                                                    && x.eFunctionType == objItem.eFunctionType);

                            //lstFnWithErrorGotoNotHandler.RemoveAt(iCounter);
                            lstFnWithErrorGotoNotHandler.RemoveAll(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                    && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                    && x.ePropertyType == objItem.ePropertyType
                                                                    && x.eFunctionType == objItem.eFunctionType);
                        }
                    }
                }

                //remove any funcion with handler
                for (int iCounter = lstFnWithErrorGotoNotHandler.Count - 1; iCounter >= 0; iCounter--)
                {
                    if (iCounter < lstFnWithErrorGotoNotHandler.Count)
                    {
                        strLog objItem = lstFnWithErrorGotoNotHandler[iCounter];

                        if (lstFnWithErrorHandlerNotGoto.Exists(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                && x.ePropertyType == objItem.ePropertyType
                                                                && x.eFunctionType == objItem.eFunctionType))
                        {
                            lstFnWithErrorHandlerNotGoto.RemoveAll(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                    && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                    && x.ePropertyType == objItem.ePropertyType
                                                                    && x.eFunctionType == objItem.eFunctionType);
                            //lstFnWithErrorGotoNotHandler.RemoveAt(iCounter);
                            lstFnWithErrorGotoNotHandler.RemoveAll(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                                                                    && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                                                                    && x.ePropertyType == objItem.ePropertyType
                                                                    && x.eFunctionType == objItem.eFunctionType);
                        }
                    }
                }

                //lstFnWithErrorGotoNotHandler & lstFnWithErrorHandlerNotGoto need to be sent off to have warnings appear in the HTML output
                //


                foreach (strLog objLog in lstFnWithErrorGotoNotHandler)
                { lstLog.Add(objLog); }

                foreach (strLog objLog in lstFnWithErrorHandlerNotGoto)
                { lstLog.Add(objLog); }


                for (int iCounter = lstFnWithErrorCode.Count - 1; iCounter >= 0; iCounter--)
                {
                    strLog objItem = lstFnWithErrorCode[iCounter];

                    if (lstFnWithErrorHandlerNotGoto.Exists(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                        && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                        && x.ePropertyType == objItem.ePropertyType
                        && x.eFunctionType == objItem.eFunctionType))
                    { lstFnWithErrorCode.RemoveAt(iCounter); }

                    if (lstFnWithErrorGotoNotHandler.Exists(x => x.sModule.Trim().ToLower() == objItem.sModule.Trim().ToLower()
                        && x.sFunction.Trim().ToLower() == objItem.sFunction.Trim().ToLower()
                        && x.ePropertyType == objItem.ePropertyType
                        && x.eFunctionType == objItem.eFunctionType))
                    { lstFnWithErrorCode.RemoveAt(iCounter); }
                }

                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                {
                    //if contains beginning or end function and is not list of functions that have on error or list of functions that have error handler or list of functions that have both on error and handler
                    predFunctionNameAndEnd = x => (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) 
                                    || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                                    && !(lstFnWithErrorCode.Exists(a => a.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                    && a.eFunctionType == x.eFunctionType
                                                                    && a.ePropertyType == x.ePropertyType)
                                        || lstFnWithErrorGotoNotHandler.Exists(b => b.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                        && b.eFunctionType == x.eFunctionType
                                                                        && b.ePropertyType == x.ePropertyType)
                                        || lstFnWithErrorHandlerNotGoto.Exists(c => c.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                        && c.eFunctionType == x.eFunctionType
                                                                        && c.ePropertyType == x.ePropertyType));
                }
                else
                {
                    //if contains beginning or end function and is not list of functions that have on error or list of functions that have error handler or list of functions that have both on error and handler
                    //and in list of functions that we are interested in.
                    predFunctionNameAndEnd = x => (lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) 
                                    && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) 
                                            || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction)))
                                    && !(lstFnWithErrorCode.Exists(a => a.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                    && a.eFunctionType == x.eFunctionType
                                                                    && a.ePropertyType == x.ePropertyType)
                                        || lstFnWithErrorGotoNotHandler.Exists(b => b.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                        && b.eFunctionType == x.eFunctionType
                                                                        && b.ePropertyType == x.ePropertyType)
                                        || lstFnWithErrorHandlerNotGoto.Exists(c => c.sFunction.Trim().ToLower() == x.sFunctionName.Trim().ToLower()
                                                                        && c.eFunctionType == x.eFunctionType
                                                                        && c.ePropertyType == x.ePropertyType));
                }

                //find beginning and end of functions, note with properties there will be multiple functions.
                foreach (ClsCodeMapper.strLine objLine in cCodeMapper.lines.FindAll(predFunctionNameAndEnd).OrderByDescending(x => x.iIndex))
                {
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                    {

                        List<string> lstNewCall = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCallNoCheck(ref lstNewCall, ref cSettings, iIndent);

                        foreach (string sNewLine in lstNewCall)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_OnError);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewCall, iNewIndex);
                    }

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                    {
                        List<string> lstNewHandler = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex - 1;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerBodyNoCheck(ref lstNewHandler, ref cSettings, iIndent, objLine.eFunctionType);

                        foreach (string sNewLine in lstNewHandler)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_ErrorHandler);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewHandler, iNewIndex);
                    }
                }
                //Add calls and add handlers

                cCodeMapper.ImplementChanges();

                cSettings = null;
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                LstCodeTop = null;
                lstLinesOfErrorCode = null;
                predErrorHandler = null;
                predFunctionNameAndEnd = null;

                lstFnWithErrorCode = null;
                lstFnWithErrorGotoNotHandler = null;
                lstFnWithErrorHandlerNotGoto = null;
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

        public void replaceErrorRoutines_OneAnywhereThenReplace(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, List<string> lstFunctionNames, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        {
            try
            {
                ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(sModuleName);
                if (!cCodeMapper.isRead)
                { cCodeMapper.readCode(); }
                cCodeMapper.fixIndex();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> LstCodeTop = new List<string>();

                List<ClsCodeMapper.strLine> lstToBeCommentedOut = new List<ClsCodeMapper.strLine>();

                Predicate<ClsCodeMapper.strLine> predErrorHandler;
                Predicate<ClsCodeMapper.strLine> predFunctionNameAndEnd;

                //set search criteria for which lines we want to look at
                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                { 
                    predErrorHandler = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError); 
                    predFunctionNameAndEnd= x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction); 
                }
                else
                { 
                    predErrorHandler = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError));
                    predFunctionNameAndEnd = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction));
                }

                List<ClsCodeMapper.strLine> lstErrorLines = cCodeMapper.lines.FindAll(predErrorHandler);

                //make sure that we only have a list of all the function name and not just "<All>" and we only need to function where there is 
                List<ClsCodeMapper.strFunctionIdentity> lstFnNames = new List<ClsCodeMapper.strFunctionIdentity>();

                foreach (ClsCodeMapper.strLine objLine in lstErrorLines)
                {
                    ClsCodeMapper.strFunctionIdentity objFn = new ClsCodeMapper.strFunctionIdentity();

                    objFn.eFunctionType = objLine.eFunctionType;
                    objFn.ePropertyType = objLine.ePropertyType;
                    objFn.sModuleName = "";
                    objFn.sName = objLine.sFunctionName;

                    lstFnNames.Add(objFn);
                }

                lstFnNames = lstFnNames.Distinct().ToList();


                //delete the lines from tempery table that are in functions where there is more than one "on error command"
                for (int iCounter = lstFnNames.Count - 1; iCounter >= 0; iCounter--) 
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = lstFnNames[iCounter];

                    bool bDelete = false;

                    if (lstErrorLines.FindAll(x => x.sFunctionName == objFnTemp.sName
                        && x.ePropertyType == objFnTemp.ePropertyType
                        && x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError)).Count == 1)
                    { bDelete = false; }
                    else
                    { bDelete = true; }

                    if (bDelete == true)
                    { lstFnNames.RemoveAt(iCounter); }
                }

                for (int iCounter = lstErrorLines.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = new ClsCodeMapper.strFunctionIdentity();

                    objFnTemp.eFunctionType = lstErrorLines[iCounter].eFunctionType;
                    objFnTemp.ePropertyType = lstErrorLines[iCounter].ePropertyType;
                    objFnTemp.sModuleName = "";
                    objFnTemp.sName = lstErrorLines[iCounter].sFunctionName;

                    if (!lstFnNames.Contains(objFnTemp))
                    { lstErrorLines.RemoveAt(iCounter); }
                }

                //loop through the lines comment any that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in lstErrorLines)
                {
                    ClsCodeMapper.strLine objTemp = objLine;

                    if (objTemp.sLabel.Trim() == "")
                    { objTemp.sText_NoComment = "'" + objTemp.sText_NoComment; }
                    else
                    {
                        objTemp.sText_NoComment = "'" + objTemp.sLabel + ": " + objTemp.sText_NoComment;
                        objTemp.sLabel = "";
                    }

                    lstToBeCommentedOut.Add(objTemp);
                }

                //loop though the list of lines that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in lstToBeCommentedOut.OrderByDescending(x => x.iIndex))
                { 
                    cCodeMapper.updateLine(objLine);

                    strLog objLog = new strLog();

                    objLog.sModule = sModuleName;
                    objLog.sFunction = objLine.sFunctionName;
                    objLog.eFunctionType = objLine.eFunctionType;
                    objLog.ePropertyType = objLine.ePropertyType;
                    objLog.lstDetails = new List<eDetails>();
                    objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                    lstLog.Add(objLog);
                }
                
                //although it shouldn't be broken but just to be sure.
                cCodeMapper.fixIndex();

                List<ClsCodeMapper.strLine> lstFnStartEnd = cCodeMapper.lines.FindAll(predFunctionNameAndEnd);

                for (int iCounter = lstFnStartEnd.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = new ClsCodeMapper.strFunctionIdentity();

                    objFnTemp.eFunctionType = lstFnStartEnd[iCounter].eFunctionType;
                    objFnTemp.ePropertyType = lstFnStartEnd[iCounter].ePropertyType;
                    objFnTemp.sModuleName = "";
                    objFnTemp.sName = lstFnStartEnd[iCounter].sFunctionName;

                    if (!lstFnNames.Contains(objFnTemp))
                    { lstFnStartEnd.RemoveAt(iCounter); }
                }

                //find beginning and end of functions, note with properties there will be multiple functions.
                foreach (ClsCodeMapper.strLine objLine in lstFnStartEnd.OrderByDescending(x => x.iIndex))
                {
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                    {
                        List<string> lstNewCall = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCallNoCheck(ref lstNewCall, ref cSettings, iIndent);

                        foreach (string sNewLine in lstNewCall)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_OnError);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        //addCode(ref lstNewCall, iNewIndex);
                        lstNewCall = null;
                    }

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                    {
                        List<string> lstNewHandler = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex - 1;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerBodyNoCheck(ref lstNewHandler, ref cSettings, iIndent, objLine.eFunctionType);

                        foreach (string sNewLine in lstNewHandler)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_ErrorHandler);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eAddNewErrorHandler);

                            lstLog.Add(objLog);
                        }
                        lstNewHandler = null;
                    }
                }
                //Add calls and add handlers

                cCodeMapper.ImplementChanges();

                cSettings = null;
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                LstCodeTop = null;
                lstToBeCommentedOut = null;

                predErrorHandler = null;
                predFunctionNameAndEnd = null;

                lstErrorLines = null;
                lstFnNames = null;
                lstFnStartEnd = null;
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

        public void replaceErrorRoutines_OneAtTopThenReplace(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, List<string> lstFunctionNames, FrmInsertCode_ErrorHandler.enumReplaceErrorHandlerActions eActions)
        {
            try
            {
                ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(sModuleName);
                if (!cCodeMapper.isRead)
                { cCodeMapper.readCode(); }
                cCodeMapper.fixIndex();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> LstCodeTop = new List<string>();

                List<ClsCodeMapper.strLine> lstToBeCommentedOut = new List<ClsCodeMapper.strLine>();

                Predicate<ClsCodeMapper.strLine> predErrorHandler;
                Predicate<ClsCodeMapper.strLine> predFunctionNameAndEnd;

                //set search criteria for which lines we want to look at
                if (lstFunctionNames.Contains(ClsDefaults.textAll))
                {
                    predErrorHandler = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError);
                    predFunctionNameAndEnd = x => x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction);
                }
                else
                {
                    predErrorHandler = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError));
                    predFunctionNameAndEnd = x => lstFunctionNames.Exists(y => y.ToLower().Trim() == x.sFunctionName.Trim().ToLower()) && (x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName) || x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction));
                }

                List<ClsCodeMapper.strLine> lstErrorLines = cCodeMapper.lines.FindAll(predErrorHandler);

                //make sure that we only have a list of all the function name and not just "<All>" and we only need to function where there is 
                List<ClsCodeMapper.strFunctionIdentity> lstFnNames = new List<ClsCodeMapper.strFunctionIdentity>();

                foreach (ClsCodeMapper.strLine objLine in lstErrorLines)
                {
                    ClsCodeMapper.strFunctionIdentity objFn = new ClsCodeMapper.strFunctionIdentity();

                    objFn.eFunctionType = objLine.eFunctionType;
                    objFn.ePropertyType = objLine.ePropertyType;
                    objFn.sModuleName = "";
                    objFn.sName = objLine.sFunctionName;

                    lstFnNames.Add(objFn);
                }

                lstFnNames = lstFnNames.Distinct().ToList();

                //delete the lines from tempery table that are in functions where there is more than one "on error command"
                for (int iCounter = lstFnNames.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = lstFnNames[iCounter];

                    bool bDelete = false;

                    if (lstErrorLines.FindAll(x => x.sFunctionName == objFnTemp.sName
                        && x.ePropertyType == objFnTemp.ePropertyType
                        && x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError)).Count == 1)
                    { bDelete = false; }
                    else
                    { bDelete = true; }

                    if (bDelete == true)
                    { lstFnNames.RemoveAt(iCounter); }
                }

                //remove any error line in our list that is not in the functions we are interested in
                for (int iCounter = lstErrorLines.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = new ClsCodeMapper.strFunctionIdentity();

                    objFnTemp.eFunctionType = lstErrorLines[iCounter].eFunctionType;
                    objFnTemp.ePropertyType = lstErrorLines[iCounter].ePropertyType;
                    objFnTemp.sModuleName = "";
                    objFnTemp.sName = lstErrorLines[iCounter].sFunctionName;

                    if (!lstFnNames.Contains(objFnTemp))
                    { lstErrorLines.RemoveAt(iCounter); }
                }

                lstFnNames = lstFnNames.Distinct().ToList();

                //delete error lines in our list where the on error is not at the top
                for (int iCounter = lstErrorLines.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strLine objErrLine = lstErrorLines[iCounter];

                    if (objErrLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError))
                    {
                        List<ClsCodeMapper.strLine> lstTopLines = cCodeMapper.lines.FindAll(x => x.sFunctionName.Trim().ToLower() == objErrLine.sFunctionName.Trim().ToLower()
                            && x.ePropertyType == objErrLine.ePropertyType
                            && x.iIndex < objErrLine.iIndex
                            && !x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName)
                            && !x.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Empty)
                            && x.lstLineType.Count > 0);
                        
                        if (lstTopLines.Exists(x => x.lstLineType.Exists(z => z == ClsCodeMapper.enumLineType.eLineType_AssignValue
                            || z == ClsCodeMapper.enumLineType.eLineType_AssignValue
                            || z == ClsCodeMapper.enumLineType.eLineType_BeginLoop
                            || z == ClsCodeMapper.enumLineType.eLineType_Call
                            || z == ClsCodeMapper.enumLineType.eLineType_DeInitialise
                            || z == ClsCodeMapper.enumLineType.eLineType_Dim
                            || z == ClsCodeMapper.enumLineType.eLineType_Else
                            || z == ClsCodeMapper.enumLineType.eLineType_EndFunction
                            || z == ClsCodeMapper.enumLineType.eLineType_EndIf
                            || z == ClsCodeMapper.enumLineType.eLineType_EndLoop
                            || z == ClsCodeMapper.enumLineType.eLineType_EndWith
                            || z == ClsCodeMapper.enumLineType.eLineType_ExitFn
                            || z == ClsCodeMapper.enumLineType.eLineType_ExitIf
                            || z == ClsCodeMapper.enumLineType.eLineType_ExitLoop
                            || z == ClsCodeMapper.enumLineType.eLineType_Goto
                            || z == ClsCodeMapper.enumLineType.eLineType_If
                            || z == ClsCodeMapper.enumLineType.eLineType_Initialise
                            || z == ClsCodeMapper.enumLineType.eLineType_Options
                            || z == ClsCodeMapper.enumLineType.eLineType_With
                            || z == ClsCodeMapper.enumLineType.eLineType_Input
                            || z == ClsCodeMapper.enumLineType.eLineType_Output)
                            ))
                        {
                            lstErrorLines.RemoveAt(iCounter);
                            ClsCodeMapper.strFunctionIdentity objFnTemp = new ClsCodeMapper.strFunctionIdentity();

                            objFnTemp.eFunctionType = lstErrorLines[iCounter].eFunctionType;
                            objFnTemp.ePropertyType = lstErrorLines[iCounter].ePropertyType;
                            objFnTemp.sModuleName = "";
                            objFnTemp.sName = lstErrorLines[iCounter].sFunctionName;

                            while (lstFnNames.Contains(objFnTemp))
                            { lstFnNames.Remove(objFnTemp); }
                        }

                        lstTopLines = null;
                    }
                }

                //remove any error line in our list that is not in the functions we are interested in
                for (int iCounter = lstErrorLines.Count - 1; iCounter >= 0; iCounter--)
                {
                    if (!lstFnNames.Exists(x => x.sName.Trim().ToLower() == lstErrorLines[iCounter].sFunctionName.Trim().ToLower()
                        && x.eFunctionType == lstErrorLines[iCounter].eFunctionType
                        && x.ePropertyType == lstErrorLines[iCounter].ePropertyType))
                    { lstErrorLines.RemoveAt(iCounter); }
                }

                //loop through the lines comment any that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in lstErrorLines)
                {
                    ClsCodeMapper.strLine objTemp = objLine;

                    if (objTemp.sLabel.Trim() == "")
                    { objTemp.sText_NoComment = "'" + objTemp.sText_NoComment; }
                    else
                    {
                        objTemp.sText_NoComment = "'" + objTemp.sLabel + ": " + objTemp.sText_NoComment;
                        objTemp.sLabel = "";
                    }

                    lstToBeCommentedOut.Add(objTemp);
                }

                //loop though the list of lines that need to be commented out
                foreach (ClsCodeMapper.strLine objLine in lstToBeCommentedOut.OrderByDescending(x => x.iIndex))
                { 
                    cCodeMapper.updateLine(objLine);

                    strLog objLog = new strLog();

                    objLog.sModule = sModuleName;
                    objLog.sFunction = objLine.sFunctionName;
                    objLog.eFunctionType = objLine.eFunctionType;
                    objLog.ePropertyType = objLine.ePropertyType;
                    objLog.lstDetails = new List<eDetails>();
                    objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                    lstLog.Add(objLog);
                }

                //although it shouldn't be broken but just to be sure.
                cCodeMapper.fixIndex();

                List<ClsCodeMapper.strLine> lstFnStartEnd = cCodeMapper.lines.FindAll(predFunctionNameAndEnd);

                for (int iCounter = lstFnStartEnd.Count - 1; iCounter >= 0; iCounter--)
                {
                    ClsCodeMapper.strFunctionIdentity objFnTemp = new ClsCodeMapper.strFunctionIdentity();

                    objFnTemp.eFunctionType = lstFnStartEnd[iCounter].eFunctionType;
                    objFnTemp.ePropertyType = lstFnStartEnd[iCounter].ePropertyType;
                    objFnTemp.sModuleName = "";
                    objFnTemp.sName = lstFnStartEnd[iCounter].sFunctionName;

                    if (!lstFnNames.Contains(objFnTemp))
                    { lstFnStartEnd.RemoveAt(iCounter); }
                }

                //find beginning and end of functions, note with properties there will be multiple functions.
                foreach (ClsCodeMapper.strLine objLine in lstFnStartEnd.OrderByDescending(x => x.iIndex))
                {
                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                    {
                        List<string> lstNewCall = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCallNoCheck(ref lstNewCall, ref cSettings, iIndent);

                        foreach (string sNewLine in lstNewCall)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_OnError);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                            lstLog.Add(objLog);
                        }
                        
                        lstNewCall = null;
                    }

                    if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                    {
                        List<string> lstNewHandler = new List<string>();
                        int iIndent = cCodeMapper.indentLevel(objLine.iIndex);
                        int iNewIndex = objLine.iIndex - 1;

                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerBodyNoCheck(ref lstNewHandler, ref cSettings, iIndent, objLine.eFunctionType);

                        foreach (string sNewLine in lstNewHandler)
                        {
                            iNewIndex++;
                            ClsCodeMapper.strLine objNewLine = objLine;

                            objNewLine.iIndex = iNewIndex;
                            objNewLine.sText_Comment = "";
                            objNewLine.sText_NoComment = sNewLine;
                            objNewLine.sText_Orig = sNewLine;
                            objNewLine.lstLineType.Clear();
                            objNewLine.lstLineType.Add(ClsCodeMapper.enumLineType.eLineType_ErrorHandler);

                            cCodeMapper.addLine(iNewIndex, ref objNewLine);

                            strLog objLog = new strLog();

                            objLog.sModule = sModuleName;
                            objLog.sFunction = objLine.sFunctionName;
                            objLog.eFunctionType = objLine.eFunctionType;
                            objLog.ePropertyType = objLine.ePropertyType;
                            objLog.lstDetails = new List<eDetails>();
                            objLog.lstDetails.Add(eDetails.eCommentOutErrorHandler);

                            lstLog.Add(objLog);
                        }
                        lstNewHandler = null;
                    }
                }
                //Add calls and add handlers

                cCodeMapper.ImplementChanges();

                cSettings = null;
                cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                LstCodeTop = null;
                lstToBeCommentedOut = null;
                lstFnStartEnd = null;
                predErrorHandler = null;
                predFunctionNameAndEnd = null;
                lstErrorLines = null;
                lstFnNames = null;
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

        public void makeUniqueLstLog()
        {
            try
            {
                int iIndex = 0;

                for (int iCounter = lstLog.Count - 1; iCounter >= 0; iCounter--)
                {
                    strLog objLog = lstLog[iCounter];
                    iIndex = 0;

                    iIndex = lstLog.FindIndex(x => x.eFunctionType == objLog.eFunctionType
                                                    && x.ePropertyType == objLog.ePropertyType
                                                    && x.sFunction == objLog.sFunction
                                                    && x.sModule == objLog.sModule);

                    if (iIndex != iCounter)
                    {
                        if (iIndex != -1)
                        {
                            strLog objLogTemp = lstLog[iIndex];

                            foreach (eDetails objDetail in lstLog[iCounter].lstDetails)
                            { objLogTemp.lstDetails.Add(objDetail); }

                            objLogTemp.lstDetails = objLogTemp.lstDetails.Distinct().ToList();

                            lstLog[iIndex] = objLogTemp;

                            lstLog.RemoveAt(iCounter);
                        }
                    }
                }

                for (int iCounter = lstLog.Count - 1; iCounter >= 0; iCounter--)
                {
                    strLog objLog = lstLog[iCounter];

                    objLog.lstDetails = objLog.lstDetails.Distinct().ToList();

                    lstLog[iCounter] = objLog;
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

        public void generateHtmlFile(ref ClsConfigReporter cConfigReporter, ref List<FrmInsertCode_ErrorHandler.strFnMod> lstEffectedFn)
        {
            try
            {
                cConfigReporter = new ClsConfigReporter();

                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId = 0;
                int iRowId = 0;
                
                string sPreviousModule = "";
                string sPreviousFunction = "";
                ClsCodeMapper.enumFunctionType ePreviousFnType = ClsCodeMapper.enumFunctionType.eFnType_None;
                ClsCodeMapper.enumFunctionPropertyType ePreviousFnPropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;
                bool bIncludedPropertyColumn = false;

                /*******************************************************
                 *   Successful error handler and "on error" inserts   *
                 *******************************************************/
                foreach (strLog objLog in lstLog.FindAll(y => !y.lstDetails.Contains(eDetails.eIssueNeedsReporting)).OrderBy(x => x.sModule).ThenBy(x => x.sFunction).ThenBy(x => x.eFunctionType).ThenBy(x => x.ePropertyType))
                {
                    if (objLog.sModule.Trim().ToLower() != sPreviousModule.Trim().ToLower())
                    {
                        sPreviousModule = objLog.sModule;

                        //if contains properties
                        //add title
                        /***************
                         *   A table   *
                         ***************/
                        if (lstLog.Exists(x => x.sModule.Trim().ToLower() == objLog.sModule.Trim().ToLower()
                            && (x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Get
                                || x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Set
                                || x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Let)))
                        {
                            bIncludedPropertyColumn = true;
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 1, 10 }, "Module: " + objLog.sModule);
                        }
                        else
                        {
                            bIncludedPropertyColumn = false;
                            //add title
                            /***************
                             *   A table   *
                             ***************/
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 1, 4 }, "Module: " + objLog.sModule);
                        }

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Routine Name";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Routine Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        if (bIncludedPropertyColumn == true) 
                        {
                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Property Type";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        }

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Details";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objLog.sFunction;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    switch (objLog.eFunctionType)
                    {
                        case ClsCodeMapper.enumFunctionType.eFnType_Function:
                            objCell.sText = "Function";
                            break;
                        case ClsCodeMapper.enumFunctionType.eFnType_Property:
                            objCell.sText = "Property";
                            break;
                        case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                            objCell.sText = "Sub";
                            break;
                        default:
                            objCell.sText = "Unknown Type";
                            break;
                    }
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    if (bIncludedPropertyColumn == true)
                    {
                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (objLog.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        {
                            switch (objLog.ePropertyType)
                            {
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                    objCell.sText = "Set";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                    objCell.sText = "Let";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                    objCell.sText = "Get";
                                    break;
                                default:
                                    objCell.sText = "Unknown";
                                    break;
                            }
                        }
                        else
                        { objCell.sText = ""; }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }

                    string sDetails = "";
                    string sHiddenDetails = "";

                    if (objLog.lstDetails.Contains(eDetails.eAddNewErrorHandler))
                    {
                        if (sDetails != "")
                        { sDetails += "& "; }
                        sDetails += "Add ";

                        if (sHiddenDetails != "")
                        { sHiddenDetails += "\n"; }
                        sHiddenDetails += "Add New Error Handler";
                    }

                    if (objLog.lstDetails.Contains(eDetails.eCommentOutErrorHandler))
                    {
                        if (sDetails != "")
                        { sDetails += "& "; }
                        sDetails += "Comment ";
                        
                        if (sHiddenDetails != "")
                        { sHiddenDetails += "\n"; }
                        sHiddenDetails += "Comment Out Error Handler";
                    }

                    if (objLog.lstDetails.Contains(eDetails.eNoPreviousErrorHandler))
                    {
                        if (sDetails != "")
                        { sDetails += "& "; }
                        sDetails += "No ";
                        
                        if (sHiddenDetails != "")
                        { sHiddenDetails += "\n"; }
                        sHiddenDetails += "No Previous Error Handler";
                    }

                    sDetails = sDetails.Trim();
                    sHiddenDetails = sHiddenDetails.Trim();

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = sDetails;
                    objCell.sHiddenText = sHiddenDetails;
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }

                /**************
                 *   Issues   *
                 **************/
                sPreviousModule = "";

                foreach (strLog objLog in lstLog.FindAll(y => y.lstDetails.Contains(eDetails.eIssueNeedsReporting)).OrderBy(x => x.sModule).ThenBy(x => x.sFunction).ThenBy(x => x.eFunctionType).ThenBy(x => x.ePropertyType))
                {
                    if (objLog.sModule.Trim().ToLower() != sPreviousModule.Trim().ToLower())
                    {
                        sPreviousModule = objLog.sModule;

                        //if contains properties
                        //add title
                        /***************
                         *   A table   *
                         ***************/
                        if (lstLog.Exists(x => x.sModule.Trim().ToLower() == objLog.sModule.Trim().ToLower()
                            && (x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Get
                                || x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Set
                                || x.ePropertyType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Let)))
                        {
                            bIncludedPropertyColumn = true;
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 1, 10 }, "Issues in Module: " + objLog.sModule);
                        }
                        else
                        {
                            bIncludedPropertyColumn = false;
                            //add title
                            /***************
                             *   A table   *
                             ***************/
                            cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 1, 4 }, "Issues in Module: " + objLog.sModule);
                        }

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Routine Name";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Routine Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        if (bIncludedPropertyColumn == true)
                        {
                            objCell.iOrder = 0;
                            objCell.bPropHtml = true;
                            objCell.sText = "Property Type";
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                            objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                            objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        }

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Details";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objLog.sFunction;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    switch (objLog.eFunctionType)
                    {
                        case ClsCodeMapper.enumFunctionType.eFnType_Function:
                            objCell.sText = "Function";
                            break;
                        case ClsCodeMapper.enumFunctionType.eFnType_Property:
                            objCell.sText = "Property";
                            break;
                        case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                            objCell.sText = "Sub";
                            break;
                        default:
                            objCell.sText = "Unknown Type";
                            break;
                    }
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    if (bIncludedPropertyColumn == true)
                    {
                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (objLog.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        {
                            switch (objLog.ePropertyType)
                            {
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                    objCell.sText = "Set";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                    objCell.sText = "Let";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                    objCell.sText = "Get";
                                    break;
                                default:
                                    objCell.sText = "Unknown";
                                    break;
                            }
                        }
                        else
                        { objCell.sText = ""; }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objLog.sIssue;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                    objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Red);

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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
    }
}
