using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.Dependencies
{
    class ClsDepenenciesVariables
    {
        public static void reportLocalVariableDependencies(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, string sVariableName, IWin32Window win)
        {
            try
            {
                /****************************  
                 *   For Global variables   *
                 ****************************/

                ClsConfigReporter cConfigReporter = new ClsConfigReporter();

                List<ClsCodeMapper.strLine> lstLineSource = cCodeMapperWrk.getLines(sModuleName);

                //ClsCodeMapper.strVariables objVariable = cCodeMapperWrk.getVariable(sModuleName, sFunctionName, ePropType, sVariableName);
                ClsCodeMapper.strVariables objVariable = cCodeMapperWrk.getVariable(sModuleName, sVariableName);

                List<ClsCodeMapper.strLine> lstLineContainsVariable = findLines(ref lstLineSource, objVariable);
                buildHtml(ref cConfigReporter, lstLineContainsVariable, objVariable);
                displayHtmlSummary(ref cConfigReporter, win);
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

        public static void reportLocalVariableDependencies(ref ClsCodeMapperWrk cCodeMapperWrk, string sModuleName, string sFunctionName, ClsCodeMapper.enumFunctionPropertyType ePropType, string sVariableName, IWin32Window win)
        {
            try
            {
                /***************************
                 *   For local variables   *
                 ***************************/

                ClsConfigReporter cConfigReporter = new ClsConfigReporter();
                List<string> lstFn = new List<string>();

                lstFn.Add(sFunctionName);

                List<ClsCodeMapper.strLine> lstLineSource = cCodeMapperWrk.getLines(sModuleName, lstFn);

                ClsCodeMapper.strVariables objVariable = cCodeMapperWrk.getVariable(sModuleName, sFunctionName, ePropType, sVariableName);

                List<ClsCodeMapper.strLine> lstLineContainsVariable = findLines(ref lstLineSource, objVariable);
                buildHtml(ref cConfigReporter, lstLineContainsVariable, objVariable);
                displayHtmlSummary(ref cConfigReporter, win);
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

        public static List<ClsCodeMapper.strLine> findLines(ref List<ClsCodeMapper.strLine> lstLinesSource, ClsCodeMapper.strVariables objVariable) 
        {
            try
            {
                List<ClsCodeMapper.strLine> lstResults = new List<ClsCodeMapper.strLine>();

                switch (objVariable.eType) 
                {
                    case ClsDataTypes.vbVarType.vbArray:
                    case ClsDataTypes.vbVarType.vbBoolean:
                    case ClsDataTypes.vbVarType.vbByte:
                    case ClsDataTypes.vbVarType.vbCurrency:
                    case ClsDataTypes.vbVarType.vbDate:
                    case ClsDataTypes.vbVarType.vbDecimal:
                    case ClsDataTypes.vbVarType.vbDouble:
                    case ClsDataTypes.vbVarType.vbEmpty:
                    case ClsDataTypes.vbVarType.vbError:
                    case ClsDataTypes.vbVarType.vbInteger:
                    case ClsDataTypes.vbVarType.vbLong:
                    case ClsDataTypes.vbVarType.vbLongLong:
                    case ClsDataTypes.vbVarType.vbSingle:
                    case ClsDataTypes.vbVarType.vbString:
                        foreach (ClsCodeMapper.strLine objLine in lstLinesSource)
                        {
                            if (ClsMiscString.containsVariable(objLine.sText_NoComment, objVariable.sName))
                            { lstResults.Add(objLine); }
                        }
                        break;
                    case ClsDataTypes.vbVarType.vbNull:
                    case ClsDataTypes.vbVarType.vbUnknown:
                    case ClsDataTypes.vbVarType.vbDataObject:
                    case ClsDataTypes.vbVarType.vbVariant:
                    case ClsDataTypes.vbVarType.vbObject:
                    case ClsDataTypes.vbVarType.vbUserDefinedType:
                        //More complicated because a with can be problematic    
                        bool bInWith = false;
                        List<string> lstWith = new List<string>();

                        foreach (ClsCodeMapper.strLine objLine in lstLinesSource)
                        {
                            if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_With))
                            {
                                if (objLine.sText_NoComment.ToUpper().Trim().StartsWith("WITH"))
                                {
                                    string sNextWith = objLine.sText_NoComment.Trim().Substring(4).Trim();
                                    lstWith.Add(sNextWith.Trim());

                                    if (sNextWith.ToUpper().Trim() == objVariable.sName.ToUpper().Trim())
                                    { 
                                        bInWith = true;
                                        lstResults.Add(objLine);
                                    }
                                    else
                                    { bInWith = false; }
                                }
                            }
                            else if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndWith))
                            {
                                if (lstWith.Count() > 0)
                                {
                                    if (lstWith.Count() > 0)
                                    {
                                        if (lstWith.Last().ToUpper().Trim() == objVariable.sName.ToUpper().Trim())
                                        { lstResults.Add(objLine); }

                                        lstWith.RemoveAt(lstWith.Count() - 1);

                                        if (lstWith.Count() > 0)
                                        {
                                            if (lstWith.Last().ToUpper().Trim() == objVariable.sName.ToUpper().Trim())
                                            { bInWith = true; }
                                            else
                                            { bInWith = false; }
                                        }
                                        else
                                        { bInWith = false; }
                                    }
                                }
                            }
                            else
                            {
                                if (ClsMiscString.containsVariable(objLine.sText_NoComment, objVariable.sName))
                                { lstResults.Add(objLine); }

                                if (bInWith)
                                {
                                    //search for a full stop without anything preseeding it
                                    if (objLine.sText_NoComment.Trim().StartsWith(".") || ClsMiscString.containsWildcard(objLine.sText_NoComment, "[ (][.][A-Za-z]"))
                                    { lstResults.Add(objLine); }
                                }
                            }
                        }
                        break;
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

        public static void buildHtml(ref ClsConfigReporter cConfigReporter, List<ClsCodeMapper.strLine> lstLines, ClsCodeMapper.strVariables objVariable)
        {
            try
            {
                List<ClsConfigReporter.strLine> lstHtml = new List<ClsConfigReporter.strLine>();
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                string sPreviousModuleName = "";
                string sPreviousFunctionName = "";
                ClsCodeMapper.enumFunctionType ePreviousFunctionType = ClsCodeMapper.enumFunctionType.eFnType_None;
                ClsCodeMapper.enumFunctionPropertyType ePreviousPropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;

                int iTableId = 0;
                int iRowId = 0;

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1 }, "Variable Dependencies");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = "Data Type";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = objVariable.sName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = cDataTypes.getName(objVariable.eType);
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);


                foreach (ClsCodeMapper.strLine objLine in lstLines.OrderBy(x => x.iOriginalLineNo))
                {
                    if (sPreviousModuleName == "" 
                        || sPreviousModuleName != objLine.sModuleName || sPreviousFunctionName != objLine.sFunctionName
                        || ePreviousFunctionType != objLine.eFunctionType || ePreviousPropType != objLine.ePropertyType)
                    {
                        /***************
                         *   A table   *
                         ***************/
                        string sTableTitle = objLine.sModuleName + " - ";

                        if (objLine.sFunctionName == "")
                        { sTableTitle += "<Not in Function>"; }
                        else
                        { sTableTitle += objLine.sFunctionName; }

                        if (objLine.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        {
                            switch(objLine.ePropertyType)
                            {
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                    sTableTitle += " (Get)";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                    sTableTitle += " (Set)";
                                    break;
                                case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                    sTableTitle += " (Let)";
                                    break;

                            }
                        }

                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 15 }, sTableTitle);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                        objCell.iOrder = 0;
                        objCell.sText = "Row Number";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.sText = "Code";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    }
                    
                    sPreviousModuleName = objLine.sModuleName;
                    sPreviousFunctionName = objLine.sFunctionName;
                    ePreviousFunctionType = objLine.eFunctionType;
                    ePreviousPropType = objLine.ePropertyType;

                    //string sTableTitle = objLine.iOriginalLineNo.ToString() + " - ";

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.sText = objLine.iOriginalLineNo.ToString();
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = objLine.sText_NoComment.Trim();
                    objCell.sHiddenText = objLine.sText_Comment.Trim();
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }

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

        public static void displayHtmlSummary(ref ClsConfigReporter cConfigReporter, IWin32Window win)
        {
            try
            {
                string sHtml = cConfigReporter.getHtml();

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Dependencies");

                frm.ShowDialog(win);

                frm = null;
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
