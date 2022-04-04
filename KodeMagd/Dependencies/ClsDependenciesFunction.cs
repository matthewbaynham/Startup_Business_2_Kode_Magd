using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KodeMagd.Misc;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Reporter;

namespace KodeMagd.Dependencies
{
    class ClsDependenciesFunction
    {
        public static void generateReport(ref ClsCodeMapperWrk cCodeMapperWrk, string sModule, string sFunctionName, ClsCodeMapper.enumFunctionType eFuntionType, ClsCodeMapper.enumFunctionPropertyType ePropertyType, IWin32Window win)
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";

                ClsConfigReporter cConfigReporter = new ClsConfigReporter();
                ClsCodeMapper.strFunctions objFunction = cCodeMapperWrk.getFunction(sModule, sFunctionName, eFuntionType, ePropertyType);

                if (string.IsNullOrEmpty(objFunction.sName))
                { 
                    bIsOk = false;
                    sMessage = "Can't find function.";
                }

                if (bIsOk)
                {
                    List<ClsCodeMapper.strLine> lstLines = searchFunctionCalls(ref cCodeMapperWrk, ref objFunction);

                    buildHtml(ref cConfigReporter, lstLines, objFunction);
                    
                    displayHtmlSummary(ref cConfigReporter, win);
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

        //public static List<ClsCodeMapper.strLine> searchFunctionCalls(ref ClsCodeMapperWrk cCodeMapperWrk, string sModule, string sFunctionName, ClsCodeMapper.enumFunctionType eFuntionType, ClsCodeMapper.enumFunctionPropertyType ePropertyType)
        public static List<ClsCodeMapper.strLine> searchFunctionCalls(ref ClsCodeMapperWrk cCodeMapperWrk, ref ClsCodeMapper.strFunctions objFunction)
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";
                List<ClsCodeMapper.strLine> lstLinesResult = new List<ClsCodeMapper.strLine>();;

                string sFunctionName = objFunction.sName;

                if (bIsOk)
                {
                    if (objFunction.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Private)
                    {
                        foreach (ClsCodeMapper.strLine objLine in cCodeMapperWrk.getLines(objFunction.sModuleName).FindAll(x => x.sText_NoComment.ToUpper().Contains(sFunctionName.ToUpper())))
                        {
                            if (ClsMiscString.containsTextNoPrefixOrSuffix(objLine.sText_NoComment.ToUpper(), sFunctionName.ToUpper().Trim()))
                            {
                                if (objFunction.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                                {
                                    if (ClsMiscString.checkPropertyCallTypeOK(objLine.sText_NoComment.ToUpper(), sFunctionName.ToUpper().Trim(), objFunction.ePropertyType))
                                    { lstLinesResult.Add(objLine); }
                                }
                                else
                                { lstLinesResult.Add(objLine); }
                            }
                        }
                    }
                    else
                    {
                        foreach (ClsCodeMapper.strModuleDetails objModuleDetails in cCodeMapperWrk.getLstModuleDetails())
                        {
                            foreach (ClsCodeMapper.strLine objLine in cCodeMapperWrk.getLines(objModuleDetails.sName).FindAll(x => x.sText_NoComment.ToUpper().Contains(sFunctionName.ToUpper())))
                            {
                                if (ClsMiscString.containsTextNoPrefixOrSuffix(objLine.sText_NoComment.ToUpper(), sFunctionName.ToUpper().Trim()))
                                {
                                    if (objFunction.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                                    {
                                        if (ClsMiscString.checkPropertyCallTypeOK(objLine.sText_NoComment.ToUpper(), sFunctionName.ToUpper().Trim(), objFunction.ePropertyType))
                                        { lstLinesResult.Add(objLine); }
                                    }
                                    else
                                    { lstLinesResult.Add(objLine); }
                                }
                            }
                        }
                    }
                }

                return lstLinesResult;
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

        public static void buildHtml(ref ClsConfigReporter cConfigReporter, List<ClsCodeMapper.strLine> lstLines, ClsCodeMapper.strFunctions objFunction)
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
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 3 }, "Variable Dependencies");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                switch(objFunction.eFunctionType)
                {
                    case ClsCodeMapper.enumFunctionType.eFnType_Function:
                        objCell.sText = "Function Name";
                        break;
                    case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                        objCell.sText = "Sub Routine Name";
                        break;
                    case ClsCodeMapper.enumFunctionType.eFnType_Property:
                        objCell.sText = "Property Name";
                        break;
                    default:
                        objCell.sText = "Name";
                        break;
                }
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = objFunction.sName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (objFunction.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Property Type";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    switch(objFunction.ePropertyType)
                    {
                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                            objCell.sText = "Get";
                            break;
                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                            objCell.sText = "Let";
                            break;
                        case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                            objCell.sText = "Set";
                            break;
                        default:
                            objCell.sText = "Unknown Property Type";
                            break;
                    }
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = "Module Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = objFunction.sModuleName;
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
                            switch (objLine.ePropertyType)
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
