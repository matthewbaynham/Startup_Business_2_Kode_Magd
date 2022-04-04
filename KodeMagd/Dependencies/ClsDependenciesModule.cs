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
    class ClsDependenciesModule
    {
        public static void buildHtml(ref ClsConfigReporter cConfigReporter, ref List<ClsCodeMapperWrk.strLinesInModule> lstModule, ClsCodeMapper.strModuleDetails objModuleDetails)
        {
            try
            {
                //ClsConfigReporter cConfigReporter = new ClsConfigReporter();

                List<ClsConfigReporter.strLine> lstHtml = new List<ClsConfigReporter.strLine>();
                //cConfigReporter.addHeader(ref lstHtml);
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId = 0;
                int iRowId = 0;

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, 2, "Module Dependencies");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = "Description";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                switch (objModuleDetails.eType)
                {
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                        objCell.sText = "Active-X Module";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                        objCell.sText = "Class Module";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                        objCell.sText = "Document Module";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                        objCell.sText = "Module";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                        objCell.sText = "Code Module";
                        break;
                    default:
                        objCell.sText = "Unknown Module type";
                        break;
                }
                objCell.sHiddenText = "Which Module is being checked for Dependencies.";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = objModuleDetails.sName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                foreach (ClsCodeMapperWrk.strLinesInModule objLinesInModule in lstModule.OrderBy(x => x.objModuleDetails.eType).ThenBy(x => x.objModuleDetails.sName))
                {
                    string sTableTitle = objLinesInModule.objModuleDetails.sName + " - ";

                    switch (objLinesInModule.objModuleDetails.eType)
                    {
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                            sTableTitle += "Active X Designer";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                            sTableTitle += "Class";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                            sTableTitle += "Document";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                            sTableTitle += "Form";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                            sTableTitle += "Code Module";
                            break;
                        default:
                            sTableTitle += "Unknown Module Type";
                            break;
                    }

                    if (objLinesInModule.lstLines.Count == 0)
                    {
                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 1 }, sTableTitle);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                        objCell.iOrder = 0;
                        objCell.sText = "No Depenencies found";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Italic);
                        objCell.lstFormatDetails.Add(ClsConfigReporter.enumFormatDetails.eFmt_Gray);

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }
                    else
                    {
                        cConfigReporter.TableAddNew(out iTableId, new List<int> { 2, 1, 3 }, sTableTitle);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                        objCell.iOrder = 0;
                        objCell.sText = "Function Name";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.sText = "Line Number";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.sText = "VBA";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        foreach (ClsCodeMapper.strLine objLine in objLinesInModule.lstLines)
                        {
                            //Add Row
                            cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                            objCell.iOrder = 0;
                            if (string.IsNullOrEmpty(objLine.sFunctionName))
                            {
                                objCell.sText = "<None>";
                                objCell.sHiddenText = "Code is outside of any functions.";
                            }
                            else
                            {
                                objCell.sText = objLine.sFunctionName;
                                objCell.sHiddenText = "";
                            }
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.sText = objLine.iOriginalLineNo.ToString();
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                            objCell.iOrder = 0;
                            objCell.sText = objLine.sText_Orig;
                            objCell.sHiddenText = "";
                            objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                            cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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
