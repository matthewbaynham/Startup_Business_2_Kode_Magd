using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using KodeMagd.Misc;

namespace KodeMagd.Reporter
{
    class ClsConfigReporterObjectModel : ClsConfigReporter
    {
        public struct strReportSpec
        {
            public bool bOnlyPublicFunctions;
            public List<string> lstModules;
            public bool bShowMemberVariablePublic;
            public bool bShowMemberVariablePrivate;
        }

        private List<string> lstReport = new List<string>();

        public string createObjectModelHtml(ref ClsCodeMapperWrk cCodeMapperWrk, strReportSpec objReportSpec)
        {
            try
            {
                //string sLine;
                List<strLine> lstHtml = new List<strLine>();

                addHeader(ref lstHtml);
                //addNextLine(ref lstHtml, "<!DCOTYPE html>\n");
                //addNextLine(ref lstHtml, "<html>\n");
                //addNextLine(ref lstHtml, "<head>\n");
                //addNextLine(ref lstHtml, "<title>" + ClsCodeEditorGUI.csCommandBarName + " - " + ClsMisc.ActiveWorkBook().Name + "</title>");
                //addNextLine(ref lstHtml, "<meta http-equiv='X-UA-Compatible' content='IE=9' >\n");
                //addNextLine(ref lstHtml, "</head>\n");
                //addNextLine(ref lstHtml, "<body>\n");
                addNextLine(ref lstHtml, "<canvas id='myCanvas' width=1000 height=1000 style='border:1px solid #d3d3d3;'>Your browser does not support HTML5 canvas.</canvas>");
                addNextLine(ref lstHtml, "<script type='text/javascript'>");
                addNextLine(ref lstHtml, "var c=document.getElementById('myCanvas');");
                addNextLine(ref lstHtml, "var ctx=c.getContext('2d');");


                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var canvasMaxSize = drawObjectModel(c, ctx);");
                addNextLine(ref lstHtml, "c.width = canvasMaxSize[0];");
                addNextLine(ref lstHtml, "c.height = canvasMaxSize[1];");
                addNextLine(ref lstHtml, "drawObjectModel(c, ctx);");
                addNextLine(ref lstHtml, "");

                addNextLine(ref lstHtml, "function drawObjectModel(c, ctx)\n");
                addNextLine(ref lstHtml, "{");
                addNextLine(ref lstHtml, "var colourVariable = \"#00FF00\";");
                addNextLine(ref lstHtml, "var colourFunction = \"#FFFF00\";");
                addNextLine(ref lstHtml, "var colourSub = \"#990099\";");
                addNextLine(ref lstHtml, "var colourProperty = \"#A80000\";");
                addNextLine(ref lstHtml, "var colourModule = \"#00003F\";");
                addNextLine(ref lstHtml, "var fontVariable = \"italic 14px Arial\";");
                addNextLine(ref lstHtml, "var fontFunction = \"16px Arial\";");
                addNextLine(ref lstHtml, "var fontModule = \"24px Arial\";");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iCanvasMaxWidth = 100;");
                addNextLine(ref lstHtml, "var iCanvasMaxHeight = 100;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iModTop = 30;");
                addNextLine(ref lstHtml, "var iModLeft = 0;"); //iModLeft = iModHorSpacing -> below
                addNextLine(ref lstHtml, "var iModHeight = 20;");
                addNextLine(ref lstHtml, "var iModWidth = 20;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iModMarginTop = 20;");
                addNextLine(ref lstHtml, "var iModMarginBottom = 20;");
                addNextLine(ref lstHtml, "var iModMarginLeft = 20;");
                addNextLine(ref lstHtml, "var iModMarginRight = 20;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iModHorSpacing = 50;");
                addNextLine(ref lstHtml, "var iModVertSpacing = 50;");
                addNextLine(ref lstHtml, "var iFunctionVertSpacing = 4;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iFunctionTop = 20;");
                addNextLine(ref lstHtml, "var iFunctionLeft = 20;");
                addNextLine(ref lstHtml, "var iFunctionHeight = 20;");
                addNextLine(ref lstHtml, "var iFunctionWidth = 20;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iFunctionMarginTop = 5;");
                addNextLine(ref lstHtml, "var iFunctionMarginBottom = 5;");
                addNextLine(ref lstHtml, "var iFunctionMarginLeft = 5;");
                addNextLine(ref lstHtml, "var iFunctionMarginRight = 5;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iTextHeight = 15;");
                addNextLine(ref lstHtml, "var iTextToModuleGap = 3;");
                addNextLine(ref lstHtml, "var iMaxHorLength = 1;");
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "var iCanvasNiceMaxToHave = window.innerWidth;"); //after the width has gone over this we move down to another row.
                addNextLine(ref lstHtml, "");
                addNextLine(ref lstHtml, "iModLeft = iModHorSpacing;");
                addNextLine(ref lstHtml, "");

                //lstReport.Add(sLine);

                foreach (ClsCodeMapper.strModuleDetails objModule in cCodeMapperWrk.getLstModuleDetails().FindAll(x => objReportSpec.lstModules.Contains(x.sName)))
                {
                    string sModuleNameText = nameFormatted(objModule);

                    addNextLine(ref lstHtml, "/*Module: " + sModuleNameText + "*/");
                    addNextLine(ref lstHtml, "iMaxHorLength = 1;");
                    addNextLine(ref lstHtml, "iFunctionTop = iModTop + iModMarginTop;");
                    addNextLine(ref lstHtml, "iFunctionLeft = iModLeft + iModMarginLeft;");
                    addNextLine(ref lstHtml, "ctx.font = fontModule;");
                    addNextLine(ref lstHtml, "ctx.fillText('" + sModuleNameText + "', iModLeft, iModTop - iTextToModuleGap);");
                    addNextLine(ref lstHtml, "iMaxHorLength = ctx.measureText('" + sModuleNameText + "').width;");
                    addNextLine(ref lstHtml, "");
                    //lstReport.Add(sLine);

                    Predicate<ClsCodeMapper.strVariables> predVariables;

                    if (objReportSpec.bShowMemberVariablePrivate && objReportSpec.bShowMemberVariablePublic)
                    { predVariables = a => string.IsNullOrEmpty(a.sFunctionName); }
                    else if (!objReportSpec.bShowMemberVariablePrivate && objReportSpec.bShowMemberVariablePublic)
                    { predVariables = a => string.IsNullOrEmpty(a.sFunctionName) && a.eScope == ClsCodeMapper.enumScopeVar.eScope_Global; }
                    else if (objReportSpec.bShowMemberVariablePrivate && !objReportSpec.bShowMemberVariablePublic)
                    { predVariables = a => string.IsNullOrEmpty(a.sFunctionName) && a.eScope == ClsCodeMapper.enumScopeVar.eScope_Module; }
                    else
                    { predVariables = a => string.IsNullOrEmpty(a.sFunctionName) && (a.eScope != ClsCodeMapper.enumScopeVar.eScope_Global || a.eScope != ClsCodeMapper.enumScopeVar.eScope_Module); }

                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapperWrk.getLstVariableDetails(objModule.sName).FindAll(predVariables).OrderBy(x => x.eScope).ThenBy(y => y.sName))
                    {
                        string sVariableNameText = "";

                        switch (objVariable.eScope)
                        {
                            case ClsCodeMapper.enumScopeVar.eScope_Function:
                                sVariableNameText += "(Function Variable): ";
                                break;
                            case ClsCodeMapper.enumScopeVar.eScope_Global:
                                sVariableNameText += "(Global): ";
                                break;
                            case ClsCodeMapper.enumScopeVar.eScope_Module:
                                sVariableNameText += "(local): ";
                                break;
                        }

                        sVariableNameText += objVariable.sName;

                        //Write out the text for the functions
                        addNextLine(ref lstHtml, "iFunctionHeight = iTextHeight + iFunctionMarginTop + iFunctionMarginBottom;");
                        addNextLine(ref lstHtml, "ctx.font = fontVariable;");
                        addNextLine(ref lstHtml, "ctx.fillText('" + sVariableNameText + "', iFunctionLeft + iFunctionMarginLeft, iFunctionTop + (iFunctionHeight - iFunctionMarginBottom));");
                        addNextLine(ref lstHtml, "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;");

                        addNextLine(ref lstHtml, "if (ctx.measureText('" + sVariableNameText + "').width > iMaxHorLength)");
                        addNextLine(ref lstHtml, "{ iMaxHorLength = ctx.measureText('" + sVariableNameText + "').width; }");
                        addNextLine(ref lstHtml, "");
                    }

                    Predicate<ClsCodeMapper.strFunctions> predFunction;
                    if (objReportSpec.bOnlyPublicFunctions)
                    { predFunction = x => x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Public; }
                    else
                    { predFunction = x => x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Public || x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Private || x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Friend; }

                    foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(objModule.sName).FindAll(predFunction).OrderBy(x => x.eFunctionType).ThenBy(y => y.sName).ThenBy(z => z.ePropertyType))
                    {
                        string sFunctionNameText = nameFormatted(objFunction);

                        //Write out the text for the functions
                        addNextLine(ref lstHtml, "iFunctionHeight = iTextHeight + iFunctionMarginTop + iFunctionMarginBottom;");
                        addNextLine(ref lstHtml, "ctx.font = fontFunction;");

                        if (objFunction.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        {
                            addNextLine(ref lstHtml, "ctx.fillText('" + sFunctionNameText + ")', iFunctionLeft + iFunctionMarginLeft, iFunctionTop + (iFunctionHeight - iFunctionMarginBottom));");
                        }
                        else
                        {
                            addNextLine(ref lstHtml, "ctx.fillText('" + sFunctionNameText + "', iFunctionLeft + iFunctionMarginLeft, iFunctionTop + (iFunctionHeight - iFunctionMarginBottom));");
                        }

                        addNextLine(ref lstHtml, "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;");

                        addNextLine(ref lstHtml, "if (ctx.measureText('" + sFunctionNameText + "').width > iMaxHorLength)");
                        addNextLine(ref lstHtml, "{ iMaxHorLength = ctx.measureText('" + sFunctionNameText + "').width; }");
                        addNextLine(ref lstHtml, "");
                    }

                    addNextLine(ref lstHtml, "iFunctionTop = iModTop + iModMarginTop;");
                    addNextLine(ref lstHtml, "iFunctionWidth = iMaxHorLength + iFunctionMarginLeft + iFunctionMarginRight;");
                    addNextLine(ref lstHtml, "iFunctionLeft = iModLeft + iModMarginLeft;");
                    addNextLine(ref lstHtml, "");

                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapperWrk.getLstVariableDetails(objModule.sName).FindAll(predVariables).OrderBy(x => x.eScope).ThenBy(y => y.sName))
                    {
                        //addNextLine(ref lstHtml, "ctx.strokeStyle = colourVariable;");
                        //addNextLine(ref lstHtml, "ctx.strokeRect(iFunctionLeft, iFunctionTop, iFunctionWidth, iFunctionHeight);");
                        addNextLine(ref lstHtml, "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;");
                        addNextLine(ref lstHtml, "");
                    }

                    foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(objModule.sName).FindAll(predFunction).OrderBy(x => x.eFunctionType).ThenBy(y => y.sName).ThenBy(z => z.ePropertyType))
                    {
                        //string sFunctionNameText = functionNameFormatted(objFunction);
                        //draw the rectangle for the functions
                        switch (objFunction.eFunctionType)
                        {
                            case ClsCodeMapper.enumFunctionType.eFnType_Function:
                                addNextLine(ref lstHtml, "ctx.strokeStyle = colourFunction;");
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                                addNextLine(ref lstHtml, "ctx.strokeStyle = colourSub;");
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Property:
                                addNextLine(ref lstHtml, "ctx.strokeStyle = colourProperty;");
                                break;

                        }

                        addNextLine(ref lstHtml, "ctx.strokeRect(iFunctionLeft, iFunctionTop, iFunctionWidth, iFunctionHeight);");
                        addNextLine(ref lstHtml, "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;");
                        addNextLine(ref lstHtml, "");

                        //lstReport.Add(sLine);
                    }

                    //draw the rectangle for the functions

                    addNextLine(ref lstHtml, "");
                    addNextLine(ref lstHtml, "iModWidth = iFunctionWidth + iModMarginLeft + iModMarginRight;");
                    addNextLine(ref lstHtml, "iModHeight = iFunctionTop + iModMarginTop + iModMarginBottom - iModTop - iFunctionHeight;");
                    addNextLine(ref lstHtml, "ctx.strokeStyle = colourModule;");
                    addNextLine(ref lstHtml, "ctx.strokeRect(iModLeft, iModTop, iModWidth, iModHeight);");
                    //addNextLine(ref lstHtml, "ctx.stroke();");
                    //lstReport.Add(sLine);




                    addNextLine(ref lstHtml, "iModLeft = iModLeft + iFunctionWidth + iModMarginLeft + iModMarginRight + iModHorSpacing;");
                    addNextLine(ref lstHtml, "");
                    addNextLine(ref lstHtml, "if (iModLeft > iCanvasMaxWidth)");
                    addNextLine(ref lstHtml, "{ iCanvasMaxWidth = iModLeft; }");
                    addNextLine(ref lstHtml, "");
                    addNextLine(ref lstHtml, "if (iModTop + iModHeight > iCanvasMaxHeight)");
                    addNextLine(ref lstHtml, "{ iCanvasMaxHeight = iModTop + iModHeight; }");
                    addNextLine(ref lstHtml, "");
                    addNextLine(ref lstHtml, "if (iModLeft > iCanvasNiceMaxToHave)");
                    addNextLine(ref lstHtml, "{");
                    addNextLine(ref lstHtml, "iModLeft = iModHorSpacing;");
                    addNextLine(ref lstHtml, "iModTop = iCanvasMaxHeight + iModVertSpacing;");
                    addNextLine(ref lstHtml, "}");
                    addNextLine(ref lstHtml, "");
                    //lstReport.Add(sLine);
                }

                addNextLine(ref lstHtml, "ctx.stroke();");
                addNextLine(ref lstHtml, "return [iCanvasMaxWidth, iCanvasMaxHeight];");
                addNextLine(ref lstHtml, "}");

                addNextLine(ref lstHtml, "</script>");
                //addNextLine(ref lstHtml, "</html>\n");
                //addNextLine(ref lstHtml, "</body>\n");
                //lstReport.Add(sLine);
                addFooter(ref lstHtml);

                string sDoc = "";

                foreach (strLine objLine in lstHtml)
                { sDoc += objLine.sLine + "\n"; }

                return sDoc;
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

        private string nameFormatted(ClsCodeMapper.strFunctions objFunction)
        {
            try
            {
                string sResult = "";
                switch (objFunction.eFunctionType)
                {
                    case ClsCodeMapper.enumFunctionType.eFnType_Function:
                        sResult = "FN: " + objFunction.sName;
                        break;
                    case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                        sResult = "Sub: " + objFunction.sName;
                        break;
                    case ClsCodeMapper.enumFunctionType.eFnType_Property:
                        sResult = "Prop: " + objFunction.sName;
                        switch (objFunction.ePropertyType)
                        {
                            case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                sResult += "(Get)";
                                break;
                            case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                sResult += "(Let)";
                                break;
                            case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                sResult += "(Set)";
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
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

        private string nameFormatted(ClsCodeMapper.strModuleDetails objModule)
        {
            try
            {
                string sResult = "";
                switch (objModule.eType)
                {
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                        sResult = "ActiveX: " + objModule.sName;
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                        sResult = "Class: " + objModule.sName;
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                        sResult = "Doc: " + objModule.sName;
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                        sResult = "Form: " + objModule.sName;
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                        sResult = "Module: " + objModule.sName;
                        break;
                    default:
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
    }
}
