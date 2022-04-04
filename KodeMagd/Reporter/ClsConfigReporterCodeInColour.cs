using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.WorkbookAnalysis;
using KodeMagd.Misc;

namespace KodeMagd.Reporter
{
    public class ClsConfigReporterCodeInColour : ClsConfigReporter
    {
        public void setCssColour(strCss objCss)
        {
            try
            {
                lstCssExtra.Add(objCss); 
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

        public string createColouredHtmlText(ref List<ClsCodeMapper.strLine> lstLines, ref List<strCss> lstCss)
        {
            try
            {
                int iIndexAssignVariables = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_AssignVariables.Trim().ToUpper());
                int iIndexComments = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_Comments.Trim().ToUpper());
                int iIndexDeclareFunctions = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_DeclareFunctions.Trim().ToUpper());
                int iIndexDeclareVariables = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_DeclareVariables.Trim().ToUpper());
                int iIndexErrors = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_Errors.Trim().ToUpper());
                int iIndexIfStatements = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_IfStatements.Trim().ToUpper());
                int iIndexLoops = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_Loops.Trim().ToUpper());
                int iIndexWith = lstCss.FindIndex(x => x.sName.Trim().ToUpper() == ClsCodeInColour.sCssName_With.Trim().ToUpper());

                string sHtml = "";

                foreach (ClsCodeMapper.strLine objLine in lstLines.OrderBy(x => x.iOrder))
                {
                    int iColourIndex = -1;
                    string sColour = "";
                    string sPaddingLeft = "";

                    if (objLine.sText_Orig.Length == objLine.sText_Orig.TrimStart().Length)
                    { sPaddingLeft = ""; }
                    else
                    {
                        int iPadding = objLine.sText_Orig.Length - objLine.sText_Orig.TrimStart().Length;

                        sPaddingLeft = " style=\"padding-left:" + (10 * iPadding).ToString() + "px;\" ";
                    }

                    if (objLine.sText_NoComment.Trim() != "")
                    {
                        foreach (ClsCodeMapper.enumLineType eLineType in objLine.lstLineType)
                        {
                            switch (ClsCodeInColour.convert(eLineType))
                            {
                                case ClsCodeInColour.enumCodeColourType.eAssigningValues:
                                    iColourIndex = iIndexAssignVariables;
                                    break;
                                //case ClsCodeInColour.enumCodeColourType.eComments:
                                //Don't do it for comments because a line could be a comment line as well as another type of line.
                                //    iColourIndex = iIndexComments;
                                //    break;
                                case ClsCodeInColour.enumCodeColourType.eDeclaringVariables:
                                    iColourIndex = iIndexDeclareVariables;
                                    break;
                                case ClsCodeInColour.enumCodeColourType.eErrors:
                                    iColourIndex = iIndexErrors;
                                    break;
                                case ClsCodeInColour.enumCodeColourType.eFunctions:
                                    iColourIndex = iIndexDeclareFunctions;
                                    break;
                                case ClsCodeInColour.enumCodeColourType.eIfStatements:
                                    iColourIndex = iIndexIfStatements;
                                    break;
                                case ClsCodeInColour.enumCodeColourType.eLoops:
                                    iColourIndex = iIndexLoops;
                                    break;
                                case ClsCodeInColour.enumCodeColourType.eWith:
                                    iColourIndex = iIndexWith;
                                    break;
                            }
                        }

                        if (iColourIndex == -1)
                        {
                            if (sPaddingLeft.Trim() == "")
                            { sHtml += ClsConfigReporterCodeInColour.prepHtmlText(objLine.sText_NoComment.Trim()); }
                            else
                            { sHtml += "<font " + sPaddingLeft + ">" + ClsConfigReporterCodeInColour.prepHtmlText(objLine.sText_NoComment.Trim()) + "</font>"; }
                        }
                        else
                        {
                            sHtml += "<font Class=\"" + ClsConfigReporterCodeInColour.prepHtmlText(lstCss[iColourIndex].sName) + "\" " + sPaddingLeft + ">";

                            //foreach (strCssStyle objCssStyle in lstCss[iColourIndex].lstCssStyles)
                            //{
                            //    sHtml += objCssStyle.sName + "=" + objCssStyle.sValue;
                            //}
                            //sHtml += ">";

                            sHtml += ClsConfigReporterCodeInColour.prepHtmlText(objLine.sText_NoComment.Trim()) + "</font>";
                        }
                    }

                    if (objLine.sText_Comment.Trim() != "")
                    {
                        string sComment = ClsConfigReporterCodeInColour.prepHtmlText(objLine.sText_Comment.Trim());

                        if (!sComment.StartsWith("'") && sComment.Trim() != "")
                        { sComment = "'" + sComment; }

                        if (iIndexComments == -1)
                        { sHtml += sComment; }
                        else
                        {
                            if (objLine.sText_NoComment.Trim() == "")
                            { sHtml += "<font Class=\"" + ClsConfigReporterCodeInColour.prepHtmlText(lstCss[iIndexComments].sName) + "\" " + sPaddingLeft + ">"; }
                            else
                            { sHtml += "<font Class=\"" + ClsConfigReporterCodeInColour.prepHtmlText(lstCss[iIndexComments].sName) + "\">"; }
                            //sHtml += "<font ";

                            //foreach (strCssStyle objCssStyle in lstCss[iIndexComments].lstCssStyles)
                            //{ sHtml += objCssStyle.sName + "=" + objCssStyle.sValue; }
                            //sHtml += ">";

                            sHtml += sComment + "</font>";
                        }
                    }

                    sHtml += "<br>\n";
                }

                return sHtml;
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
/*
        public string createObjectModelHtml(ref ClsCodeMapperWrk cCodeMapperWrk, List<ClsCodeMapper.strFunctions> lstFunctions)
        {
            try
            {
                string sResult = "";

                foreach (ClsCodeMapper.strFunctions objFn in lstFunctions)//Note: only the module name and function name are set.
                {
                    List<ClsCodeMapper.strLine> lstLists = cCodeMapperWrk.getLines(objFn.sModuleName, new List<string> { objFn.sName });

                    if (lstLists.Count > 0)
                    {
                        string sFunctionType = "";
                        string sPropertyType = "";

                        switch (lstLists[0].eFunctionType)
                        {
                            case ClsCodeMapper.enumFunctionType.eFnType_Function:
                                sFunctionType = "Function";
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Sub:
                                sFunctionType = "Sub";
                                break;
                            case ClsCodeMapper.enumFunctionType.eFnType_Property:
                                sFunctionType = "Property";
                                switch (lstLists[0].ePropertyType)
                                {
                                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
                                        sPropertyType = "Get";
                                        break;
                                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
                                        sPropertyType = "Let";
                                        break;
                                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
                                        sPropertyType = "Set";
                                        break;
                                }
                                break;
                            default:
                                sFunctionType = "Unknown";
                                break;
                        }




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
*/
    }
}
