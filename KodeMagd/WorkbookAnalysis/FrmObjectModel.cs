using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.WorkbookAnalysis
{
    public partial class FrmObjectModel : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        private string sTextAll = "<All>";

        //public struct strReportSpec
        //{
        //    public bool bOnlyPublicFunctions;
        //}

        public FrmObjectModel()
        {
            InitializeComponent();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                generate();
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

        public void generate()
        {
            try
            {
                ClsConfigReporterObjectModel cConfigReporterObjectModel = new ClsConfigReporterObjectModel();
                bool bIsOk = true;
                string sMessage = "";
                ClsConfigReporterObjectModel.strReportSpec objReportSpec = new ClsConfigReporterObjectModel.strReportSpec();

                //Scripting.FileSystemObject fso = new Scripting.FileSystemObject();
                List<string> lstText = new List<string>();

                //string sFullPath = txtOutputPath.Text;

                //if (fso.FileExists(sFullPath))
                //{
                //    bIsOk = false;
                //    sMessage= "File already exists.\n\rCancelling operation.";
                //}

                objReportSpec.bOnlyPublicFunctions = chkLstOnlyPublicFunctions.Checked;
                objReportSpec.lstModules = selectedModules();
                
                if (bIsOk)
                {
                    string sHtml = cConfigReporterObjectModel.createObjectModelHtml(ref cCodeMapperWrk, objReportSpec);

                    //Scripting.TextStream tsFile = fso.CreateTextFile(sFullPath, true, true);

                    //foreach (string sText in lstText)
                    //{ tsFile.WriteLine(sText); }

                    //tsFile.Close();

                    //tsFile = null;


                    //string sHtml = "";

                    foreach (string sText in lstText)
                    { sHtml += sText; }

                    FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Object Model");

                    frm.ShowDialog(this);

                    frm = null;

                    this.Close();
                }
                else
                { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

                //fso = null;
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

        //public void createHtml5Report(ref List<string> lstReport, strReportSpec objReportSpec)
        //{
        //    try
        //    {
        //        string sLine;

        //        sLine = "<!DCOTYPE html>\n";
        //        sLine += "<html>\n";
        //        sLine += "<head>\n";
        //        sLine += "<title>" + ClsCodeEditorGUI.csCommandBarName + " - " + ClsMisc.ActiveWorkBook().Name + "</title>";
        //        sLine += "<meta http-equiv='X-UA-Compatible' content='IE=9' >\n";
        //        sLine += "</head>\n";
        //        sLine += "<body>\n";
        //        sLine += "<canvas id='myCanvas' width=1000 height=1000 style='border:1px solid #d3d3d3;'>Your browser does not support HTML5 canvas.</canvas>\n";
        //        sLine += "<script type='text/javascript'>\n";
        //        sLine += "var c=document.getElementById('myCanvas');\n";
        //        sLine += "var ctx=c.getContext('2d');\n";


        //        sLine += "\n";
        //        sLine += "var canvasMaxSize = drawObjectModel(c, ctx);\n";
        //        sLine += "c.width = canvasMaxSize[0];\n";
        //        sLine += "c.height = canvasMaxSize[1];\n";
        //        sLine += "drawObjectModel(c, ctx);\n";
        //        sLine += "\n";

        //        sLine += "function drawObjectModel(c, ctx)\n";
        //        sLine += "{\n";
        //        sLine += "var colourFunction = \"#FFFF00\";\n";
        //        sLine += "var colourSub = \"#990099\";\n";
        //        sLine += "var colourProperty = \"#A80000\";\n";
        //        sLine += "var colourModule = \"#00003F\";\n";
        //        sLine += "var fontFunction = \"16px Arial\";\n";
        //        sLine += "var fontModule = \"24px Arial\";\n";
        //        sLine += "\n";
        //        sLine += "var iCanvasMaxWidth = 100;\n";
        //        sLine += "var iCanvasMaxHeight = 100;\n";
        //        sLine += "\n";
        //        sLine += "var iModTop = 30;\n";
        //        sLine += "var iModLeft = 0;\n"; //iModLeft = iModHorSpacing -> below
        //        sLine += "var iModHeight = 20;\n";
        //        sLine += "var iModWidth = 20;\n";
        //        sLine += "\n";
        //        sLine += "var iModMarginTop = 20;\n";
        //        sLine += "var iModMarginBottom = 20;\n";
        //        sLine += "var iModMarginLeft = 20;\n";
        //        sLine += "var iModMarginRight = 20;\n";
        //        sLine += "\n";
        //        sLine += "var iModHorSpacing = 50;\n";
        //        sLine += "var iModVertSpacing = 50;\n";
        //        sLine += "var iFunctionVertSpacing = 4;\n";
        //        sLine += "\n";
        //        sLine += "var iFunctionTop = 20;\n";
        //        sLine += "var iFunctionLeft = 20;\n";
        //        sLine += "var iFunctionHeight = 20;\n";
        //        sLine += "var iFunctionWidth = 20;\n";
        //        sLine += "\n";
        //        sLine += "var iFunctionMarginTop = 5;\n";
        //        sLine += "var iFunctionMarginBottom = 5;\n";
        //        sLine += "var iFunctionMarginLeft = 5;\n";
        //        sLine += "var iFunctionMarginRight = 5;\n";
        //        sLine += "\n";
        //        sLine += "var iTextHeight = 15;\n";
        //        sLine += "var iMaxHorLength = 1;\n";
        //        sLine += "\n";
        //        sLine += "var iCanvasNiceMaxToHave = 1000;\n"; //after the width has gone over this we move down to another row.
        //        sLine += "\n";
        //        sLine += "iModLeft = iModHorSpacing;\n";
        //        sLine += "\n";

        //        lstReport.Add(sLine);

        //        foreach (ClsCodeMapper.strModuleDetails objModule in cCodeMapperWrk.getLstModuleDetails())
        //        {
        //            string sModuleNameText = functionNameFormatted(objModule);

        //            sLine = "/*Module: " + sModuleNameText + "*/\n";
        //            sLine += "iMaxHorLength = 1;\n";
        //            sLine += "iFunctionTop = iModTop + iModMarginTop;\n";
        //            sLine += "iFunctionLeft = iModLeft + iModMarginLeft;\n";
        //            sLine += "ctx.font = fontModule;\n";
        //            sLine += "ctx.fillText('" + sModuleNameText + "', iModLeft, iModTop);\n";
        //            sLine += "iMaxHorLength = ctx.measureText('" + sModuleNameText + "').width;\n";
        //            sLine += "\n";
        //            lstReport.Add(sLine);

        //            Predicate<ClsCodeMapper.strFunctions> predFunction;
        //            if (objReportSpec.bOnlyPublicFunctions)
        //            { predFunction = x => x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Public; }
        //            else
        //            { predFunction = x => x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Public || x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Private || x.eScope == ClsCodeMapper.enumScopeFn.eScopeFn_Friend; }

        //            foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(objModule.sName).FindAll(predFunction).OrderBy(x => x.eFunctionType).ThenBy(y => y.ePropertyType))
        //            {
        //                string sFunctionNameText = functionNameFormatted(objFunction);


        //                //Write out the text for the functions
        //                sLine = "iFunctionHeight = iTextHeight + iFunctionMarginTop + iFunctionMarginBottom;\n";
        //                sLine += "ctx.font = fontFunction;\n";

        //                if (objFunction.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
        //                {
        //                    sLine += "ctx.fillText('" + sFunctionNameText + ")', iFunctionLeft + iFunctionMarginLeft, iFunctionTop + (iFunctionHeight - iFunctionMarginBottom));\n";
        //                }
        //                else
        //                {
        //                    sLine += "ctx.fillText('" + sFunctionNameText + "', iFunctionLeft + iFunctionMarginLeft, iFunctionTop + (iFunctionHeight - iFunctionMarginBottom));\n";
        //                }
        //                //sLine += "iFunctionHeight = ctx.measureText('" + sFunctionName + "').height;\n";

        //                sLine += "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;\n";

        //                sLine += "if (ctx.measureText('" + sFunctionNameText + "').width > iMaxHorLength)\n";
        //                sLine += "{ iMaxHorLength = ctx.measureText('" + sFunctionNameText + "').width; }\n";
        //                sLine += "\n";

        //                lstReport.Add(sLine);

        //            }

        //            sLine = "iFunctionTop = iModTop + iModMarginTop;\n";
        //            sLine += "iFunctionWidth = iMaxHorLength + iFunctionMarginLeft + iFunctionMarginRight;\n";
        //            sLine += "iFunctionLeft = iModLeft + iModMarginLeft;\n";
        //            sLine += "\n";

        //            lstReport.Add(sLine);

        //            foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapperWrk.getLstFunctions(objModule.sName).FindAll(predFunction).OrderBy(x => x.eFunctionType).ThenBy(y => y.ePropertyType))
        //            {
        //                //string sFunctionNameText = functionNameFormatted(objFunction);
        //                //draw the rectangle for the functions
        //                switch (objFunction.eFunctionType)
        //                {
        //                    case ClsCodeMapper.enumFunctionType.eFnType_Function:
        //                        sLine = "ctx.strokeStyle = colourFunction;\n";
        //                        break;
        //                    case ClsCodeMapper.enumFunctionType.eFnType_Sub:
        //                        sLine = "ctx.strokeStyle = colourSub;\n";
        //                        break;
        //                    case ClsCodeMapper.enumFunctionType.eFnType_Property:
        //                        sLine = "ctx.strokeStyle = colourProperty;\n";
        //                        break;
                        
        //                }

        //                sLine += "ctx.strokeRect(iFunctionLeft, iFunctionTop, iFunctionWidth, iFunctionHeight);\n";
        //                sLine += "iFunctionTop = iFunctionTop + iFunctionHeight + iFunctionVertSpacing;\n";
        //                sLine += "\n";

        //                lstReport.Add(sLine);
        //            }

        //            //draw the rectangle for the functions

        //            sLine = "\n";
        //            sLine += "iModWidth = iFunctionWidth + iModMarginLeft + iModMarginRight;\n";
        //            sLine += "iModHeight = iFunctionTop + iFunctionHeight + iFunctionVertSpacing + iModMarginTop + iModMarginBottom - iModTop;\n";
        //            sLine += "ctx.strokeStyle = colourModule;\n";
        //            sLine += "ctx.strokeRect(iModLeft, iModTop, iModWidth, iModHeight);\n";
        //            //sLine += "ctx.stroke();\n";
        //            lstReport.Add(sLine);



                    
        //            sLine = "iModLeft = iModLeft + iFunctionWidth + iModMarginLeft + iModMarginRight + iModHorSpacing;\n";
        //            sLine += "\n";
        //            sLine += "if (iModLeft > iCanvasMaxWidth)\n";
        //            sLine += "{ iCanvasMaxWidth = iModLeft; }\n";
        //            sLine += "\n";
        //            sLine += "if (iModTop + iModHeight > iCanvasMaxHeight)\n";
        //            sLine += "{ iCanvasMaxHeight = iModTop + iModHeight; }\n";
        //            sLine += "\n";
        //            sLine += "if (iModLeft > iCanvasNiceMaxToHave)\n";
        //            sLine += "{ \n";
        //            sLine += "iModLeft = iModHorSpacing;\n";
        //            sLine += "iModTop = iCanvasMaxHeight;\n";
        //            sLine += "}\n";
        //            sLine += "\n";
        //            lstReport.Add(sLine);
        //        }

        //        sLine = "ctx.stroke();\n";
        //        sLine += "return [iCanvasMaxWidth, iCanvasMaxHeight];\n";
        //        sLine += "}\n";

        //        sLine += "</script>\n";
        //        sLine += "</html>\n";
        //        sLine += "</body>\n";
        //        lstReport.Add(sLine);
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

        //private string functionNameFormatted(ClsCodeMapper.strModuleDetails objModule)
        //{
        //    try
        //    {
        //        string sResult = "";
        //        switch (objModule.eType)
        //        {
        //            case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
        //                sResult = "ActiveX: " + objModule.sName;
        //                break;
        //            case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
        //                sResult = "Class: " + objModule.sName;
        //                break;
        //            case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
        //                sResult = "Doc: " + objModule.sName;
        //                break;
        //            case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
        //                sResult = "Form: " + objModule.sName;
        //                break;
        //            case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
        //                sResult = "Module: " + objModule.sName;
        //                break;
        //            default:
        //                break;
        //        }

        //        return sResult;
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
                
        //        return string.Empty;
        //    }
        //}

        //private string functionNameFormatted(ClsCodeMapper.strFunctions objFunction)
        //{
        //    try
        //    {
        //        string sResult = "";
        //        switch (objFunction.eFunctionType)
        //        {
        //            case ClsCodeMapper.enumFunctionType.eFnType_Function:
        //                sResult = "FN: " + objFunction.sName;
        //                break;
        //            case ClsCodeMapper.enumFunctionType.eFnType_Sub:
        //                sResult = "Sub: " + objFunction.sName;
        //                break;
        //            case ClsCodeMapper.enumFunctionType.eFnType_Property:
        //                sResult = "Prop: " + objFunction.sName;
        //                switch (objFunction.ePropertyType)
        //                {
        //                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Get:
        //                        sResult += "(Get)";
        //                        break;
        //                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Let:
        //                        sResult += "(Let)";
        //                        break;
        //                    case ClsCodeMapper.enumFunctionPropertyType.ePropType_Set:
        //                        sResult += "(Set)";
        //                        break;
        //                    default:
        //                        break;
        //                }
        //                break;
        //            default:
        //                break;
        //        }

        //        return sResult;
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
                
        //        return string.Empty;
        //    }
        //}

        private void FrmModuleMap_Load(object sender, EventArgs e)
        {
            try
            {
                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                chkLstOnlyPublicFunctions.Checked = false;
                chkIncludeMemberVariables.Checked = true;

                fillChkLstModules();
                chkMemberVariable_VisibleCheck();

                int iIndexTextAll = chkLstModules.Items.IndexOf(sTextAll);
                chkLstModules.SetSelected(iIndexTextAll, true);
                chkLstModules.SetItemChecked(iIndexTextAll, true);

                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref lblMemberVariables);
                ClsDefaults.FormatControl(ref chkLstModules);
                
                ClsDefaults.FormatControl(ref chkLstOnlyPublicFunctions);
                ClsDefaults.FormatControl(ref chkIncludeMemberVariables);
                ClsDefaults.FormatControl(ref chkIncludeMemberVariablePublic);
                ClsDefaults.FormatControl(ref chkIncludeMemberVariablePrivate);
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblMemberVariables, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkLstModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(chkLstOnlyPublicFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkIncludeMemberVariables, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkIncludeMemberVariablePublic, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkIncludeMemberVariablePrivate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
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

        private void fillChkLstModules()
        {
            try
            {
                chkLstModules.Items.Add(sTextAll);

                foreach (ClsCodeMapper.strModuleDetails objModule in cCodeMapperWrk.getLstModuleDetails().OrderBy(x => x.eType).ThenBy(y => y.sName))
                {
                    string sTemp = "";

                    switch(objModule.eType)
                    {
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                            sTemp += "ActiveX: ";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                            sTemp += "Class: ";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                            sTemp += "Document: ";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                            sTemp += "Form: ";
                            break;
                        case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                            sTemp += "Module: ";
                            break;
                    }

                    sTemp += objModule.sName;

                    chkLstModules.Items.Add(sTemp);
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
        private void checkAllSelected()
        {
            try
            {
                if (chkLstModules.SelectedItem == chkLstModules.Items[chkLstModules.Items.IndexOf(sTextAll)])
                {
                    if (chkLstModules.CheckedItems.Contains(sTextAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < chkLstModules.Items.Count; iIndex++)
                        {
                            if (iIndex != chkLstModules.Items.IndexOf(sTextAll))
                            { chkLstModules.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (chkLstModules.CheckedItems.Count == 0 | (chkLstModules.CheckedItems.Count == 1 & chkLstModules.CheckedItems.Contains(sTextAll)))
                    { bAnySelected = false; }
                    else
                    { bAnySelected = true; }

                    if (bAnySelected)
                    { chkLstModules.SetItemChecked(chkLstModules.Items.IndexOf(sTextAll), false); }
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

        private List<string> selectedModules() 
        {
            try
            {
                List<string> lstResults = new List<string>();

                if (chkLstModules.CheckedItems.Contains(sTextAll))
                {
                    foreach (string sTemp in chkLstModules.Items)
                    {
                        if (sTemp != sTextAll)
                        {
                            string sTemp2 = "";
                            if (sTemp.StartsWith("ActiveX: "))
                            { sTemp2 = sTemp.Substring("ActiveX: ".Length); }
                            else if (sTemp.StartsWith("Class: "))
                            { sTemp2 = sTemp.Substring("Class: ".Length); }
                            else if (sTemp.StartsWith("Document: "))
                            { sTemp2 = sTemp.Substring("Document: ".Length); }
                            else if (sTemp.StartsWith("Form: "))
                            { sTemp2 = sTemp.Substring("Form: ".Length); }
                            else if (sTemp.StartsWith("Module: "))
                            { sTemp2 = sTemp.Substring("Module: ".Length); }
                            else
                            { sTemp2 = sTemp; }

                            lstResults.Add(sTemp2);
                        }
                    }
                }
                else
                {
                    foreach (string sTemp in chkLstModules.CheckedItems)
                    {
                        string sTemp2 = "";
                        if (sTemp.StartsWith("ActiveX: "))
                        { sTemp2 = sTemp.Substring("ActiveX: ".Length); }
                        else if (sTemp.StartsWith("Class: "))
                        { sTemp2 = sTemp.Substring("Class: ".Length); }
                        else if (sTemp.StartsWith("Document: "))
                        { sTemp2 = sTemp.Substring("Document: ".Length); }
                        else if (sTemp.StartsWith("Form: "))
                        { sTemp2 = sTemp.Substring("Form: ".Length); }
                        else if (sTemp.StartsWith("Module: "))
                        { sTemp2 = sTemp.Substring("Module: ".Length); }
                        else 
                        { sTemp2 = sTemp; }

                        lstResults.Add(sTemp2);
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

                return new List<string>();
            }
        }

        private void chkLstModules_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelected();
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

        private void chkLstModules_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelected();
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

        private void chkIncludeMemberVariables_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                chkMemberVariable_VisibleCheck();
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

        private void chkMemberVariable_VisibleCheck()
        {
            try
            {
                bool bShow;
                
                bShow = chkIncludeMemberVariables.Checked;

                chkIncludeMemberVariablePrivate.Visible = bShow;
                chkIncludeMemberVariablePublic.Visible = bShow;

                chkIncludeMemberVariablePrivate.Checked = true;
                chkIncludeMemberVariablePublic.Checked = true;
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

        private void FrmObjectModel_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);

                cControlPosition.positionControl(ref lblMemberVariables);
                cControlPosition.positionControl(ref chkLstModules);

                cControlPosition.positionControl(ref chkLstOnlyPublicFunctions);
                cControlPosition.positionControl(ref chkIncludeMemberVariables);
                cControlPosition.positionControl(ref chkIncludeMemberVariablePublic);
                cControlPosition.positionControl(ref chkIncludeMemberVariablePrivate);
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
