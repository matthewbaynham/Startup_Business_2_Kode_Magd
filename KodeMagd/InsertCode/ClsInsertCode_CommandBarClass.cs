using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using KodeMagd.Misc;
using Office = Microsoft.Office.Core;

namespace KodeMagd.InsertCode
{
    public class ClsInsertCode_CommandBarClass : ClsInsertCode
    {
        private const string csSampleCodeModulePrefix = "SampleCode_";
        private const string csPrefixProperty = "prop";
        public enum enumMenuType
        { 
            eMenuRibbonAddin,
            eMenuRightClick
        }

        private enum enumText 
        { 
            eText_Declare,
            eText_Initialise,
            eText_PropertyGet
        }
 
        public struct strCommandControl 
        {
            public string sVariableName;
            public string sFullPath;
            public string sFullPathParent;
            public string sCaption;
            public string sOnAction;
            public string sTooltipText;
            public Office.MsoControlType eType;
            public List<string> lstCmbValues;
        }

        List<strCommandControl> lstCommandBarControls = new List<strCommandControl>();
        private string sClassName;
        private bool bPutSampleCallInOwnNewMod;
        private enumMenuType eMenuType;

        public List<strCommandControl> CommandBarControls
        {
            get 
            {
                try
                {
                    return lstCommandBarControls;
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
                    return new List<strCommandControl>();
                }
            }
        }

        public string SampleCodeModulePrefix
        {
            get
            {
                try
                {
                    return csSampleCodeModulePrefix;
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

        public string className 
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
                    return "";
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

        public bool PutSampleCallInOwnNewMod
        {
            get
            {
                try
                {
                    return bPutSampleCallInOwnNewMod;
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
                    return true;
                }
            }
            set
            {
                try
                {
                    bPutSampleCallInOwnNewMod = value;
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

        public enumMenuType menuType
        {
            get
            {
                try
                {
                    return eMenuType;
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
                    return enumMenuType.eMenuRightClick;
                }
            }
            set
            {
                try
                {
                    eMenuType = value;
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

        public void addControl(strCommandControl objControl) 
        {
            try 
            {
                lstCommandBarControls.Add(objControl);
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

        public void editControl(strCommandControl objControl)
        {
            try
            {
                if (lstCommandBarControls.Any(x => x.sFullPath == objControl.sFullPath))
                {
                    int iIndexControl = lstCommandBarControls.FindIndex(x => x.sFullPath == objControl.sFullPath);

                    lstCommandBarControls[iIndexControl] = objControl;
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


        public void deleteControl(string sFullPath)
        {
            try
            {
                for (int iCounter = lstCommandBarControls.Count - 1; iCounter >= 0; iCounter--)
                {
                    strCommandControl objTemp = lstCommandBarControls[iCounter];
                    if (objTemp.sFullPathParent == sFullPath)
                    { deleteControl(objTemp.sFullPath); }
                }
                
                //delete this node
                if (lstCommandBarControls.Any(x => x.sFullPath == sFullPath))
                {
                    int iIndexControl = lstCommandBarControls.FindIndex(x => x.sFullPath == sFullPath);

                    lstCommandBarControls.RemoveAt(iIndexControl);
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

        public void deleteControl(strCommandControl objControl)
        {
            try
            {
                if (lstCommandBarControls.Any(x => x.sFullPath == objControl.sFullPath))
                {
                    int iIndexControl = lstCommandBarControls.FindIndex(x => x.sFullPath == objControl.sFullPath);

                    lstCommandBarControls.RemoveAt(iIndexControl);
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

        public strCommandControl findControl(string sFullPath) 
        {
            try
            {
                //strCommandControl objResult = lstCommandBarControls.Find(x => x.iId == iId);
                strCommandControl objResult = lstCommandBarControls.Find(x => x.sFullPath == sFullPath);

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

                strCommandControl objError;

                objError.eType = Office.MsoControlType.msoControlCustom;
                //objError.iId = ClsMisc.gciError;
                //objError.iParentId = ClsMisc.gciError;
                objError.sVariableName = "";
                objError.sFullPath = "";
                objError.sFullPathParent = "";
                objError.sCaption = "";
                objError.sOnAction = "";
                objError.sTooltipText = "";
                objError.lstCmbValues = new List<string>();
                objError.lstCmbValues.Clear();

                return objError;
            }
        }

        public bool isExistsComboBox() 
        { 
            try
            {
                //bool bIsfound;

                //strCommandControl objControl = lstCommandBarControls.Find(X => X.eType == Office.MsoControlType.msoControlComboBox);
                bool bIsfound = lstCommandBarControls.Exists(X => X.eType == Office.MsoControlType.msoControlComboBox);

                //if (objControl = null)
                //{ bIsfound = false; }
                //else
                //{ bIsfound = true; }

                return bIsfound;
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

                strCommandControl objError;

                return false;
            }
        }


        //public strCommandControl getParent(int iParentId)
        public strCommandControl getParent(string sFullPathChild)
        {
            try
            {
                int iPos = sFullPathChild.LastIndexOf('\\');
                string sFullPath = ClsMiscString.Left(ref sFullPathChild, iPos);

                strCommandControl objResult = lstCommandBarControls.Find(x => x.sFullPath == sFullPath);

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

                strCommandControl objError;

                objError.eType = Office.MsoControlType.msoControlCustom;
                //objError.iId = ClsMisc.gciError;
                //objError.iParentId = ClsMisc.gciError;
                objError.sVariableName = "";
                objError.sFullPath = "";
                objError.sFullPathParent = "";
                objError.sCaption = "";
                objError.sOnAction = "";
                objError.sTooltipText = "";
                objError.lstCmbValues = new List<string>();
                objError.lstCmbValues.Clear();

                return objError;
            }
        }

        public void allocateVariableNames() 
        {
            try
            {
                List<string> lstVariables = new List<string>();
                lstVariables.Clear();

                for (int iCounter = 0; iCounter < lstCommandBarControls.Count; iCounter++)
                {
                    strCommandControl objTemp = lstCommandBarControls[iCounter];
                    string sPrefix = "";
                    switch (objTemp.eType)
                    {
                        case Office.MsoControlType.msoControlActiveX:
                            sPrefix = "actX";
                            break;
                        case Office.MsoControlType.msoControlAutoCompleteCombo:
                        case Office.MsoControlType.msoControlComboBox:
                        case Office.MsoControlType.msoControlGraphicCombo:
                            sPrefix = "cmb";
                            break;
                        case Office.MsoControlType.msoControlButton:
                        case Office.MsoControlType.msoControlButtonDropdown:
                        case Office.MsoControlType.msoControlButtonPopup:
                            sPrefix = "btn";
                            break;
                        case Office.MsoControlType.msoControlCustom:
                            sPrefix = "cust";
                            break;
                        case Office.MsoControlType.msoControlEdit:
                            sPrefix = "edt";
                            break;
                        case Office.MsoControlType.msoControlExpandingGrid:
                            sPrefix = "grd";
                            break;
                        case Office.MsoControlType.msoControlGauge:
                            sPrefix = "gug";
                            break;
                        case Office.MsoControlType.msoControlDropdown:
                        case Office.MsoControlType.msoControlGenericDropdown:
                        case Office.MsoControlType.msoControlGraphicDropdown:
                        case Office.MsoControlType.msoControlGraphicPopup:
                        case Office.MsoControlType.msoControlPopup:
                        case Office.MsoControlType.msoControlOCXDropdown:
                        case Office.MsoControlType.msoControlSplitButtonMRUPopup:
                        case Office.MsoControlType.msoControlSplitButtonPopup:
                        case Office.MsoControlType.msoControlSplitDropdown:
                            sPrefix = "mnu";
                            break;
                        case Office.MsoControlType.msoControlGrid:
                        case Office.MsoControlType.msoControlSplitExpandingGrid:
                            sPrefix = "grd";
                            break;
                        case Office.MsoControlType.msoControlLabel:
                        case Office.MsoControlType.msoControlLabelEx:
                            sPrefix = "lbl";
                            break;
                        case Office.MsoControlType.msoControlPane:
                        case Office.MsoControlType.msoControlWorkPane:
                            sPrefix = "Pane";
                            break;
                        case Office.MsoControlType.msoControlSpinner:
                            sPrefix = "spn";
                            break;
                        default:
                            sPrefix = "";
                            break;
                    }

                    objTemp.sVariableName = sPrefix + ClsMiscString.makeValidVarName(objTemp.sCaption);
                    lstVariables.Add(objTemp.sVariableName);
                    lstCommandBarControls[iCounter] = objTemp;
                }

                var duplicateItems = lstVariables.GroupBy(x => x).Where(x => x.Count() > 1).Select(x => x.Key);

                foreach (string sDupe in duplicateItems) 
                {
                    int iSuffex = 1;

                    for (int iCounter = 0; iCounter < lstCommandBarControls.Count; iCounter++)
                    {
                        strCommandControl objTemp = lstCommandBarControls[iCounter];

                        if (objTemp.sVariableName == sDupe)
                        { 
                            objTemp.sVariableName += iSuffex.ToString();
                            iSuffex++;
                        }

                        lstCommandBarControls[iCounter] = objTemp;
                    }
                }
                lstVariables = null;
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

        public void generateToolbarClass()
        { 
            try 
            {
                allocateVariableNames();

                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = 0;

                VBA.VBComponent vbComp = addModule(sClassName, VBA.vbext_ComponentType.vbext_ct_ClassModule);

                /*
                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }
                */
                lstCodeTop.Add("Option Explicit");
                lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase);

                lstCode.Add(cSettings.Indent(iIndent));

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'When the code to initialise the commandbar has been run then be very careful");
                lstCode.Add(cSettings.Indent(iIndent) + "'if any VBA crashes of if someone hits the stop button in the VBA editor");
                lstCode.Add(cSettings.Indent(iIndent) + "'window, then the connection between the commandbar and the VBA commandbar objects");
                lstCode.Add(cSettings.Indent(iIndent) + "'will stop.  So you will have to run the initialise code again.");
                lstCode.Add(cSettings.Indent(iIndent));

                loopThroughControls(ref cSettings, iIndent, ref lstCode, 0, "", enumText.eText_Declare);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Get Properties are only genereated for ComboBox Controls");
                lstCode.Add(cSettings.Indent(iIndent) + "'If any part of your application needs to know the value inside a ComboBox on the toolbar please use the Get Properties");
                lstCode.Add(cSettings.Indent(iIndent));
                if (eMenuType == enumMenuType.eMenuRightClick) 
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Private cmdRightClickMenuBar As CommandBar");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                loopThroughControls(ref cSettings, iIndent, ref lstCode, 0, "", enumText.eText_PropertyGet);

                lstCode.Add(cSettings.Indent(iIndent));

                lstCode.Add(cSettings.Indent(iIndent) + "Public Sub createToolBar()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                lstCode.Add(cSettings.Indent(iIndent));

                if (eMenuType == enumMenuType.eMenuRightClick) 
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "'Get the right click menu bar");
                    lstCode.Add(cSettings.Indent(iIndent) + "Set cmdRightClickMenuBar = Application.CommandBars(\"Cell\")");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                loopThroughControls(ref cSettings, iIndent, ref lstCode, 0, "", enumText.eText_Initialise);
                lstCode.Add(cSettings.Indent(iIndent));
                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Public Sub deleteToolBar()");
                if (cSettings.IndentFirstLevel) { iIndent++; }
                addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent) + "Dim cmd As CommandBar");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "For Each cmd in ThisWorkbook.Application.CommandBars");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If cmd.Name = \"" + sClassName + "\" Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "cmd.Delete");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next cmd");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set cmd = Nothing");
                lstCode.Add(cSettings.Indent(iIndent));

                addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                if (cSettings.IndentFirstLevel) { iIndent--; }
                lstCode.Add(cSettings.Indent(iIndent) + "End Sub");

                this.addCode(ref lstCode, ref vbComp);

                if (lstCodeTop.Count > 0)
                {
                    lstCodeTop.Add("");
                    this.addCode(ref lstCodeTop, ref vbComp, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                //cCodeMapper = null;
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

        private void loopThroughControls(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, int iLevel, string sParentPath, enumText eText)
        {
            try
            {
                foreach (strCommandControl ctrl in lstCommandBarControls)
                {
                    int iNumBackslashes = ctrl.sFullPath.Count(x => x == '\\');

                    if (iNumBackslashes == iLevel & ctrl.sFullPathParent == sParentPath) 
                    {
                        switch (eText)
                        {
                            case enumText.eText_Declare:
                                if (iLevel == 0)
                                {
                                    switch (eMenuType)
                                    {
                                        case enumMenuType.eMenuRibbonAddin:
                                            generateCommandBarDeclare(ref cSettings, iIndent, ref lstCode, ctrl);
                                            break;
                                        case enumMenuType.eMenuRightClick:
                                            generateCommandControlDeclare(ref cSettings, iIndent, ref lstCode, ctrl);
                                            break;
                                    }
                                }
                                else
                                { generateCommandControlDeclare(ref cSettings, iIndent, ref lstCode, ctrl); }
                                break;
                            case enumText.eText_Initialise:
                                if (iLevel == 0)
                                { 
                                    switch (eMenuType) {
                                        case enumMenuType.eMenuRibbonAddin:
                                            generateCommandBarInitialise(ref cSettings, iIndent, ref lstCode, ctrl);
                                            break;
                                        case enumMenuType.eMenuRightClick:
                                            generateCommandControlInitialise(ref cSettings, iIndent, ref lstCode, ctrl);
                                            break;
                                    }
                                }
                                else
                                { generateCommandControlInitialise(ref cSettings, iIndent, ref lstCode, ctrl); }
                                break;
                            case enumText.eText_PropertyGet:
                                generateGets(ref cSettings, iIndent, ref lstCode, ctrl);
                                break;
                            default:
                                break;
                        }

                        loopThroughControls(ref cSettings, iIndent, ref lstCode, iLevel + 1, ctrl.sFullPath, eText);
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

        private void generateCommandControlDeclare(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, strCommandControl ctrl)
        {
            try
            {
                string sTypeName = "";
                
                switch (ctrl.eType)
                {
                    case Office.MsoControlType.msoControlPopup:
                        sTypeName = "CommandBarPopUp";
                        break;
                    case Office.MsoControlType.msoControlButton:
                        sTypeName = "CommandBarButton";
                        break;
                    case Office.MsoControlType.msoControlComboBox:
                        sTypeName = "CommandBarComboBox";
                        break;
                    default:
                        sTypeName = ctrl.eType.ToString().Replace("msoControl", "CommandBar");
                        break;
                }

                lstCode.Add(cSettings.Indent(iIndent) + "Private " + ctrl.sVariableName + " As " + sTypeName);
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

        private void generateCommandBarDeclare(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, strCommandControl ctrl)
        {
            try
            {
                lstCode.Add(cSettings.Indent(iIndent) + "Private " + ctrl.sVariableName + " As CommandBar");
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

        private void generateGets(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, strCommandControl ctrl)
        {
            try
            {
                //only applies to some controls which have a value
                bool bIsRequired;

                switch (ctrl.eType)
                {
                    case Office.MsoControlType.msoControlActiveX:
                    case Office.MsoControlType.msoControlButton:
                    case Office.MsoControlType.msoControlButtonDropdown:
                    case Office.MsoControlType.msoControlButtonPopup:
                    case Office.MsoControlType.msoControlCustom:
                    case Office.MsoControlType.msoControlEdit:
                    case Office.MsoControlType.msoControlExpandingGrid:
                    case Office.MsoControlType.msoControlDropdown:
                    case Office.MsoControlType.msoControlGenericDropdown:
                    case Office.MsoControlType.msoControlGraphicDropdown:
                    case Office.MsoControlType.msoControlGraphicPopup:
                    case Office.MsoControlType.msoControlPopup:
                    case Office.MsoControlType.msoControlOCXDropdown:
                    case Office.MsoControlType.msoControlSplitButtonMRUPopup:
                    case Office.MsoControlType.msoControlSplitButtonPopup:
                    case Office.MsoControlType.msoControlSplitDropdown:
                    case Office.MsoControlType.msoControlGrid:
                    case Office.MsoControlType.msoControlSplitExpandingGrid:
                    case Office.MsoControlType.msoControlLabel:
                    case Office.MsoControlType.msoControlLabelEx:
                    case Office.MsoControlType.msoControlPane:
                    case Office.MsoControlType.msoControlWorkPane:

                    case Office.MsoControlType.msoControlSpinner:
                    case Office.MsoControlType.msoControlAutoCompleteCombo:
                    case Office.MsoControlType.msoControlGraphicCombo:
                    case Office.MsoControlType.msoControlGauge:
                        bIsRequired = false;
                        break;

                    case Office.MsoControlType.msoControlComboBox:
                    
                        bIsRequired = true;
                        break;
                    default:
                        bIsRequired = false;
                        break;
                }

                if (bIsRequired)
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Property Get " + csPrefixProperty + ctrl.sVariableName + "() as String");
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + csPrefixProperty + ctrl.sVariableName + " = " + ctrl.sVariableName + ".Text");
                    lstCode.Add(cSettings.Indent(iIndent));
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Property);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Property");
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

        private void generateCommandControlInitialise(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, strCommandControl ctrl) 
        {
            try 
            {
                strCommandControl ctrlParent = this.getParent(ctrl.sFullPath);

                lstCode.Add(cSettings.Indent(iIndent));

                if (!ctrl.sFullPath.Contains('\\') & eMenuType == enumMenuType.eMenuRightClick)
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ctrl.sVariableName + " = cmdRightClickMenuBar.Controls.Add(Type:=" + ctrl.eType.ToString() + ", Temporary:=True)"); }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + "Set " + ctrl.sVariableName + " = " + ctrlParent.sVariableName + ".Controls.Add(Type:=" + ctrl.eType.ToString() + ")"); }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "With " + ctrl.sVariableName);
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + ".OnAction = \"" + ctrl.sOnAction +  "\"");

                if (ctrl.eType == Office.MsoControlType.msoControlComboBox)
                { lstCode.Add(cSettings.Indent(iIndent) + ".Text = \"" + ctrl.sCaption + "\""); }
                else
                { lstCode.Add(cSettings.Indent(iIndent) + ".Caption = \"" + ctrl.sCaption + "\""); }
                
                lstCode.Add(cSettings.Indent(iIndent) + ".TooltipText = \"" + ctrl.sTooltipText + "\"");

                if (ctrl.eType == Office.MsoControlType.msoControlButton)
                { lstCode.Add(cSettings.Indent(iIndent) + ".Style = msoButtonCaption"); }
                
                lstCode.Add(cSettings.Indent(iIndent) + ".Visible = true");
                lstCode.Add(cSettings.Indent(iIndent) + ".Enabled = true");
                if (ctrl.eType == Office.MsoControlType.msoControlComboBox) 
                {
                    foreach (string sItem in ctrl.lstCmbValues)
                    { lstCode.Add(cSettings.Indent(iIndent) + "Call .AddItem(\"" + sItem + "\")"); }
                }

                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End With");
                lstCode.Add(cSettings.Indent(iIndent));
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

        private void generateCommandBarInitialise(ref ClsSettings cSettings, int iIndent, ref List<string> lstCode, strCommandControl ctrl)
        {
            try
            {
                strCommandControl ctrlParent = this.getParent(ctrl.sFullPath);


                //lstCode.Add(cSettings.Indent(iIndent) + "Set " + ctrl.sVariableName + " = " + ctrlParent.sVariableName + ".Controls.Add(Type:=" + ctrl.eType.ToString() + ")");

                lstCode.Add(cSettings.Indent(iIndent) + "Call deleteToolbar()");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set " + ctrl.sVariableName + " = ThisWorkbook.Application.CommandBars.Add(Name:=\"" + sClassName + "\")");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "With " + ctrl.sVariableName);
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + ".Visible = True");
                lstCode.Add(cSettings.Indent(iIndent) + ".Enabled = True");
                //lstCode.Add(cSettings.Indent(iIndent) + ".Caption = \"" + ctrl.sCaption + "\"");
                //lstCode.Add(cSettings.Indent(iIndent) + ".TooltipText = \"" + ctrl.sTooltipText + "\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End With");
                lstCode.Add(cSettings.Indent(iIndent));
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

        public void generateSampleCode(ref ClsCodeMapper cCodeMapper) 
        { 
            try
            {
                allocateVariableNames();

                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCodeCall = new List<string>();
                List<string> lstCodeDeclare = new List<string>();
                int iIndent = 0;

                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent(); 

                /*
                if (bPutSampleCallInOwnNewMod)
                { 
                    vbComp = addModule(csSampleCodeModulePrefix + sClassName, VBA.vbext_ComponentType.vbext_ct_StdModule);

                    cCodeMapper = cCodeMapperWrk.getCodeMapper(csSampleCodeModulePrefix + sClassName);
                }
                else
                { vbComp = ClsMisc.ActiveVBComponent(); }
                */

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeDeclare.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeDeclare.Add("Option Base " + cSettings.defaultOptionBase); }

                lstCodeDeclare.Add(cSettings.Indent(iIndent));

                /*
                 * Will be put at the top of module after the call has been added in the current position
                 * Note: has to be after because I don't want to change the current location
                 */

                lstCodeDeclare.Add(cSettings.Indent(iIndent) + "public cToolbar as " + sClassName);
                lstCodeDeclare.Add(cSettings.Indent(iIndent));

                /*
                 * This code is calling the toolbar code
                 * 
                 */

                lstCodeCall.Add(cSettings.Indent(iIndent));
                addTitleComment(ref lstCodeCall, ref cSettings, iIndent);

                if (!cCodeMapper.cursorIsInFunction)
                {
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    lstCodeCall.Add(cSettings.Indent(iIndent) + "Public Sub " + getNextSampleFunctionName(ref cCodeMapper, "_Initialize"));
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCodeCall, ref cSettings, iIndent);
                }
                lstCodeCall.Add(cSettings.Indent(iIndent) + "'Initialize Toolbar");
                lstCodeCall.Add(cSettings.Indent(iIndent)); 
                lstCodeCall.Add(cSettings.Indent(iIndent) + "Set cToolbar = New " + sClassName);
                lstCodeCall.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                { lstCodeCall.Add(cSettings.Indent(iIndent) + "Call cToolbar.deleteToolBar() 'It's always a good idea to delete first just encase it's already there"); }
                else
                { lstCodeCall.Add(cSettings.Indent(iIndent) + "Call cToolbar.deleteToolBar()"); }
                lstCodeCall.Add(cSettings.Indent(iIndent) + "Call cToolbar.createToolBar()");
                lstCodeCall.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCodeCall, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCodeCall.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                }
                lstCodeCall.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    lstCodeCall.Add(cSettings.Indent(iIndent) + "Public Sub " + getNextSampleFunctionName(ref cCodeMapper, "_Deinitialize"));
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCodeCall, ref cSettings, iIndent);
                }
                lstCodeCall.Add(cSettings.Indent(iIndent));
                lstCodeCall.Add(cSettings.Indent(iIndent) + "'Deinitialize Toolbar");
                lstCodeCall.Add(cSettings.Indent(iIndent) + "Set cToolbar = New " + sClassName);
                lstCodeCall.Add(cSettings.Indent(iIndent));
                lstCodeCall.Add(cSettings.Indent(iIndent) + "Call cToolbar.deleteToolBar()");
                lstCodeCall.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCodeCall, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCodeCall.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                }

                if (this.lstCommandBarControls.Count(x => x.eType == Office.MsoControlType.msoControlComboBox) > 0)
                {
                    if (!cCodeMapper.cursorIsInFunction)
                    {
                        lstCodeCall.Add(cSettings.Indent(iIndent));
                        lstCodeCall.Add(cSettings.Indent(iIndent) + "Public Sub " + getNextSampleFunctionName(ref cCodeMapper, "_Value_in_ComboBox"));
                        if (cSettings.IndentFirstLevel) { iIndent++; }
                        addErrorHandlerCall(ref lstCodeCall, ref cSettings, iIndent);
                    }

                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    foreach (strCommandControl ctrl in this.lstCommandBarControls)
                    {
                        if (ctrl.eType == Office.MsoControlType.msoControlComboBox)
                        { lstCodeCall.Add(cSettings.Indent(iIndent) + "Debug.Print cToolbar." + csPrefixProperty + ctrl.sVariableName); }
                    }
                    lstCodeCall.Add(cSettings.Indent(iIndent));
                    if (!cCodeMapper.cursorIsInFunction)
                    {
                        addErrorHandlerBody(ref lstCodeCall, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                        if (cSettings.IndentFirstLevel) { iIndent--; }
                        lstCodeCall.Add(cSettings.Indent(iIndent) + "End Sub");
                        lstCodeCall.Add(cSettings.Indent(iIndent));
                    }
                }

                this.addCode(ref lstCodeCall, ref vbComp);

                this.addCode(ref lstCodeDeclare, ref vbComp, enumPosition.ePosBeginningAfterOptions);

                cSettings = null;
                cCodeMapper = null;
                lstCodeCall = null;
                lstCodeDeclare = null;
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

        public bool isAllSubMenusAreSubMenus() 
        {
            try 
            {
                bool bIsOk = true;
                /*
                 * 1) loop through all nodes
                 * 2) find the parent of each node
                 * 3) check the parent is a submenu NOT a button or combobox
                 */

                foreach (strCommandControl objNode in lstCommandBarControls)
                {
                    strCommandControl objParent = findControl(objNode.sFullPathParent);

                    if (!string.IsNullOrEmpty(objParent.sFullPath))
                    {
                        if (objParent.eType != Office.MsoControlType.msoControlPopup)
                        { bIsOk = false; }
                    }
                }

                return bIsOk;
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
}
