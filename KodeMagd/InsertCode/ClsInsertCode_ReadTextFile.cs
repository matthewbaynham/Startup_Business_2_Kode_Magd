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
using Office = Microsoft.Office.Core;
using KodeMagd.Misc;
using KodeMagd.Settings;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_ReadTextFile : ClsInsertCode
    {
        private string sFullPath;
        private bool bUsingWith;
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

        public string FullPath
        {
            get
            {
                try
                {
                    return sFullPath;
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
                    sFullPath = value;
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

        public void generateCode(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = "";

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Read_Text_File");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Dim fso as Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set fso = new Scripting.FileSystemObject");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If fso.FileExists(\"" + sFullPath + "\") Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent));

                if (bUsingWith)
                { 
                    sWithTemp = "";
                    lstCode.Add(cSettings.Indent(iIndent) + "With ts");
                    iIndent++;
                }
                else
                { sWithTemp = "ts"; }

                lstCode.Add(cSettings.Indent(iIndent) + "Do While " + sWithTemp + ".AtEndOfStream");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "sLine = " + sWithTemp + ".ReadLine");
                lstCode.Add(cSettings.Indent(iIndent));
                if (cSettings.UserTips == true)
                { lstCode.Add(cSettings.Indent(iIndent) + "'This outputs the line of data into the Immediate window"); }
                lstCode.Add(cSettings.Indent(iIndent) + "Debug.Print sLine");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Loop");
                
                if (bUsingWith)
                {
                    sWithTemp = "";
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                }
                
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "Set fso = nothing");
                lstCode.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                this.addCode(ref lstCode);

                if (lstCodeTop.Count > 0)
                {
                    lstCodeTop.Add("");
                    this.addCode(ref lstCodeTop, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                //cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeTop = null;
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
