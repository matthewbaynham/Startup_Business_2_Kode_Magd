using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_ConnectionString : ClsInsertCode
    {
        private bool bDimVariable = false;
        private string sVariableName = "";
        private string sConnectionString = "";
        private string sFunctionName = "";


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

        public bool dimVariable
        {
            get
            {
                try
                {
                    return bDimVariable;
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
            set
            {
                try
                {
                    bDimVariable = value;
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

        public string variableName
        {
            get
            {
                try
                {
                    return sVariableName;
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
            set
            {
                try
                {
                    sVariableName = value;
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

        public string connectionString
        {
            get 
            {
                try
                {
                    return sConnectionString;
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
            set
            {
                try
                {
                    sConnectionString = value;
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
                List<string> lstCodeExtraFn = new List<string>();
                int iIndent = 0;

                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();

                List<string> lstCodeOptions = new List<string>();

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeOptions.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeOptions.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper,  "Connection_String");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                lstCode.Add(cSettings.Indent(iIndent));
                if (this.bDimVariable)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim " + sVariableName + " As String");
                    lstCode.Add(cSettings.Indent(iIndent));
                }
                
                lstCode.Add(cSettings.Indent(iIndent) + sVariableName + " = " + ClsMisc.replaceReturnCharInQuotedTxtWithConst(ClsMiscString.addQuotes(sConnectionString)));
                lstCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }

                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                }
                lstCode.Add(cSettings.Indent(iIndent));

                this.addCode(ref lstCode, ref vbComp);
                this.addCode(ref lstCodeOptions, ref vbComp, enumPosition.ePosBeginningAfterOptions);

                cSettings = null;
                //cCodeMapper = null;
                cDataTypes = null;
                lstCode = null;
                lstCodeExtraFn = null;
                vbComp = null;
                lstCodeOptions = null;
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
