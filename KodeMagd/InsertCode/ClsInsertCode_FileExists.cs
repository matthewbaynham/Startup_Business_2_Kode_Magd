using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_FileExists : ClsInsertCode
    {
        public enum enumType
        {
            eTyp_HardCoded,
            eTyp_Variable,
            eTyp_Unknown
        }

        private enumType eType = enumType.eTyp_Unknown;
        private string sPath = "";
        private string sVariableName = "";
        private string sName = "";

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

        public string name
        {
            get
            {
                try
                { return sName; }
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
                { sName = value; }
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

        public string path
        {
            get
            {
                try
                { return sPath; }
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
                { sPath = value; }
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
                { return sVariableName; }
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
                { sVariableName = value; }
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

        public enumType type
        {
            get
            {
                try
                { return eType; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumType.eTyp_Unknown;
                }
            }
            set
            {
                try
                { eType = value; }
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


        /*
        Dim fso As Scripting.FileSystemObject

        Set fso = New Scripting.FileSystemObject

        If fso.FileExists("") Then
            MsgBox "File Exists.", vbInformation, "File Exists"
        Else
            MsgBox "File not Found.", vbCritical, "Not Found"
        End If

        Set fso = Nothing
         */


        public void generateCode(ref ClsCodeMapper cCodeMapper)
        {
            try
            {
                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                //List<string> lstCode = new List<string>();
                //List<string> lstCodeTop = new List<string>();
                ClsLinesOutputRapper cCode = new ClsLinesOutputRapper();
                ClsLinesOutputRapper cCodeTop = new ClsLinesOutputRapper();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();

                string sFsoName = ClsMiscString.makeValidVarName(name, "fso");
                string sFilePathName = ClsMiscString.makeValidVarName(name, "s");
 
                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { cCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { cCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                cCode.Add(cSettings.Indent(iIndent));

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_File_Exists");
                    cCode.Add(cSettings.Indent(iIndent));
                    cCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref cCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref cCode, ref cSettings, iIndent);

                cCode.Add(cSettings.Indent(iIndent) + "'Dimension Variables/Objects");

                if (!ClsMiscString.isValidVariableName(sFsoName) || !ClsMiscString.isValidVariableName(sFilePathName))
                {
                    cCode.Add(cSettings.Indent(iIndent));
                    cCode.Add(cSettings.Indent(iIndent) + "'WARNING:");
                    cCode.Add(cSettings.Indent(iIndent) + "'Check that the \"name\" that you have entered is OK");
                    cCode.Add(cSettings.Indent(iIndent) + "'" + ClsCodeEditorGUI.csCommandBarName + " will remove spaces for you before it is used as a suffix in the Variable name,");
                    cCode.Add(cSettings.Indent(iIndent) + "'but check there is nothing else wrong with it.");
                    cCode.Add(cSettings.Indent(iIndent));
                }

                cCode.Add(cSettings.Indent(iIndent) + "Dim " + sFsoName + " As Scripting.FileSystemObject");
                switch (this.eType)
                {
                    case enumType.eTyp_HardCoded:
                        cCode.Add(cSettings.Indent(iIndent) + "Dim " + sFilePathName + " As String");
                        break;
                }
                cCode.Add(cSettings.Indent(iIndent));
                cCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = new Scripting.FileSystemObject");
                cCode.Add(cSettings.Indent(iIndent));
                switch (this.eType)
                {
                    case enumType.eTyp_HardCoded:
                        cCode.Add(cSettings.Indent(iIndent) + sFilePathName + " = " + ClsMiscString.addQuotes(sPath));
                        break;
                }
                cCode.Add(cSettings.Indent(iIndent));
                string sTempVar = "";
                switch (this.eType)
                {
                    case enumType.eTyp_HardCoded:
                        sTempVar = sFilePathName;
                        break;
                    case enumType.eTyp_Variable:
                        sTempVar = sVariableName;
                        break;
                }
                cCode.Add(cSettings.Indent(iIndent) + "If " + sFsoName + ".fileExists(" + sTempVar + ") then ");
                iIndent++;
                if (cSettings.UserTips)
                { cCode.Add(cSettings.Indent(iIndent) + "'Write all your code here if the file does exist"); }
                cCode.Add(cSettings.Indent(iIndent) + "MsgBox \"File Exists.\" & VbCrLf  & VbCrLf & " + sTempVar + ", vbInformation, \"File Exists\"");
                iIndent--;
                cCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                if (cSettings.UserTips)
                { cCode.Add(cSettings.Indent(iIndent) + "'Write all your code here if the file does NOT exist"); }
                cCode.Add(cSettings.Indent(iIndent) + "MsgBox \"File not Found.\" & VbCrLf  & VbCrLf & " + sTempVar + ", vbCritical, \"Not Found\"");
                iIndent--;
                cCode.Add(cSettings.Indent(iIndent) + "End If");
                cCode.Add(cSettings.Indent(iIndent));
                cCode.Add(cSettings.Indent(iIndent) + "set " + sFsoName + " = nothing");
                cCode.Add(cSettings.Indent(iIndent));
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref cCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    cCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    cCode.Add(cSettings.Indent(iIndent));
                    cCode.Add(cSettings.Indent(iIndent));
                }

                this.addCode(ref cCode);

                if (cCodeTop.Count > 0)
                {
                    cCodeTop.Add("");
                    this.addCode(ref cCodeTop, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                //cCodeMapper = null;
                cCode = null;
                cCodeTop = null;
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
