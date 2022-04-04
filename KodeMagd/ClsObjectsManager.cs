using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using KodeMagd.Misc;

/*
 Keep track of which modules exist and which functions are in which modules and whether they are public, private or friend
 */

namespace KodeMagd
{
    class ClsObjectsManager
    {
        private enum enumType 
        {
            eType_Sub,
            eType_Function,
            eType_Property
        }

        private Excel.Application app;
        private Excel.Workbook wrk;
        private int iModuleNumber;
        private int iFunctionNumber;
        private struct strFunction
        {
            public int iLine_Start;
            public int iLine_End;
            public string Name;
            public string ModuleName;
            public enumType eType;
        }
        private List<strFunction> lstFunctions = new List<strFunction>();

        private VBA.VBProject vbProj;

        public ClsObjectsManager()
        {
            try
            {
                app = Globals.ThisAddIn.Application;
                wrk = app.ActiveWorkbook;

                if (wrk.HasVBProject == true)
                { vbProj = wrk.VBProject; }
                else
                { vbProj = null; }
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

        ~ClsObjectsManager()
        {
            try
            {
                app = null;
                wrk = null;
                lstFunctions = null;
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

        public bool has_Code
        {
            get
            {
                try
                {
                    bool bResult;

                    bResult = wrk.HasVBProject;

                    return bResult;
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

        public string module_Name()
        {
            try
            {
                return module_Name(iModuleNumber);
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
                return ex.Message;
            }
        }

        public string module_Name(int iModuleNo)
        {
            try
            {
                VBA.VBComponent vbComp = vbProj.VBComponents.Item(iModuleNo + 1);
                read_Module(vbComp);
                return vbComp.Name;
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
                return ex.Message;
            }
        }

        public string module_TypeName()
        {
            try
            {
                return module_TypeName(iModuleNumber);
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
                return ex.Message;
            }
        }

        public string module_TypeName(int iModuleNo)
        {
            try
            {
                string sResult;
                VBA.VBComponent vbComp = vbProj.VBComponents.Item(iModuleNo + 1);
                
                switch (vbComp.Type) {
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                        sResult = "ActiveX Designer";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                        sResult = "Class";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                        sResult = "Document";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                        sResult = "Form";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                        sResult = "Module";
                        break;
                    default:
                        sResult = "Unknown";
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
                return ex.Message;
            }
        }

        public void module_Move_First()
        {
            try
            {
                iModuleNumber = 0;
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

        public void module_Move_Last()
        { 
            try
            {
                iModuleNumber = wrk.VBProject.VBComponents.Count; 
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

        public void module_Move_Next()
        {
            try
            {
                if (iModuleNumber < wrk.VBProject.VBComponents.Count)
                { iModuleNumber++; }
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

        public void module_Move_Previous()
        {
            try
            {
                if (iModuleNumber > 0)
                { iModuleNumber--; }
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

        public void read_Module(string sName)
        {
            try
            {
                foreach (VBA.VBComponent vbComp in vbProj.VBComponents)
                {
                    if (vbComp.Name == sName)
                    {
                        read_Module(vbComp);
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

        public bool module_BOF
        {
            get {
                try
                {
                    bool bResult;

                    if (iModuleNumber > 0) 
                    { bResult = false; }
                    else
                    { bResult = true; }

                    return bResult;
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

        public bool module_EOF
        {
            get {
                try 
                {
                    bool bResult;

                    if (iModuleNumber < wrk.VBProject.VBComponents.Count)
                    { bResult = false; }
                    else
                    { bResult = true; }

                    return bResult;
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

        public void edit_Module(string sName, int iLine)
        {
            try
            {
                foreach (VBA.VBComponent vbComp in vbProj.VBComponents)
                {
                    if (vbComp.Name == sName)
                    {
                        edit_Module(vbComp, iLine);
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

        public void edit_Module(VBA.VBComponent vbComp, int iLine) {
            try {
                string sCodeLine = "'Adding a comment";
                VBA.CodeModule objCode = vbComp.CodeModule;

                objCode.InsertLines(iLine, sCodeLine); 

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

        public void read_Module(VBA.VBComponent vbComp) 
        {
            try
            {
            /*
             Reads all the code in the module and establishes what functions, sub and properties are in the module 
             */
                lstFunctions.Clear();

                VBA.CodeModule objCode = vbComp.CodeModule;
                int iPreviousFunctionLine = objCode.CountOfLines;

                for (int iLine = objCode.CountOfLines; iLine > 0; iLine--)
                {
                    string sLine = objCode.get_Lines(iLine + 1, 1);

                    string sTemp = sLine.Trim();
                    //if (sTemp.Substring(0, 6).ToLower() == "private")
                    if (sTemp.Trim().ToUpper().StartsWith("private ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //sTemp = sTemp.Substring(7).Trim();
                        sTemp = ClsMiscString.Right(ref sTemp, sTemp.Length - 7).Trim();
                    }
                    //else if (sTemp.Substring(0, 5).ToLower() == "public")
                    else if (sTemp.Trim().ToUpper().StartsWith("public ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //sTemp = sTemp.Substring(6).Trim();
                        sTemp = ClsMiscString.Right(ref sTemp, sTemp.Length - 6).Trim();
                    }
                    //else if (sTemp.Substring(0, 5).ToLower() == "friend")
                    else if (sTemp.Trim().ToUpper().StartsWith("friend ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //sTemp = sTemp.Substring(6).Trim();
                        sTemp = ClsMiscString.Right(ref sTemp, sTemp.Length - 6).Trim();
                    }

                    /*'function', 'Sub', 'property'*/
                    //if (sTemp.Substring(0, 3).ToLower() == "sub")
                    if (sTemp.Trim().ToUpper().StartsWith("sub ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        strFunction sFnTemp = new strFunction();

                        sFnTemp.Name = sTemp;
                        sFnTemp.ModuleName = vbComp.Name;
                        sFnTemp.iLine_Start = iLine;
                        sFnTemp.eType = enumType.eType_Sub;
                        sFnTemp.iLine_End = iPreviousFunctionLine;
                        
                        lstFunctions.Add(sFnTemp);

                        iPreviousFunctionLine = iLine - 1;
                    }
                    //else if (sTemp.Substring(0, 7).ToLower() == "function")
                    else if (sTemp.Trim().ToUpper().StartsWith("function ", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //MessageBox.Show(sLine);
                        strFunction sFnTemp = new strFunction();

                        sFnTemp.Name = sTemp;
                        sFnTemp.ModuleName = vbComp.Name;
                        sFnTemp.iLine_Start = iLine;
                        sFnTemp.eType = enumType.eType_Function;
                        sFnTemp.iLine_End = iPreviousFunctionLine;

                        lstFunctions.Add(sFnTemp);

                        iPreviousFunctionLine = iLine - 1;
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

        public bool function_BOF
        {
            get
            {
                try
                {
                    bool bResult;

                    if (iFunctionNumber > 0)
                    { bResult = false; }
                    else
                    { bResult = true; }

                    return bResult;
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

        public bool function_EOF
        {
            get
            {
                try
                {
                    bool bResult;

                    if (iFunctionNumber < lstFunctions.Count)
                    { bResult = false; }
                    else
                    { bResult = true; }

                    return bResult;
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

        public bool function_exists
        {
            get 
            {
                try{
                    bool bResult;

                    if (lstFunctions.Count > 0)
                    { bResult = true; }
                    else
                    { bResult = false; }

                    return bResult;
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

        public void function_Move_First()
        {
            try
            {
                iFunctionNumber = 0;
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

        public void function_Move_Last()
        {
            try
            {
                iFunctionNumber = lstFunctions.Count;
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

        public void function_Move_Next()
        {
            try
            {
                if (iFunctionNumber < lstFunctions.Count)
                { iFunctionNumber++; }
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

        public void function_Move_Previous()
        {
            try
            {
                if (iFunctionNumber > 0)
                { iFunctionNumber--; }
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

        public string function_Name()
        {
            try
            {
                string sResult = function_Name(iFunctionNumber);

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
                return "Error";
            }
        }

        public string function_Name(int iIndex)
        {
            try
            {
                string sResult = lstFunctions[iIndex].Name;

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
                return "Error";
            }
        }

        public string function_TypeName()
        {
            try
            {
                string sResult = function_TypeName(iFunctionNumber);

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
                return "Error";
            }
        }

        public string function_TypeName(int iIndex)
        {
            try
            {
                string sResult;

                switch (lstFunctions[iIndex].eType)
                {
                    case enumType.eType_Function:
                        sResult = "Function";
                        break;
                    case enumType.eType_Property:
                        sResult = "Property";
                        break;
                    case enumType.eType_Sub:
                        sResult = "Sub";
                        break;
                    default:
                        sResult = "Unknown";
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
                return "Error";
            }
        }



        public int function_LineStart()
        {
            try
            {
                int iResult = function_LineStart(iFunctionNumber);

                return iResult;
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
                return -1;
            }
        }

        public int function_LineStart(int iIndex)
        {
            try
            {
                int iResult = lstFunctions[iIndex].iLine_Start;

                return iResult;
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
                return -1;
            }
        }

        public int function_LineEnd()
        {
            try
            {
                int iResult = function_LineEnd(iFunctionNumber);

                return iResult;
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
                return -1;
            }
        }

        public int function_LineEnd(int iIndex)
        {
            try
            {
                int iResult = lstFunctions[iIndex].iLine_End;

                return iResult;
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
                return -1;
            }
        }

        public VBA.VBComponent code_module(string sName) 
        { 
            try 
            {
                VBA.VBComponent vbTemp;

                vbTemp = null;

                foreach (VBA.VBComponent vbComp in vbProj.VBComponents)
                {
                    if (vbComp.Name == sName)
                    {
                        vbTemp = vbComp;
                    }

                }
                return vbTemp;
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
                return null;
            }
        }
    }
}
