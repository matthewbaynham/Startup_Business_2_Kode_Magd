using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Reflection;
using Microsoft;
using Microsoft.VisualBasic;
using KodeMagd.Misc;
using Microsoft.Win32;

namespace KodeMagd
{
    class ClsTestStuff
    {
        /*
        public ClsTestStuff() { 
        
        }
        
        public void test_Split() {
            //string sTestString = "matthew:baynham";
            string sTestString = "matthewbaynham";
            const char csDelimiter = ':';

            string[] lstLines = sTestString.Split(csDelimiter);

            foreach (string sTemp in lstLines)
            {
                MessageBox.Show(sTemp);
            }
        }
        
        public void addCode()
        {
            try
            {
                Excel.Application app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Workbook wrk = app.ActiveWorkbook;

                string sName = wrk.Name;

                MessageBox.Show(sName);

                VBA.VBProject vbProj;

                if (wrk.HasVBProject == true)
                {
                    vbProj = wrk.VBProject;

                    foreach (VBA.VBComponent vbComp in vbProj.VBComponents)
                    {
                        string sCompName = vbComp.Name;

                        MessageBox.Show(sCompName);
                        
                        VBA.CodeModule objCode = vbComp.CodeModule;

                        for (int iLine = 0; iLine < objCode.CountOfLines; iLine++) {
                            string sLine = objCode.get_Lines(iLine + 1, 1);
                            
                            //string sLine = objCode.Lines[iLine, 1].ToString();
                            
                            MessageBox.Show(sLine);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(Ex.Message.GetHashCode().ToString());
                //if (Ex.Message.GetHashCode() == 182848584) {
                    string sInstructions;

                    sInstructions = "Please change your security settings.\n" +
                                    "\n" +
                                    "(in English)\n" +
                                    "File\n" +
                                    "Options\n" +
                                    "Trust Center\n" +
                                    "Trust Center Settings...\n" +
                                    "Macro Settings\n" +
                                    "Trust Access to the VBA project object model.";
                    
                    string sHelpPath = "http://office.microsoft.com/en-001/help/enable-or-disable-macros-in-office-documents-HA010031071.aspx";
                    
                    MessageBox.Show(sInstructions,
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Error, 
                                    MessageBoxDefaultButton.Button1, 
                                    0, 
                                    sHelpPath);
                    //}
                MessageBox.Show(text:ex.Message, caption:"Error", buttons:MessageBoxButtons.OK, icon:MessageBoxIcon.Error);
            }
        }
        */
        public void getCursorPosition()
        {
            Excel.Application app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook wrk = app.ActiveWorkbook;

            VBA.VBProject vbProj;

            vbProj = wrk.VBProject;

            foreach (VBA.VBComponent vbComp in vbProj.VBComponents)
            {
                VBA.CodeModule mod = vbComp.CodeModule;
                VBA.CodePane cp = mod.CodePane;

                int iStartLine;
                int iStartColumn;
                int iEndLine;
                int iEndColumn;

                cp.GetSelection(out iStartLine, out iStartColumn, out iEndLine, out iEndColumn);

                string sMessage = "";

                sMessage += vbComp.Name;

                sMessage += " Cursor Position: ";


                sMessage += "StartLine: " + iStartLine.ToString();
                sMessage += "StartColumn: " + iStartColumn.ToString();
                sMessage += "EndLine: " + iEndLine.ToString();
                sMessage += "EndColumn: " + iEndColumn.ToString();

                MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle());

//                VBA.CodePane CP = vbProj.VBComponents();
            }




        }
        /*
        public void topWindow() 
        {
            try 
            {
                string sMessage;

                Excel.Application app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                Excel.Workbook wrk = app.ActiveWorkbook;
                VBA.VBProject vbProj = wrk.VBProject;
                VBA.VBComponent cmpResult = vbProj.VBE.SelectedVBComponent;

                //sMessage = app.Windows.get_Item(app.Windows.Count).Caption;
                //sMessage = app.VBE.ActiveWindow.Caption;
                
                //sMessage = vbProj.VBE.CodePanes.Item(vbProj.VBE.CodePanes.Count).Window.Caption;

                //sMessage = vbProj.VBE.CodePanes.Count.ToString(); // = 0
                //sMessage = app.VBE.CodePanes.Count.ToString(); // = 0
                sMessage = app.VBE.Windows.Count.ToString();

                VBA.Window win = app.VBE.Windows.Item(1);
                sMessage = app.VBE.Windows.Item(1).Caption;
                MessageBox.Show(sMessage);

                if (app.VBE.Windows.Item(1).Type == VBA.vbext_WindowType.vbext_wt_CodeWindow) 
                { 
                
                }
                MessageBox.Show(sMessage);


                //sMessage = app.VBE.ActiveWindow.Caption;
                //MessageBox.Show(sMessage);
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
        
        public void ProjectWindow() 
        {
            Excel.Application app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook wrk = app.ActiveWorkbook;
            VBA.VBProject vbProj = wrk.VBProject;
            VBA.VBComponent cmpResult = vbProj.VBE.SelectedVBComponent;
            ClsTextFileLog cLog = new ClsTextFileLog();
            //VBA.VBComponent vbCompResult;

            bool bIsAllWindowsClosed = true;

            //if (cmpResult == null)
            //{ cLog.LOG("cmpResult == null", "", ""); }
            //else
            //{ cLog.LOG("cmpResult", cmpResult.Name, ""); }

            //app.VBE.ActiveCodePane. 

            foreach (VBA.Window win in vbProj.VBE.Windows)
            {
                if (win.Type == VBA.vbext_WindowType.vbext_wt_CodeWindow & win.WindowState == VBA.vbext_WindowState.vbext_ws_Maximize)
                { 
                    //foreach (VBA.VBComponent cmpResult in win.VBE.)
                    
                     
                }
                
                
                //cLog.LOG(win.Caption, win.Type.ToString(), win.WindowState.ToString());
                if (win.Type == VBA.vbext_WindowType.vbext_wt_CodeWindow & win.WindowState == VBA.vbext_WindowState.vbext_ws_Normal)
                { bIsAllWindowsClosed = false; }


            }

            //cLog.LOG("Finished", "Finished", "Finished");

            //cLog.Close();

        }
        */
        public void lookAtRegistry() 
        { 
            try 
            {
                string sRegKeyParent = @"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\CLSID";

                Microsoft.Win32.RegistryKey regKy = Microsoft.Win32.Registry.LocalMachine.OpenSubKey (sRegKeyParent );
                //(Microsoft.Win32.RegistryKey)

                foreach (string sRegKeyName in regKy.GetSubKeyNames())
                { 
                    Microsoft.Win32.RegistryKey regKyTemp = (Microsoft.Win32.RegistryKey)Microsoft.Win32.Registry.LocalMachine.GetValue(sRegKeyName);

                    string sMessage = regKyTemp.Name;

                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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


        public void getControlType()
        {
            try
            {
                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();
                List<VBA.Forms.Control> lstResult = new List<VBA.Forms.Control>();

                if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                {
//                    foreach (var c in vbComp.Designer.Controls.OfType<ComboBox>()) c.Text = string.Empty;

                    foreach (VBA.Forms.Control c in vbComp.Designer.Controls)
                    {
                        if (c is VBA.Forms.ComboBox) 
                        { MessageBox.Show(c.Name + " is Combobox"); };

                        if (c is VBA.Forms.ListBox)
                        { MessageBox.Show(c.Name + " is ListBox"); };
                    } 

                    foreach (VBA.Forms.Control ctrl in vbComp.Designer.Controls)
                    {
                        bool bIsComboBox = true;

                        try
                        { VBA.Forms.ComboBox cmb = (VBA.Forms.ComboBox)ctrl; }
                        catch
                        { bIsComboBox = false; }

                        bool bIsListBox = true;

                        try
                        { VBA.Forms.ListBox cmb = (VBA.Forms.ListBox)ctrl; }
                        catch
                        { bIsListBox = false; }

                        bool bIsTextBox = true;

                        try
                        { VBA.Forms.TextBox cmb = (VBA.Forms.TextBox)ctrl; }
                        catch
                        { bIsTextBox = false; }

                        bool bIsCommandButton = true;

                        try
                        { VBA.Forms.CommandButton cmb = (VBA.Forms.CommandButton)ctrl; }
                        catch
                        { bIsCommandButton = false; }

                        
                        
                        
                        //ctrl.cre
                        //MessageBox.Show(ctrl.Events.Count.ToString());

                        string sTemp = "Control Name: " + ctrl.Name + "\n\r";

                        sTemp += "Is ComboBox: " + bIsComboBox.ToString() + "\n\r";
                        sTemp += "Is ListBox: " + bIsListBox.ToString() + "\n\r";
                        sTemp += "Is TextBox: " + bIsTextBox.ToString() + "\n\r";
                        sTemp += "Is CommandButton: " + bIsCommandButton.ToString() + "\n\r";

                        MessageBox.Show(sTemp, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        //vbComp.Designer.pro
                        //ctrl.Parent.p
                        //ctrl.GetHashCode;
                        //    ctrl.GetHashCode
                        //MessageBox.Show(ctrl.GetType().ToString());

                        //MessageBox.Show(ctrl.);
                    }
                }

//                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                //VBA.Forms.
                //wrk.VBProject.VBComponents.
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

        public static void excelDialogTest() 
        {
            try
            {
                Excel.Application app = ClsMisc.ActiveApplication();
                //Excel.Dialog dlg = app.Dialogs[Excel.XlBuiltInDialog.xlDialogChartType];
                Excel.Dialog dlg = app.Dialogs[Excel.XlBuiltInDialog.xlDialogVbaMakeAddin];

                //bool bResult = dlg.Show(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //bool bResult = dlg.Show();

                //bool bResult = app.Dialogs[Excel.XlBuiltInDialog.xlDialogChartType].Show();
                //bool bResult = app.Dialogs[Excel.XlBuiltInDialog.xlDialogChartSourceData].Show();
/*
                int iRow = 1;
                int iColumn = 1;

                bool bResult = app.Dialogs[Excel.XlBuiltInDialog.xlDialogTable].Show(iRow, iColumn);
*/

                //Excel.Chart cht = new Excel.Chart();

                //Excel.Chart cht = app.Charts.Add();
                //cht.ChartType = Excel.XlChartType.xlLine;

                //cht.Name
                /*
                Microsoft.Office.Interop.Excel.Dialog
                //bool bResult = app.Dialogs[Excel.XlBuiltInDialog.xlDialogAddChartAutoformat].Show(cht.Name, "");
                bool iTop = false;
                bool iLeft = true;
                bool iBottom = false;
                bool iRight = true;
                
                bool bResult = app.Dialogs[Excel.XlBuiltInDialog.xlDialogCreateNames].Show(iTop, iLeft, iBottom, iRight);
                */

                //Excel.Application app = ClsMisc.ActiveApplication();
                Excel.Range rng = app.InputBox("Select cell(s)", 8); 


                //xlDialogCreateNames top, left, bottom, right


                MessageBox.Show("OK", ClsDefaults.messageBoxTitle());

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

        public static void testAddInCollection()
        {
            try
            {
                Excel.Application app = ClsMisc.ActiveApplication();

                foreach (Excel.AddIn objAddIn in app.AddIns)
                { MessageBox.Show(objAddIn.Name); }

                app = null;
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

        public static void testReferencesCollection()
        {
            try
            {
                Excel.Application app = ClsMisc.ActiveApplication();

                //foreach (VBA.Reference objRef in app.VBE.ActiveVBProject.References)
                //{ MessageBox.Show(objRef.Name); }

                //foreach (VBA.Reference objRef in Microsoft.Build.Tasks.ManagedCompiler)
                //{ MessageBox.Show(objRef.Name); }

                List<string> lstGUID = new List<string>();
                List<string> lstRef = new List<string>();
                List<string> lstAss = new List<string>();

                foreach (Assembly ass in AppDomain.CurrentDomain.GetAssemblies())
                {
                    //lstAss.Add(ass.GetName().ToString());
                    lstAss.Add(ass.ManifestModule.ScopeName);
                    foreach (Type t in ass.GetTypes())
                    { 
                        lstGUID.Add(t.GUID.ToString());
                        lstRef.Add(t.AssemblyQualifiedName);
                    }                
                }

                lstGUID = lstGUID.Distinct().ToList();
                lstRef = lstRef.Distinct().ToList();
                lstAss = lstAss.Distinct().ToList();

                foreach (string sTemp in lstAss)
                { MessageBox.Show(sTemp, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }

                app = null;
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

        //public struct strRegistryAssemblies 
        //{
        //    public string sPath;
        //    public string sAssembly;
        //    public string sClass;
        //    public string sRuntimeVersion;
        //}

        //public static void loopThroughRegisterAssemblies() 
        //{
        
        //    List<string> lstMasterFolders = new List<string>();
        //    List<strRegistryAssemblies> lstItemsFound = new List<strRegistryAssemblies>();

        //    //lstMasterFolders.Add("HKCR\\CLSID\\{myGUID}\\InprocServer32");
        //    //lstMasterFolders.Add("Computer\\HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Classes\\CLSID");
        //    //lstMasterFolders.Add("HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Classes\\CLSID");
        //    lstMasterFolders.Add("SOFTWARE\\Wow6432Node\\Classes\\CLSID");
            
        //    foreach (string sMaterFolder in lstMasterFolders)
        //    {
        //        RegistryKey keyMaster = Registry.LocalMachine.OpenSubKey(sMaterFolder);

        //        foreach (string sKeyChild in keyMaster.GetSubKeyNames())
        //        {
        //            //string sKeyChildFull = keyMaster.Name + "\\" + sKeyChild;

        //            RegistryKey keyChild = keyMaster.OpenSubKey(sKeyChild);

        //            List<string> lstChildKey = keyChild.GetSubKeyNames().ToList();

        //            //InprocServer32 key exists
        //            if (lstChildKey.Contains("InprocServer32"))
        //            {
        //                RegistryKey keyInprocServer32 = keyChild.OpenSubKey("InprocServer32");

        //                //List lstInprocServer32 = keyInprocServer32.GetSubKeyNames().ToList();

        //                strRegistryAssemblies objAssembly = new strRegistryAssemblies();

        //                List<string> lstInprocServer32ValueNames = keyInprocServer32.GetValueNames().ToList();

        //                if (lstInprocServer32ValueNames.Contains("") & lstInprocServer32ValueNames.Contains("Assembly") & lstInprocServer32ValueNames.Contains("Class") & lstInprocServer32ValueNames.Contains("RuntimeVersion"))
        //                {
        //                    objAssembly.sPath = Convert.ToString(keyInprocServer32.GetValue(""));
        //                    objAssembly.sAssembly = Convert.ToString(keyInprocServer32.GetValue("Assembly"));
        //                    objAssembly.sClass = Convert.ToString(keyInprocServer32.GetValue("Class"));
        //                    objAssembly.sRuntimeVersion = Convert.ToString(keyInprocServer32.GetValue("RuntimeVersion"));

        //                    lstItemsFound.Add(objAssembly);
        //                }
        //            }

        //            //inside InprocServer32
        //            //(Default) contains path address
        //            //Assembly contains comma delimited string
        //            //Class contains object name
        //            //RuntimeVersion
        //        }
        //    }

        //    MessageBox.Show("Finished - Items found " + lstItemsFound.Count.ToString());

        //    //Must have subfolder "InprocServer32"
        //    //Must have GUID


        //    //System.ComponentModel.

        //        /*
        //         * pick a few directories in the registry to loop through
        //         * loop through directories and any folder that matches a set of criteria will be noted
        //         */

        
        //}

    }
}
