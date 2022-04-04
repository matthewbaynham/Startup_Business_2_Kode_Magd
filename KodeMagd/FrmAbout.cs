using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Misc;
using KodeMagd.InsertCode;
using KodeMagd.Settings;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Diagnostics;
using KodeMagd.License;
using IntelliLock.Licensing;
using System.Security.Principal;
using System.Security.Permissions;

namespace KodeMagd
{
    public partial class FrmAbout : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();

        public FrmAbout()
        {
            try
            {
                InitializeComponent();
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

        private void FrmAbout_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref pnlKodeMagdImage);
                ClsDefaults.FormatControl(ref lblKodeMagd);
                ClsDefaults.FormatControl(ref rtbInfo);

                ClsDefaults.FormatControl(ref lblLicenseID);
                ClsDefaults.FormatControl(ref txtLicenseID);
                ClsDefaults.FormatControl(ref lblMachineID);
                ClsDefaults.FormatControl(ref txtMachineID);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(pnlKodeMagdImage, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(rtbInfo, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblLicenseID, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtLicenseID, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblMachineID, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtMachineID, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                ClsCodeMapper cCodeMapper = new ClsCodeMapper();

                cCodeMapper.readCode();

                string sResult = "";

                if (cCodeMapper.cursorIsInFunction)
                { 
                    sResult = "In Function: ";
                    sResult += cCodeMapper.cursorInFunctionName;
                }
                else
                { sResult = "Not In Function"; }

                fillRtfAbout();
                fillIDs();

                lblVersion.Text = "Version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
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

        private void btnTestStuff_Click(object sender, EventArgs e)
        {
            try
            {
                //Create a temp file
                string sTempFilePath = ClsMisc.getTempDirectory();

                if (!sTempFilePath.EndsWith("\\"))
                { sTempFilePath += "\\"; }

                sTempFilePath += ClsMisc.getRandomNewFileName(sTempFilePath, ".bat");

                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();
                Scripting.TextStream ts = fso.CreateTextFile(sTempFilePath, true, true);

                String sCmdExePath = System.Environment.GetEnvironmentVariable("COMSPEC", EnvironmentVariableTarget.Machine);

                ts.WriteLine("@echo off");
                ts.WriteLine("");
                ts.WriteLine(":: BatchGotAdmin");
                ts.WriteLine(":-------------------------------------");
                ts.WriteLine("REM  --> Check for permissions");
                ts.WriteLine(">nul 2>&1 \"%SYSTEMROOT%\\system32\\cacls.exe\" \"%SYSTEMROOT%\\system32\\config\\system\"");
                ts.WriteLine("");
                ts.WriteLine("REM --> If error flag set, we do not have admin.");
                ts.WriteLine("if '%errorlevel%' NEQ '0' (");
                ts.WriteLine("    echo Requesting administrative privileges...");
                ts.WriteLine("    goto UACPrompt");
                ts.WriteLine(") else ( goto gotAdmin )");
                ts.WriteLine("");
                ts.WriteLine(":UACPrompt");
                ts.WriteLine("    echo Set UAC = CreateObject^(\"Shell.Application\"^) > \"%temp%\\getadmin.vbs\"");
                ts.WriteLine("    set params = %*:\"=\"\"");
                ts.WriteLine("    echo UAC.ShellExecute \"" + sCmdExePath + "\", \"/c %~s0 %params%\", \"\", \"runas\", 1 >> \"%temp%\\getadmin.vbs\"");
                ts.WriteLine("");
                ts.WriteLine("    \"%temp%\\getadmin.vbs\"");
                ts.WriteLine("    del \"%temp%\\getadmin.vbs\"");
                ts.WriteLine("    exit /B");
                ts.WriteLine("");
                ts.WriteLine(":gotAdmin");
                ts.WriteLine("    pushd \"%CD%\"");
                ts.WriteLine("    CD /D \"%~dp0\"");
                ts.WriteLine("");
                ts.WriteLine(":--------------------------------------");
                ts.WriteLine("");
                ts.WriteLine("copy \"C:\\Misc\\testfile.txt\" \"C:\\Program Files (x86)\\Adobe\\\"");




                /*
                var wi = WindowsIdentity.GetCurrent();
                var wp = new WindowsPrincipal(wi);
 
                if (!wp.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    // It is not possible to launch a ClickOnce app as administrator directly, so instead we launch the
                    // app as administrator in a new process.
                    var processInfo = new ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase);
 
                    // The following properties run the new process as administrator
                    processInfo.UseShellExecute = true;
                    processInfo.Verb = "runas";
                    processInfo.
                    // Start the new process
                    try
                    { Process.Start(processInfo); }
                    catch (Exception)
                    {
                        // The user did not allow the application to run as administrator
                        MessageBox.Show("Sorry, this application must be run as Administrator.");
                    }
 
                    // Shut down the current process
                    //Process.Shutdown();
                }
                else
                {
                    // We are running as administrator
 
                    // Do normal startup stuff...
                }
                */

                /*
                var wp = WindowsPrincipal(WindowsIdentity.GetCurrent());

                if (!WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
                    {
                    // It is not possible to launch a ClickOnce app as administrator directly, so instead we launch the
                    // app as administrator in a new process.
                    var processInfo = new ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase);

                    // The following properties run the new process as administrator
                    processInfo.UseShellExecute = true;
                    processInfo.Verb = "runas";

                    // Start the new process
                    try
                    {Process.Start(processInfo); }
                    catch (Exception)
                    {
                        // The user did not allow the application to run as administrator
                        MessageBox.Show("Sorry, this application must be run as Administrator.");
                    }

                    // Shut down the current process
                    Application.Current.Shutdown();
                }
                else
                {
                // We are running as administrator

                // Do normal startup stuff...
                }
                 
                */ 
                 



                /*
                WindowsPrincipal pricipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
                bool hasAdministrativeRight = pricipal.IsInRole(WindowsBuiltInRole.Administrator);
                if (!hasAdministrativeRight)
                {
                    // relaunch the application with admin rights
                    string fileName = Assembly.GetExecutingAssembly().Location;
                    ProcessStartInfo processInfo = new ProcessStartInfo();
                    processInfo.Verb = "runas";
                    processInfo.FileName = fileName;
 
                    try
                    {
                        Process.Start(processInfo);
                    }
                    catch (Win32Exception)
                    {
                        // This will be thrown if the user cancels the prompt
                    }
 
                    return;
                }                 
                */

                /*
                 
                    process.Start(...);


                    process.StandardInput.WriteLine("Dir xxxxx");
                    process.StandardInput.WriteLine("Dir yyyyy");
                    process.StandardInput.WriteLine("Dir zzzzzz");
                    process.StandardInput.WriteLine("other command(s)");
                 */

                /*

                System.Diagnostics.Process myProcess = new System.Diagnostics.Process();

                //myProcess.StartInfo = @"C:\MyScript.bat";

                //myProcess.StartInfo.UseShellExecute = true;
                //myProcess.StartInfo.RedirectStandardInput = true;
                //myProcess.StartInfo.RedirectStandardOutput = true;

                String sCmdExePath = System.Environment.GetEnvironmentVariable("COMSPEC", EnvironmentVariableTarget.Machine);


                myProcess.StartInfo = new System.Diagnostics.ProcessStartInfo(sCmdExePath);
                myProcess.StartInfo.Arguments = "";//String.Format(@"/c g++ ""C:\Alps\{0}\Debug\Main.cpp""", project_name);
                //myProcess.StartInfo.WorkingDirectory = "\"C:\\Program Files (x86)\\Adobe\\\"";
                myProcess.StartInfo.WorkingDirectory = "\"C:\\Misc\\\"";
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.StartInfo.ErrorDialog = true;
                myProcess.StartInfo.FileName = sCmdExePath;

                myProcess.StartInfo.UseShellExecute = false;
                myProcess.StartInfo.RedirectStandardInput = true;

                myProcess.Start();

                myProcess.StandardInput.WriteLine("copy \"C:\\Misc\\testfile.txt\" \"C:\\Program Files (x86)\\Adobe\\\"");

                //myProcess.StandardInput.WriteLine("@echo off");

                //:: BatchGotAdmin
                //:-------------------------------------
                //REM  --> Check for permissions
                myProcess.StandardInput.WriteLine(">nul 2>&1 \"%SYSTEMROOT%\\system32\\cacls.exe\" \"%SYSTEMROOT%\\system32\\config\\system\"");

                //REM --> If error flag set, we do not have admin.
                myProcess.StandardInput.WriteLine("if '%errorlevel%' NEQ '0' (");
                myProcess.StandardInput.WriteLine("    echo Requesting administrative privileges...");
                myProcess.StandardInput.WriteLine("    goto UACPrompt");
                myProcess.StandardInput.WriteLine(") else ( goto gotAdmin )");

                myProcess.StandardInput.WriteLine(":UACPrompt");
                myProcess.StandardInput.WriteLine("    echo Set UAC = CreateObject^(\"Shell.Application\"^) > \"%temp%\\getadmin.vbs\"");
                myProcess.StandardInput.WriteLine("    set params = %*:\"=\"\"");
                myProcess.StandardInput.WriteLine("    echo UAC.ShellExecute \"cmd.exe\", \"/c %~s0 %params%\", \"\", \"runas\", 1 >> \"%temp%\\getadmin.vbs\"");

                myProcess.StandardInput.WriteLine("    \"%temp%\\getadmin.vbs\"");
                myProcess.StandardInput.WriteLine("    del \"%temp%\\getadmin.vbs\"");
                myProcess.StandardInput.WriteLine("    exit /B");

                myProcess.StandardInput.WriteLine(":gotAdmin");
                myProcess.StandardInput.WriteLine("    pushd \"%CD%\"");
                myProcess.StandardInput.WriteLine("    CD /D \"%~dp0\"");

                myProcess.StandardInput.WriteLine(":--------------------------------------");


                myProcess.StandardInput.WriteLine("copy \"C:\\Misc\\testfile.txt\" \"C:\\Program Files (x86)\\Adobe\\\"");

                myProcess.Start();

                myProcess.Close();



                */

                /*
                bool bIsOK = true;
                string sErrorMessage = "";

                string sDestinationPath = "C:\\Program Files (x86)\\Adobe\\";
                string sFilePath = "C:\\Misc\\testfile.txt";

                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                try
                {
                    //[assembly: PermissionSetAttribute(SecurityAction.RequestMinimum, Name = "FullTrust")]

                    Assembly.GetExecutingAssembly().PermissionSet.AddPermission(new SecurityPermission(SecurityPermissionFlag.AllFlags));

                    Scripting.Folder fld = fso.GetFolder(sDestinationPath);

                    fso.

                    fso.CopyFile(sFilePath, sDestinationPath, true);
                }
                catch (Exception ex)
                {
                    bIsOK = false;
                    sErrorMessage = ex.Message;
                }

                if (bIsOK)
                { MessageBox.Show("Finished"); }
                else
                { MessageBox.Show(sErrorMessage); }
                */

                /*
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                MessageBox.Show("Assemble location: " + path);
                */


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




        private void lblKodeMagd_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(ClsDefaults.website));
                lblKodeMagd.LinkVisited = true;
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

        private void pnlKodeMagdImage_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(ClsDefaults.website));
                lblKodeMagd.LinkVisited = true;
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

        private void fillIDs()
        {
            try
            {
                txtMachineID.Text = ClsIntellilock.hardwareID();
                txtLicenseID.Text = EvaluationMonitor.CurrentLicense.HardwareID;
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

        private void fillRtfAbout()
        {
            try
            {
                string sRftAbout = "";
                 
                sRftAbout += "{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang2057{\\fonttbl{\\f0\\fdecor\\fprq2\\fcharset0 Imprint MT Shadow;}{\\f1\\fscript\\fprq2\\fcharset0 Comic Sans MS;}{\\f2\\fnil\\fcharset0 Calibri;}}";
                sRftAbout += "{\\colortbl ;\\red255\\green0\\blue0;}";
                sRftAbout += "{\\*\\generator Msftedit 5.41.21.2510;}\\viewkind4\\uc1\\pard\\sl240\\slmult1\\qc\\f0\\fs24  \\fs32 Kode Magd\\par";
                sRftAbout += "\\pard\\sl240\\slmult1\\qj\\f1\\fs20 This software is protected by copyright.\\par";
                sRftAbout += "\\par";
                sRftAbout += " Developed by:\\par";
                sRftAbout += "\\fs24 Matthew Baynham\\par";
                sRftAbout += " Baynham Coding UG\\par";
                sRftAbout += "\\par";
                sRftAbout += "\\fs20 License Status:\\par";
                if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.EvaluationMode) 
                { sRftAbout += "\\fs24 Trial Period\\par"; }
                else if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.Licensed 
                            && ClsIntellilock.hardwareID() == EvaluationMonitor.CurrentLicense.HardwareID)
                { sRftAbout += "\\fs24 Full License\\par"; }
                else if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.Licensed
                            && ClsIntellilock.hardwareID() != EvaluationMonitor.CurrentLicense.HardwareID)
                {
                    sRftAbout += "\\cf1\\ul\\b\\fs24 LICENSE INVALID\\cf0\\ulnone\\b0\\par";
                    sRftAbout += "\\cf1\\ul\\b\\fs24 License ID does not match machine ID\\cf0\\ulnone\\b0\\par";
                    sRftAbout += "\\cf1\\ul\\b\\fs24 License ID: " + EvaluationMonitor.CurrentLicense.HardwareID + "\\cf0\\ulnone\\b0\\par";
                    sRftAbout += "\\cf1\\ul\\b\\fs24 Machine ID: " + ClsIntellilock.hardwareID() + "\\cf0\\ulnone\\b0\\par";
                }
                else
                {
                    sRftAbout += "\\cf1\\ul\\b\\fs24 INVALID - LICENSE REQUIRED\\cf0\\ulnone\\b0\\par";
                    sRftAbout += "\\pard\\sl240\\slmult1\\lang9\\f2\\fs22\\par";
                }

                if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.EvaluationMode)
                {
                    //if (EvaluationMonitor.CurrentLicense.TrialRestricted)
                    //{
                        sRftAbout += "\\par";
                        //sRftAbout += "\\fs20 License Type:\\par";
                        //sRftAbout += "\\cf1\\ul\\b\\fs24 Trial Period\\cf0\\ulnone\\b0\\par";
                        //sRftAbout += "\\pard\\sl240\\slmult1\\lang9\\f2\\fs22\\par";

                        if (!EvaluationMonitor.CurrentLicense.ExpirationDays_Enabled)
                        { sRftAbout += "\\cf1\\ul\\b\\fs24 EXPIRED\\cf0\\ulnone\\b0\\par"; }

                        sRftAbout += "\\fs24 " + EvaluationMonitor.CurrentLicense.ExpirationDays_Current.ToString() + " of " + EvaluationMonitor.CurrentLicense.ExpirationDays.ToString() + " Days \\par";
                    //}
                }

                sRftAbout += "\\par";
                 
                /* Check first if a valid license file is found */
                
                if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.EvaluationMode ||
                    (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.Licensed 
                    && ClsIntellilock.hardwareID() == EvaluationMonitor.CurrentLicense.HardwareID))
                {
                    if (EvaluationMonitor.CurrentLicense.LicenseInformation.Count > 0)
                    {
                        sRftAbout += "\\fs20 License Info:\\par";

                        /* Read additional license information */
                        for (int i = 0; i < EvaluationMonitor.CurrentLicense.LicenseInformation.Count; i++)
                        {
                            string sKey = EvaluationMonitor.CurrentLicense.LicenseInformation.GetKey(i).ToString();
                            string sValue = EvaluationMonitor.CurrentLicense.LicenseInformation.GetByIndex(i).ToString();

                            if (sKey == "License Type" && sValue == "Full License")
                            {
                                if (ClsIntellilock.hardwareID() != EvaluationMonitor.CurrentLicense.HardwareID)
                                {
                                    sValue = "Full License but wrong Machine ID";
                                }
                            }
                            
                            sRftAbout += "\\fs20 " + sKey + ":\\par";
                            sRftAbout += "\\fs20 " + sValue + "\\par";

                            sRftAbout += "\\par";
                        }
                    }
                }

                sRftAbout += "\\pard\\sl240\\slmult1\\lang9\\f2\\fs22\\par";
                sRftAbout += "\\par";
                sRftAbout += "}";

                rtbInfo.Rtf = sRftAbout;
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

        private void fillRtfAbout(string sRftAbout)
        {
            try
            {
                rtbInfo.Rtf = sRftAbout;
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

        private void FrmAbout_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref pnlKodeMagdImage);
                cControlPosition.positionControl(ref rtbInfo);

                cControlPosition.positionControl(ref lblLicenseID);
                cControlPosition.positionControl(ref txtLicenseID);
                cControlPosition.positionControl(ref lblMachineID);
                cControlPosition.positionControl(ref txtMachineID);
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
