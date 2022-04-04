using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Misc;

namespace KodeMagd.License
{
    public partial class FrmLoadingLicense : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private string sLicensePath;

        public FrmLoadingLicense(string sPath)
        {
            try
            {
                InitializeComponent();

                sLicensePath = sPath;
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

        private void FrmLoadingLicense_Load(object sender, EventArgs e)
        {
            try
            {
                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref lblTitle, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref txtDosCopy, false, ClsDefaults.enumSpecialEffect.eEff_DosLook);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(lblTitle, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtDosCopy, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                loadLicense();
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

        private void loadLicense()
        {
            try
            {
                if (sLicensePath!="")
                {
                    string sFileName = ClsMisc.getFileName(sLicensePath);

                    bool bCopyFailed = false;
                    string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                    UriBuilder uri = new UriBuilder(codeBase);
                    string path = Uri.UnescapeDataString(uri.Path);

                    string sDestinationPath = ClsMisc.getDirectory(path);
                    sDestinationPath = sDestinationPath.Replace("/", "\\");

                    if (!sDestinationPath.EndsWith("\\"))
                    { sDestinationPath += "\\"; }

                    Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                    string sDosCopyCommand = "copy \"" + sLicensePath + "\" \"" + sDestinationPath + "\"";
                    txtDosCopy.Text = sDosCopyCommand;

                    try
                    { 
                        fso.CopyFile(sLicensePath, sDestinationPath, true);

                        if (fso.FileExists(sDestinationPath + sFileName))
                        {
                            Scripting.File flSource = fso.GetFile(sLicensePath);
                            Scripting.File flDestination = fso.GetFile(sDestinationPath + sFileName);

                            if (flSource.DateCreated != flDestination.DateCreated)
                            { bCopyFailed = true; }
                            else if (flSource.Size != flDestination.Size)
                            { bCopyFailed = true; }
                            else if (flSource.Type != flDestination.Type)
                            { bCopyFailed = true; }
                            //else if (flSource.Attributes != flDestination.Attributes)
                            //{ bCopyFailed = true; }
                        }
                        else
                        { bCopyFailed = true; }
                    }
                    catch(Exception e)
                    { bCopyFailed = true; }

                    if (bCopyFailed)
                    {
                        //Create a temp file
                        string sTempFilePath = ClsMisc.getTempDirectory();

                        if (!sTempFilePath.EndsWith("\\"))
                        { sTempFilePath += "\\"; }

                        sTempFilePath += ClsMisc.getRandomNewFileName(sTempFilePath, ".bat");

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
                        ts.WriteLine(sDosCopyCommand);
                        ts.WriteLine("PAUSE");

                        ts.Close();

                        ts = null;

                        //Run script
                        //System.Diagnostics.Process.Start(sTempFilePath);

                        System.Diagnostics.Process proc = new System.Diagnostics.Process();
                        proc.StartInfo.FileName = sTempFilePath;
                        proc.StartInfo.RedirectStandardError = true;
                        proc.StartInfo.RedirectStandardOutput = true;
                        proc.StartInfo.UseShellExecute = false;


                        proc.Start();
                        proc.WaitForExit();
                        //output1 = proc.StandardError.ReadToEnd();
                        //proc.WaitForExit();
                        //output2 = proc.StandardOutput.ReadToEnd();
                        //proc.WaitForExit();
                    }

                    fso = null;
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

        private void FrmLoadingLicense_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblTitle);
                cControlPosition.positionControl(ref lblWarning);
                cControlPosition.positionControl(ref txtDosCopy);
                cControlPosition.positionControl(ref btnClose);
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

        private void txtDosCopy_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                bool bIsOK = true;

                if (e.Modifiers == Keys.Enter) // || e.Modifiers == Keys.Tab 
                { bIsOK = true; }
                else if (e.KeyValue == 13) // || e.KeyValue == 9)
                { bIsOK = true; }
                else if (e.KeyValue == 37 || e.KeyValue == 38 || e.KeyValue == 39 || e.KeyValue == 40)
                { bIsOK = true; }
                else if (e.Modifiers == Keys.Left || e.Modifiers == Keys.Right || e.Modifiers == Keys.Up || e.Modifiers == Keys.Down)
                { bIsOK = true; }
                else if (e.KeyValue == 17 || e.Modifiers == Keys.Control || e.Modifiers == Keys.ControlKey)
                { bIsOK = true; }
                else if ((e.KeyValue == 67 || e.Modifiers == Keys.C) && e.Control == true)
                { bIsOK = true; }
                else if ((e.KeyValue == 65 || e.Modifiers == Keys.A) && e.Control == true)
                { bIsOK = true; }
                else
                { bIsOK = false; }

                /*
                Debug.Print("\r\n"
                    + "\r\ne.KeyCode: " + e.KeyCode.ToString()
                    + "\r\ne.KeyData: " + e.KeyData.ToString()
                    + "\r\ne.KeyValue: " + e.KeyValue.ToString()
                    + "\r\ne.Modifiers: " + e.Modifiers.ToString()
                    + "\r\ne.Control: " + e.Control.ToString()
                    + "\r\ne.Shift: " + e.Shift.ToString()
                    + "\r\nbIsOK: " + bIsOK.ToString());
                */

                if (!bIsOK)
                { e.SuppressKeyPress = true; }
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
