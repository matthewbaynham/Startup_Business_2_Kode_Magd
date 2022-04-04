using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using IntelliLock.Licensing;
using KodeMagd.Misc;
using MySql.Data.MySqlClient;
using KodeMagd.MessageWebsite;

namespace KodeMagd.License
{
    public partial class FrmLicense : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();

        public FrmLicense()
        {
            try
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;
                this.Text = ClsDefaults.formTitle;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnOpenLicenseFile);
                ClsDefaults.FormatControl(ref btnOpenSynchronisationFileAndMessage);

                ClsDefaults.FormatControl(ref lblStep01);
                ClsDefaults.FormatControl(ref lblStep02);
                ClsDefaults.FormatControl(ref lblStep03);
                ClsDefaults.FormatControl(ref lblStep04);
                ClsDefaults.FormatControl(ref lblStep05);
                ClsDefaults.FormatControl(ref lblStep06);
                ClsDefaults.FormatControl(ref lblStep07);

                ClsDefaults.FormatControl(ref ssStatus);
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

        //private void btnOpenSynchronisationFile_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        bool bIsOk = true;
        //        string sErrorMessage = "";

        //        string sSynchronisationKey = openSynchronisationFile(ref bIsOk , ref sErrorMessage);
        //        string sHardwareID = "";

        //        if (bIsOk)
        //        {
        //            if (sSynchronisationKey == "")
        //            {
        //                bIsOk = false;
        //                sErrorMessage = "Failed to Open Synchronisation File";
        //            }
        //        }

        //        if (bIsOk)
        //        {
        //            sHardwareID = getHardwareID(ref bIsOk, ref sErrorMessage);

        //            if (sHardwareID == "") 
        //            {
        //                bIsOk = false;
        //                sErrorMessage = "Failed to get hardware ID";
        //            }
        //        }
                
        //        if (bIsOk)
        //        {
        //            sendDetailToDB(ref bIsOk, ref sErrorMessage, sSynchronisationKey, sHardwareID);
        //        }

        //        if (bIsOk)
        //        {
        //            MessageBox.Show("Computer is Syncronised with website.", ClsDefaults.formTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        }
        //        else
        //        {
        //            MessageBox.Show(sErrorMessage, ClsDefaults.formTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        }
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

        private void btnOpenLicenseFile_Click(object sender, EventArgs e)
        {
            try
            {
                bool bIsOK = true;
                string sErrorMessage = "";

                //string sResult = "";
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                dlgKeyFile = new OpenFileDialog();

                dlgKeyFile.DefaultExt = "license";
                dlgKeyFile.ShowDialog();

                string sFileName = dlgKeyFile.FileName;

                if (sFileName.Trim() == "")
                {
                    bIsOK = false;
                    sErrorMessage = "Cancelled";
                }

                if (!fso.FileExists(sFileName))
                {
                    bIsOK = false;
                    sErrorMessage = "Can't find file " + sFileName;
                }

                if (bIsOK)
                {
                    FrmLoadingLicense frm = new FrmLoadingLicense(sFileName);

                    frm.ShowDialog();

                    frm = null;

                    /*
                    string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                    UriBuilder uri = new UriBuilder(codeBase);
                    string path = Uri.UnescapeDataString(uri.Path);

                    string sDestinationPath = ClsMisc.getDirectory(path);

                    try
                    {
                        fso.CopyFile(sFileName, sDestinationPath, true);
                    }
                    catch (Exception ex) 
                    {
                        bIsOK = false;
                        sErrorMessage = ex.Message;
                    }
                    */ 
                }

                if (bIsOK)
                {
                    MessageBox.Show("Please close down Excel and reopen.\n\rYou can find the details of your license if you click the left most button on the toolbar labelled \"Kode Magd\"", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(sErrorMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private string openSynchronisationFile(ref bool bIsOK, ref string sErrorMessage)
        {
            try
            {
                string sResult = "";
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                dlgKeyFile = new OpenFileDialog();

                //dlgKeyFile.DefaultExt = "key";
                dlgKeyFile.Filter = "Key Files (*.key)|*.key";
                
                dlgKeyFile.ShowDialog();

                string sFileName = dlgKeyFile.FileName;

                if (sFileName.Trim() == "")
                {
                    bIsOK = false;
                    sErrorMessage = "Cancelled";
                }

                if (!fso.FileExists(sFileName))
                {
                    bIsOK = false;
                    sErrorMessage = "Can't find file " + sFileName;
                }

                try
                {
                    if (fso.GetFile(sFileName).Size < 50)
                    {
                        bIsOK = false;
                        sErrorMessage = "Please check the file " + sFileName + "\n\ris the Synchronisation file you download from the Kode Magd website.";
                    }
                }
                catch (Exception e)
                {
                    bIsOK = false;
                    sErrorMessage = "Couldn't check the file\n\r" 
                        + sFileName + "\n\r" 
                        + e.Message;
                }

                if (bIsOK)
                {
                    //get text from file
                    Scripting.TextStream ts = fso.OpenTextFile(sFileName, Scripting.IOMode.ForReading, false, Scripting.Tristate.TristateMixed);

                    sResult = ts.ReadAll();

                    ts.Close();

                    ts = null;
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

                return "";
            }
        }

        private string getHardwareID(ref bool bIsOK, ref string sErrorMessage)
        {
            try
            {
                return HardwareID.GetHardwareID(true, true, false, true, true, true);
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

        //private void sendDetailToDB_old(ref bool bIsOK, ref string sErrorMessage, string sSynchronisationKey, string sHardwareID) 
        //{
        //    try
        //    {
        //        ADODB.Connection conn = new ADODB.Connection();
        //        ADODB.Command cmd = new ADODB.Command();
        //        ADODB.Parameter parSynchronisationKey = new ADODB.Parameter();
        //        ADODB.Parameter parHardwareID = new ADODB.Parameter();
        //        ADODB.Recordset rstError = new ADODB.Recordset();
        //        ClsWebsiteDetails cWebsiteDetails = new ClsWebsiteDetails();
        //        object objResult = null;
        //        object objParameters = null;

        //        conn.ConnectionString = cWebsiteDetails.connectionString;

        //        try
        //        {
        //            conn.Open();
        //        }
        //        catch 
        //        {
        //            bIsOK = false;
        //            sErrorMessage = "Can't connection to website, please check the connection.";
        //        }

        //        if (bIsOK)
        //        {
        //            try
        //            {
        //                cmd.ActiveConnection = conn;
        //                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc;
        //                cmd.CommandText = "identify_machine";

        //                parHardwareID = cmd.CreateParameter("sKey", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 128, sSynchronisationKey);
        //                cmd.Parameters.Append(parHardwareID);

        //                parSynchronisationKey = cmd.CreateParameter("shardwareid", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1024, sHardwareID);
        //                cmd.Parameters.Append(parSynchronisationKey);

        //                objParameters = cmd.Parameters;

        //                //cmd.Execute(out objResult, ref objParameters, (int)ADODB.CommandTypeEnum.adCmdStoredProc);

        //                rstError.Open(cmd, Type.Missing, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, (int)ADODB.CommandTypeEnum.adCmdStoredProc);

        //                if (!(rstError.BOF && rstError.EOF))
        //                {
        //                    while (!rstError.EOF)
        //                    {
        //                        if (rstError.Fields[0].Value.Trim() != "")
        //                        {
        //                            bIsOK = false;
        //                            sErrorMessage = rstError.Fields[0].Value.trim();
        //                        }
        //                        rstError.MoveNext();
        //                    }
        //                }
        //            }
        //            catch
        //            {
        //                bIsOK = false;
        //                sErrorMessage = "Couldn't update website.";
        //            }
        //        }

        //        conn = null;
        //        cmd = null;
        //        parSynchronisationKey = null;
        //        parHardwareID = null;
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

        //private void sendDetailToDB(ref bool bIsOK, ref string sErrorMessage, string sSynchronisationKey, string sHardwareID)
        //{
        //    try
        //    {
        //        ClsWebsiteDetails cWebsiteDetails = new ClsWebsiteDetails();
        //        MySqlConnection con = new MySqlConnection();
        //        con.ConnectionString = cWebsiteDetails.mysql_connectionString;

        //        try
        //        {
        //            con.Open();
        //        }
        //        catch
        //        {
        //            bIsOK = false;
        //            sErrorMessage = "Can't connection to website, please check the connection.";
        //            sErrorMessage += "\n\r";
        //            sErrorMessage += "Connection String: " + con.ConnectionString;
        //        }

        //        if (bIsOK)
        //        {
        //            try
        //            {
        //                MySqlCommand cmd = new MySqlCommand("call dblic.identify_machine( @sKey , @shardwareid )", con);
        //                MySqlParameter parKey = cmd.CreateParameter();
        //                parKey.DbType = System.Data.DbType.String;
        //                parKey.Direction = System.Data.ParameterDirection.Input;
        //                parKey.MySqlDbType = MySqlDbType.String;
        //                parKey.Size = 128;
        //                parKey.ParameterName = "@sKey";
        //                parKey.Value = sSynchronisationKey;

        //                cmd.Parameters.Add(parKey);

        //                MySqlParameter parHardwareId = cmd.CreateParameter();
        //                parHardwareId.DbType = System.Data.DbType.String;
        //                parHardwareId.Direction = System.Data.ParameterDirection.Input;
        //                parHardwareId.MySqlDbType = MySqlDbType.String;
        //                parHardwareId.Size = 1024;
        //                parHardwareId.ParameterName = "@shardwareid";
        //                parHardwareId.Value = sHardwareID;

        //                cmd.Parameters.Add(parHardwareId);

        //                cmd.ExecuteNonQueryAsync();

        //                parKey = null;
        //                parHardwareId = null;
        //                cmd = null;
        //            }
        //            catch
        //            {
        //                bIsOK = false;
        //                sErrorMessage = "Couldn't update website.";
        //            }
        //        }
        //        con = null;



        //        //ADODB.Connection conn = new ADODB.Connection();
        //        //ADODB.Command cmd = new ADODB.Command();
        //        //ADODB.Parameter parSynchronisationKey = new ADODB.Parameter();
        //        //ADODB.Parameter parHardwareID = new ADODB.Parameter();
        //        //ADODB.Recordset rstError = new ADODB.Recordset();
        //        //object objResult = null;
        //        //object objParameters = null;

        //        //conn.ConnectionString = cWebsiteDetails.connectionString;

        //        //try
        //        //{
        //        //    conn.Open();
        //        //}
        //        //catch
        //        //{
        //        //    bIsOK = false;
        //        //    sErrorMessage = "Can't connection to website, please check the connection.";
        //        //}

        //        //if (bIsOK)
        //        //{
        //        //    try
        //        //    {
        //        //        cmd.ActiveConnection = conn;
        //        //        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc;
        //        //        cmd.CommandText = "identify_machine";

        //        //        parHardwareID = cmd.CreateParameter("sKey", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 128, sSynchronisationKey);
        //        //        cmd.Parameters.Append(parHardwareID);

        //        //        parSynchronisationKey = cmd.CreateParameter("shardwareid", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1024, sHardwareID);
        //        //        cmd.Parameters.Append(parSynchronisationKey);

        //        //        objParameters = cmd.Parameters;

        //        //        //cmd.Execute(out objResult, ref objParameters, (int)ADODB.CommandTypeEnum.adCmdStoredProc);

        //        //        rstError.Open(cmd, Type.Missing, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, (int)ADODB.CommandTypeEnum.adCmdStoredProc);

        //        //        if (!(rstError.BOF && rstError.EOF))
        //        //        {
        //        //            while (!rstError.EOF)
        //        //            {
        //        //                if (rstError.Fields[0].Value.Trim() != "")
        //        //                {
        //        //                    bIsOK = false;
        //        //                    sErrorMessage = rstError.Fields[0].Value.trim();
        //        //                }
        //        //                rstError.MoveNext();
        //        //            }
        //        //        }
        //        //    }
        //        //    catch
        //        //    {
        //        //        bIsOK = false;
        //        //        sErrorMessage = "Couldn't update website.";
        //        //    }
        //        //}

        //        //conn = null;
        //        //cmd = null;
        //        //parSynchronisationKey = null;
        //        //parHardwareID = null;
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

        private void FrmLicense_Load(object sender, EventArgs e)
        {
            try
            {

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

        private void btnOpenSynchronisationFileAndMessage_Click(object sender, EventArgs e)
        {
            try
            {
                bool bIsOk = true;
                string sErrorMessage = "";

                string sSynchronisationKey = openSynchronisationFile(ref bIsOk, ref sErrorMessage);
                string sHardwareID = "";

                if (bIsOk)
                {
                    if (sSynchronisationKey == "")
                    {
                        bIsOk = false;
                        sErrorMessage = "Failed to Open Synchronisation File";
                    }
                }

                if (bIsOk)
                {
                    sHardwareID = getHardwareID(ref bIsOk, ref sErrorMessage);

                    if (sHardwareID == "")
                    {
                        bIsOk = false;
                        sErrorMessage = "Failed to get hardware ID";
                    }
                }

                string sVersionNo = Assembly.GetExecutingAssembly().GetName().Version.ToString();

                if (bIsOk)
                {
                    ClsMessageWebsite.sendInfo(ref bIsOk, ref sErrorMessage, sHardwareID, sSynchronisationKey, sVersionNo);

                    //sendDetailToDB(ref bIsOk, ref sErrorMessage, sSynchronisationKey, sHardwareID);
                }

                if (bIsOk)
                {
                    MessageBox.Show("Computer is Syncronised with website.", ClsDefaults.formTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(sErrorMessage, ClsDefaults.formTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

    }
}
