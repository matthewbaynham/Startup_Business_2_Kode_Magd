using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using IntelliLock.Licensing;
using System.Windows.Forms;

namespace KodeMagd.License
{
    class ClsIntellilock
    {
        public bool locked 
        {
            get 
            {
                try 
                {
                    bool bResult;

                    if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.EvaluationMode)
                    { bResult = false; }
                    else if (IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.Licensed && hardwareID() == EvaluationMonitor.CurrentLicense.HardwareID)
                    { bResult = false; }
                    else
                    {
                        string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                        UriBuilder uri = new UriBuilder(codeBase);
                        string sPath = Uri.UnescapeDataString(uri.Path);
                        if (sPath.ToUpper().StartsWith("C:/visual studio 2010/Projects/KodeMagd/KodeMagd/bin/Debug".ToUpper()))
                        { bResult = false; }
                        else
                        { bResult = true; }
                    
                    }

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

        public static string hardwareID()
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

                return string.Empty;
            }
        }

        public void status(ref string sResult)
        {
            try
            {
                List<string> lstTemp = new List<string>();

                this.status(ref lstTemp);
                sResult = "";

                foreach (string sTemp in lstTemp)
                { sResult += sTemp + "\n\r"; }

                while (sResult.EndsWith("\n\r"))
                { sResult = sResult.Substring(0, sResult.Length - 2); }
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

        public void status(ref ListBox txtResult)
        {
            try
            {
                List<string> lstTemp = new List<string>();

                this.status(ref lstTemp);
                txtResult.Items.Clear();

                foreach (string sTemp in lstTemp)
                { txtResult.Items.Add(sTemp); }
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

        public void status(ref List<string> lstResult)
        {
            try
            {
                lstResult = new List<string>();
                /* Check first if a valid license file is found */
                if (EvaluationMonitor.CurrentLicense.LicenseStatus == IntelliLock.Licensing.LicenseStatus.Licensed)
                {
                    /* Read additional license information */
                    for (int iCounter = 0; iCounter < EvaluationMonitor.CurrentLicense.LicenseInformation.Count; iCounter++)
                    {
                        string sKey = EvaluationMonitor.CurrentLicense.LicenseInformation.GetKey(iCounter).ToString();
                        string sValue = EvaluationMonitor.CurrentLicense.LicenseInformation.GetByIndex(iCounter).ToString();

                        lstResult.Add(sKey + ": " + sValue);
                    }
                }

                if (EvaluationMonitor.CurrentLicense.TrialRestricted == true)
                { lstResult.Add("TrialRestricted"); }
                else
                { lstResult.Add("Not TrialRestricted"); }

                lstResult.Add("Status: " + EvaluationMonitor.CurrentLicense.LicenseStatus.ToString());

                if (EvaluationMonitor.CurrentLicense.ExpirationDays_Enabled == true)
                { lstResult.Add("Days elapsed " + EvaluationMonitor.CurrentLicense.ExpirationDays_Current.ToString() + " out of " + EvaluationMonitor.CurrentLicense.ExpirationDays.ToString()); }

                foreach (Object objInfo in EvaluationMonitor.CurrentLicense.LicenseInformation)
                { lstResult.Add(objInfo.ToString()); }
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

        public int daysLeft
        {
            get 
            {
                try
                {
                    return EvaluationMonitor.CurrentLicense.ExpirationDays_Current;
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

                    return 0;
                }
            }
        }

        public bool isPaidFor
        {
            get 
            {
                try
                {
                    return IntelliLock.Licensing.EvaluationMonitor.CurrentLicense.LicenseStatus == LicenseStatus.Licensed;
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
        }

    }
}
