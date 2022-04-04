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


namespace KodeMagd.Misc
{
    public class ClsReferences
    {
        public enum enumFilterType 
        {
            eFilt_Outlook,
            eFilt_Access,
            eFilt_ADO,
            eFilt_Scripting,
            eFilt_None
        }

        public struct strAsssembly
        {
            public string sName;
            public string sVersion;
            public string sPath;
            public string sGUID;
            public string sWinXX;
        }

        public struct strRegistryAssemblies
        {
            public string sGUID;
            public string sPath;
            public string sAssembly;
            public string sClass;
            public string sRuntimeVersion;
        }

        public struct strRegistryAssembliesOverview //dao, Version=10.0.4504.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35
        {
            public string sName;
            public string sVersion;
            public string sCulture;
            public string sPublicKeyToken;
        }

        private List<strAsssembly> lstAssembliesTypeLib;
        private List<string> lstRegistryKeysParent;
        private List<strRegistryAssemblies> lstAssemblies;

        public ClsReferences(enumFilterType eFilterType, ref StatusStrip ss)
        {
            try
            {
                lstRegistryKeysParent = new List<string>();
                lstAssemblies = new List<strRegistryAssemblies>();
                lstAssembliesTypeLib = new List<strAsssembly>();

                addRegistriesFromSettings(ref ss);

                //addRegisteryKeyParentDir("TypeLib", ref ss);
                //addRegisteryKeyParentDir("SOFTWARE\\Classes\\TypeLib", ref ss);
                //addRegisteryKeyParentDir("SOFTWARE\\Classes\\Wow6432Node\\TypeLib", ref ss);
                //addRegisteryKeyParentDir("SOFTWARE\\Wow6432Node\\Classes\\TypeLib", ref ss);

                //findAssemblies();
                //findAssembliesOverview();
                findLibType(eFilterType, ref ss);
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

        ~ClsReferences()
        {
            try
            {
                lstRegistryKeysParent = null;
                lstAssemblies = null;
                lstAssembliesTypeLib = null;
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

        public void addRegistriesFromSettings(ref StatusStrip ss)
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                List<string> lstRegistriesDir = cSettings.registryDirForReferences;

                lstRegistriesDir = lstRegistriesDir.Distinct().ToList();

                foreach(string sRegistryDir in lstRegistriesDir)
                { addRegisteryKeyParentDir(sRegistryDir, ref ss); }
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

        private void addRegisteryKeyParentDir(string sKey, ref StatusStrip ss)
        {
            try
            {
                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                lstRegistryKeysParent.Add(sKey);

                lstRegistryKeysParent = lstRegistryKeysParent.Distinct().ToList();
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

        public List<strRegistryAssemblies> assemblies
        {
            get
            {
                try
                {
                    return lstAssemblies;
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

                    return new List<strRegistryAssemblies>();
                }
            }
        }

        public List<strAsssembly> assembliesTypeLib
        {
            get
            {
                try
                {
                    return lstAssembliesTypeLib;
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

                    return new List<strAsssembly>();
                }
            }
        }

        public void addReference(ref VBA.VBProject vbProj, Guid objGuid, ref bool bIsOk, ref string sMessage)
        {
            string sGuid = objGuid.ToString();  

            if (lstAssemblies.Exists(x => x.sGUID == sGuid))
            {
                //int iPosAss = lstAssemblies.FindIndex(x => x.sGUID == sGuid);
                strRegistryAssemblies objAss = lstAssemblies.Find(x => x.sGUID == sGuid);

                string sAss = objAss.sAssembly;

                List<string> lstAss = sAss.Split(',').ToList();


            }
            else
            {
            
            }
            //vbProj.References.AddFromGuid(sGuid, , );
        }


        /*
        GUID (name = GUID)
        --> 1.0 (Name = Version, Value = Name)
        --> --> 0
        --> --> --> win32 (Value = path)
        --> --> --> win64 (Value = path)
        --> --> FLAGS
        --> --> HELPDIR
         */

        public void findLibType(enumFilterType eFilter, ref StatusStrip ss)
        {
            try
            {

                /*
                HKEY_CLASSES_ROOT\TypeLib\{420B2830-E718-11CF-893D-00A0C9054228}\1.0
                HKEY_CLASSES_ROOT\Wow6432Node\TypeLib\{420B2830-E718-11CF-893D-00A0C9054228}\1.0
                HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib\{420B2830-E718-11CF-893D-00A0C9054228}\1.0
                HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Wow6432Node\TypeLib\{420B2830-E718-11CF-893D-00A0C9054228}\1.0
                HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Classes\TypeLib\{420B2830-E718-11CF-893D-00A0C9054228}\1.0

                 */

                /*
                 * list
                 * 1) Name
                 * 2) Version
                 * 3) Path
                 * 4) GUID
                 */

                lstAssembliesTypeLib.Clear();

                //ss.Refresh();
                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                foreach (string sMaterFolder in lstRegistryKeysParent)
                {
                    //ss.Refresh();
                    ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                    RegistryKey keyMaster = Registry.LocalMachine.OpenSubKey(sMaterFolder);
                    if (keyMaster != null)
                    {
                        foreach (string sKeyGUID in keyMaster.GetSubKeyNames())
                        {
                            //ss.Refresh();
                            ClsDefaults.changeStatusStrip_ProgressBar(ref ss);
                            
                            RegistryKey keyGUID = keyMaster.OpenSubKey(sKeyGUID);
                            if (keyGUID != null)
                            {
                                List<string> lstGUIDSubKeys = keyGUID.GetSubKeyNames().ToList();

                                //error cantains multiple keys one for each version
                                //foreach (string sVersion in lstGUID)

                                foreach (string sVersion in lstGUIDSubKeys)
                                {
                                    RegistryKey keyVersion = keyGUID.OpenSubKey(sVersion);
                                    if (keyVersion != null)
                                    {
                                        List<string> lstVersionSubKeys = keyVersion.GetSubKeyNames().ToList();
                                        List<string> lstVersionValueNames = keyVersion.GetValueNames().ToList();

                                        //if (lstVersion.Contains("0") & lstVersion.Contains("FLAGS") & lstVersion.Contains("HELPDIR"))
                                        if (lstVersionSubKeys != null)
                                        {
                                            if (lstVersionSubKeys.Contains("0")) //HELPDIR is optional so assume FLAGS is optional as well
                                            {
                                                RegistryKey key0 = keyVersion.OpenSubKey("0");
                                                if (key0 != null)
                                                {
                                                    List<string> lst0 = key0.GetSubKeyNames().ToList();

                                                    foreach (string sWinXX in lst0)
                                                    {
                                                        RegistryKey keyWinXX = key0.OpenSubKey(sWinXX);
                                                        if (keyWinXX != null)
                                                        {
                                                            List<string> lstWinXXValueNames = keyWinXX.GetValueNames().ToList();

                                                            strAsssembly objAssembly = new strAsssembly();

                                                            /*
                                                             * If keyWinXX contains all this shit then it's different
                                                             * 
                                                             ?keyWinXX.GetValueNames();
                                                            {string[2]}
                                                                [0]: ""
                                                                [1]: "PrimaryInteropAssemblyName"
                                                            ?keyWinXX.GetValue("PrimaryInteropAssemblyName");
                                                            "ADODB, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
                                                             */

                                                            if (lstVersionValueNames.Contains(""))
                                                            { objAssembly.sName = keyVersion.GetValue("").ToString(); }
                                                            else
                                                            { objAssembly.sName = ""; }

                                                            objAssembly.sVersion = ClsMiscString.lastDelimitedValue(keyVersion.Name, '\\');

                                                            if (lstWinXXValueNames.Contains(""))
                                                            { objAssembly.sPath = keyWinXX.GetValue("").ToString(); }
                                                            else
                                                            { objAssembly.sPath = ""; }

                                                            objAssembly.sGUID = ClsMiscString.lastDelimitedValue(keyGUID.Name, '\\');
                                                            objAssembly.sWinXX = ClsMiscString.lastDelimitedValue(keyWinXX.Name, '\\');

                                                            bool bExclude = false;

                                                            switch (eFilter)
                                                            {
                                                                case enumFilterType.eFilt_Access:
                                                                    if (!objAssembly.sName.ToLower().Contains("Microsoft".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Access".ToLower()))
                                                                    { bExclude = true; }
                                                                    break;
                                                                case enumFilterType.eFilt_ADO:
                                                                    //Microsoft ActiveX Data Objects 6.0 Library
                                                                    if (!objAssembly.sName.ToLower().Contains("Microsoft".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("ActiveX".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Data".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Objects".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Library".ToLower()))
                                                                    { bExclude = true; }
                                                                    break;
                                                                case enumFilterType.eFilt_None:
                                                                    break;
                                                                case enumFilterType.eFilt_Outlook:
                                                                    if (!objAssembly.sName.ToLower().Contains("Microsoft".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Outlook".ToLower()))
                                                                    { bExclude = true; }
                                                                    break;
                                                                case enumFilterType.eFilt_Scripting:
                                                                    if (!objAssembly.sName.ToLower().Contains("Microsoft".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Scripting".ToLower()))
                                                                    { bExclude = true; }
                                                                    if (!objAssembly.sName.ToLower().Contains("Runtime".ToLower()))
                                                                    { bExclude = true; }
                                                                    break;
                                                                default:
                                                                    break;
                                                            }



                                                            if (!bExclude)
                                                            { lstAssembliesTypeLib.Add(objAssembly); }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                lstAssembliesTypeLib = lstAssembliesTypeLib.Distinct().ToList();
                lstAssembliesTypeLib = lstAssembliesTypeLib.OrderBy(x => x.sName).ToList();

                //return lstAssemblies;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message + "\n\r\n\r";

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                //return new List<strRegistryAssemblies>();
            }
        }

        public string stripCurvyBrackets(string sText) 
        {
            try
            {
                string sTemp = sText.Trim();

                if (ClsMiscString.Left(ref sTemp, 1) == "{")
                { sTemp = ClsMiscString.Right(ref sTemp, sTemp.Length - 1); }

                if (ClsMiscString.Right(ref sTemp, 1) == "}")
                { sTemp = ClsMiscString.Left(ref sTemp, sTemp.Length - 1); }

                return sTemp;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message + "\n\r\n\r";

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                return string.Empty;
            }
        }

        public static string referenceStatus(ref VBA.VBProject vbProj, string sGUID, string sName)
        {
            try
            {
                /*
                 Come back to later
                 */


                string sStatus = "";

                foreach (VBA.Reference objRef in vbProj.References)
                {
                    if (sGUID == objRef.Guid)
                    { sStatus = objRef.Name; }
                }

                return sStatus;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message + "\n\r\n\r";

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                return string.Empty;
            }
        }


    }
}
