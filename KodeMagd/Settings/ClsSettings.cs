using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Misc;
using KodeMagd.InsertCode;

namespace KodeMagd
{
    public class ClsSettings : ApplicationSettingsBase
    {
        private char cDelimiter = '\t';

        public struct strInfo
        {
            public string sValue;
            public DateTime dtFirstUsed;
            public DateTime dtLastUsed;
        }


        [UserScopedSetting()]
        [DefaultSettingValue("213.95.189.213")]
        public string WebsiteAddress
        {
            get { return ((string)this["WebsiteAddress"]); }
            set { this["WebsiteAddress"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("1974")]
        public int WebsitePortNo
        {
            get { return ((int)this["WebsitePortNo"]); }
            set { this["WebsitePortNo"] = (int)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("OK")]
        public string WebsiteConfirmationReply
        {
            get { return ((string)this["WebsiteConfirmationReply"]); }
            set { this["WebsiteConfirmationReply"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("4")]
        public int IndentSize
        {
            get { return ((int)this["IndentSize"]); }
            set { this["IndentSize"] = (int)value; }
        }
        
        [UserScopedSetting()]
        [DefaultSettingValue("1")]
        public string defaultOptionBase 
        {
            get { return ((string)this["DefaultOptionBase"]); }
            set { this["DefaultOptionBase"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("false")]
        public bool IndentFirstLevel
        {
            get { return ((bool)this["IndentFirstLevel"]); }
            set { this["IndentFirstLevel"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("true")]
        public bool UserTips
        {
            get { return ((bool)this["UserTips"]); }
            set { this["UserTips"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("true")]
        public bool SplitConcatinatedLines
        {
            get { return ((bool)this["SplitConcatinatedLines"]); }
            set { this["SplitConcatinatedLines"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("true")]
        public bool SetFocusActivePane
        {
            get { return ((bool)this["SetFocusActivePane"]); }
            set { this["SetFocusActivePane"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("false")]
        public bool InsertErrorHandlers
        {
            get { return ((bool)this["InsertErrorHandlers"]); }
            set { this["InsertErrorHandlers"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("true")]
        public bool UseWith
        {
            get { return ((bool)this["UseWith"]); }
            set { this["UseWith"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("true")]
        public bool AddReferencesAutomatically
        {
            get { return ((bool)this["AddReferencesAutomatically"]); }
            set { this["AddReferencesAutomatically"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string RecentConnectionStrings
        {
            get { return ((string)this["RecentConnectionStrings"]); }
            set { this["RecentConnectionStrings"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("false")]
        public bool InsertCode_Format_EveryParameterOnNewLine
        {
            get { return ((bool)this["InsertCode_Format_EveryParameterOnNewLine"]); }
            set { this["InsertCode_Format_EveryParameterOnNewLine"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("false")]
        public bool InsertCode_Format_UseMaxCharParLine
        {
            get { return ((bool)this["InsertCode_Format_UseMaxCharParLine"]); }
            set { this["InsertCode_Format_UseMaxCharParLine"] = (bool)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("80")]
        public int InsertCode_Format_CharCutOffPoint
        {
            get { return ((int)this["InsertCode_Format_CharCutOffPoint"]); }
            set { this["InsertCode_Format_CharCutOffPoint"] = (int)value; }
        }

        //FormatCutTextChar
        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string FormatVarDimTypeString
        {
            get { return ((string)this["FormatVarDimTypeString"]); }
            set { this["FormatVarDimTypeString"] = (string)value; }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string DocumentDirectory
        {
            get { return ((string)this["DocumentDirectory"]); }
            set { this["DocumentDirectory"] = (string)value; }
        }

        public ClsCodeMapper.enumVarDimType FormatVarDimType
        {
            get
            {
                string sTemp = this.FormatVarDimTypeString;
                ClsCodeMapper.enumVarDimType eResult = ClsCodeMapper.enumVarDimType.eVarDim_Nothing;

                foreach (ClsCodeMapper.enumVarDimType eTemp in Enum.GetValues(typeof(ClsCodeMapper.enumVarDimType)))
                {
                    if (eTemp.ToString() == sTemp)
                    { eResult = eTemp; }
                }

                return eResult;
            }
            set
            {
                ClsCodeMapper.enumVarDimType eTemp = (ClsCodeMapper.enumVarDimType)value;

                this.FormatVarDimTypeString = eTemp.ToString();
            }
        }

        //enumFormatLineCutMethodology

        [UserScopedSetting()]
        [DefaultSettingValue("")]
        public string FormatLineCutMethodologyString
        {
            get { return ((string)this["FormatLineCutMethodologyString"]); }
            set { this["FormatLineCutMethodologyString"] = (string)value; }
        }

        public ClsInsertCode.enumFormatLineCutMethodology FormatLineCutMethodology
        {
            get {
                string sTemp = this.FormatLineCutMethodologyString;
                ClsInsertCode.enumFormatLineCutMethodology eResult = ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_None;

                foreach (ClsInsertCode.enumFormatLineCutMethodology eTemp in Enum.GetValues(typeof(ClsInsertCode.enumFormatLineCutMethodology)))
                {
                    if (eTemp.ToString() == sTemp)
                    { eResult = eTemp; }
                }

                return eResult; 
            }
            set {
                ClsInsertCode.enumFormatLineCutMethodology eTemp = (ClsInsertCode.enumFormatLineCutMethodology)value;

                this.FormatLineCutMethodologyString = eTemp.ToString();
            }
        }

        [UserScopedSetting()]
        [DefaultSettingValue("TypeLib\tSOFTWARE\\Classes\\TypeLib\tSOFTWARE\\Classes\\Wow6432Node\\TypeLib\tSOFTWARE\\Wow6432Node\\Classes\\TypeLib")]
        public string RegistryDirReferences
        {
            get { return ((string)this["RegistryDirReferences"]); }
            set { this["RegistryDirReferences"] = (string)value; }
        }
        //TypeLib
        //SOFTWARE\\Classes\\TypeLib
        //SOFTWARE\\Classes\\Wow6432Node\\TypeLib
        //SOFTWARE\\Wow6432Node\\Classes\\TypeLib

        public string Indent(int iLevels) 
        {
            try 
            {
                string sTemp;

                if (iLevels > 0)
                { sTemp = new string(' ', iLevels * this.IndentSize); }
                else
                { sTemp = ""; }

                return sTemp;
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

        public List<strInfo> UsedConnectionString
        {
            get
            {
                try
                {
                    string sTemp = this.RecentConnectionStrings;

                    List<strInfo> lstTemp = convertConnectionLst(sTemp);

                    return lstTemp;
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
            set
            {
                try
                {
                    List<strInfo> lstTemp = value;

                    string sTemp = convertConnectionLst(lstTemp);

                    this.RecentConnectionStrings = sTemp;
                    this.Save();
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

        public void removeUsedConnectionString(string sConnectionString)
        {
            try
            {
                List<strInfo> lstTemp = this.UsedConnectionString;

                if (lstTemp.Exists(a => a.sValue.Trim().ToLower() == sConnectionString.Trim().ToLower()))
                {
                    int iIndex = lstTemp.FindIndex(a => a.sValue.Trim().ToLower() == sConnectionString.Trim().ToLower());

                    lstTemp.RemoveAt(iIndex);
                }

                this.UsedConnectionString = lstTemp;

                this.Save();
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

        public void addUsedConnectionString(string sConnectionString)
        {
            try
            {
                List<strInfo> lstTemp = this.UsedConnectionString;

                if (lstTemp.Any(a => a.sValue == sConnectionString))
                {
                    int iIndex = lstTemp.FindIndex(a => a.sValue == sConnectionString);

                    strInfo objTemp = new strInfo();

                    objTemp.sValue = sConnectionString;
                    objTemp.dtFirstUsed = lstTemp[iIndex].dtFirstUsed;
                    objTemp.dtLastUsed = DateTime.Now;

                    lstTemp.RemoveAt(iIndex);
                    lstTemp.Add(objTemp);
                }
                else
                {
                    strInfo objTemp = new strInfo();

                    objTemp.sValue = sConnectionString;
                    objTemp.dtFirstUsed = DateTime.Now;
                    objTemp.dtLastUsed = DateTime.Now;

                    lstTemp.Add(objTemp);
                }

                this.UsedConnectionString = lstTemp;

                this.Save();
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

        public static string convertConnectionLst(List<strInfo> lst)
        {
            try
            {
                string sResult = "";

                foreach (strInfo objTemp in lst)
                {
                    sResult += csTagItemBegin;

                    sResult += csTagValueBegin;
                    sResult += objTemp.sValue;
                    sResult += csTagValueEnd;

                    sResult += csTagStartDateBegin;
                    sResult += objTemp.dtFirstUsed;
                    sResult += csTagStartDateEnd;

                    sResult += csTagLastDateBegin;
                    sResult += objTemp.dtLastUsed;
                    sResult += csTagLastDateEnd;

                    sResult += csTagItemEnd;
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

        public const string csTagItemBegin = "\t<item>\t";
        public const string csTagItemEnd = "\t</item>\t";
        public const string csTagValueBegin = "\t<value>\t";
        public const string csTagValueEnd = "\t</value>\t";
        public const string csTagStartDateBegin = "\t<date first used>\t";
        public const string csTagStartDateEnd = "\t</date first used>\t";
        public const string csTagLastDateBegin = "\t<date last used>\t";
        public const string csTagLastDateEnd = "\t</date last used>\t";

        public static List<strInfo> convertConnectionLst(string sText)
        {
            try
            {
                List<strInfo> lstResult = new List<strInfo>();
                bool bIsEnd = false;

                if (!string.IsNullOrEmpty(sText))
                {
                    string sTemp = sText;

                    while (!bIsEnd)
                    {
                        strInfo objTemp = new strInfo();

                        if (sTemp.Contains(csTagItemBegin)
                            & sTemp.Contains(csTagItemEnd)
                            & sTemp.Contains(csTagValueBegin)
                            & sTemp.Contains(csTagValueEnd)
                            & sTemp.Contains(csTagStartDateBegin)
                            & sTemp.Contains(csTagStartDateEnd)
                            & sTemp.Contains(csTagLastDateBegin)
                            & sTemp.Contains(csTagLastDateEnd))
                        {
                            int iPosItemBegin = sTemp.IndexOf(csTagItemBegin);
                            int iPosItemEnd = sTemp.IndexOf(csTagItemEnd);

                            string sItem = sTemp.Substring(iPosItemBegin + csTagItemBegin.Length, iPosItemEnd - (iPosItemBegin + csTagItemBegin.Length));

                            int iPosValueBegin = sItem.IndexOf(csTagValueBegin);
                            int iPosValueEnd = sItem.IndexOf(csTagValueEnd);
                            int iPosStartDateBegin = sItem.IndexOf(csTagStartDateBegin);
                            int iPosStartDateEnd = sItem.IndexOf(csTagStartDateEnd);
                            int iPosLastDateBegin = sItem.IndexOf(csTagLastDateBegin);
                            int iPosLastDateEnd = sItem.IndexOf(csTagLastDateEnd);


                            string sValue = sItem.Substring(iPosValueBegin + csTagValueBegin.Length, iPosValueEnd - (iPosValueBegin + csTagValueBegin.Length));
                            string sStartDate = sItem.Substring(iPosStartDateBegin + csTagStartDateBegin.Length, iPosStartDateEnd - (iPosStartDateBegin + csTagStartDateBegin.Length));
                            string sLastDate = sItem.Substring(iPosLastDateBegin + csTagLastDateBegin.Length, iPosLastDateEnd - (iPosLastDateBegin + csTagLastDateBegin.Length));

                            sTemp = ClsMiscString.Right(ref sTemp, sTemp.Length - (iPosItemEnd + csTagItemEnd.Length));

                            DateTime dtFirstDate;
                            DateTime dtLastDate;

                            if (!DateTime.TryParse(sStartDate, out dtFirstDate))
                            { dtFirstDate = DateTime.MinValue; }

                            if (!DateTime.TryParse(sLastDate, out dtLastDate))
                            { dtLastDate = DateTime.MinValue; }


                            objTemp.sValue = sValue;
                            objTemp.dtFirstUsed = dtFirstDate;
                            objTemp.dtLastUsed = dtLastDate;

                            lstResult.Add(objTemp);

                            if (string.IsNullOrEmpty(sTemp))
                            { bIsEnd = true; }
                        }
                        else
                        { bIsEnd = true; }
                    }
                }

                return lstResult;
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

        public void addRegistryDirForReferences(string sDir) 
        {
            try
            {
                List<string> lstTemp = new List<string>();

                string sRawValue = RegistryDirReferences;

                lstTemp = sRawValue.Split(cDelimiter).ToList();

                lstTemp.Add(sDir);

                lstTemp = lstTemp.Distinct().ToList();

                RegistryDirReferences = ClsMisc.joinStrings(lstTemp, cDelimiter);
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

        public void removeRegistryDirForReferences(string sDir)
        {
            try
            {
                List<string> lstTemp = new List<string>();

                string sRawValue = RegistryDirReferences;

                lstTemp = sRawValue.Split(cDelimiter).ToList();

                if (lstTemp.Contains(sDir, StringComparer.OrdinalIgnoreCase))
                {
                    int iIndex = lstTemp.FindIndex(s => s.Equals(sDir, StringComparison.OrdinalIgnoreCase));
                    
                    lstTemp.RemoveAt(iIndex);

                    lstTemp = lstTemp.Distinct().ToList();

                    RegistryDirReferences = ClsMisc.joinStrings(lstTemp, cDelimiter);
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

        public List<string> registryDirForReferences
        {
            get
            {
                try
                {
                    List<string> lstResults = new List<string>();

                    string sRawValue = this.RegistryDirReferences;

                    lstResults = sRawValue.Split(cDelimiter).ToList();

                    lstResults = lstResults.Distinct().ToList();

                    return lstResults;
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
                    return new List<string>();
                }
            }
        }
    }
}
