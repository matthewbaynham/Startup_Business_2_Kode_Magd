using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace KodeMagd.Misc
{
    class ClsMiscString
    {
        /*
        public static int stringCountChar(string sText, char cChar)
        {
            try
            {
                int iResult = 0;
                int iPos = 0;

                //sText.Count(x => x == cChar);

                if (sText.Contains(cChar))
                {
                    while (iPos >= 0)
                    {
                        iPos = sText.IndexOf(cChar, iPos + 1);
                        iResult++;
                    }
                }

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
                return ClsMisc.gciError;
            }
        }
        */

        public static string Left(string sText, int iSize)
        {
            try
            {
                string sResult;

                if (sText == "")
                { sResult = ""; }
                else
                {
                    if (iSize < 0)
                    { sResult = ""; }
                    else if (sText.Length < iSize)
                    { sResult = sText; }
                    else
                    { sResult = sText.Substring(0, iSize); }
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
                return string.Empty;
            }
        }

        public static string Left(ref string sText, int iSize)
        {
            try
            {
                string sResult;

                if (sText == "")
                { sResult = ""; }
                else
                {
                    if (iSize < 0)
                    { sResult = ""; }
                    else if (sText.Length < iSize)
                    { sResult = sText; }
                    else
                    { sResult = sText.Substring(0, iSize); }
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
                return string.Empty;
            }
        }

        public static string Right(string sText, int iSize)
        {
            try
            {
                string sResult;

                if (iSize < 0)
                { sResult = ""; }
                else if (iSize > sText.Length)
                { sResult = sText; }
                else
                { sResult = sText.Substring(sText.Length - iSize, iSize); }

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
                return string.Empty;
            }
        }

        public static string Right(ref string sText, int iSize)
        {
            try
            {
                string sResult;

                if (iSize < 0)
                { sResult = ""; }
                else if (iSize > sText.Length)
                { sResult = sText; }
                else
                { sResult = sText.Substring(sText.Length - iSize, iSize); }

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
                return string.Empty;
            }
        }

        public static string addQuotes(string sText)
        {
            try
            {
                string sTemp;

                if (string.IsNullOrEmpty(sText))
                { sTemp = "\"\""; }
                else
                {
                    string sTwoDoubleQuotes = ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString();
                    string sFourDoubleQuotes = ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString();
                    string sSixDoubleQuotes = ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString() + ClsMisc.gccChar_DoubleQuote.ToString();

                    sTemp = ClsMisc.gccChar_DoubleQuote + sText.Replace(ClsMisc.gccChar_DoubleQuote.ToString(), sTwoDoubleQuotes) + ClsMisc.gccChar_DoubleQuote;
                    //sTemp = ClsMisc.gccChar_DoubleQuote + sTemp.Replace(sFourDoubleQuotes, sSixDoubleQuotes) + ClsMisc.gccChar_DoubleQuote;
                }
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

        public static string makeValidVarName(string sText, string sPrefix)
        {
            try
            {
                string sTemp = makeValidVarName(sPrefix + " " + sText);

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

        public static string makeValidVarName(string sText)
        {
            try
            {
                TextInfo tiTemp = new CultureInfo("").TextInfo;

                string sTemp = tiTemp.ToTitleCase(sText).Trim();

                List<char> lstInvalidChar = new List<char>{' ', '.', '!', '@', '&', '$', '#', '!', '"', '#', '$', '%', '&', 
                                                      '\u0027', '(', ')', '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', 
                                                           '>', '?', '@', '[', '\\', ']', '^', '_', '`', '`', '{', '|', '}', 
                                                           '~', ' ', '¡', '¢', '£', '¤', '¥', '¦', '§', '¨', '©', 'ª', 'ª', 
                                                           '«', '¬', '®', ',', '¯', '°', '±', '²', '³', '´', 'µ', '¶', '·', 
                                                           '¸', '¹', 'º', '»', '¼', '½', '¾', '¿', 'À', 'Á', 'Â', 'ǁ', 'ǂ', 
                                                           'ǀ', '˛'};

                foreach (char cTemp in lstInvalidChar)
                {
                    while (sTemp.Contains(cTemp))
                    {
                        int iPos = sTemp.IndexOf(cTemp);
                        sTemp = sTemp.Remove(iPos, 1);
                    }
                }

                if (sTemp.Length >= 1) 
                {
                    string sLeftChar = ClsMiscString.Left(ref sTemp, 1);
                    string sRemainderOfString = ClsMiscString.Right(ref sTemp, sTemp.Length - 1);

                    sTemp = sLeftChar.ToLower() + sRemainderOfString;
                }

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

        public static string LstToText(List<string> lst, int iNewLineEveryXChar) 
        { 
            try
            {
                string sResult = "";
                int iNewLineStartingPosition = 0;

                for (int iCounter = 0; iCounter < lst.Count ; iCounter++) 
                {
                    if (iCounter == lst.Count - 1 )
                    { sResult += lst[iCounter] + "."; }
                    else if (iCounter == lst.Count - 2)
                    { sResult += lst[iCounter] + " and "; }
                    else
                    { sResult += lst[iCounter] + ", "; }

                    if ((sResult.Length - iNewLineStartingPosition) > iNewLineEveryXChar)
                    { 
                        sResult += Environment.NewLine;
                        iNewLineStartingPosition = sResult.Length;
                    }
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
                return string.Empty;
            }
        }

        public static int indentSize(string sText) 
        {
            try
            {
                int iResult = 0;
                int iPos = 0;
                bool bFinished = false;

                if (string.IsNullOrEmpty(sText))
                { iResult = 0; }
                else
                {
                    while (!bFinished)
                    {
                        if (sText.Substring(iPos, 1) == " ")
                        {
                            iPos++;
                            if (iPos > sText.Length)
                            {
                                bFinished = true;
                                iResult = sText.Length;
                            }
                        }
                        else
                        {
                            bFinished = true;
                            iResult = iPos;
                        }
                    }
                }

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
                return 0;
            }
        }

        public static string replaceMultiLineCharWithVbConst(string sText)
        {
            try
            {
                string sResult = sText;
                sResult = sResult.Replace("\n\r", "\" & vbCrLf & \"");
                sResult = sResult.Replace("\r\n", "\" & vbCrLf & \"");
                sResult = sResult.Replace("\n", "\" & vbCr & \"");
                sResult = sResult.Replace("\r", "\" & vbLf & \"");

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
                return string.Empty;
            }
        }

        public static string lastDelimitedValue(string sText, char cDemlimiter) 
        {
            try
            {
                string[] arr = sText.Split(cDemlimiter);
                string sResult = "";

                if (arr != null)
                {
                    List<string> lst = arr.ToList();

                    sResult = lst[lst.Count - 1];
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
                return string.Empty;
            }
        }

        public static string removeCurvyBrackets(string sText) 
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
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        public static bool containsVariable(string sText, string sVariable)
        {
            try
            {
                bool bResult = false;
                //bool bIsGoToLabel = false;

                if (sText.ToLower().Contains(sVariable.ToLower().Trim()))
                {
                    if (sText.Contains('"'))
                    {
                        bResult = false;
                        int iPos = 0;

                        while (iPos >= 0 && !bResult)
                        {
                            iPos = sText.ToLower().IndexOf(sVariable.ToLower().Trim(), iPos + 1);

                            if (iPos >= 0)
                            {
                                if (ClsMiscString.Left(ref sText, iPos).Count(x => x == '"') % 2 == 1)
                                {
                                    //number of double quotes is odd
                                    //text is in brackets keep looking
                                }
                                else
                                {
                                    //number of double quotes is even
                                    //text is not in brackets, found the variable
                                    bResult = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        bResult = true;
                    }
                }
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

        public static bool containsTextNoPrefixOrSuffix(string sText, string sSubText)
        {
            try
            {
                /*
                 * sText must contain sString and not have any connecting charactors
                 * e.g. 
                 * "ma tth ew" "tth" = OK
                 * "ma tth(ew" "tth" = OK
                 * "ma_tth ew" "tth" = fail
                 * "matthew" "tth" = fail
                 * 
                 */
                bool bResult = false;
                bool bIsfound = false;
                bool bIsFinished = false;
                int iPos = 0;

                Regex rgx = new Regex("[A-Za-z0-9_]");

                iPos = sText.IndexOf(sSubText, 0);

                while (!bIsFinished && !bIsfound)
                {
                    if (iPos < 0)
                    { bIsFinished = true; }
                    else
                    {
                        bIsfound = true;
                        if (sText.Substring(0, iPos).Count(x => x == '"') % 2 != 1)
                        {
                            if (iPos > 0)
                            {
                                char cBefore = sText[iPos - 1];

                                if (rgx.IsMatch(cBefore.ToString()))
                                { bIsfound = false; }
                            }

                            if (iPos + sSubText.Length < sText.Length)
                            {
                                char cAfter = sText[iPos + sSubText.Length];

                                if (rgx.IsMatch(cAfter.ToString()))
                                { bIsfound = false; }
                            }
                        }
                    }

                    iPos = sText.IndexOf(sSubText, iPos + 1);
                }

                if (bIsfound)
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


        public static int indexTextNoPrefixOrSuffix(string sText, string sSubText)
        {
            try
            {
                /*
                 * sText must contain sString and not have any connecting charactors
                 * e.g. 
                 * "ma tth ew" "tth" = OK
                 * "ma tth(ew" "tth" = OK
                 * "ma_tth ew" "tth" = fail
                 * "matthew" "tth" = fail
                 * 
                 */
                int iResult = -1;
                bool bIsfound = false;
                bool bIsFinished = false;
                int iPos = 0;

                Regex rgx = new Regex("[A-Za-z0-9_]");

                iPos = sText.IndexOf(sSubText, 0);

                while (!bIsFinished && !bIsfound)
                {
                    if (iPos < 0)
                    { bIsFinished = true; }
                    else
                    {
                        bIsfound = true;
                        if (sText.Substring(0, iPos).Count(x => x == '"') % 2 != 1)
                        {
                            if (iPos > 0)
                            {
                                char cBefore = sText[iPos - 1];

                                if (rgx.IsMatch(cBefore.ToString()))
                                { bIsfound = false; }
                            }

                            if (iPos + sSubText.Length < sText.Length)
                            {
                                char cAfter = sText[iPos + sSubText.Length];

                                if (rgx.IsMatch(cAfter.ToString()))
                                { bIsfound = false; }
                            }
                        }

                        iResult = iPos;
                    }

                    iPos = sText.IndexOf(sSubText, iPos + 1);
                }

                if (!bIsfound)
                { iResult = -1; }

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

        public static bool checkPropertyCallTypeOK(string sLine, string sFunctionName, ClsCodeMapper.enumFunctionPropertyType ePropDesired)
        {
            try
            {
                /*
                 * sText must contain sString and not have any connecting charactors
                 * e.g. 
                 * "ma tth ew" "tth" = OK
                 * "ma tth(ew" "tth" = OK
                 * "ma_tth ew" "tth" = fail
                 * "matthew" "tth" = fail
                 * 
                 */

                /*
                 * Find a call for a class property and check if it's a Let Get Set
                 */
                
                bool bResult = false;
                bool bIsfound = false;
                bool bIsFinished = false;
                int iPos = 0;

                Regex rgx = new Regex("[A-Za-z0-9_]");

                iPos = sLine.IndexOf(sFunctionName, 0);

                while (!bIsFinished && !bIsfound)
                {
                    if (iPos < 0)
                    { bIsFinished = true; }
                    else
                    {
                        bIsfound = true;
                        if (sLine.Substring(0, iPos).Count(x => x == '"') % 2 != 1)
                        {
                            if (iPos > 0)
                            {
                                char cBefore = sLine[iPos - 1];

                                if (rgx.IsMatch(cBefore.ToString()))
                                { bIsfound = false; }
                            }

                            if (iPos + sFunctionName.Length < sLine.Length)
                            {
                                char cAfter = sLine[iPos + sFunctionName.Length];

                                if (rgx.IsMatch(cAfter.ToString()))
                                { bIsfound = false; }
                            }
                        }

                        if (bIsfound)
                        {
                            /*check it's the correct type of property*/
                            /* this to check
                             * ====-==-=====
                             * if it's the declare line line
                             * the call is followed immediately by an equals => let or set
                             * the call is preceded immediately by an equals => get
                             * if it's a parameter to a function e.g. "debug.print cls.fred" => get
                             * if it's preceded by ( or , or := then it's a parameter in a call => get
                             */

                            ClsCodeMapper.enumFunctionPropertyType ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_NA;

                            //if it's the declare line line
                            if (containsTextNoPrefixOrSuffix(sLine.ToLower(), "property"))
                            {
                                if (sLine.ToLower().Trim().StartsWith("property let ")
                                    || sLine.ToLower().Trim().StartsWith("public property let ")
                                    || sLine.ToLower().Trim().StartsWith("private property let ")
                                    || sLine.ToLower().Trim().StartsWith("friend property let ")
                                    || sLine.ToLower().Trim().StartsWith("public static property let ")
                                    || sLine.ToLower().Trim().StartsWith("private static property let ")
                                    || sLine.ToLower().Trim().StartsWith("friend static property let "))
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Let; }

                                if (sLine.ToLower().Trim().StartsWith("property get ")
                                    || sLine.ToLower().Trim().StartsWith("public property get ")
                                    || sLine.ToLower().Trim().StartsWith("private property get ")
                                    || sLine.ToLower().Trim().StartsWith("friend property get ")
                                    || sLine.ToLower().Trim().StartsWith("public static property get ")
                                    || sLine.ToLower().Trim().StartsWith("private static property get ")
                                    || sLine.ToLower().Trim().StartsWith("friend static property get "))
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get; }
                                
                                if (sLine.ToLower().Trim().StartsWith("property set ")
                                    || sLine.ToLower().Trim().StartsWith("public property set ")
                                    || sLine.ToLower().Trim().StartsWith("private property set ")
                                    || sLine.ToLower().Trim().StartsWith("friend property set ")
                                    || sLine.ToLower().Trim().StartsWith("public static property set ")
                                    || sLine.ToLower().Trim().StartsWith("private static property set ")
                                    || sLine.ToLower().Trim().StartsWith("friend static property set "))
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Set; }
                            }

                            //known cases
                            if (sLine.ToLower().Trim().StartsWith("msgbox") || sLine.ToLower().Trim().StartsWith("dubug.print"))
                            { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get; }

                            if (containsTextNoPrefixOrSuffix(sLine, "="))
                            {
                                if (sLine.ToLower().Trim().EndsWith(sFunctionName.ToLower()))
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get; }

                                int iPosFunctionName = indexTextNoPrefixOrSuffix(sLine, sFunctionName);
                                int iPosEquals = indexTextNoPrefixOrSuffix(sLine, "=");

                                if (iPosFunctionName < iPosEquals)
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Let; }
                                else
                                { ePropType = ClsCodeMapper.enumFunctionPropertyType.ePropType_Get; }
                            }

                            if (ePropDesired == ClsCodeMapper.enumFunctionPropertyType.ePropType_Get 
                                && ePropType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Get)
                            { bIsfound = true; }
                            else if ((ePropDesired == ClsCodeMapper.enumFunctionPropertyType.ePropType_Let || ePropDesired == ClsCodeMapper.enumFunctionPropertyType.ePropType_Set)
                                && (ePropType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Let || ePropType == ClsCodeMapper.enumFunctionPropertyType.ePropType_Set))
                            { bIsfound = true; }
                            else
                            { bIsfound = false; }
                        }




                    }

                    iPos = sLine.IndexOf(sFunctionName, iPos + 1);
                }

                if (bIsfound)
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

        public static bool containsWildcard(string sText, string sWildcard)
        {
            try
            {
                bool bResult = false;
                //bool bIsGoToLabel = false;

                Regex rgx = new Regex(sWildcard);

                if (rgx.IsMatch(sText.ToLower()))
                {
                    if (sText.Contains('"'))
                    {
                        bResult = false;
                        int iPos = 0;

                        while (iPos >= 0 && !bResult)
                        {
                            //iPos = sText.ToLower().IndexOf(sVariable.ToLower().Trim(), iPos + 1);
                            Match mtch = rgx.Match(sText.ToLower(), iPos + 1);
                            iPos = mtch.Index;

                            if (iPos >= 0)
                            {
                                if (ClsMiscString.Left(ref sText, iPos).Count(x => x == '"') % 2 == 1)
                                {
                                    //number of double quotes is odd
                                    //text is in brackets keep looking
                                }
                                else
                                {
                                    //number of double quotes is even
                                    //text is not in brackets, found the variable
                                    bResult = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        bResult = true;
                    }
                }
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

        public static int indexLastPosition(string sText, string sVariable, int iLeftChars)
        {
            try
            {
                bool bIsInQuotes = false;
                int iResult = -1;

                int iPos = sText.Length - 1;

                string sLeftChars = "";
                
                if (sText.Trim() == "")
                { iResult = -1; }
                else if (ClsMiscString.Right(ref sText, sText.Length - iLeftChars).Contains("\""))
                {
                    bool bIsFinished = false;

                    while (!bIsFinished)
                    {
                        if (sText.Substring(iPos, 1) == "\"")
                        { bIsInQuotes = !bIsInQuotes; }

                        if (!bIsInQuotes)
                        {
                            if (iPos <= iLeftChars)
                            { iPos--; }
                            else
                            { bIsFinished = true; }
                        }
                        else
                        {
                            if (iPos == 0)
                            { bIsFinished = true; }
                        }
                    }

                    Debug.Print("End loop - iPos: " + iPos.ToString());

                    sLeftChars = ClsMiscString.Left(ref sText, iLeftChars);
                    iResult = indexLastPosition(sLeftChars, sVariable);
                }
                else
                {
                    if (iLeftChars > 0)
                    {
                        sLeftChars = ClsMiscString.Left(ref sText, iLeftChars);
                        iResult = indexLastPosition(sLeftChars, sVariable);
                    }
                    else
                    { iResult = -1; }
                }

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
        
        public static int indexLastPosition(string sText, string sVariable)
        {
            try
            {
                bool bIsFinished = false;
                //int iResult = -1;
                bool bIsFound = false;
                bool bIsInQuotes = false;
                bool bIsOk = true;

                int iPos = sText.TrimEnd().Length - 1;

                while (!bIsFinished)
                {
                    if (sText.Trim() == "")
                    { bIsFound = false; }
                    else if (sText.Substring(iPos, 1) == "\"")
                    { bIsInQuotes = !bIsInQuotes; }
                    else
                    {
                        if (!bIsInQuotes)
                        {
                            if (iPos <= sText.Length - sVariable.Trim().Length)
                            {
                                if (sText.Substring(iPos, sVariable.Trim().Length).ToLower() == sVariable.ToLower().Trim())
                                {
                                    bool bBeforeIsOk;
                                    bool bAfterIsOk;

                                    if (iPos == 0)
                                    { bBeforeIsOk = true; }
                                    else
                                    {
                                        string sCharBefore = sText.Substring(iPos - 1, 1);

                                        if (Regex.IsMatch(sCharBefore, "[a-zA-Z0-9_]"))
                                        { bBeforeIsOk = false; }
                                        else
                                        { bBeforeIsOk = true; }

                                    }

                                    if (sText.TrimEnd().Length == iPos + sVariable.Trim().Length)
                                    { bAfterIsOk = true; }
                                    else
                                    {
                                        string sCharAfter = sText.Substring(iPos + sVariable.Trim().Length, 1);

                                        if (Regex.IsMatch(sCharAfter, "[a-zA-Z0-9_]"))
                                        { bAfterIsOk = false; }
                                        else
                                        { bAfterIsOk = true; }
                                    }

                                    bIsFound = bBeforeIsOk & bAfterIsOk;
                                    if (bIsFound)
                                    { bIsFinished = true; }
                                }
                            }
                        }
                    }

                    if (iPos <= 0)
                    { bIsFinished = true; }

                    if (!bIsFinished)
                    { iPos--; }
                }

                if (!bIsFound)
                { iPos = -1; }

                return iPos;
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

        public static string ingoreNull(string sText) 
        {
            try
            {
                string sResult;

                if (sText == null)
                { sResult = ""; }
                else
                { sResult = sText; }

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
                return string.Empty;
            }
        }

        public static bool containsOutsideQuotes(string sText, string sSubText) 
        {
            try
            {
                bool bResult = false;

                if (sText.Contains('"'))
                {
                    int iPos = 0;
                    bool bIsFinished = false;
                    bool bIsInQuotes = false;

                    while (bIsFinished)
                    {
                        char cCurr = sText[iPos];

                        if (cCurr == '"')
                        { bIsInQuotes = !bIsInQuotes; }
                        
                        if (!bIsInQuotes)
                        {
                            if (sText.Substring(iPos, sSubText.Length).ToLower() == sSubText.ToLower())
                            {
                                bResult = true;
                                bIsFinished = true;
                            }
                        }

                        if (iPos >= sText.Length - sSubText.Length - 1)
                        { bIsFinished = true; }
                        else
                        { iPos++; }
                    }
                }
                else
                {
                    if (sText.ToLower().Contains(sSubText.ToLower()))
                    { bResult = true; }
                    else
                    { bResult = false; }
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

        public static void RemoveDoubleSpaces(ref string sLine) 
        {
            try
            {
                sLine = sLine.TrimEnd();

                int iPos = sLine.Length - 1;
                bool bIsFinished = false;
                int iIndent = sLine.Length - sLine.TrimStart().Length;
                bool bIsInQuotes = false;

                if (sLine.Trim().Contains("  ") || sLine.Trim().Contains(" )") || sLine.Trim().Contains(" }") || sLine.Trim().Contains(" ]"))
                {
                    while (!bIsFinished)
                    {
                        char cCurrChar = sLine[iPos];

                        if (cCurrChar == '"')
                        { bIsInQuotes = !bIsInQuotes; }

                        if (iIndent + 1 >= iPos)
                        { bIsFinished = true; }
                        else
                        {
                            if (!bIsInQuotes)
                            {
                                if (cCurrChar == ' ')
                                {
                                    if (iPos < sLine.Length - 1)
                                    {
                                        char cPrevChar = sLine[iPos + 1];

                                        string sBefore = "";
                                        string sAfter = "";

                                        switch (cPrevChar)
                                        {
                                            case ' ':
                                                sBefore = ClsMiscString.Left(ref sLine, iPos);
                                                sAfter = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 1);

                                                sLine = sBefore + sAfter;
                                                break;
                                            case ',':
                                            case ')':
                                            case '}':
                                            case ']':
                                                sBefore = ClsMiscString.Left(ref sLine, iPos);
                                                sAfter = ClsMiscString.Right(ref sLine, sLine.Length - iPos - 1);

                                                sLine = sBefore + sAfter;
                                                break;
                                        }
                                    }
                                }
                            }
                        }

                        iPos--;
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

        public static bool isValidVariableName(string sName)
        { 
            try
            {
                bool bIsValid = true;

                if (!new Regex("^[a-zA-Z0-9_]*$").IsMatch(sName.Trim()))
                {
                    List<char> lstInvalidChar = new List<char>{'.', '!', '@', '&', '$', '#', '!', '"', '#', '$', '%', '&', 
                                                '\u0027', '(', ')', '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', 
                                                    '>', '?', '@', '[', '\\', ']', '^', '`', '`', '{', '|', '}', 
                                                    '~', ' ', '¡', '¢', '£', '¤', '¥', '¦', '§', '¨', '©', 'ª', 'ª', 
                                                    '«', '¬', '®', ',', '¯', '°', '±', '²', '³', '´', 'µ', '¶', '·', 
                                                    '¸', '¹', 'º', '»', '¼', '½', '¾', '¿', 'À', 'Á', 'Â', 'ǁ', 'ǂ', 
                                                    'ǀ', '˛'};// removed this one '_'

                    foreach (char cInvalidChar in lstInvalidChar)
                    {
                        if (sName.Contains(cInvalidChar))
                        { bIsValid = false; }
                    }
                }

                return bIsValid;
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

        public static string removeNonAlphaNumeric(string sText)
        {
            try
            {
                string sTemp = sText;

                foreach (char cChar in sText.ToList().FindAll(x => !char.IsLetterOrDigit(x)).Distinct())
                { sTemp = sTemp.Replace(cChar.ToString(), ""); }

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

        public static string nextFunctionName(ref ClsCodeMapper cCodeMapper, string sPrefix)
        {
            try
            {
                int iCounter = 1;
                string sResult = "";
                string sTempName;

                if (cCodeMapper.getLstFunctionNames().FindAll(x => x.Trim().ToUpper() == sPrefix.Trim().ToUpper()).Count == 0)
                { sResult = sPrefix; }
                else
                {
                    sTempName = sPrefix;
                    while (cCodeMapper.getLstFunctionNames().FindAll(x => x.Trim().ToUpper() == sTempName.Trim().ToUpper()).Count != 0)
                    {
                        iCounter++;
                        sTempName = sPrefix + iCounter;
                    }

                    sResult = sTempName;
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

                return string.Empty;
            }
        }
    }
}
