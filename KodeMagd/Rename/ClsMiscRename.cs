using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using KodeMagd.Misc;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Rename
{
    class ClsMiscRename
    {
        public static void RenameVariable(string sNewName, string sOldName, ref string sLine)
        {
            try
            {
                if (ClsMiscString.containsVariable(sLine, sOldName))
                {
                    bool bIsFinished = false;
                    string sTemp = sLine;
                    int iPos = sTemp.TrimEnd().Length;

                    int iLookAtChars = sTemp.Length;

                    while (!bIsFinished)
                    {
                        //iPos = sTemp.LastIndexOf(sOldName, iPos);

                        iPos = ClsMiscString.indexLastPosition(ClsMiscString.Left(ref sTemp, iLookAtChars), sOldName, iPos);

                        if (iPos == -1)
                        { bIsFinished = true; }
                        else
                        {
                            bool bBeginningOK;

                            if (iPos == 0)
                            { bBeginningOK = true; }
                            else
                            {
                                string sBefore = sTemp.Substring(iPos - 1, 1);

                                if (Regex.IsMatch(sBefore, "[0-9a-zA-Z_]"))
                                { bBeginningOK = false; }
                                else
                                { bBeginningOK = true; }
                            }

                            bool bEndingOK;

                            if (iPos == sTemp.Length - sOldName.Length)
                            { bEndingOK = true; }
                            else
                            {
                                string sAfter = sTemp.Substring(iPos + sOldName.Length, 1);

                                if (Regex.IsMatch(sAfter, "[0-9a-zA-Z_]"))
                                { bEndingOK = false; }
                                else
                                { bEndingOK = true; }
                            }

                            if (bBeginningOK & bEndingOK)
                            {
                                //find replace
                                string sBeforeVariable;
                                string sAfterVariable;

                                sBeforeVariable = ClsMiscString.Left(ref sTemp, iPos);
                                sAfterVariable = ClsMiscString.Right(ref sTemp, sTemp.Length - iPos - sOldName.Length);

                                sTemp = sBeforeVariable + sNewName + sAfterVariable;
                            }
                            else
                            { iLookAtChars = iPos - 1; }
                        }
                    }

                    sLine = sTemp;
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
