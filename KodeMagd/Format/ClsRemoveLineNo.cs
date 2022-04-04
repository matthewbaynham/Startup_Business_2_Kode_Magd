using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using KodeMagd.Misc;

namespace KodeMagd.Format
{
    class ClsRemoveLineNo
    {
        public static void remove(ref string sLine)
        {
            try
            {
                string sResult = "";
                bool bNumericPrefix = false;
                
                sResult = sLine;

                if (sResult.Trim().Length > 0)
                {
                    if (Regex.IsMatch(ClsMiscString.Left(sResult.Trim(), 1), "[0-9]"))
                    { bNumericPrefix = true; }
                }

                if (bNumericPrefix)
                {
                    int iIndentSize = sResult.Length - sResult.TrimStart().Length;

                    bool bIsFinished = false;

                    while (!bIsFinished)
                    {
                        if (sResult.Trim().Length > 0)
                        {
                            if (Regex.IsMatch(ClsMiscString.Left(sResult.Trim(), 1), "[0-9]"))
                            { sResult = ClsMiscString.Right(sResult.Trim(), sResult.Trim().Length - 1); }
                            else
                            { bIsFinished = true; }
                        }
                        else
                        { bIsFinished = true; }
                    }

                    sResult = sResult.PadLeft(iIndentSize + sResult.Length, ' ');
                }

                sLine = sResult;
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
