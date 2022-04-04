using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    class ClsCommonCode
    {
        public static void sheetExists(ref List<string> lstCode, ref int iIndent, ClsSettings cSettings, string sSheetName, string sVarName_IsFound, string sVarName_Sht)
        {
            try
            {
        
                lstCode.Add(cSettings.Indent(iIndent) + sVarName_IsFound + " = False");
                lstCode.Add(cSettings.Indent(iIndent) + "For Each " + sVarName_Sht + " In ThisWorkbook.Worksheet");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(" + sVarName_Sht + ".Name)) = Trim(UCase(\"" + sSheetName + "\")) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + sVarName_IsFound + " = True");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next " + sVarName_Sht);
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

        public static void namedRangeExists(ref List<string> lstCode, ref int iIndent, ClsSettings cSettings, string sName, string sVarName_IsFound, string sVarName_Name)
        {
            try
            {

                lstCode.Add(cSettings.Indent(iIndent) + sVarName_IsFound + " = False");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "For Each " + sVarName_Name + " In ThisWorkbook.Names");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "If Trim(UCase(" + sVarName_Name + ".Name)) = Trim(UCase(\"" + sName + "\")) Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + sVarName_IsFound + " = True");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Next " + sVarName_Name);
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
