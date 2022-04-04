using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using KodeMagd.Misc;
using KodeMagd.InsertCode;
using KodeMagd.Format;
using System.Diagnostics;

namespace KodeMagd
{
    class ClsCodeCleaner : ClsCodeMapper
    {
        public enum enumCleaningType 
        { 
            eClean_All,
            eClean_Indenting,
            eClean_SplitLines,
            eClean_SetLineLength,
            eClean_RemoveLineNo,
            eClean_DimSpacing
        }

        public void cleanModule()
        {
            try
            {
                cleanModule(enumCleaningType.eClean_All);
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

        public void cleanModule(enumCleaningType eCleanType) 
        {
            try
            {
                VBA.VBComponent cmpResult;

                cmpResult = ClsMisc.ActiveVBComponent();

                if (cmpResult == null)
                { MessageBox.Show("No code window is active", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                else
                { cleanModule(cmpResult, eCleanType); }
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

        public void cleanModule(VBA.VBComponent vbComp, enumCleaningType eCleanType)
        {
            try
            {
                readCode(vbComp);

                if (eCleanType == enumCleaningType.eClean_All || eCleanType == enumCleaningType.eClean_Indenting)
                { Indenting(); }

                if (eCleanType == enumCleaningType.eClean_All || eCleanType == enumCleaningType.eClean_RemoveLineNo)
                { removeLineNo(); }

                if (eCleanType == enumCleaningType.eClean_All || eCleanType == enumCleaningType.eClean_SplitLines)
                { splitLines(); }

                if (eCleanType == enumCleaningType.eClean_All || eCleanType == enumCleaningType.eClean_SetLineLength)
                { setLineLength(); }

                if (eCleanType == enumCleaningType.eClean_All || eCleanType == enumCleaningType.eClean_DimSpacing)
                { alignVariableDim(); }

                ImplementChanges(vbComp);
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
