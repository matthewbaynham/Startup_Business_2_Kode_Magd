using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using Scripting;

namespace KodeMagd
{
    class ClsTextFileLog
    {
        private string sFullPath;
        private const string csDelimiter = "\t";
        //TextWriter tw;
        Scripting.FileSystemObject FSO = new Scripting.FileSystemObject();
        Scripting.TextStream tsLog;

        public ClsTextFileLog()
        {
            try
            {
                DateTime dtNow = new DateTime();

                dtNow = DateTime.Now;

                sFullPath = "C:\\visual studio 2010\\Logs\\";
                sFullPath += dtNow.ToString("yyyyMMdd HHmmss");
                sFullPath += ".xls";

                //tsLog = FSO.OpenTextFile(sFullPath, IOMode.ForWriting, true, Tristate.TristateFalse);

                //LOG("Title", "Sub Title", "Text");
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

        ~ClsTextFileLog()
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

        public void Close()
        {
            try
            {
                if (tsLog == null)
                {
                    tsLog.Close();
                    tsLog = null;
                    FSO = null;
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

        public void LOG(string sTitle, string sSubTitle, string sText) 
        {
            try
            {
                string sLine;

                sLine = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                sLine += csDelimiter;
                sLine += sTitle;
                sLine += csDelimiter;
                sLine += sSubTitle;
                sLine += csDelimiter;
                sLine += sText;

                // write a line of text to the file
                //tw.WriteLine(sLine);
                tsLog.WriteLine(sLine);
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
