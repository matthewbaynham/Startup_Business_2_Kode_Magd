using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    class ClsConvert
    {
        public static ADODB.DataTypeEnum DataTypeEnum(string sText)
        {
            try
            {
                ADODB.DataTypeEnum eResult = ADODB.DataTypeEnum.adError;

                foreach (ADODB.DataTypeEnum eTemp in Enum.GetValues(typeof(ADODB.DataTypeEnum))) 
                {
                    if (sText.Trim().ToUpper() == eTemp.ToString().Trim().ToUpper())
                    { eResult = eTemp; }
                }

                return eResult;
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

                return ADODB.DataTypeEnum.adError;
            }
        }

        
        public static ADODB.CursorTypeEnum CursorTypeEnum(string sText) 
        {
            try
            {
                ADODB.CursorTypeEnum eResult = ADODB.CursorTypeEnum.adOpenForwardOnly;

                foreach (ADODB.CursorTypeEnum eTemp in Enum.GetValues(typeof(ADODB.CursorTypeEnum)))
                {
                    if (eTemp.ToString() == sText)
                    { eResult = eTemp; }
                }

                return eResult;
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

                return ADODB.CursorTypeEnum.adOpenForwardOnly;
            }
        }

        public static ADODB.LockTypeEnum LockTypeEnum(string sText)
        {
            try
            {
                ADODB.LockTypeEnum eResult = ADODB.LockTypeEnum.adLockReadOnly;

                foreach (ADODB.LockTypeEnum eTemp in Enum.GetValues(typeof(ADODB.LockTypeEnum)))
                {
                    if (eTemp.ToString() == sText)
                    { eResult = eTemp; }
                }

                return eResult;
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

                return ADODB.LockTypeEnum.adLockReadOnly;
            }
        }
    }
}
