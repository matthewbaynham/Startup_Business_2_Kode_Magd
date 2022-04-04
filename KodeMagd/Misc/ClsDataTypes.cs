using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.InsertCode;

namespace KodeMagd.Misc
{
    public class ClsDataTypes
    {
        public enum enumGeneralDateType 
        {
            eString,
            eNumber,
            eDate,
            eBool,
            eUnknown
        }

        public enum vbVarType
        {
            vbEmpty = 0,
            vbNull = 1,
            vbInteger = 2,
            vbLong = 3,
            vbSingle = 4,
            vbDouble = 5,
            vbCurrency = 6,
            vbDate = 7,
            vbString = 8,
            vbObject = 9,
            vbError = 10,
            vbBoolean = 11,
            vbVariant = 12,
            vbDataObject = 13,
            vbDecimal = 14,
            vbByte = 17,
            vbLongLong = 20, //(defined only on implementations that support a LongLong value type)
            vbUserDefinedType = 36,
            vbArray = 8192,
            vbUnknown
        }

        private struct strDataType
        {
            public vbVarType eType;
            public string sName;
            public string sVbaName;
            public bool bIsCommon;
            public enumGeneralDateType eGeneralType;
        }

        private struct strAdoDataType
        {
            public ADODB.DataTypeEnum eType;
            public string sName;
            public bool bIsCommon;
            public enumGeneralDateType eGeneralType;
        }

        private List<strDataType> lstTypes = new List<strDataType>();
        private List<strAdoDataType> lstAdoTypes = new List<strAdoDataType>();

        public ClsDataTypes() 
        {
            try
            {
                lstTypes.Clear();
                strDataType objTemp = new strDataType();

                objTemp.sName = "Null";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbNull;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Empty";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbEmpty;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Integer";
                objTemp.sVbaName = "Integer";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbInteger;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Long";
                objTemp.sVbaName = "Long";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbLong;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Single";
                objTemp.sVbaName = "Single";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbSingle;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Double";
                objTemp.sVbaName = "Double";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbDouble;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Currency";
                objTemp.sVbaName = "Currency";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbCurrency;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Date";
                objTemp.sVbaName = "Date";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbDate;
                objTemp.eGeneralType = enumGeneralDateType.eDate;
                lstTypes.Add(objTemp);

                objTemp.sName = "String";
                objTemp.sVbaName = "String";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbString;
                objTemp.eGeneralType = enumGeneralDateType.eString;
                lstTypes.Add(objTemp);

                objTemp.sName = "Object";
                objTemp.sVbaName = "Object";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbObject;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Error";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbError;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Boolean";
                objTemp.sVbaName = "Boolean";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbBoolean;
                objTemp.eGeneralType = enumGeneralDateType.eBool;
                lstTypes.Add(objTemp);

                objTemp.sName = "Variant";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbVariant;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Data Object";
                objTemp.sVbaName = "Data Object";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbDataObject;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Decimal";
                objTemp.sVbaName = "Double";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbDecimal;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Byte";
                objTemp.sVbaName = "Byte";
                objTemp.bIsCommon = true;
                objTemp.eType = vbVarType.vbByte;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "Long Long";
                objTemp.sVbaName = "Long";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbLongLong;
                objTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstTypes.Add(objTemp);

                objTemp.sName = "User Defined Type";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbUserDefinedType;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Array";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbArray;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);

                objTemp.sName = "Unknown";
                objTemp.sVbaName = "Variant";
                objTemp.bIsCommon = false;
                objTemp.eType = vbVarType.vbUnknown;
                objTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstTypes.Add(objTemp);


                lstAdoTypes.Clear();
                strAdoDataType objAdoTemp = new strAdoDataType();

                objAdoTemp.sName = "adArray";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adArray;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adBigInt";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adBigInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adBinary";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adBinary;
                objAdoTemp.eGeneralType = enumGeneralDateType.eBool;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adBoolean";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adBoolean;
                objAdoTemp.eGeneralType = enumGeneralDateType.eBool;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adBSTR";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adBSTR;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adChapter";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adChapter;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adChar";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adCurrency";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adCurrency;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDate";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDate;
                objAdoTemp.eGeneralType = enumGeneralDateType.eDate;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDBDate";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDBDate;
                objAdoTemp.eGeneralType = enumGeneralDateType.eDate;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDBTime";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDBTime; 
                objAdoTemp.eGeneralType = enumGeneralDateType.eDate;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDBTimeStamp";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDBTimeStamp;
                objAdoTemp.eGeneralType = enumGeneralDateType.eDate;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDecimal";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDecimal;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adDouble";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adDouble;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adEmpty";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adEmpty;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adError";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adError;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adFileTime";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adFileTime;
                objAdoTemp.eGeneralType = enumGeneralDateType.eDate;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adGUID";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adGUID;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adIDispatch";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adIDispatch;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adInteger";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adInteger;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adIUnknown";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adIUnknown;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adLongVarBinary";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adLongVarBinary;
                objAdoTemp.eGeneralType = enumGeneralDateType.eBool;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adLongVarChar";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adLongVarChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adLongVarWChar";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adLongVarWChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adNumeric";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adNumeric;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adPropVariant";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adPropVariant;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adSingle";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adSingle;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adSmallInt";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adSmallInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adTinyInt";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adTinyInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adUnsignedBigInt";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adUnsignedBigInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adUnsignedInt";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adUnsignedInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adUnsignedSmallInt";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adUnsignedSmallInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adUnsignedTinyInt";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adUnsignedTinyInt;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adUserDefined";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adUserDefined;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adVarBinary";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adVarBinary;
                objAdoTemp.eGeneralType = enumGeneralDateType.eBool;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adVarChar";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adVarChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adVariant";
                objAdoTemp.bIsCommon = false;
                objAdoTemp.eType = ADODB.DataTypeEnum.adVariant;
                objAdoTemp.eGeneralType = enumGeneralDateType.eUnknown;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adVarNumeric";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adVarNumeric;
                objAdoTemp.eGeneralType = enumGeneralDateType.eNumber;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adVarWChar";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adVarWChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);

                objAdoTemp.sName = "adWChar";
                objAdoTemp.bIsCommon = true;
                objAdoTemp.eType = ADODB.DataTypeEnum.adWChar;
                objAdoTemp.eGeneralType = enumGeneralDateType.eString;
                lstAdoTypes.Add(objAdoTemp);
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

        ~ClsDataTypes() 
        {
            try
            {
                lstAdoTypes = null;
                lstTypes = null;
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

        public List<string> commonDataTypes() 
        { 
            try 
            {
                List<string> lstResult = new List<string>();

                foreach (strDataType strTemp in lstTypes) 
                { 
                    if (strTemp.bIsCommon == true)
                    { lstResult.Add(strTemp.sName); }
                }

                lstResult.Sort();

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

        public List<string> allDataTypes()
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (strDataType strTemp in lstTypes)
                { lstResult.Add(strTemp.sName); }

                lstResult.Sort();

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

        public string getName(ClsDataTypes.vbVarType eType)
        {
            try
            {
                string sResult = "";

                strDataType objTemp = lstTypes.Find(r => r.eType == eType);

                sResult = objTemp.sVbaName;

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

        public enumGeneralDateType getGeneralType(ClsDataTypes.vbVarType eType)
        {
            try
            {
                enumGeneralDateType eResult = enumGeneralDateType.eUnknown;

                strDataType objTemp = lstTypes.Find(r => r.eType == eType);

                if (objTemp.eType == null)
                { eResult = enumGeneralDateType.eUnknown; }
                else
                { eResult = objTemp.eGeneralType; }

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

                return enumGeneralDateType.eUnknown;
            }
        }

        public enumGeneralDateType getGeneralType(ADODB.DataTypeEnum eType)
        {
            try
            {
                enumGeneralDateType eResult = enumGeneralDateType.eUnknown;

                strAdoDataType objTemp = lstAdoTypes.Find(r => r.eType == eType);

                if (objTemp.eType == null)
                { eResult = enumGeneralDateType.eUnknown; }
                else
                { eResult = objTemp.eGeneralType; }

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

                return enumGeneralDateType.eUnknown;
            }
        }

        public enumGeneralDateType getGeneralType(string sName)
        {
            try
            {
                enumGeneralDateType eResult = enumGeneralDateType.eUnknown;

                strDataType objTemp = lstTypes.Find(r => r.sName == sName);

                if (objTemp.eType == null)
                { eResult = enumGeneralDateType.eUnknown; }
                else
                { eResult = objTemp.eGeneralType; }

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

                return enumGeneralDateType.eUnknown;
            }
        }

        public vbVarType getDataType(string sName)
        {
            try
            {
                vbVarType eResult = vbVarType.vbUnknown;

                strDataType objTemp = lstTypes.Find(r => r.sName == sName);

                if (objTemp.eType == null)
                { eResult = vbVarType.vbUnknown; }
                else
                { eResult = objTemp.eType; }

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

                return vbVarType.vbUnknown;
            }
        }

        public static ADODB.DataTypeEnum getAdodbDataType(string sText)
        {
            try
            {
                ADODB.DataTypeEnum eResult = ADODB.DataTypeEnum.adIUnknown;
 
                foreach (ADODB.DataTypeEnum eType in Enum.GetValues(typeof(ADODB.DataTypeEnum)))
                {
                    if (sText == eType.ToString())
                    { eResult = eType; }
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

                return ADODB.DataTypeEnum.adIUnknown;
            }
        }

        public vbVarType getDataType(ADODB.DataTypeEnum eType)
        {
            try
            {
                vbVarType eResult = vbVarType.vbUnknown;

                switch(eType)
                {
                    case ADODB.DataTypeEnum.adArray:
                        eResult=vbVarType.vbArray;
                        break;
                    case ADODB.DataTypeEnum.adBigInt:
                    case ADODB.DataTypeEnum.adInteger:
                    case ADODB.DataTypeEnum.adSmallInt:
                    case ADODB.DataTypeEnum.adTinyInt:
                    case ADODB.DataTypeEnum.adUnsignedBigInt:
                    case ADODB.DataTypeEnum.adUnsignedInt:
                    case ADODB.DataTypeEnum.adUnsignedSmallInt:
                    case ADODB.DataTypeEnum.adUnsignedTinyInt:
                        eResult = vbVarType.vbLong;
                        break;
                    case ADODB.DataTypeEnum.adDate:
                    case ADODB.DataTypeEnum.adDBDate:
                    case ADODB.DataTypeEnum.adDBTime:
                    case ADODB.DataTypeEnum.adDBTimeStamp:
                    case ADODB.DataTypeEnum.adFileTime:
                        eResult = vbVarType.vbDate;
                        break;
                    case ADODB.DataTypeEnum.adBSTR:
                    case ADODB.DataTypeEnum.adChar:
                    case ADODB.DataTypeEnum.adLongVarChar:
                    case ADODB.DataTypeEnum.adLongVarWChar:
                    case ADODB.DataTypeEnum.adVarChar:
                    case ADODB.DataTypeEnum.adVarWChar:
                    case ADODB.DataTypeEnum.adWChar:
                        eResult = vbVarType.vbString;
                        break;
                    case ADODB.DataTypeEnum.adBoolean:
                        eResult = vbVarType.vbBoolean;
                        break;
                    case ADODB.DataTypeEnum.adBinary:
                    case ADODB.DataTypeEnum.adVarBinary:
                    case ADODB.DataTypeEnum.adLongVarBinary:
                        eResult = vbVarType.vbByte;
                        break;
                    case ADODB.DataTypeEnum.adDecimal:
                    case ADODB.DataTypeEnum.adNumeric:
                    case ADODB.DataTypeEnum.adVarNumeric:
                        eResult = vbVarType.vbDecimal;
                        break;
                    case ADODB.DataTypeEnum.adDouble:
                        eResult = vbVarType.vbDouble;
                        break;
                    case ADODB.DataTypeEnum.adSingle:
                        eResult = vbVarType.vbSingle;
                        break;
                    case ADODB.DataTypeEnum.adCurrency:
                        eResult = vbVarType.vbCurrency;
                        break;
                    case ADODB.DataTypeEnum.adPropVariant:
                    case ADODB.DataTypeEnum.adVariant:
                        eResult = vbVarType.vbVariant;
                        break;
                    case ADODB.DataTypeEnum.adError:
                        eResult = vbVarType.vbError;
                        break;
                    case ADODB.DataTypeEnum.adEmpty:
                        eResult = vbVarType.vbEmpty;
                        break;
                    case ADODB.DataTypeEnum.adUserDefined:
                        eResult = vbVarType.vbUserDefinedType;
                        break;
                    case ADODB.DataTypeEnum.adIUnknown:
                    case ADODB.DataTypeEnum.adChapter:
                    case ADODB.DataTypeEnum.adGUID:
                    case ADODB.DataTypeEnum.adIDispatch:
                        eResult = vbVarType.vbUnknown;
                        break;
                    default:
                        eResult = vbVarType.vbUnknown;
                        break;
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

                return vbVarType.vbUnknown;
            }
        }

        public string typePrefix(vbVarType eType)
        { 
            try
            {
                string sPrefix = "";

                switch (eType)
                {
                    case ClsDataTypes.vbVarType.vbByte:
                        sPrefix = "byt";
                        break;
                    case ClsDataTypes.vbVarType.vbCurrency:
                        sPrefix = "dte";
                        break;
                    case ClsDataTypes.vbVarType.vbDecimal:
                    case ClsDataTypes.vbVarType.vbDouble:
                        sPrefix = "d";
                        break;
                    case ClsDataTypes.vbVarType.vbInteger:
                        sPrefix = "i";
                        break;
                    case ClsDataTypes.vbVarType.vbLong:
                    case ClsDataTypes.vbVarType.vbLongLong:
                        sPrefix = "l";
                        break;
                    case ClsDataTypes.vbVarType.vbSingle:
                        sPrefix = "s";
                        break;
                    case ClsDataTypes.vbVarType.vbString:
                        sPrefix = "s";
                        break;
                    case ClsDataTypes.vbVarType.vbVariant:
                        sPrefix = "var";
                        break;
                    default:
                        sPrefix = "p";
                        break;
                }

                return sPrefix;
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

        public static bool typeCheck(ADODB.DataTypeEnum eType, string sText) 
        { 
            try
            {
                bool bIsOk;

                switch (eType)
                { 
                    case ADODB.DataTypeEnum.adBigInt:
                    case ADODB.DataTypeEnum.adInteger:
                    case ADODB.DataTypeEnum.adSmallInt:
                    case ADODB.DataTypeEnum.adTinyInt:
                    case ADODB.DataTypeEnum.adUnsignedBigInt:
                    case ADODB.DataTypeEnum.adUnsignedInt:
                    case ADODB.DataTypeEnum.adUnsignedSmallInt:
                    case ADODB.DataTypeEnum.adUnsignedTinyInt:
                        int iResult;
        
                        if (int.TryParse(sText, out iResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adBinary:
                    case ADODB.DataTypeEnum.adLongVarBinary:
                    case ADODB.DataTypeEnum.adVarBinary:
                        byte binResult;
        
                        if (System.Byte.TryParse(sText, out binResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adBoolean:
                        bool boolResult;
        
                        if (bool.TryParse(sText, out boolResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adChar:
                    case ADODB.DataTypeEnum.adLongVarChar:
                    case ADODB.DataTypeEnum.adLongVarWChar:
                    case ADODB.DataTypeEnum.adVarChar:
                    case ADODB.DataTypeEnum.adVarWChar:
                    case ADODB.DataTypeEnum.adWChar:
                        //char chrResult;
        
                        //if (char.TryParse(sText, out chrResult)) 
                        //{ bIsOk = true; }
                        //else
                        //{ bIsOk = false; }
                        bIsOk = true;

                        break;
                    case ADODB.DataTypeEnum.adDecimal:
                    case ADODB.DataTypeEnum.adDouble:
                    case ADODB.DataTypeEnum.adCurrency:
                    case ADODB.DataTypeEnum.adNumeric:
                    case ADODB.DataTypeEnum.adVarNumeric:
                        decimal decResult;
        
                        if (decimal.TryParse(sText, out decResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adSingle:
                        Single sglResult;
        
                        if (Single.TryParse(sText, out sglResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adDate:
                    case ADODB.DataTypeEnum.adDBDate:
                    case ADODB.DataTypeEnum.adDBTime:
                    case ADODB.DataTypeEnum.adDBTimeStamp:
                    case ADODB.DataTypeEnum.adFileTime:
                        DateTime dtResult;
        
                        if (DateTime.TryParse(sText, out dtResult)) 
                        { bIsOk = true; }
                        else
                        { bIsOk = false; }
                        
                        break;
                    case ADODB.DataTypeEnum.adArray:
                    case ADODB.DataTypeEnum.adBSTR:
                    case ADODB.DataTypeEnum.adChapter:
                    case ADODB.DataTypeEnum.adEmpty:
                    case ADODB.DataTypeEnum.adGUID:
                    case ADODB.DataTypeEnum.adIDispatch:
                    case ADODB.DataTypeEnum.adIUnknown:
                    case ADODB.DataTypeEnum.adPropVariant:
                    case ADODB.DataTypeEnum.adUserDefined:
                    case ADODB.DataTypeEnum.adVariant:
                    case ADODB.DataTypeEnum.adError:
                        bIsOk = true;
                        break;
                        bIsOk = false;
                        break;
                    default:
                        bIsOk = true;
                        break;
                }

                return bIsOk;
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

        public static int getDataTypeSize(ADODB.DataTypeEnum eType) 
        {
            try
            {
                int iSize;

                switch (eType) 
                {
                    case ADODB.DataTypeEnum.adBigInt:
                        iSize = 8; 
                        break;
                    case ADODB.DataTypeEnum.adBinary:	
                        iSize = 50; 
                        break;
                    case ADODB.DataTypeEnum.adBoolean:	
                        iSize = 1; 
                        break;
                    case ADODB.DataTypeEnum.adBSTR:	
                        iSize =  0; 
                        break;
                    case ADODB.DataTypeEnum.adChapter:
                        iSize =  0;
                        break;
                    case ADODB.DataTypeEnum.adChar:
                        iSize = 0;
                        break;
                    case ADODB.DataTypeEnum.adCurrency:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adDate:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adDBDate:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adDBTime:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adDBTimeStamp:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adDecimal:	
                        iSize = 8; 
                        break;
                    case ADODB.DataTypeEnum.adDouble:
                        iSize = 8;
                        break;
                    case ADODB.DataTypeEnum.adEmpty:
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adError:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adFileTime:	
                        iSize = 8; 
                        break;
                    case ADODB.DataTypeEnum.adGUID:	
                        iSize = 16; 
                        break;
                    case ADODB.DataTypeEnum.adIDispatch:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adInteger:	
                        iSize = 4; 
                        break;
                    case ADODB.DataTypeEnum.adIUnknown:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adLongVarBinary:	
                        iSize = 2147483647; 
                        break;
                    case ADODB.DataTypeEnum.adLongVarChar:	
                        iSize = 2147483647; 
                        break;
                    case ADODB.DataTypeEnum.adLongVarWChar:	
                        iSize = 1073741823; 
                        break;
                    case ADODB.DataTypeEnum.adNumeric:	
                        iSize = 9; 
                        break;
                    case ADODB.DataTypeEnum.adPropVariant:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adSingle:	
                        iSize = 4; 
                        break;
                    case ADODB.DataTypeEnum.adSmallInt:	
                        iSize = 2; 
                        break;
                    case ADODB.DataTypeEnum.adTinyInt:	
                        iSize = 1; 
                        break;
                    case ADODB.DataTypeEnum.adUnsignedBigInt:	
                        iSize = 8; 
                        break;
                    case ADODB.DataTypeEnum.adUnsignedInt:	
                        iSize = 4; 
                        break;
                    case ADODB.DataTypeEnum.adUnsignedSmallInt:	
                        iSize = 2; 
                        break;
                    case ADODB.DataTypeEnum.adUnsignedTinyInt:	
                        iSize = 1; 
                        break;
                    case ADODB.DataTypeEnum.adUserDefined:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adVarBinary:	
                        iSize = 50; 
                        break;
                    case ADODB.DataTypeEnum.adVarChar:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adVariant:	
                        iSize = 8016; 
                        break;
                    case ADODB.DataTypeEnum.adVarNumeric:	
                        iSize = 8; 
                        break;
                    case ADODB.DataTypeEnum.adVarWChar:	
                        iSize = 0; 
                        break;
                    case ADODB.DataTypeEnum.adWChar:	
                        iSize = 0; 
                        break;
                    default:
                        iSize = 0;
                        break;
                }

                return iSize;
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

        public static string getDataTypeText(ADODB.DataTypeEnum eType)
        {
            try
            {
                string sResult = "";

                switch (eType)
                {
                    case ADODB.DataTypeEnum.adBigInt:
                        sResult = "BigInt";
                        break;
                    case ADODB.DataTypeEnum.adBinary:
                        sResult = "Binary";
                        break;
                    case ADODB.DataTypeEnum.adBoolean:
                        sResult = "Boolean";
                        break;
                    case ADODB.DataTypeEnum.adBSTR:
                        sResult = "BSTR";
                        break;
                    case ADODB.DataTypeEnum.adChapter:
                        sResult = "Chapter";
                        break;
                    case ADODB.DataTypeEnum.adChar:
                        sResult = "Char";
                        break;
                    case ADODB.DataTypeEnum.adCurrency:
                        sResult = "Currency";
                        break;
                    case ADODB.DataTypeEnum.adDate:
                        sResult = "Date";
                        break;
                    case ADODB.DataTypeEnum.adDBDate:
                        sResult = "DBDate";
                        break;
                    case ADODB.DataTypeEnum.adDBTime:
                        sResult = "DBTime";
                        break;
                    case ADODB.DataTypeEnum.adDBTimeStamp:
                        sResult = "DBTimeStamp";
                        break;
                    case ADODB.DataTypeEnum.adDecimal:
                        sResult = "Decimal";
                        break;
                    case ADODB.DataTypeEnum.adDouble:
                        sResult = "Double";
                        break;
                    case ADODB.DataTypeEnum.adEmpty:
                        sResult = "Empty";
                        break;
                    case ADODB.DataTypeEnum.adError:
                        sResult = "Error";
                        break;
                    case ADODB.DataTypeEnum.adFileTime:
                        sResult = "FileTime";
                        break;
                    case ADODB.DataTypeEnum.adGUID:
                        sResult = "GUID";
                        break;
                    case ADODB.DataTypeEnum.adIDispatch:
                        sResult = "IDispatch";
                        break;
                    case ADODB.DataTypeEnum.adInteger:
                        sResult = "Integer";
                        break;
                    case ADODB.DataTypeEnum.adIUnknown:
                        sResult = "IUnknown";
                        break;
                    case ADODB.DataTypeEnum.adLongVarBinary:
                        sResult = "LongVarBinary";
                        break;
                    case ADODB.DataTypeEnum.adLongVarChar:
                        sResult = "LongVarChar";
                        break;
                    case ADODB.DataTypeEnum.adLongVarWChar:
                        sResult = "LongVarWChar";
                        break;
                    case ADODB.DataTypeEnum.adNumeric:
                        sResult = "Numeric";
                        break;
                    case ADODB.DataTypeEnum.adPropVariant:
                        sResult = "PropVariant";
                        break;
                    case ADODB.DataTypeEnum.adSingle:
                        sResult = "Single";
                        break;
                    case ADODB.DataTypeEnum.adSmallInt:
                        sResult = "SmallInt";
                        break;
                    case ADODB.DataTypeEnum.adTinyInt:
                        sResult = "TinyInt";
                        break;
                    case ADODB.DataTypeEnum.adUnsignedBigInt:
                        sResult = "UnsignedBigInt";
                        break;
                    case ADODB.DataTypeEnum.adUnsignedInt:
                        sResult = "UnsignedInt";
                        break;
                    case ADODB.DataTypeEnum.adUnsignedSmallInt:
                        sResult = "UnsignedSmallInt";
                        break;
                    case ADODB.DataTypeEnum.adUnsignedTinyInt:
                        sResult = "UnsignedTinyInt";
                        break;
                    case ADODB.DataTypeEnum.adUserDefined:
                        sResult = "UserDefined";
                        break;
                    case ADODB.DataTypeEnum.adVarBinary:
                        sResult = "VarBinary";
                        break;
                    case ADODB.DataTypeEnum.adVarChar:
                        sResult = "VarChar";
                        break;
                    case ADODB.DataTypeEnum.adVariant:
                        sResult = "Variant";
                        break;
                    case ADODB.DataTypeEnum.adVarNumeric:
                        sResult = "VarNumeric";
                        break;
                    case ADODB.DataTypeEnum.adVarWChar:
                        sResult = "VarWChar";
                        break;
                    case ADODB.DataTypeEnum.adWChar:
                        sResult = "WChar";
                        break;
                    default:
                        sResult = "";
                        break;
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

        public static string convertModuleType(Microsoft.Vbe.Interop.vbext_ComponentType eType)
        {
            try
            {
                string sResult = "";

                switch (eType)
                {
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                        sResult = "ActiveX";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                        sResult = "Class";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                        sResult = "Document";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                        sResult = "Form";
                        break;
                    case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                        sResult = "Standard Module";
                        break;
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

        public static Microsoft.Vbe.Interop.vbext_ComponentType convertModuleType(string sType)
        {
            try
            {
                Microsoft.Vbe.Interop.vbext_ComponentType eResult = Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule;

                foreach (Microsoft.Vbe.Interop.vbext_ComponentType eTemp in Enum.GetValues(typeof(Microsoft.Vbe.Interop.vbext_ComponentType)))
                {
                    if (convertModuleType(eTemp) == sType)
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
                return Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule;
            }
        }

        public static List<ADODB.DataTypeEnum> NotSupportedAdoDataTypes()
        {
            try
            {
                List<ADODB.DataTypeEnum> lstResult = new List<ADODB.DataTypeEnum>();

                lstResult.Add(ADODB.DataTypeEnum.adArray);

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
                return new List<ADODB.DataTypeEnum>();
            }
        }

        public static enumGeneralDateType textToGeneralType(string sText)
        {
            try
            {
                enumGeneralDateType eResult = enumGeneralDateType.eUnknown;
                string sTemp = sText.Trim().ToUpper();

                //strings
                if (sTemp.StartsWith("STRING")
                    || sTemp.StartsWith("CHAR")
                    || sTemp.StartsWith("NCHAR")
                    || sTemp.StartsWith("WCHAR")
                    || sTemp.StartsWith("VARCHAR")
                    || sTemp.StartsWith("VARCHAR2")
                    || sTemp.StartsWith("VARWCHAR")
                    || sTemp.StartsWith("WVARCHAR")
                    || sTemp.StartsWith("NVARCHAR"))
                { eResult = ClsDataTypes.enumGeneralDateType.eString; }

                //numbers
                if (sTemp == "FLOAT"
                    || sTemp == "DOUBLE"
                    || sTemp == "SINGLE"
                    || sTemp == "NUMBER"
                    || sTemp == "INT"
                    || sTemp == "INT16"
                    || sTemp == "INT32"
                    || sTemp == "INTEGER"
                    || sTemp == "LONG"
                    || sTemp == "LONGLONG"
                    || sTemp == "LONG LONG"
                    || sTemp == "BYTE"
                    || sTemp == "DECIMAL"
                    || sTemp == "SHORT"
                    || sTemp == "ULONG"
                    || sTemp == "USHORT"
                    || sTemp == "UINTEGER")
                { eResult = ClsDataTypes.enumGeneralDateType.eNumber; }

                //dates
                if (sTemp == "DATETIME"
                    || sTemp == "DATETIME2"
                    || sTemp == "DATE"
                    || sTemp == "TIME"
                    || sTemp == "TIMESTAMP")
                { eResult = ClsDataTypes.enumGeneralDateType.eDate; }

                //boolean
                if (sTemp == "BOOL"
                    || sTemp == "BOOLEAN")
                { eResult = ClsDataTypes.enumGeneralDateType.eBool; }

                return eResult;
                /*
                Visual Basic type                       :Common language runtime type structure         :Nominal storage allocation         :Value range
                Boolean                                 :Boolean                                        :Depends on implementing            :True or False
                Byte                                    :Byte                                           :1 byte                             :0 through 255 (unsigned)
                Char (single character)                 :Char                                           :2 bytes                            :0 through 65535 (unsigned)
                Date                                    :DateTime                                       :8 bytes                            :0:00:00 (midnight) on January 1, 0001 through 11:59:59 PM on December 31, 9999
                Decimal                                 :Decimal                                        :16 bytes                           :0 through +/-79,228,162,514,264,337,593,543,950,335 (+/-7.9...E+28) † with no decimal point; 0 through +/-7.9228162514264337593543950335 with 28 places to the right of the decimal;
                Double (double-precision floating-point):Double                                         :8 bytes                            :-1.79769313486231570E+308 through -4.94065645841246544E-324 † for negative values;4.94065645841246544E-324 through 1.79769313486231570E+308 † for positive values
                Integer                                 :Int32                                          :4 bytes                            :-2,147,483,648 through 2,147,483,647 (signed)
                Long (long integer)                     :Int64                                          :8 bytes                            :-9,223,372,036,854,775,808 through 9,223,372,036,854,775,807 (9.2...E+18 †) (signed)
                Object                                  :Object (class)                                 :4 bytes on 32-bit platform         :8 bytes on 64-bit platform;Any type can be stored in a variable of type Object
                SByte                                   :SByte                                          :1 byte                             :-128 through 127 (signed)
                Short (short integer)                   :Int16                                          :2 bytes                            :-32,768 through 32,767 (signed)
                Single (single-precision floating-point):Single                                         :4 bytes                            :-3.4028235E+38 through -1.401298E-45 † for negative values;1.401298E-45 through 3.4028235E+38 † for positive values
                String (variable-length)                :String (class)                                 :Depends on implementing platform   :0 to approximately 2 billion Unicode characters
                UInteger                                :UInt32                                         :4 bytes                            :0 through 4,294,967,295 (unsigned)
                ULong                                   :UInt64                                         :8 bytes                            :0 through 18,446,744,073,709,551,615 (1.8...E+19 †) (unsigned)
                User-Defined (structure)                :(inherits from ValueType)                      :Depends on implementing platform   :Each member of the structure has a range determined by its data type and independent of the ranges of the other members
                UShort                                  :UInt16                                         :2 bytes                            :0 through 65,535 (unsigned) 
                 */

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
                return enumGeneralDateType.eUnknown;
            }
        }
    }
}
