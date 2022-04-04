using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using VBA = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;
using KodeMagd.Misc;
using Microsoft.Win32;
using KodeMagd.InsertCode;
using System.Diagnostics;

namespace KodeMagd.Misc
{
    class ClsMisc
    {
        public const char gccChar_SingleQuote = '\'';
        public const char gccChar_DoubleQuote = '"';
        public const string gcsAppName = "Kode Magd";
        public const string gcsAppWebsite = "http://www.kodemagd.de";
        public const int gciError = -1;

        //public enum vbVarType
        //{ 
        //    vbEmpty = 0,
        //    vbNull = 1,
        //    vbInteger = 2,
        //    vbLong = 3,
        //    vbSingle = 4,
        //    vbDouble = 5,
        //    vbCurrency = 6,
        //    vbDate = 7,
        //    vbString = 8,
        //    vbObject = 9,
        //    vbError = 10,
        //    vbBoolean = 11,
        //    vbVariant = 12,
        //    vbDataObject = 13,
        //    vbDecimal = 14,
        //    vbByte = 17,
        //    vbLongLong = 20, //(defined only on implementations that support a LongLong value type)
        //    vbUserDefinedType = 36,
        //    vbArray = 8192,
        //    vbUnknown
        //}

        public static ClsDataTypes.vbVarType getVBA_VarType(string sTypeName)
        {
            try
            {
                ClsDataTypes.vbVarType eResult;

                switch (sTypeName.ToUpper().Trim())
                {
                    case "BOOLEAN":
                        eResult = ClsDataTypes.vbVarType.vbBoolean;
                        break;
                    case "BYTE":
                        eResult = ClsDataTypes.vbVarType.vbByte;
                        break;
                    case "CURRENCY":
                        eResult = ClsDataTypes.vbVarType.vbCurrency;
                        break;
                    case "DATAOBJECT":
                        eResult = ClsDataTypes.vbVarType.vbDataObject;
                        break;
                    case "DATE":
                        eResult = ClsDataTypes.vbVarType.vbDate;
                        break;
                    case "DECIMAL":
                        eResult = ClsDataTypes.vbVarType.vbDecimal;
                        break;
                    case "ERROR":
                        eResult = ClsDataTypes.vbVarType.vbError;
                        break;
                    case "DOUBLE":
                        eResult = ClsDataTypes.vbVarType.vbDouble;
                        break;
                    case "INTEGER":
                        eResult = ClsDataTypes.vbVarType.vbInteger;
                        break;
                    case "LONG":
                        eResult = ClsDataTypes.vbVarType.vbLong;
                        break;
                    case "OBJECT":
                        eResult = ClsDataTypes.vbVarType.vbObject;
                        break;
                    case "SINGLE":
                        eResult = ClsDataTypes.vbVarType.vbSingle;
                        break;
                    case "STRING":
                        eResult = ClsDataTypes.vbVarType.vbString;
                        break;
                    case "VARIANT":
                        eResult = ClsDataTypes.vbVarType.vbVariant;
                        break;
                    default:
                        eResult = ClsDataTypes.vbVarType.vbUnknown;
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
                return ClsDataTypes.vbVarType.vbError;
            }
        }

        public static Excel.Application ActiveApplication()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;

                return app;
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

        public static bool checkCodeIsActive()
        {
            try
            {
                Excel.Application app = ActiveApplication();
                Excel.Workbook wrk = ActiveWorkBook(app);
                VBA.CodePane cpResult = app.VBE.ActiveCodePane;

                bool bResult;

                if (app.VBE.ActiveCodePane == null || app.VBE.ActiveVBProject == null || app.VBE.SelectedVBComponent == null)
                { bResult = false; }
                else
                { bResult = true; }

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

        public static bool checkFormIsActive()
        {
            try
            {
                Excel.Application app = ActiveApplication();
                Excel.Workbook wrk = ActiveWorkBook(app);
                VBA.CodePane cpResult = app.VBE.ActiveCodePane;

                bool bResult;

                if (app.VBE.ActiveVBProject == null || app.VBE.SelectedVBComponent == null)
                { bResult = false; }
                else
                {
                    if (app.VBE.SelectedVBComponent.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
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

        public static VBA.CodePane ActiveVBCodePane()
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                Excel.Application app = ActiveApplication();
                Excel.Workbook wrk = ActiveWorkBook(app);
                VBA.CodePane cpResult = app.VBE.ActiveCodePane;

                if (cpResult == null)
                {
                    /* do something here to make sure that the object is set to something.  */
                }

                if (cSettings.SetFocusActivePane)
                { cpResult.Window.SetFocus(); }

                cSettings = null;

                return cpResult;
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

        public static VBA.CodePane VBCodePane(string sName)
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                Excel.Application app = ActiveApplication();
                Excel.Workbook wrk = ActiveWorkBook(app);

                //VBA.CodePane cpResult = app.VBE.CodePanes.Item(sName);
                VBA.CodePane cpResult = null;

                //wrk.VBProject.VBComponents. 
                bool bIsFound= false;
                int iVbCompIndex = 0;
                for(int iComp = 1; iComp <= wrk.VBProject.VBComponents.Count; iComp++)
                {
                    VBA.VBComponent cTemp = wrk.VBProject.VBComponents.Item(iComp);
                    if (cTemp.Name.Trim().ToUpper() == sName.Trim().ToUpper())
                    { 
                        bIsFound = true; 
                        iVbCompIndex = iComp;
                    }
                }
 
                if (bIsFound)
                { cpResult = wrk.VBProject.VBComponents.Item(iVbCompIndex).CodeModule.CodePane;}


                /*
                for(int iPane = 1; iPane < app.VBE.CodePanes.Count; iPane++)
                {
                    VBA.CodePane cpTemp = app.VBE.CodePanes.Item(iPane);
                    cpTemp.CodeModule.
                }
                */

                if (cpResult == null)
                {
                    /* do something here to make sure that the object is set to something.  */
                }

                if (cSettings.SetFocusActivePane)
                { cpResult.Window.SetFocus(); }

                cSettings = null;

                return cpResult;
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

        public static Excel.Workbook ActiveWorkBook()
        {
            try
            {
                Excel.Workbook wrk = Globals.ThisAddIn.Application.ActiveWorkbook;

                return wrk;
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

        public static Excel.Workbook ActiveWorkBook(Excel.Application app)
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                Excel.Workbook wrk = app.ActiveWorkbook;

                if (wrk == null)
                {




                    /* do something here to make sure that the object is set to something.  */
                }

                cSettings = null;

                return wrk;
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


        public static VBA.VBComponent ActiveVBComponent()
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                Excel.Application app = ActiveApplication();
                Excel.Workbook wrk = ActiveWorkBook(app);
                VBA.VBComponent cmpResult = app.VBE.SelectedVBComponent;

                if (cmpResult == null)
                { /* do something here to make sure that the object is set to something.  */ }

                if (cSettings.SetFocusActivePane)
                { cmpResult.CodeModule.CodePane.Window.SetFocus(); }

                cSettings = null;

                return cmpResult;
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

        public static Excel.Range ActiveRange() 
        { 
            try
            {
                Excel.Application app = ActiveApplication();
                Excel.Range rng = app.Selection;

                return rng;
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

        //public static List<string> splitButNotInQuotes(string sLine, char cDelimiter)
        //{
        //    try
        //    {
        //        const char ccQuote = '"';
        //        List<string> lstResult = new List<string>();
        //        bool bIsString = false;
        //        string sNewLine = "";
        //        bool bIsFirstItteration = true;

        //        /* WARNING: Need to check if the first sTempOuter is a label */

        //        if (sLine.Contains(ccQuote))
        //        {
        //            string sLineNo = getLineNumbers(sLine);
        //            if (sLineNo != "")
        //            { sLine = stripLineNumbers(sLine); }

        //            string sLabel = getLineLabels(sLine);
        //            if (sLabel.Trim() != "")
        //            { sLine = stripLineLabels(sLine); }

        //            foreach (string sTempOuter in sLine.Split(ccQuote))
        //            {
        //                if (!bIsFirstItteration)
        //                { sNewLine += ccQuote; }
        //                else
        //                {
        //                    if (sLineNo != "")
        //                    { sNewLine += sLineNo + " "; }
        //                    if (sLabel != "")
        //                    { sNewLine += sLabel + ": "; }
        //                }
        //                bIsFirstItteration = false;

        //                if (bIsString)
        //                { sNewLine += sTempOuter; }
        //                else
        //                {
        //                    if (sTempOuter.Contains(cDelimiter))
        //                    {
        //                        int iCounter = 1;
        //                        //                                int iItemsCount = sTempOuter.Count(,)
        //                        int iItemsCount = sTempOuter.Count(f => f == cDelimiter);

        //                        foreach (string sTempInner in sTempOuter.Split(cDelimiter))
        //                        {
        //                            /*
        //                             BUG: if last time around this loop then don't add to lstResult
        //                             */
        //                            sNewLine += sTempInner;

        //                            if (iCounter < iItemsCount)
        //                            {
        //                                lstResult.Add(sNewLine);
        //                                sNewLine = "";
        //                            }
        //                            iCounter++;
        //                        }
        //                    }
        //                    else
        //                    { sNewLine += sTempOuter; }
        //                }
        //                bIsString = !bIsString;
        //            }

        //            if (sNewLine != "")
        //            { lstResult.Add(sNewLine); }
        //        }
        //        else
        //        { lstResult.Add(sLine); }

        //        return lstResult;
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //        return null;
        //    }
        //}

        public static List<string> splitButNotInQuotes(ref string sLine, char cDelimiter)
        {
            try
            {
                const char ccQuote = '"';
                List<string> lstResult = new List<string>();
                bool bIsInQuotes = false;
                string sNewLine = "";
                bool bIsFirstItteration = true;
                bool bIsFinished = false;
                string sNextLine = "";
                bool bAddLine = false;
                int iIndent = sLine.Length - sLine.TrimStart().Length;


                /* WARNING: Need to check if the first sTempOuter is a label */

                int iPos = 0;

                if (sLine.Contains(":"))
                {
                    if (sLine.Trim().ToUpper().StartsWith("REM", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //is comment
                        lstResult.Add(sLine);
                        bIsFinished = true;
                    }

                    while (!bIsFinished)
                    {
                        char cCurrChar = sLine[iPos];
                        char cNextChar = ' ';

                        if (sLine.Length > iPos + 1)
                        { cNextChar = sLine[iPos + 1]; }

                        if (cCurrChar == '"')
                        { 
                            bIsInQuotes = !bIsInQuotes;
                            sNextLine += cCurrChar.ToString();
                        }
                        else if (!bIsInQuotes && cCurrChar == '\'')
                        {
                            //is comment
                            lstResult.Add(sNextLine.TrimStart().PadLeft(iIndent + sNextLine.TrimStart().Length, ' ') + ClsMiscString.Right(sLine, sLine.Length - iPos));
                            bIsFinished = true;
                        }
                        else if (!bIsInQuotes && cCurrChar == cDelimiter && cNextChar != '=')
                        {
                            if (bIsFirstItteration)
                            {
                                bIsFirstItteration = false;

                                bool bIsLabel = true;

                                if (sNextLine.Trim().Contains(' ')
                                    | sNextLine.Trim().Contains('(')
                                    | sNextLine.Trim().Contains(')')
                                    | sNextLine.Trim().Contains(','))
                                { bIsLabel = false; }
                                else if (ClsMisc.isKeyword(sNextLine))
                                { bIsLabel = false; }
                                else
                                { bIsLabel = true; }

                                if (bIsLabel)
                                {
                                    bAddLine = false;
                                    sNextLine = "";
                                }
                                else
                                { bAddLine = true; }
                            }
                            else
                            { bAddLine = true; }

                            if (bAddLine)
                            {
                                lstResult.Add(sNextLine.Trim().PadLeft(iIndent + sNextLine.Trim().Length, ' '));
                                sNextLine = "";
                            }
                            else
                            { sNextLine = ""; }
                        }
                        else
                        { sNextLine += cCurrChar.ToString(); }

                        if (iPos >= sLine.Length - 1)
                        { bIsFinished = true; }
                        else
                        { iPos++; }
                    }

                    if (sNextLine.Trim() != "")
                    { lstResult.Add(sNextLine.Trim().PadLeft(iIndent + sNextLine.Trim().Length, ' ')); }
                }
                else
                { lstResult.Add(sLine.Trim().PadLeft(iIndent + sLine.Trim().Length, ' ')); }

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

        public static bool isKeyword(string sText)
        {
            try
            {
                bool bIsFound;

                //List<string> lstReservedWords = new List<string> {"-",  "#CONST",  "#ELSE",  "#ELSEIF",  "#END",  "#IF",  "&",  "&=",  "*",  "*=",  "/",  "/=",
                //                                                    "\\",  "\\=",  "^",  "^=",  "+",  "+=",  "=",  "-=",  "ADD",  "ADDHANDLER",  "ADDRESSOF",  "ALIAS",
                //                                                    "ALL",  "ALPHANUMERIC",  "ALTER",  "AND",  "ANDALSO",  "ANY",  "APPLICATION",  "AS",  "ASC",  "ASSISTANT",  "AUTOINCREMENT",  "AVG",
                //                                                    "BETWEEN",  "BINARY",  "BIT",  "BOOLEAN",  "BY",  "BYREF",  "BYTE",  "BYVAL",  "CALL",  "CASE",  "CATCH",  "CBOOL",
                //                                                    "CBYTE",  "CCHAR",  "CDATE",  "CDBL",  "CDEC",  "CHAR",  "CHARACTER",  "CINT",  "CLASS",  "CLNG",  "COBJ",  "COLUMN",
                //                                                    "COMPACTDATABASE",  "CONST",  "CONSTRAINT",  "CONTAINER",  "CONTINUE",  "COUNT",  "COUNTER",  "CREATE",  "CREATEDATABASE",  "CREATEFIELD",  "CREATEGROUP",  "CREATEINDEX",
                //                                                    "CREATEOBJECT",  "CREATEPROPERTY",  "CREATERELATION",  "CREATETABLEDEF",  "CREATEUSER",  "CREATEWORKSPACE",  "CSBYTE",  "CSHORT",  "CSNG",  "CSTR",  "CTYPE",  "CUINT",
                //                                                    "CULNG",  "CURRENCY",  "CURRENTUSER",  "CUSHORT",  "DATABASE",  "DATE",  "DATETIME",  "DECIMAL",  "DECLARE",  "DEFAULT",  "DELEGATE",  "DELETE",
                //                                                    "DESC",  "DESCRIPTION",  "DIM",  "DIRECTCAST",  "DISALLOW",  "DISTINCT",  "DISTINCTROW",  "DO",  "DOCUMENT",  "DOUBLE",  "DROP",  "EACH",
                //                                                    "ECHO",  "ELSE",  "ELSEIF",  "END",  "ENDIF",  "ENUM",  "EQV",  "ERASE",  "ERROR",  "EVENT",  "EXISTS",  "EXIT",
                //                                                    "FALSE",  "FIELD",  "FIELDS",  "FILLCACHE",  "FINALLY",  "FLOAT",  "FLOAT4",  "FLOAT8",  "FOR",  "FOREIGN",  "FORM",  "FORMS",
                //                                                    "FRIEND",  "FROM",  "FULL",  "FUNCTION",  "GENERAL",  "GET",  "GETOBJECT",  "GETOPTION",  "GETTYPE",  "GLOBAL",  "GOSUB",  "GOTO",
                //                                                    "GOTOPAGE",  "GROUP",  "GROUP BY",  "GUID",  "HANDLES",  "HAVING",  "IDLE",  "IEEEDOUBLE",  "IEEESINGLE",  "IF",  "IGNORE",  "IMP",
                //                                                    "IMPLEMENTS",  "IMPORTS",  "IN",  "INDEX",  "INDEXES",  "INHERITS",  "INNER",  "INSERT",  "INSERTTEXT",  "INT",  "INTEGER",  "INTEGER1",
                //                                                    "INTEGER2",  "INTEGER4",  "INTERFACE",  "INTO",  "IS",  "ISNOT",  "JOIN",  "KEY",  "LASTMODIFIED",  "LEFT",  "LET",  "LEVEL",
                //                                                    "LIB",  "LIKE",  "LOGICAL",  "LOGICAL1",  "LONG",  "LONGBINARY",  "LONGTEXT",  "LOOP",  "MACRO",  "MATCH",  "MAX",  "ME",
                //                                                    "MEMO",  "MIN",  "MOD",  "MODULE",  "MONEY",  "MOVE",  "MUSTINHERIT",  "MUSTOVERRIDE",  "MYBASE",  "MYCLASS",  "NAME",  "NAMESPACE",
                //                                                    "NARROWING",  "NEW",  "NEWPASSWORD",  "NEXT",  "NO",  "NOT",  "NOTE",  "NOTHING",  "NOTINHERITABLE",  "NOTOVERRIDABLE",  "NULL",  "NUMBER",
                //                                                    "NUMERIC",  "OBJECT",  "OF",  "OFF",  "OLEOBJECT",  "ON",  "OPENRECORDSET",  "OPERATOR",  "OPTION",  "OPTIONAL",  "OR",  "ORDER",
                //                                                    "ORELSE",  "ORIENTATION",  "OUTER",  "OVERLOADS",  "OVERRIDABLE",  "OVERRIDES",  "OWNERACCESS",  "PARAMARRAY",  "PARAMETER",  "PARAMETERS",  "PARTIAL",  "PERCENT",
                //                                                    "PIVOT",  "PRIMARY",  "PRIVATE",  "PROCEDURE",  "PROPERTY",  "PROTECTED",  "PUBLIC",  "QUERIES",  "QUERY",  "QUIT",  "RAISEEVENT",  "READONLY",
                //                                                    "REAL",  "RECALC",  "RECORDSET",  "REDIM",  "REFERENCES",  "REFRESH",  "REFRESHLINK",  "REGISTERDATABASE",  "RELATION",  "REM",  "REMOVEHANDLER",  "REPAINT",
                //                                                    "REPAIRDATABASE",  "REPORT",  "REPORTS",  "REQUERY",  "RESUME",  "RETURN",  "RIGHT",  "SBYTE",  "SCREEN",  "SECTION",  "SELECT",  "SET",
                //                                                    "SETFOCUS",  "SETOPTION",  "SHADOWS",  "SHARED",  "SHORT",  "SINGLE",  "SMALLINT",  "SOME",  "SQL",  "STATIC",  "STDEV",  "STDEVP",
                //                                                    "STEP",  "STOP",  "STRING",  "STRUCTURE",  "SUB",  "SUM",  "SYNCLOCK",  "TABLE",  "TABLEDEF",  "TABLEDEFS",  "TABLEID",  "TEXT",
                //                                                    "THEN",  "THROW",  "TIME",  "TIMESTAMP",  "TO",  "TOP",  "TRANSFORM",  "TRUE",  "TRY",  "TRYCAST",  "TYPE",  "TYPEOF",
                //                                                    "UINTEGER",  "ULONG",  "UNION",  "UNIQUE",  "UPDATE",  "USER",  "USHORT",  "USING",  "VALUE",  "VALUES",  "VAR",  "VARBINARY",
                //                                                    "VARCHAR",  "VARIANT",  "VARP",  "VERSION",  "WEND",  "WHEN",  "WHERE",  "WHILE",  "WIDENING",  "WITH",  "WITHEVENTS",  "WORKSPACE",
                //                                                    "WRITEONLY",  "XOR",  "YEAR",  "YES",  "YESNO"};
                List<string> lstSomeKeywords = new List<string>{"#ELSE", "#END", "DEFAULT", "DO", "ELSE", "END", "ENDIF", "EXIT", "LOOP", "NEXT", "REM", "RESUME", "WEND", 
                                                                "APPACTIVATE", "BEEP", "CALL", "CHDIR", "CHDRIVE", "CLOSE", "CONST", "DATE", "DECLARE", "DELETESETTING", 
                                                                "DIM", "DO", "LOOP", "END", "ERASE", "ERROR", "EXIT", "FILECOPY", "FOR", "EACH", "NEXT", "FUNCTION", 
                                                                "GET", "GOSUB", "RETURN", "GOTO", "IF", "THEN", "ELSE", "INPUT", "KILL", "LET", "LINE", "INPUT", "LOAD", 
                                                                "LOCK", "UNLOCK", "MID", "MKDIR", "NAME", "ON", "ERROR", "GOSUB", "GOTO", "OPEN", "OPTION", "BASE", "COMPARE", 
                                                                "EXPLICIT", "PRIVATE", "PRINT", "PRIVATE", "PROPERTY", "GET", "LET", "SET", "PUBLIC", "PUT", "RAISEEVENT", 
                                                                "RANDOMIZE", "REDIM", "REM", "RESET", "RESUME", "RMDIR", "SAVESETTING", "SEEK", "SELECT", "CASE", "SENDKEYS", 
                                                                "SET", "SETATTR", "STATIC", "STOP", "SUB", "TIME", "TYPE", "UNLOAD", "WHILE", "WEND", "WIDTH", "WITH", "WRITE", 
                                                                "ABS", "ARRAY", "ASC", "ATN", "CBOOL", "CBYTE", "CCUR", "CDATE", "CDBL", "CDEC", "CHOOSE", 
                                                                "CHR", "CINT", "CLNG", "COS", "CREATEOBJECT", "CSNG", "CSTR", "CURDIR", "CVAR", "CVDATE", 
                                                                "CVERR", "DATE", "DATEADD", "DATEDIFF", "DATEPART", "DATESERIAL", "DATEVALUE", "DAY", 
                                                                "DIR", "DOEVENTS", "EOF", "ERROR", "EXP", "FILEATTR", "FILEDATETIME", "FILELEN", "FIX", 
                                                                "FORMAT", "FORMATCURRENCY", "FORMATDATETIME", "FORMATNUMBER", "FORMATPERCENT", "FREEFILE", 
                                                                "GETALL", "GETATTR", "GETOBJECT", "GETSETTING", "HEX", "HOUR", "IIF", "INPUT", "INPUTBOX", 
                                                                "INSTR", "INSTRREV", "INT", "ISARRAY", "ISDATE", "ISEMPTY", "ISERROR", "ISMISSING", "ISNULL", 
                                                                "ISNUMERIC", "ISOBJECT", "JOIN", "LBOUND", "LCASE", "LEFT", "LEN", "LOC", "LOF", "LOG", "LTRIM", 
                                                                "MID", "MIDB", "MINUTE", "MONTH", "MONTHNAME", "MSGBOX", "NOW", "OCT", "REPLACE", "RGB", "RIGHT", 
                                                                "RND", "ROUND", "RTRIM", "SECOND", "SEEK", "SGN", "SHELL", "SIN", "SPACE", "SPLIT", "SQR", 
                                                                "STR", "STRCOMP", "STRCONV", "STRING", "STRREVERSE", "SWITCH", "TAB", "TAN", "TIME", "TIMER", 
                                                                "TIMESERIAL", "TIMEVALUE", "TRIM", "TYPENAME", "UBOUND", "UCASE", "VAL", "VARTYPE", "WEEKDAY", 
                                                                "WEEKDAY", "NAME", "YEAR"};

                if (lstSomeKeywords.Exists(x => x.Trim().ToUpper() == sText.Trim().ToUpper()))
                { bIsFound = true; }
                else
                { bIsFound = false; }

                return bIsFound;
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

        public static string stripLineLabels(ref ClsCodeMapper.strLine objLine, ref string sLine)
        {
            try
            {
                bool bHasLineNo;

                string sTemp = stripLineNumbers(sLine);

                if (sTemp.Length == sLine.Length)
                { bHasLineNo = false; }
                else
                { bHasLineNo = true; }

                if (sTemp.Trim().ToUpper().StartsWith(objLine.sLabel.Trim().ToUpper(), StringComparison.CurrentCultureIgnoreCase))
                {
                    int iPos = sTemp.Trim().ToUpper().IndexOf(objLine.sLabel.Trim().ToUpper());

                    sTemp = ClsMiscString.Right(sTemp.Trim(), sTemp.Trim().Length - iPos - objLine.sLabel.Trim().Length);

                    int iIndent = sLine.IndexOf(sTemp);

                    sTemp = sTemp.PadLeft(sTemp.Length + iIndent, ' ');
                }

                if (bHasLineNo)
                { sTemp = objLine.sLineNo + " " + sTemp; }

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

        public static string getLineLabels(string sLine)
        {
            try
            {
                string sResult = "";
                //if the string has a : then check the text before the : 
                //if it's a keyword then it might not be a label else it is a label (label as in goto statement)
                //if (sLine.Contains(':'))
                bool bDoCheck;

                if (ClsMisc.charFirstPosNotInQuotes(ref sLine, ':') != -1)
                {
                    //do more checks here

                    //if (charFirstPosNotInQuotes(sLine, ':') != -1)

                    if (sLine.Contains("'"))
                    {
                        if (ClsMisc.charFirstPosNotInQuotes(ref sLine, '\'') < ClsMisc.charFirstPosNotInQuotes(ref sLine, ':'))
                        { bDoCheck = false; }
                        else
                        { bDoCheck = true; }
                    }
                    else
                    { bDoCheck = true; }

                }
                else
                { bDoCheck = false; }

                if (bDoCheck)
                {
                    sLine = sLine.Trim();
                    string sLineNo = getLineNumbers(sLine);
                    int iPosLineNo = sLine.IndexOf(sLineNo);
                    sLine = ClsMiscString.Right(ref sLine, sLine.Length - sLineNo.Length - iPosLineNo);  

                    //int iPosColon = sLine.IndexOf(':');
                    int iPosColon = ClsMisc.charFirstPosNotInQuotes(ref sLine, ':');

                    string sPrefix = ClsMiscString.Left(ref sLine, iPosColon).Trim();

                    if (!(sPrefix.Trim().Contains(' ') && sPrefix.Trim().Contains('"') && sPrefix.Trim().Contains('\'') && sPrefix.Trim().Contains('+') && sPrefix.Trim().Contains('&')))
                    {
                        List<string> lstSomeKeywords = new List<string>{"#ELSE", "#END", "DEFAULT", "DO", "ELSE", "END", "ENDIF", "EXIT", "LOOP", "NEXT", "REM", "RESUME", "WEND", 
                                                                        "APPACTIVATE", "BEEP", "CALL", "CHDIR", "CHDRIVE", "CLOSE", "CONST", "DATE", "DECLARE", "DELETESETTING", 
                                                                        "DIM", "DO", "LOOP", "END", "ERASE", "ERROR", "EXIT", "FILECOPY", "FOR", "EACH", "NEXT", "FUNCTION", 
                                                                        "GET", "GOSUB", "RETURN", "GOTO", "IF", "THEN", "ELSE", "INPUT", "KILL", "LET", "LINE", "INPUT", "LOAD", 
                                                                        "LOCK", "UNLOCK", "MID", "MKDIR", "NAME", "ON", "ERROR", "GOSUB", "GOTO", "OPEN", "OPTION", "BASE", "COMPARE", 
                                                                        "EXPLICIT", "PRIVATE", "PRINT", "PRIVATE", "PROPERTY", "GET", "LET", "SET", "PUBLIC", "PUT", "RAISEEVENT", 
                                                                        "RANDOMIZE", "REDIM", "REM", "RESET", "RESUME", "RMDIR", "SAVESETTING", "SEEK", "SELECT", "CASE", "SENDKEYS", 
                                                                        "SET", "SETATTR", "STATIC", "STOP", "SUB", "TIME", "TYPE", "UNLOAD", "WHILE", "WEND", "WIDTH", "WITH", "WRITE", 
                                                                        "ABS", "ARRAY", "ASC", "ATN", "CBOOL", "CBYTE", "CCUR", "CDATE", "CDBL", "CDEC", "CHOOSE", 
                                                                        "CHR", "CINT", "CLNG", "COS", "CREATEOBJECT", "CSNG", "CSTR", "CURDIR", "CVAR", "CVDATE", 
                                                                        "CVERR", "DATE", "DATEADD", "DATEDIFF", "DATEPART", "DATESERIAL", "DATEVALUE", "DAY", 
                                                                        "DIR", "DOEVENTS", "EOF", "ERROR", "EXP", "FILEATTR", "FILEDATETIME", "FILELEN", "FIX", 
                                                                        "FORMAT", "FORMATCURRENCY", "FORMATDATETIME", "FORMATNUMBER", "FORMATPERCENT", "FREEFILE", 
                                                                        "GETALL", "GETATTR", "GETOBJECT", "GETSETTING", "HEX", "HOUR", "IIF", "INPUT", "INPUTBOX", 
                                                                        "INSTR", "INSTRREV", "INT", "ISARRAY", "ISDATE", "ISEMPTY", "ISERROR", "ISMISSING", "ISNULL", 
                                                                        "ISNUMERIC", "ISOBJECT", "JOIN", "LBOUND", "LCASE", "LEFT", "LEN", "LOC", "LOF", "LOG", "LTRIM", 
                                                                        "MID", "MIDB", "MINUTE", "MONTH", "MONTHNAME", "MSGBOX", "NOW", "OCT", "REPLACE", "RGB", "RIGHT", 
                                                                        "RND", "ROUND", "RTRIM", "SECOND", "SEEK", "SGN", "SHELL", "SIN", "SPACE", "SPLIT", "SQR", 
                                                                        "STR", "STRCOMP", "STRCONV", "STRING", "STRREVERSE", "SWITCH", "TAB", "TAN", "TIME", "TIMER", 
                                                                        "TIMESERIAL", "TIMEVALUE", "TRIM", "TYPENAME", "UBOUND", "UCASE", "VAL", "VARTYPE", "WEEKDAY", 
                                                                        "WEEKDAY", "NAME", "YEAR"};

                        if (!lstSomeKeywords.Contains(sPrefix.Trim().ToUpper(), StringComparer.OrdinalIgnoreCase))
                        {
                            //Is a Label
                            sResult = ClsMiscString.Right(ref sPrefix, sPrefix.Length);
                        }
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

        public static string getLineNumbers(string sLine)
        {
            try
            {
                string sResults = "";

                if (sLine.Trim() != "")
                {
                    if (Regex.IsMatch(sLine.Trim(), @"\d+"))
                    {
                        //string begins with a number (i.e. a line number)
                        //need to strip it off.
                        Match mtch = Regex.Match(sLine.Trim(), @"^\d+");
                        if (mtch.Groups.Count > 0)
                        {
                            Group grp = mtch.Groups[0];
                            sResults = grp.Value;
                        }
                    }
                }

                return sResults;
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

        public static string stripLineNumbers(string sLine)
        {
            try
            {
                string sLineNo = getLineNumbers(sLine);
                string sResult;

                if (string.IsNullOrEmpty(sLineNo))
                { sResult = sLine; }
                else
                { sResult = ClsMiscString.Right(sLine.Trim(), sLine.Trim().Length - sLineNo.Length); }

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

        public static bool referenceOK(VBA.VBProject vbProj, string sName)
        {
            try
            {
                bool bIsOk = false;

                foreach (VBA.Reference refTemp in vbProj.References)
                {
                    string sRefName = refTemp.Name;

                    if (sRefName.ToUpper().Contains(sName.ToUpper()))
                    { bIsOk = true; }
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

        public static bool isTextDataType(ADODB.DataTypeEnum eType)
        {
            try
            {
                bool bResult;

                if (eType == ADODB.DataTypeEnum.adBSTR
                    || eType == ADODB.DataTypeEnum.adChar
                    || eType == ADODB.DataTypeEnum.adVarChar
                    || eType == ADODB.DataTypeEnum.adVarWChar
                    || eType == ADODB.DataTypeEnum.adWChar)
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
                return true;
            }
        }

        public static bool isTextDataType(ClsDataTypes.vbVarType eType)
        {
            try
            {
                bool bResult;

                if (eType == ClsDataTypes.vbVarType.vbString)
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
                return true;
            }
        }

        public static int getDefaultSize(ADODB.DataTypeEnum eDataType)
        {
            try
            {
                int iResult;

                switch (eDataType)
                {
                    case ADODB.DataTypeEnum.adSmallInt:
                        iResult = 2;
                        break;
                    case ADODB.DataTypeEnum.adInteger:
                        iResult = 4;
                        break;
                    case ADODB.DataTypeEnum.adTinyInt:
                        iResult = 1;
                        break;
                    case ADODB.DataTypeEnum.adUnsignedTinyInt:
                        iResult = 1;
                        break;
                    case ADODB.DataTypeEnum.adUnsignedSmallInt:
                        iResult = 2;
                        break;
                    case ADODB.DataTypeEnum.adUnsignedInt:
                        iResult = 4;
                        break;
                    case ADODB.DataTypeEnum.adBigInt:
                        iResult = 8;
                        break;
                    case ADODB.DataTypeEnum.adUnsignedBigInt:
                        iResult = 8;
                        break;
                    default:
                        iResult = 0;
                        break;
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

        public static bool moduleExists(string sName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                bool bResult = moduleExists(wrk, sName);

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


        public static bool moduleExists(Excel.Workbook wrk, string sName)
        {
            try
            {
                bool bIsFound = false;

                foreach (VBA.VBComponent VBComp in wrk.VBProject.VBComponents)
                {
                    if (VBComp.Name.Trim().ToUpper() == sName.Trim().ToUpper())
                    { bIsFound = true; }
                }

                return bIsFound;
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

        public static string GetFolder(string sFullPath)
        {
            try
            {
                string sFolder = sFullPath;

                if (sFullPath.Contains('\\'))
                {
                    int iPos = sFullPath.LastIndexOf('\\');
                    sFolder = ClsMiscString.Left(ref sFullPath, iPos);
                }

                return sFolder;
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

        public static List<string> CommonDateFormats()
        {
            try
            {
                List<string> lstResult = new List<string>();

                lstResult.Add("mm-dd-yyyy");
                lstResult.Add("dd-mm-yyyy");
                lstResult.Add("yyyymmdd");
                lstResult.Add("yyyy-mm-dd");

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

        public static List<string> listForms()
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                List<string> lstResult = listForms(wrk);
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

        public static List<string> listForms(Excel.Workbook wrk)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                    { lstResult.Add(vbComp.Name); }
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

        public static bool isExistsForm(string sName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                bool bResult = isExistsForm(wrk, sName);
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

        public static bool isExistsForm(Excel.Workbook wrk, string sName)
        {
            try
            {
                bool bIsFound = false;

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                    {
                        if (vbComp.Name == sName)
                        { bIsFound = true; }
                    }
                }

                return bIsFound;
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

        public static bool isExistsModule(string sName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                bool bResult = isExistsModule(wrk, sName);
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

        public static bool isExistsModule(Excel.Workbook wrk, string sName)
        {
            try
            {
                bool bIsFound = false;

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Name == sName)
                    { bIsFound = true; }
                }

                return bIsFound;
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

        public static List<string> getListboxesNames()
        {
            try
            {
                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();
                List<string> lstResult = getListboxesNames(vbComp);

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

        public static List<string> getListboxesNames(VBA.VBComponent vbComp)
        {
            try
            {
                List<string> lstResult = new List<string>();

                if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                {
                    foreach (VBA.Forms.Control ctrl in vbComp.Designer.Controls)
                    {
                        if (ctrl is VBA.Forms.ListBox)
                        { lstResult.Add(ctrl.Name); }
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

        public static List<string> getComboBoxesNames()
        {
            try
            {
                VBA.VBComponent vbComp = ClsMisc.ActiveVBComponent();
                List<string> lstResult = getComboBoxesNames(vbComp);

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

        public static List<string> getForms()
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                List<string> lstResult = getForms(wrk);

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

        public static List<string> getForms(Excel.Workbook wrk)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                    { lstResult.Add(vbComp.Name); }
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

        public static List<string> getComboBoxesNames(VBA.VBComponent vbComp)
        {
            try
            {
                List<string> lstResult = new List<string>();

                if (vbComp.Type == VBA.vbext_ComponentType.vbext_ct_MSForm)
                {
                    foreach (VBA.Forms.Control ctrl in vbComp.Designer.Controls)
                    {
                        if (ctrl is VBA.Forms.ComboBox)
                        { lstResult.Add(ctrl.Name); }
                    }
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

        public static List<string> getListboxesNames(Excel.Workbook wrk, string sFormName)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Name == sFormName)
                    { lstResult = getListboxesNames(vbComp); }
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

        public static List<string> getComboBoxesNames(Excel.Workbook wrk, string sFormName)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (VBA.VBComponent vbComp in wrk.VBProject.VBComponents)
                {
                    if (vbComp.Name == sFormName)
                    { lstResult = getComboBoxesNames(vbComp); }
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

        public static List<string> getListboxesNames(string sFormName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                List<string> lstResult = getListboxesNames(wrk, sFormName);

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
        public static List<string> getComboBoxesNames(string sFormName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                List<string> lstResult = getComboBoxesNames(wrk, sFormName);

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

        public static List<string> namedRanges()
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                List<string> lstResult = namedRanges(wrk);

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

        public static List<string> namedRanges(Excel.Workbook wrk)
        {
            try
            {
                List<string> lstResult = new List<string>();

                foreach (Excel.Name objName in wrk.Application.Names)
                { lstResult.Add(objName.Name); }

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

        public static Excel.Range getRange(string sName)
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                Excel.Range rng = getRange(wrk, sName);

                return rng;
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

        public static Excel.Range getRange(Excel.Workbook wrk, string sName) 
        { 
            try 
            {
                bool bIsFound = false;
                Excel.Range rng = null;

                foreach (Excel.Name nme in wrk.Application.Names) 
                {
                    if (nme.Name == sName)
                    {
                        bIsFound = true;
                        rng = nme.RefersToRange;
                    }
                }

                if (bIsFound)
                { return rng; }
                else
                { return null; }
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

        public static ADODB.ParameterDirectionEnum getAdodbDirection(string sText)
        {
            try
            {
                ADODB.ParameterDirectionEnum eResult = ADODB.ParameterDirectionEnum.adParamUnknown;

                foreach (ADODB.ParameterDirectionEnum eDirection in Enum.GetValues(typeof(ADODB.ParameterDirectionEnum)))
                {
                    if (sText == eDirection.ToString())
                    { eResult = eDirection; }
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

                return ADODB.ParameterDirectionEnum.adParamUnknown;
            }
        }

        public static ADODB.DataTypeEnum getAdodbDataType(string sText)
        {
            try
            {
                ADODB.DataTypeEnum eResult = ADODB.DataTypeEnum.adIUnknown;

                foreach (ADODB.DataTypeEnum eDataType in Enum.GetValues(typeof(ADODB.DataTypeEnum)))
                {
                    if (sText == eDataType.ToString())
                    { eResult = eDataType; }
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

        public static bool checkSheetName(string sName, ref List<string> lstMessages)
        {
            try 
            {
                bool bIsOk = true;

                if (sName.Length > 31) 
                { lstMessages.Add("Name is to long (max 31 char)"); }

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

        public static int getVersionMajor(string sVersion) 
        {
            try
            {
                int iResult = -1;

                List<string> lst = sVersion.Split('.').ToList();

                if (lst.Count == 2)
                {
                    if (!int.TryParse(lst[0], out iResult))
                    { iResult = -1; }
                }
                else
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

                return 0;
            }
        }
        
        public static int getVersionMinor(string sVersion)
        {
            try
            {
                int iResult = -1;

                List<string> lst = sVersion.Split('.').ToList();

                if (lst.Count == 2)
                {
                    if (!int.TryParse(lst[1], out iResult))
                    { iResult = -1; }
                }
                else
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

                return 0;
            }
        }

        public static string joinStrings(List<string> lst, char cDelimiter)
        {
            try 
            {
                string sTemp = "";

                foreach (string sItem in lst)
                { sTemp += sItem + cDelimiter.ToString(); }

                if (sTemp.Length > 0)
                {
                    if (ClsMiscString.Right(ref sTemp, 1) == cDelimiter.ToString()) 
                    { sTemp = ClsMiscString.Left(ref sTemp, sTemp.Length - 1); } 
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

        public static bool shtIsExists(ref Excel.Workbook wrk, string sShtName)
        {
            try
            {
                bool bResult = false;

                foreach (Excel.Worksheet sht in wrk.Worksheets) 
                {
                    if (sht.Name.ToString().ToUpper() == sShtName.ToUpper())
                    { bResult = true; }
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

        public static string newWorksheetName(ref Excel.Workbook wrk, string sSuffix) 
        { 
            try
            {
                string sResult;
                int iCounter = 1;
                string sTemp = sSuffix;

                while (ClsMisc.shtIsExists(ref wrk, sTemp))
                {
                    iCounter++;
                    if (sSuffix.Length + iCounter.ToString().Length + 2 > 31)
                    { sTemp = ClsMiscString.Left(ref sSuffix , 31 - (2 + iCounter.ToString().Length)) + "(" + iCounter.ToString() + ")"; }
                    else
                    { sTemp = sSuffix + "(" + iCounter.ToString() + ")"; }
                }

                sResult = sTemp;

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

        public static string Convert_FormatLineCutMethodology(ClsInsertCode.enumFormatLineCutMethodology eTemp) 
        {
            try
            {
                string sResult;

                switch (eTemp)
                {
                    case ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_None:
                        sResult = "None";
                        break;
                    case ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_AfterXChar:
                        sResult = "Cut line after X char";
                        break;
                    case ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_AtFirstBracket:
                        sResult = "Cut line after first bracket";
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

        public static ClsInsertCode.enumFormatLineCutMethodology Convert_FormatLineCutMethodology(string sTemp)
        {
            try
            {
                ClsInsertCode.enumFormatLineCutMethodology eResult = ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_None;

                foreach (ClsInsertCode.enumFormatLineCutMethodology eTemp in Enum.GetValues(typeof(ClsInsertCode.enumFormatLineCutMethodology)))
                {
                    if (sTemp.ToUpper().Trim() == Convert_FormatLineCutMethodology(eTemp).ToUpper().Trim())
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

                return ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_None;
            }
        }

        public static string Convert_FormatLineVarDim(ClsCodeMapper.enumVarDimType eTemp)
        {
            try
            {
                string sResult;

                switch (eTemp)
                {
                    case ClsCodeMapper.enumVarDimType.eVarDim_Nothing:
                        sResult = "None";
                        break;
                    case ClsCodeMapper.enumVarDimType.eVarDim_InLine:
                        sResult = "Line up Data Type";
                        break;
                    case ClsCodeMapper.enumVarDimType.eVarDim_OneSpace:
                        sResult = "Min space to Data Type";
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

        public static ClsCodeMapper.enumVarDimType Convert_FormatLineVarDim(string sTemp)
        {
            try
            {
                ClsCodeMapper.enumVarDimType eResult = ClsCodeMapper.enumVarDimType.eVarDim_Nothing;

                foreach (ClsCodeMapper.enumVarDimType eTemp in Enum.GetValues(typeof(ClsCodeMapper.enumVarDimType)))
                {
                    if (sTemp.ToUpper().Trim() == Convert_FormatLineVarDim(eTemp).ToUpper().Trim())
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

                return ClsCodeMapper.enumVarDimType.eVarDim_Nothing;
            }
        }

        public static int topRowNo(DataGridViewSelectedCellCollection rng) 
        {
            try
            {
                int iResult = -1;

                if (rng == null)
                { iResult = 1; }
                else
                {
                    foreach (DataGridViewTextBoxCell cell in rng)
                    {
                        if (iResult == -1 || cell.RowIndex < iResult)
                        { iResult = cell.RowIndex; }
                    }

                    if (iResult == -1)
                    { iResult = 1; }
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

                return 1;
            }
        }

        public static bool validVariableNameCheck(string sName, out string sErrorMessage) 
        {
            try 
            {
                bool bIsOK = true;

                sErrorMessage = string.Empty;

                if (sName == null)
                {
                    bIsOK = false;
                    sErrorMessage = "Invalid variable name: can't use empty text.";
                }

                if (sName.Trim() == "")
                {
                    bIsOK = false;
                    sErrorMessage = "Invalid variable name: can't use empty text.";
                }

                if (bIsOK) 
                {
                    if (sName.Length > 255)
                    {
                        bIsOK = false;
                        sErrorMessage = "Invalid variable name: too long.";
                    }
                
                }

                if (bIsOK)
                {
                    Regex rgxCheckFirstDigit = new Regex("[a-z]");

                    if (!rgxCheckFirstDigit.Match(ClsMiscString.Left(sName.Trim(), 1).ToLower()).Success)
                    {
                        bIsOK = false;
                        sErrorMessage = "Invalid variable name: Variable name must start with a letter.";
                    }
                }

                if (bIsOK)
                {
                    Regex rgxCheckAllChars = new Regex(@"[^0-9A-Za-z_]");

                    if (rgxCheckAllChars.Match(ClsMiscString.Left(sName.Trim(), 1).ToLower()).Success)
                    {
                        bIsOK = false;
                        sErrorMessage = "Invalid variable name: Variable name can only have alphanumeric characters or underscores.";
                    }
                }

                if (bIsOK)
                {
                    List<char> lstInvalidChar = new List<char>{'.', '!', '@', '&', '$', '#', '!', '"', '#', '$', '%', '&', 
                                                    '\'', '(', ')', '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', 
                                                    '>', '?', '@', '[', '\\', ']', '^', '_', '`', '`', '{', '|', '}', 
                                                    '~', ' ', '¡', '¢', '£', '¤', '¥', '¦', '§', '¨', '©', 'ª', 'ª', 
                                                    '«', '¬', '®', ',', '¯', '°', '±', '²', '³', '´', 'µ', '¶', '·', 
                                                    '¸', '¹', 'º', '»', '¼', '½', '¾', '¿', 'À', 'Á', 'Â', 'ǁ', 'ǂ', 
                                                    'ǀ', '˛'};

                    foreach (char cInvalidChar in lstInvalidChar)
                    {
                        if (sName.Contains(cInvalidChar))
                        {
                            bIsOK = false;
                            sErrorMessage = "Contains invalid charactor ";
                            switch (cInvalidChar)
                            {
                                case '\'':
                                    sErrorMessage += "<Single Quote>";
                                    break;
                                case '"':
                                    sErrorMessage += "<Double Quote>";
                                    break;
                                case ' ':
                                    sErrorMessage += "<Space>";
                                    break;
                                default:
                                    sErrorMessage += "'" + cInvalidChar.ToString() + "'";
                                    break;
                            }
                        }
                    }
                }

                if (bIsOK)
                {
                    List<string> lstReservedWords = new List<string> {"-",  "#CONST",  "#ELSE",  "#ELSEIF",  "#END",  "#IF",  "&",  "&=",  "*",  "*=",  "/",  "/=",
                                                                    "\\",  "\\=",  "^",  "^=",  "+",  "+=",  "=",  "-=",  "ADD",  "ADDHANDLER",  "ADDRESSOF",  "ALIAS",
                                                                    "ALL",  "ALPHANUMERIC",  "ALTER",  "AND",  "ANDALSO",  "ANY",  "APPLICATION",  "AS",  "ASC",  "ASSISTANT",  "AUTOINCREMENT",  "AVG",
                                                                    "BETWEEN",  "BINARY",  "BIT",  "BOOLEAN",  "BY",  "BYREF",  "BYTE",  "BYVAL",  "CALL",  "CASE",  "CATCH",  "CBOOL",
                                                                    "CBYTE",  "CCHAR",  "CDATE",  "CDBL",  "CDEC",  "CHAR",  "CHARACTER",  "CINT",  "CLASS",  "CLNG",  "COBJ",  "COLUMN",
                                                                    "COMPACTDATABASE",  "CONST",  "CONSTRAINT",  "CONTAINER",  "CONTINUE",  "COUNT",  "COUNTER",  "CREATE",  "CREATEDATABASE",  "CREATEFIELD",  "CREATEGROUP",  "CREATEINDEX",
                                                                    "CREATEOBJECT",  "CREATEPROPERTY",  "CREATERELATION",  "CREATETABLEDEF",  "CREATEUSER",  "CREATEWORKSPACE",  "CSBYTE",  "CSHORT",  "CSNG",  "CSTR",  "CTYPE",  "CUINT",
                                                                    "CULNG",  "CURRENCY",  "CURRENTUSER",  "CUSHORT",  "DATABASE",  "DATE",  "DATETIME",  "DECIMAL",  "DECLARE",  "DEFAULT",  "DELEGATE",  "DELETE",
                                                                    "DESC",  "DESCRIPTION",  "DIM",  "DIRECTCAST",  "DISALLOW",  "DISTINCT",  "DISTINCTROW",  "DO",  "DOCUMENT",  "DOUBLE",  "DROP",  "EACH",
                                                                    "ECHO",  "ELSE",  "ELSEIF",  "END",  "ENDIF",  "ENUM",  "EQV",  "ERASE",  "ERROR",  "EVENT",  "EXISTS",  "EXIT",
                                                                    "FALSE",  "FIELD",  "FIELDS",  "FILLCACHE",  "FINALLY",  "FLOAT",  "FLOAT4",  "FLOAT8",  "FOR",  "FOREIGN",  "FORM",  "FORMS",
                                                                    "FRIEND",  "FROM",  "FULL",  "FUNCTION",  "GENERAL",  "GET",  "GETOBJECT",  "GETOPTION",  "GETTYPE",  "GLOBAL",  "GOSUB",  "GOTO",
                                                                    "GOTOPAGE",  "GROUP",  "GROUP BY",  "GUID",  "HANDLES",  "HAVING",  "IDLE",  "IEEEDOUBLE",  "IEEESINGLE",  "IF",  "IGNORE",  "IMP",
                                                                    "IMPLEMENTS",  "IMPORTS",  "IN",  "INDEX",  "INDEXES",  "INHERITS",  "INNER",  "INSERT",  "INSERTTEXT",  "INT",  "INTEGER",  "INTEGER1",
                                                                    "INTEGER2",  "INTEGER4",  "INTERFACE",  "INTO",  "IS",  "ISNOT",  "JOIN",  "KEY",  "LASTMODIFIED",  "LEFT",  "LET",  "LEVEL",
                                                                    "LIB",  "LIKE",  "LOGICAL",  "LOGICAL1",  "LONG",  "LONGBINARY",  "LONGTEXT",  "LOOP",  "MACRO",  "MATCH",  "MAX",  "ME",
                                                                    "MEMO",  "MIN",  "MOD",  "MODULE",  "MONEY",  "MOVE",  "MUSTINHERIT",  "MUSTOVERRIDE",  "MYBASE",  "MYCLASS",  "NAME",  "NAMESPACE",
                                                                    "NARROWING",  "NEW",  "NEWPASSWORD",  "NEXT",  "NO",  "NOT",  "NOTE",  "NOTHING",  "NOTINHERITABLE",  "NOTOVERRIDABLE",  "NULL",  "NUMBER",
                                                                    "NUMERIC",  "OBJECT",  "OF",  "OFF",  "OLEOBJECT",  "ON",  "OPENRECORDSET",  "OPERATOR",  "OPTION",  "OPTIONAL",  "OR",  "ORDER",
                                                                    "ORELSE",  "ORIENTATION",  "OUTER",  "OVERLOADS",  "OVERRIDABLE",  "OVERRIDES",  "OWNERACCESS",  "PARAMARRAY",  "PARAMETER",  "PARAMETERS",  "PARTIAL",  "PERCENT",
                                                                    "PIVOT",  "PRIMARY",  "PRIVATE",  "PROCEDURE",  "PROPERTY",  "PROTECTED",  "PUBLIC",  "QUERIES",  "QUERY",  "QUIT",  "RAISEEVENT",  "READONLY",
                                                                    "REAL",  "RECALC",  "RECORDSET",  "REDIM",  "REFERENCES",  "REFRESH",  "REFRESHLINK",  "REGISTERDATABASE",  "RELATION",  "REM",  "REMOVEHANDLER",  "REPAINT",
                                                                    "REPAIRDATABASE",  "REPORT",  "REPORTS",  "REQUERY",  "RESUME",  "RETURN",  "RIGHT",  "SBYTE",  "SCREEN",  "SECTION",  "SELECT",  "SET",
                                                                    "SETFOCUS",  "SETOPTION",  "SHADOWS",  "SHARED",  "SHORT",  "SINGLE",  "SMALLINT",  "SOME",  "SQL",  "STATIC",  "STDEV",  "STDEVP",
                                                                    "STEP",  "STOP",  "STRING",  "STRUCTURE",  "SUB",  "SUM",  "SYNCLOCK",  "TABLE",  "TABLEDEF",  "TABLEDEFS",  "TABLEID",  "TEXT",
                                                                    "THEN",  "THROW",  "TIME",  "TIMESTAMP",  "TO",  "TOP",  "TRANSFORM",  "TRUE",  "TRY",  "TRYCAST",  "TYPE",  "TYPEOF",
                                                                    "UINTEGER",  "ULONG",  "UNION",  "UNIQUE",  "UPDATE",  "USER",  "USHORT",  "USING",  "VALUE",  "VALUES",  "VAR",  "VARBINARY",
                                                                    "VARCHAR",  "VARIANT",  "VARP",  "VERSION",  "WEND",  "WHEN",  "WHERE",  "WHILE",  "WIDENING",  "WITH",  "WITHEVENTS",  "WORKSPACE",
                                                                    "WRITEONLY",  "XOR",  "YEAR",  "YES",  "YESNO"};


                    if (lstReservedWords.Exists(x => x.Trim().ToLower() == sName.Trim().ToLower()))
                    {
                        bIsOK = false;
                        sErrorMessage = "Invalid variable name Keyword used";
                    }
                }

                return bIsOK;
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

                sErrorMessage = "Code Crashed";

                return false;
            }
        }

        public static ADODB.CommandTypeEnum getAdoCommandTypeEnum(string sText)
        {
            try
            {
                ADODB.CommandTypeEnum eResult = ADODB.CommandTypeEnum.adCmdUnknown;
                bool bIsFound = false;

                foreach (ADODB.CommandTypeEnum eTemp in Enum.GetValues(typeof(ADODB.CommandTypeEnum)))
                {
                    if (sText.ToUpper().Trim() == eTemp.ToString().ToUpper().Trim())
                    { 
                        eResult = eTemp;
                        bIsFound = true;
                    }
                }

                if (bIsFound == true)
                { return eResult; }
                else
                { return ADODB.CommandTypeEnum.adCmdUnknown; }
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

                return ADODB.CommandTypeEnum.adCmdUnknown;
            }
        }

        public static void removeReturnChar(ref List<string> lstLines)
        {
            try
            {
                Predicate<string> predReturnChar = x => x.Contains('\n') || x.Contains('\r');

                int iIndex = lstLines.FindIndex(predReturnChar);

                while (iIndex != -1)
                {
                    if (lstLines[iIndex].Length == 1)
                    {
                        lstLines[iIndex] = "";
                        lstLines.Insert(iIndex, "");
                    }
                    else
                    {
                        string sTemp = lstLines[iIndex];

                        int iPos = sTemp.ToList().FindIndex(y => y == '\n' || y == '\r');

                        string sBefore = ClsMiscString.Left(ref sTemp, iPos);
                        string sAfter = ClsMiscString.Right(ref sTemp, lstLines[iIndex].Length - iPos - 1);

                        lstLines.Insert(iIndex, sBefore);
                        lstLines[iIndex + 1] = sAfter;
                    }

                    iIndex = lstLines.FindIndex(predReturnChar);
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

        //public static void removeReturnChar(ref ClsLinesOutputRapper cLines)
        //{
        //    try
        //    {
        //        Predicate<string> predReturnChar = x => x.Contains('\n') || x.Contains('\r');

        //        int iIndex = cLines.FindIndex(predReturnChar);

        //        while (iIndex != -1)
        //        {
        //            if (cLines[iIndex].Length == 1)
        //            {
        //                cLines[iIndex] = "";
        //                cLines.Insert(iIndex, "");
        //            }
        //            else
        //            {
        //                string sTemp = cLines[iIndex];

        //                int iPos = sTemp.ToList().FindIndex(y => y == '\n' || y == '\r');

        //                string sBefore = ClsMiscString.Left(ref sTemp, iPos);
        //                string sAfter = ClsMiscString.Right(ref sTemp, cLines[iIndex].Length - iPos - 1);

        //                cLines.Insert(iIndex, sBefore);
        //                cLines[iIndex + 1] = sAfter;
        //            }

        //            iIndex = cLines.FindIndex(predReturnChar);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        public static string replaceReturnCharInQuotedTxtWithConst(string sText)
        {
            try
            {
                while (sText.Contains("\n\r"))
                { sText = sText.Replace("\n\r", "\" & vbCrLf & \""); }

                while (sText.Contains("\r\n"))
                { sText = sText.Replace("\r\n", "\" & vbCrLf & \""); }

                while (sText.Contains('\r'))
                { sText = sText.Replace("\r", "\" & vbCr & \""); }

                while (sText.Contains('\n'))
                { sText = sText.Replace("\n", "\" & vbLf & \""); }

                return sText;
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

        public static int findFirstLineExcludingOptions(ref VBA.CodePane objCodePane)
        {
            try
            {
                int iResult = 1;
                bool bIsFinished = false;
                int iLine = 1;

                while (!bIsFinished) 
                {
                    bool bIgnoreThisLine = false;
                    string sLine = objCodePane.CodeModule.get_Lines(iLine, 1);

                    if (sLine.Trim().ToUpper().StartsWith("OPTION "))
                    { bIgnoreThisLine = true; }

                    if (sLine.Trim().ToUpper().StartsWith("'"))
                    { bIgnoreThisLine = true; }

                    if (sLine.Trim().ToUpper().StartsWith("REM"))
                    { bIgnoreThisLine = true; }

                    if (sLine.Trim() == "")
                    { bIgnoreThisLine = true; }

                    if (sLine.ToUpper().Contains(" LIB ")
                        && sLine.ToUpper().Contains('"')
                        && (sLine.Trim().ToUpper().StartsWith("DECLARE SUB ")
                        || sLine.Trim().ToUpper().StartsWith("DECLARE FUNCTION ")
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC DECLARE SUB ")
                        || sLine.Trim().ToUpper().StartsWith("PUBLIC DECLARE FUNCTION ")
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE DECLARE SUB ")
                        || sLine.Trim().ToUpper().StartsWith("PRIVATE DECLARE FUNCTION ")))
                    { bIgnoreThisLine = true; }

                    if (!bIgnoreThisLine)
                    {
                        bIsFinished = true;
                        if (iLine > 1)
                        { iResult = iLine - 1; }
                        else
                        { iResult = 1; }
                    }

                    if (iLine >= objCodePane.CountOfVisibleLines)
                    { 
                        bIsFinished = true;
                        iResult = iLine;
                    }

                    iLine++;
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

                return 1;
            }
        }
        
        public static void guessDataTypeFromVariableName(ref ClsCodeMapper cCode, string sVaribleName, ref DataGridViewComboBoxCell ComboCellString) 
        {
            try
            {
                string sResult = "";
                string sCmbValue;

                if (ComboCellString.Value == null) 
                { sCmbValue = ""; }
                else
                { sCmbValue = ComboCellString.Value.ToString().ToUpper().Trim(); }
                List<ClsCodeMapper.strVariables> lstVar = cCode.lstVariables(sVaribleName);

                if (lstVar.Count > 0)
                {
                    string sTempDataType = "";
                    
                    ClsDataTypes.enumGeneralDateType eVaribleDataTypeGen = ClsDataTypes.enumGeneralDateType.eUnknown;

                    foreach (ClsCodeMapper.strVariables objVar in lstVar)
                    {
                        eVaribleDataTypeGen = ClsDataTypes.textToGeneralType(sCmbValue);
                    }

                    
                    if (ComboCellString.Value == null)
                    {}
                    else
                    {}


                    //ClsDataTypes.enumGeneralDateType eTextEnteredDataTypeGen = ClsDataTypes.textToGeneralType(objVar.sDatatype);
                }

                //Finish this so that selecting a varible in a combo will return a datatype in another combo

                //return sResult;
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

                //return "";
            }
        }

        public static string getFileName(string sFullPath)
        {
            try
            {
                string sResult = "";

                int iPosStart = getDirectory(sFullPath).Length + 1;

                if (iPosStart < sFullPath.Length)
                { sResult = sFullPath.Substring(iPosStart); }

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


        public static string getDirectory(string sFullPath) 
        {
            try
            {
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                string sPath = sFullPath;
                bool bIsFinished = false;

                while (!bIsFinished)
                {
                    if (fso.FolderExists(sPath))
                    {
                        bIsFinished = true;
                    }
                    else
                    {
                        if (sPath.Contains("/"))
                        {
                            int iPos = sPath.LastIndexOf("/");
                            sPath = sPath.Substring(0, iPos);
                        }
                        else if (sPath.Contains("\\"))
                        {
                            int iPos = sPath.LastIndexOf("\\");
                            sPath = sPath.Substring(0, iPos);
                        }
                        else
                        {
                            sPath = "";
                            bIsFinished = true;
                        }
                    }
                }

                fso = null;

                return sPath;
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

        public static string randomText(int iNoChar)
        {
            try
            {
                string sResult = "";
                Random rnd = new Random();

                for (int iCounter = 0; iCounter < iNoChar; iCounter++)
                {
                    int iAscii = rnd.Next(65, 91);

                    sResult += (char)iAscii;
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

        public static string getRandomNewFileName(string sDirectory, string sExtension)
        {
            try
            {
                string sFileName = randomText(20);

                if (!sDirectory.EndsWith("\\"))
                { sDirectory+="\\"; }

                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();

                while (fso.FileExists(sDirectory + sFileName + sExtension))
                {
                    sFileName = randomText(20);
                }

                fso = null;

                return sFileName + sExtension;
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

        public static string getTempDirectory()
        {
            try
            {
                string sResult = Environment.GetEnvironmentVariable("TEMP");

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

        public static int charFirstPosNotInQuotes(ref string sText, char sChar)
        {
            try
            {
                int iResult = 0;
                const char cDoubleQuote = '"';

                if (sText.Contains(cDoubleQuote))
                {
                    int iPosFirstQuote = sText.IndexOf(cDoubleQuote);
                    int iPosChar = sText.IndexOf(sChar);

                    if (iPosChar < iPosFirstQuote)
                    {
                        iResult = iPosChar;
                    }
                    else
                    {
                        bool bIsInQuotes = false;
                        bool bIsFinished = false;

                        int iPos = 0;

                        while (!bIsFinished)
                        {
                            if (sText[iPos] == sChar)
                            {
                                if (!bIsInQuotes)
                                {
                                    iResult = iPos;
                                    bIsFinished = true;
                                }
                            }
                            else if (sText[iPos] == cDoubleQuote)
                            { bIsInQuotes = !bIsInQuotes; }

                            iPos++;
                            if (iPos == sText.Length)
                            { 
                                bIsFinished = true;
                                iResult = -1;
                            }
                        }
                    }
                }
                else
                { iResult = sText.IndexOf(sChar); }

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

        public static int hexToDec(string sHex)
        {
            try
            {
                int iResult = 0;

                for (int iDigit = 0; iDigit < sHex.Count(); iDigit++)
                {
                    int iMultiplier = (int)Math.Pow(16, (sHex.Count() - iDigit - 1));

                    switch (sHex[iDigit])
                    {
                        case '0':
                            iResult += 0 * iMultiplier;
                            break;
                        case '1':
                            iResult += 1 * iMultiplier;
                            break;
                        case '2':
                            iResult += 2 * iMultiplier;
                            break;
                        case '3':
                            iResult += 3 * iMultiplier;
                            break;
                        case '4':
                            iResult += 4 * iMultiplier;
                            break;
                        case '5':
                            iResult += 5 * iMultiplier;
                            break;
                        case '6':
                            iResult += 6 * iMultiplier;
                            break;
                        case '7':
                            iResult += 7 * iMultiplier;
                            break;
                        case '8':
                            iResult += 8 * iMultiplier;
                            break;
                        case '9':
                            iResult += 9 * iMultiplier;
                            break;
                        case 'A':
                            iResult += 10 * iMultiplier;
                            break;
                        case 'B':
                            iResult += 11 * iMultiplier;
                            break;
                        case 'C':
                            iResult += 12 * iMultiplier;
                            break;
                        case 'D':
                            iResult += 13 * iMultiplier;
                            break;
                        case 'E':
                            iResult += 14 * iMultiplier;
                            break;
                        case 'F':
                            iResult += 15 * iMultiplier;
                            break;
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

                return -1;
            }
        }


        public static string decToHex(int iDec)
        {
            try
            {
                string sResult = "";
                int iRemaineder = iDec;

                for (int iCounter = 4; iCounter >= 0; iCounter--)
                {
                    int iCurrentAmount = (int)Math.Pow((double)16, (double)iCounter);
                    string sThisDigit = "";

                    if (iCurrentAmount <= iRemaineder)
                    {
                        for (int iDigit = 15; iDigit >= 0; iDigit--)
                        {
                            if (iDigit * iCurrentAmount <= iRemaineder)
                            {
                                if (sThisDigit == "")
                                {
                                    switch(iDigit)
                                    {
                                        case 10:
                                            sThisDigit = "A";
                                            break;
                                        case 11:
                                            sThisDigit = "B";
                                            break;
                                        case 12:
                                            sThisDigit = "C";
                                            break;
                                        case 13:
                                            sThisDigit = "D";
                                            break;
                                        case 14:
                                            sThisDigit = "E";
                                            break;
                                        case 15:
                                            sThisDigit = "F";
                                            break;
                                        default:
                                            sThisDigit = iDigit.ToString();
                                            break;
                                    }

                                    iRemaineder -= iDigit * iCurrentAmount;
                                }
                            }


                        }

                    }

                    if (sThisDigit == "" && sResult != "")
                    { sResult += "0"; }

                    if (sThisDigit != "")
                    { sResult += sThisDigit; }
                }



                /*
                string sResult = "";
                int iRunnerTotal = 0;

                for (int iCounter = 4; iCounter >= 0; iCounter--)
                {
                    int iDivider = (int)Math.Pow((double)16, (double)iCounter);
                    int iRemainer = iDec % iDivider;

                    if (iRemainer != iDec && sResult == "")
                    {
                        //iRunnerTotal += iRemainer 
                        switch ((iDec - iRemainer) / 16)
                        {
                            case 0:
                                sResult += "0";
                                break;
                            case 1:
                                sResult += "1";
                                break;
                            case 2:
                                sResult+="2";
                                break;
                            case 3:
                                sResult+="3";
                                break;
                            case 4:
                                sResult+="4";
                                break;
                            case 5:
                                sResult+="5";
                                break;
                            case 6:
                                sResult+="6";
                                break;
                            case 7:
                                sResult+="7";
                                break;
                            case 8:
                                sResult+="8";
                                break;
                            case 9:
                                sResult+="9";
                                break;
                            case 10:
                                sResult+="A";
                                break;
                            case 11:
                                sResult+="B";
                                break;
                            case 12:
                                sResult+="C";
                                break;
                            case 13:
                                sResult+="D";
                                break;
                            case 14:
                                sResult+="E";
                                break;
                            case 15:
                                sResult+="F";
                                break;
                        }
                    }
                }
                */

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

        public static Color convertRGBColour(string sRgb)
        {
            try
            {
                Color cResult = new Color();
                bool bIsOk = true;

                if (sRgb.StartsWith("#") && sRgb.Length == 7)
                { sRgb = sRgb.Substring(1, 6); }
                else if (sRgb.Length != 6)
                { bIsOk = false; }

                if (bIsOk)
                {
                    string sRed = sRgb.Substring(0, 2);
                    string sGreen = sRgb.Substring(2, 2);
                    string sBlue = sRgb.Substring(4, 2);

                    int iRed = hexToDec(sRed);
                    int iGreen = hexToDec(sGreen);
                    int iBlue = hexToDec(sBlue);

                    cResult = Color.FromArgb(255, iRed, iGreen, iBlue);
                }

                return cResult;
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

                return new Color();
            }
        }

        public static string convertColourRGB(Color objColor)
        {
            try
            {
                string sResult = "";

                int iRed = objColor.R;
                int iGreen = objColor.G;
                int iBlue = objColor.B;

                string sRed = decToHex(iRed).PadLeft(2, '0');
                string sGreen = decToHex(iGreen).PadLeft(2, '0');
                string sBlue = decToHex(iBlue).PadLeft(2, '0');

                string sColour = "#" + sRed + sGreen + sBlue;

                return sColour;
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

        /*
                try
                {
                    string sResult = "#";

                    sResult += colTest.R.ToString("X").PadLeft(2, '0');
                    sResult += colTest.G.ToString("X").PadLeft(2, '0');
                    sResult += colTest.B.ToString("X").PadLeft(2, '0');

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
         */
        public static void removeEmptyLinesAtBeginningAndEnd(ref List<ClsCodeMapper.strLine> lstLines)
        {
            try
            {
                Predicate<ClsCodeMapper.strLine> predNoneEmptyLines = x => x.sText_NoComment.Trim() != "" || x.sText_Comment.Trim() != "";

                if (lstLines.Exists(predNoneEmptyLines))
                {
                    int iFirstLine = lstLines.FindAll(predNoneEmptyLines).Min(y => y.iOrder);
                    int iLastLine = lstLines.FindAll(predNoneEmptyLines).Max(y => y.iOrder);

                    lstLines = lstLines.FindAll(x => x.iOrder >= iFirstLine && x.iOrder <= iLastLine);
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
