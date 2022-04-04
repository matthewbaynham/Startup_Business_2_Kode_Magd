using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.InsertCode;
using KodeMagd.Misc;

namespace KodeMagd.InsertCode
{
    class ClsInsertCode_SparkLines : ClsInsertCode
    {
        private enumDirection eDirection;
        private bool bDestinationShtIsNew;
        private string sDestinationShtName;
        private string sSourceShtName;
        private string sDestinationRange;
        private bool bDestinationIsNamedRange;
        private bool bDestinationCreateNamedRange;
        private string sSourceRange;
        private bool bSourceIsNamedRange;
        private string sModuleName;
        private string sFunctionName;

        public enum enumDirection 
        {
            eSpkDir_Column,
            eSpkDir_ColumnStacked100,
            eSpkDir_Line,
            eSpkDir_Unknown
        }

        public string functionName
        {
            get
            {
                try
                { return sFunctionName; }
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

        public string moduleName
        {
            get
            {
                try
                { return sModuleName; }
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

        public enumDirection Direction
        {
            get
            {
                try
                { return eDirection; }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return enumDirection.eSpkDir_Unknown;
                }
            }
            set
            {
                try
                { eDirection = value; }
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

        public bool DestinationShtIsNew
        {
            get
            {
                try
                { return bDestinationShtIsNew; }
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
            set
            {
                try
                { bDestinationShtIsNew = value; }
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

        public string DestinationShtName
        {
            get
            {
                try
                { return sDestinationShtName; }
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
            set
            {
                try
                { sDestinationShtName = value; }
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

        public string SourceShtName
        {
            get
            {
                try
                { return sSourceShtName; }
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
            set
            {
                try
                { sSourceShtName = value; }
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

        public string SourceRange
        {
            get
            {
                try
                { return sSourceRange; }
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
            set
            {
                try
                { sSourceRange = value; }
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

        public bool SourceIsNamedRange
        {
            get
            {
                try
                { return bSourceIsNamedRange; }
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
            set
            {
                try
                { bSourceIsNamedRange = value; }
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

        public string DestinationRange
        {
            get
            {
                try
                { return sDestinationRange; }
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
            set
            {
                try
                { sDestinationRange = value; }
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

        public bool DestinationIsNamedRange
        {
            get
            {
                try
                { return bDestinationIsNamedRange; }
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
            set
            {
                try
                { bDestinationIsNamedRange = value; }
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

        public void generateCode(ref ClsCodeMapper cCodeMapper) 
        {
            try
            {
                /*
                 Possible Source's
                 a) Named Range
                 b) Range Address given
                  
                 Possible Destination
                 a) New Sheet
                    i) create named range
                    ii) Range Address given
                  
                  
                 b) Existing Sheet 
                    i) Named Range
                    ii) Create Named Range
                    iii) Range Address given
                 
                 */

                //ClsCodeMapper cCodeMapper = new ClsCodeMapper();
                //cCodeMapper.readCode();
                ClsSettings cSettings = new ClsSettings();
                List<string> lstCode = new List<string>();
                List<string> lstCodeTop = new List<string>();
                ClsDataTypes cDataTypes = new ClsDataTypes();
                int iIndent = cCodeMapper.cursorCurrentIndentLevel();
                string sWithTemp = "";
                //bool bIsOK;

                sModuleName = cCodeMapper.ModuleDetails.sName;

                if (!cCodeMapper.hasOptionExplicit)
                { lstCodeTop.Add("Option Explicit"); }

                if (!cCodeMapper.hasOptionBase)
                { lstCodeTop.Add("Option Base " + cSettings.defaultOptionBase); }

                if (!cCodeMapper.cursorIsInFunction)
                {
                    sFunctionName = getNextSampleFunctionName(ref cCodeMapper, "_Spark_Lines");
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Public Sub " + sFunctionName);
                    if (cSettings.IndentFirstLevel) { iIndent++; }
                    addErrorHandlerCall(ref lstCode, ref cSettings, iIndent);
                }
                else
                { sFunctionName = cCodeMapper.currentFunctionName(); }

                addTitleComment(ref lstCode, ref cSettings, iIndent);

                /*
                 Dim everything
                 */

                lstCode.Add(cSettings.Indent(iIndent) + "Dim spkGrp As Excel.SparklineGroup");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsFoundDestinationSheet As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsSourceSheet As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsOK As Boolean");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sErrorMessage As String");
                lstCode.Add(cSettings.Indent(iIndent) + "Dim sht as Excel.Worksheet"); 

                if (DestinationIsNamedRange)
                { 
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim bIsFoundDestinationRng as Boolean");
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim rngDestination as Excel.Range");
                }

                /*
                 Check that sheets and named ranges either exist or don't exist
                 */

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = True");
                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Check destination sheet exists");
                ClsCommonCode.sheetExists(ref lstCode, ref iIndent, cSettings, sDestinationShtName, "bIsFoundDestinationSheet", "sht");
                lstCode.Add(cSettings.Indent(iIndent));

                if (bDestinationIsNamedRange)
                {ClsCommonCode.namedRangeExists(ref lstCode, ref iIndent, cSettings, sDestinationRange, "bIsFoundDestinationRng", "rngDestination");}

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "'Check source sheet exists");
                ClsCommonCode.sheetExists(ref lstCode, ref iIndent, cSettings, sSourceShtName, "bIsFoundSourceSheet", "sht");
                lstCode.Add(cSettings.Indent(iIndent));

                if (bSourceIsNamedRange)
                { ClsCommonCode.namedRangeExists(ref lstCode, ref iIndent, cSettings, sSourceRange, "bIsFoundSourceRng", "rngSource"); }
                lstCode.Add(cSettings.Indent(iIndent));

                if (bDestinationShtIsNew)
                {
                    /*
                     Check new sheet doesn't exist
                     */
                    
                    lstCode.Add(cSettings.Indent(iIndent) + "If bIsFoundDestinationSheet Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't create Destination sheet.  Sheet already exists.\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                    lstCode.Add(cSettings.Indent(iIndent));

                    lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "Set sht = ThisWorkbook.Worksheet.Add()");

                    if (UsingWith)
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + "With sht");
                        iIndent++;
                        sWithTemp = "";
                    }
                    else
                    { sWithTemp = "sht"; }

                    lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".Name = \"" + sDestinationShtName + "\"");

                    if (UsingWith)
                    {
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End With");
                        sWithTemp = "";
                    }
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                else
                {
                    /*
                     Check existing sheet does exist
                     */
                    lstCode.Add(cSettings.Indent(iIndent) + "If not bIsFoundDestinationSheet Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't find Destination sheet.\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                    lstCode.Add(cSettings.Indent(iIndent));

                    if (bDestinationIsNamedRange)
                    {
                        if (bDestinationCreateNamedRange)
                        {
                            /*
                             check named range doesn't already exist
                             */
                            lstCode.Add(cSettings.Indent(iIndent) + "If bIsFoundDestinationNamedRange Then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't create Destination Named Range, because it already exists.\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                        }
                        else
                        {
                            /*
                             check destination named range exists
                             */
                            lstCode.Add(cSettings.Indent(iIndent) + "If Not bIsFoundDestinationNamedRange Then");
                            iIndent++;
                            lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = true");
                            lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't output to Destination Named Range, because it doesn't exist.\"");
                            iIndent--;
                            lstCode.Add(cSettings.Indent(iIndent) + "End If");
                        }
                    }
                }

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsFoundSourceSheet Then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't find source Sheet\"");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));

                if (bSourceIsNamedRange)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "If bIsFoundSourceSheet Then");
                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "bIsOk = False");
                    lstCode.Add(cSettings.Indent(iIndent) + "sErrorMessage = \"Can't find source Sheet\"");
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End If");
                }
                else
                {
                    lstCode.Add(cSettings.Indent(iIndent));
                }


                /*
                 if we are going to create the destination Range and it's a named range then...
                 */

                if (bDestinationIsNamedRange)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "Dim rngDestination as Excel.Range");
                    lstCode.Add(cSettings.Indent(iIndent));
                    if (UsingWith)
                    {
                        lstCode.Add(cSettings.Indent(iIndent) + "With ThisWorkbook.Worksheets(\"" + DestinationShtName + "\")");
                        iIndent++;
                        sWithTemp = "";
                    }
                    else
                    { sWithTemp = "ThisWorkbook.Worksheets(\"" + DestinationShtName + "\")"; }

                    iIndent++;
                    lstCode.Add(cSettings.Indent(iIndent) + "Set rngDestination = " + sWithTemp + ".Range(\"" + DestinationRange + "\")");
                    iIndent--;
                    if (UsingWith)
                    {
                        iIndent--;
                        lstCode.Add(cSettings.Indent(iIndent) + "End With");
                        sWithTemp = "";
                    }
                    lstCode.Add(cSettings.Indent(iIndent));
                    lstCode.Add(cSettings.Indent(iIndent) + "Set nmRng = ThisWorkbook.Names.Add(\"" + DestinationRange + "\", rngDestination) 'bugger");
                }

                /*
                 if everything is OK then create the spark group
                 */

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + "If bIsOk then");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Set spkGrp = " + sWithTemp + ".Range(" + sDestinationRange + ").SparklineGroups.Add(xlSparkLine, \"" + sSourceShtName + "!" + sSourceRange + "\")");
                lstCode.Add(cSettings.Indent(iIndent));

                if (UsingWith)
                {
                    lstCode.Add(cSettings.Indent(iIndent) + "With ThisWorkbook.Worksheets(\"" + sSourceShtName + "\")");
                    iIndent++;
                    sWithTemp = "";
                }
                else
                { sWithTemp = "ThisWorkbook.Worksheets(\"" + sSourceShtName + "\")"; }

                string sTemp = cSettings.Indent(iIndent) + sWithTemp + ".Type = ";

                switch(eDirection)
                {
                    case enumDirection.eSpkDir_Column:
                        sTemp += "xlSparkColumn";
                        break;
                    case enumDirection.eSpkDir_ColumnStacked100:
                        sTemp += "xlSparkColumnStacked100";
                        break;
                    case enumDirection.eSpkDir_Line:
                        sTemp += "xlSparkLine";
                        break;
                    case enumDirection.eSpkDir_Unknown:
                        sTemp += "Unknown";
                        break;
                    default:
                        //bIsOK = false;
                        break;
                }
                
                lstCode.Add(sTemp);

                lstCode.Add(cSettings.Indent(iIndent));
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".DisplayBlanksAs = xlInterpolated");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".DisplayHidden = False");
                lstCode.Add(cSettings.Indent(iIndent) + sWithTemp + ".LineWeight = 4");

                if (UsingWith)
                {
                    iIndent--;
                    lstCode.Add(cSettings.Indent(iIndent) + "End With");
                    sWithTemp = "";
                }

                lstCode.Add(cSettings.Indent(iIndent));
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "Else");
                iIndent++;
                lstCode.Add(cSettings.Indent(iIndent) + "Msgbox sErrorMessage");
                iIndent--;
                lstCode.Add(cSettings.Indent(iIndent) + "End If");
                lstCode.Add(cSettings.Indent(iIndent));

                /*
                 finish up
                 */
                if (!cCodeMapper.cursorIsInFunction)
                {
                    addErrorHandlerBody(ref lstCode, ref cSettings, iIndent, ClsCodeMapper.enumFunctionType.eFnType_Sub);
                    if (cSettings.IndentFirstLevel) { iIndent--; }
                    lstCode.Add(cSettings.Indent(iIndent) + "End Sub");
                    lstCode.Add(cSettings.Indent(iIndent));
                }

                this.addCode(ref lstCode);

                if (lstCodeTop.Count > 0)
                {
                    lstCodeTop.Add("");
                    this.addCode(ref lstCodeTop, enumPosition.ePosBeginningAfterOptions);
                }

                cSettings = null;
                //cCodeMapper = null;
                lstCode = null;
                lstCodeTop = null;
                cDataTypes = null;
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
