using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using Office = Microsoft.Office.Core;
using KodeMagd.InsertCode;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_SparkLines : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        public FrmInsertCode_SparkLines()
        {
            try
            {
                InitializeComponent();
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

        private void FrmInsertCode_SparkLines_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnSource);
                ClsDefaults.FormatControl(ref btnDestination);

                ClsDefaults.FormatControl(ref lblSource);
                ClsDefaults.FormatControl(ref lblSourceNamedRange);
                ClsDefaults.FormatControl(ref lblSourceRange);
                ClsDefaults.FormatControl(ref lblSourceShtName);

                ClsDefaults.FormatControl(ref txtSourceShtName);
                ClsDefaults.FormatControl(ref txtSourceRange);

                ClsDefaults.FormatControl(ref cmbSourceNamedRange);

                ClsDefaults.FormatControl(ref chkSourceNamedRange);

                ClsDefaults.FormatControl(ref lblDestination);
                ClsDefaults.FormatControl(ref lblDestinationNamedRange);
                ClsDefaults.FormatControl(ref lblDestinationShtName);

                ClsDefaults.FormatControl(ref txtDestinationRange);
                ClsDefaults.FormatControl(ref txtDestinationShtName);
                ClsDefaults.FormatControl(ref cmbDestinationNameRange);
                ClsDefaults.FormatControl(ref cmbDestinationShtName);

                ClsDefaults.FormatControl(ref chkDestinationNamedRange);
                ClsDefaults.FormatControl(ref chkDestinationCreateNamedRange);
                ClsDefaults.FormatControl(ref chkNewSheet);

                ClsDefaults.FormatControl(ref grpDirection);
                ClsDefaults.FormatControl(ref optColumn);
                ClsDefaults.FormatControl(ref optColumnStacked100);
                ClsDefaults.FormatControl(ref optLine);

                ClsDefaults.FormatControl(ref ssStatus);

                chkSourceNamedRange.Checked = false;
                chkNewSheet.Checked = false;
                chkDestinationNamedRange.Checked = false;

                chkSourceNamedRangeChanged();

                fillcmbDesinationShtName();
                chkDestinationNamedRangeChanged();
                chkDestinationCreateNamedRangeCheck();

                chkNewSheetCheck();

                cCodeMapper.readCode();
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

        private void btnClose_Click(object sender, EventArgs e)
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

        private void btnSource_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application app = ClsMisc.ActiveApplication();
                string sSelection = "";
                string sSheetName = "";
                Excel.Range rng;

                Object objResult = app.InputBox("Select Data Source:", "Data Source:", "", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                bool bResult = false;

                if (bool.TryParse(objResult.ToString(), out bResult))
                { 
                    sSelection = "No selection";
                    sSheetName = "";
                }
                else
                {
                    rng = (Excel.Range)objResult;
                    sSelection = rng.Address;
                    sSheetName = rng.Worksheet.Name;
                }
                txtSourceShtName.Text = sSheetName;
                txtSourceRange.Text = sSelection;
                cmbSourceNamedRange.Text = "";

                this.Activate();

                app = null;
                rng = null;
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

        private void chkNamedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                chkSourceNamedRangeChanged();
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

        private void chkSourceNamedRangeChanged()
        {
            try
            {
                if (chkSourceNamedRange.Checked)
                {
                    fillCmbSourceNamedRange();

                    txtSourceRange.Enabled = false;
                    btnSource.Visible = false;
                    cmbSourceNamedRange.Enabled = true;
                }
                else
                {
                    txtSourceRange.Enabled = true;
                    btnSource.Visible = true;
                    cmbSourceNamedRange.Enabled = false;
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

        private void chkDestinationNamedRangeChanged()
        {
            try
            {
                if (chkDestinationNamedRange.Checked)
                {
                    fillCmbDestinationNamedRange();

                    txtDestinationRange.Enabled = false;
                    btnDestination.Visible = false;
                    cmbDestinationNameRange.Enabled = true;
                }
                else
                {
                    txtDestinationRange.Enabled = true;
                    btnDestination.Visible = true;
                    cmbDestinationNameRange.Enabled = false;
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

        public void fillCmbSourceNamedRange()
        {
            try
            {
                //Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                List<string> lstTemp = new List<string>();

                foreach (Excel.Name nmTemp in ClsMisc.ActiveWorkBook().Names)
                { lstTemp.Add(nmTemp.Name); }

                lstTemp.Sort();

                cmbSourceNamedRange.Items.Clear();
                foreach (string sTemp in lstTemp)
                { cmbSourceNamedRange.Items.Add(sTemp); }
                
                lstTemp = null;
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

        public void fillCmbDestinationNamedRange()
        {
            try
            {
                List<string> lstTemp = new List<string>();

                foreach (Excel.Name nmTemp in ClsMisc.ActiveWorkBook().Names)
                { lstTemp.Add(nmTemp.Name); }

                lstTemp.Sort();

                cmbDestinationNameRange.Items.Clear();
                foreach (string sTemp in lstTemp)
                { cmbDestinationNameRange.Items.Add(sTemp); }

                lstTemp = null;
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

        private void btnDestination_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application app = ClsMisc.ActiveApplication();
                string sSelection = "";
                string sShtName = "";
                Excel.Range rng;

                Object objResult = app.InputBox("Select Destination:", "Destination:", "", Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                bool bResult = false;

                if (bool.TryParse(objResult.ToString(), out bResult))
                { sSelection = "No selection"; }
                else
                {
                    rng = (Excel.Range)objResult;
                    sSelection = rng.Address;
                    sShtName = rng.Worksheet.Name;
                }

                txtDestinationRange.Text = sSelection;
                txtDestinationShtName.Text = sShtName;
                cmbDestinationShtName.Text = sShtName;

                validationDestination();

                this.Activate();

                app = null;
                rng = null;
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

        public void checkRangesShape() 
        {
            try
            {
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                Excel.Range rngSource;
                Excel.Range rngDestination;
                ClsInsertCode_SparkLines.enumDirection eDirection = ClsInsertCode_SparkLines.enumDirection.eSpkDir_Unknown;
                bool bIsOk = true;

                if (!optColumn.Checked & !optColumnStacked100.Checked & optLine.Checked)
                { eDirection = ClsInsertCode_SparkLines.enumDirection.eSpkDir_Line; }
                else if (!optColumn.Checked & optColumnStacked100.Checked & !optLine.Checked)
                { eDirection = ClsInsertCode_SparkLines.enumDirection.eSpkDir_ColumnStacked100; }
                else if (optColumn.Checked & !optColumnStacked100.Checked & !optLine.Checked)
                { eDirection = ClsInsertCode_SparkLines.enumDirection.eSpkDir_Column; }
                else
                { 
                    eDirection = ClsInsertCode_SparkLines.enumDirection.eSpkDir_Unknown;
                    bIsOk = false;
                }

                if (bIsOk)
                {
                    switch (eDirection)
                    {
                        case ClsInsertCode_SparkLines.enumDirection.eSpkDir_Column:
                            //check the number columns in the source equals the number of columns in the destination

                            break;
                        case ClsInsertCode_SparkLines.enumDirection.eSpkDir_ColumnStacked100:
                            break;
                        case ClsInsertCode_SparkLines.enumDirection.eSpkDir_Line:
                            //check the number lines in the source equals the number of lines in the destination
                            break;
                        default:
                            bIsOk = false;
                            break;
                    }
                }


                if (bIsOk) 
                { 
                }


                wrk = null;
                rngSource = null;
                rngDestination = null;
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

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                ClsInsertCode_SparkLines cInsertCode_SparkLines = new ClsInsertCode_SparkLines();
                
                if (chkSourceNamedRange.Checked) 
                { cInsertCode_SparkLines.SourceRange = cmbSourceNamedRange.Text; }
                else
                { cInsertCode_SparkLines.SourceRange = txtSourceRange.Text; }

                cInsertCode_SparkLines.SourceShtName = txtSourceShtName.Text;
                cInsertCode_SparkLines.SourceIsNamedRange = chkSourceNamedRange.Checked;

                cInsertCode_SparkLines.DestinationIsNamedRange = chkDestinationNamedRange.Checked;
                cInsertCode_SparkLines.DestinationShtIsNew = chkNewSheet.Checked;
                cInsertCode_SparkLines.DestinationShtName = txtDestinationShtName.Text; 

                cInsertCode_SparkLines.generateCode(ref cCodeMapper);

                configHtmlSummary(ref cInsertCode_SparkLines);
                displayHtmlSummary();

                cInsertCode_SparkLines = null;

                //cInsertCode_SparkLines = null;

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

        private void btnColour_Click(object sender, EventArgs e)
        {

        }

        private void chkNewSheet_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                validationDestination();
                
                //fillcmbDesinationShtName();
                //chkNewSheetCheck();
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

        private void chkNewSheetCheck()
        {
            try
            {
                bool bNew;

                if (chkNewSheet.Checked)
                { bNew = true; }
                else
                { bNew = false; }

                if (bNew)
                {
                    btnDestination.Visible = false;
                    cmbDestinationShtName.Visible = false;
                    txtDestinationShtName.Visible = true;
                    if (chkDestinationNamedRange.Checked)
                    { chkDestinationCreateNamedRange.Enabled = true; }
                    else
                    { chkDestinationCreateNamedRange.Enabled = false; }
                }
                else
                {
                    btnDestination.Visible = false;
                    cmbDestinationShtName.Visible = true;
                    txtDestinationShtName.Visible = false;
                    chkDestinationNamedRange.Enabled = true;
                    chkDestinationCreateNamedRange.Enabled = true;
                }

                validationCheck();
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

        private void validationCheck()
        {
            try
            {
                //destination is a existing range on new sheet.

                checkRangesShape();
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

        private void chkDestinationNamedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                validationDestination();

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

        private void cmbSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                bool bIsFound = false;
                string sShtName = "";
                string sAddress = "";

                foreach (Excel.Name nmRng in ClsMisc.ActiveWorkBook().Names) 
                {
                    if (nmRng.Name == cmbSourceNamedRange.Text) 
                    { 
                        bIsFound = true;
                        sShtName = nmRng.RefersToRange.Worksheet.Name;
                        sAddress = nmRng.RefersToRange.Address;
                    }
                }

                if (bIsFound)
                {
                    txtSourceShtName.Text = sShtName;
                    txtSourceRange.Text = sAddress;
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

        private void fillcmbDesinationShtName() 
        {
            try
            {
                List<string> lstSheets = new List<string>();

                foreach (Excel.Worksheet sht in ClsMisc.ActiveWorkBook().Worksheets)
                { lstSheets.Add(sht.Name); }

                lstSheets.Sort();
                
                cmbDestinationShtName.Items.Clear();
                foreach (string sShtName in lstSheets)
                { cmbDestinationShtName.Items.Add(sShtName); }

                lstSheets = null;
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

        private void cmbDestinationNameRange_TextChanged(object sender, EventArgs e)
        {
            try
            {




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

        private void cmbDestinationNameRange_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                bool bIsFound = false;
                string sShtName = "";
                string sAddress = "";

                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                foreach (Excel.Name nmRng in wrk.Names)
                {
                    if (nmRng.Name == cmbDestinationNameRange.Text)
                    {
                        bIsFound = true;
                        sShtName = nmRng.RefersToRange.Worksheet.Name;
                        sAddress = nmRng.RefersToRange.Address;
                    }
                }

                if (bIsFound)
                {
                    txtDestinationShtName.Text = sShtName;
                    txtDestinationRange.Text = sAddress;
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

        private void chkDestinationCreateNamedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                chkDestinationCreateNamedRangeCheck();
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

    
        private void chkDestinationCreateNamedRangeCheck()
        {
            try
            {
                if (chkDestinationCreateNamedRange.Checked)
                {
                    cmbDestinationNameRange.Visible = false;
                    txtDestinationNameRange.Visible = true;
                }
                else
                {
                    cmbDestinationNameRange.Visible = true;
                    txtDestinationNameRange.Visible = false;
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

        private void validationDestination() 
        {
            try
            {
                fillcmbDesinationShtName();
                chkNewSheetCheck();

                fillCmbDestinationNamedRange();

                if (chkDestinationNamedRange.Checked)
                {
                    btnDestination.Visible = false;
                    cmbDestinationNameRange.Enabled = true;
                    txtDestinationNameRange.Enabled = true;
                    txtDestinationRange.Enabled = false;
                }
                else
                {
                    btnDestination.Visible = true;
                    cmbDestinationNameRange.Enabled = false;
                    txtDestinationNameRange.Enabled = false;
                    txtDestinationRange.Enabled = true;
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

        private void configHtmlSummary(ref ClsInsertCode_SparkLines cInsertCode_SparkLines)
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId;
                int iRowId = 0;

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, 2, "Auto generated code is located.");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = "Description";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = cInsertCode_SparkLines.moduleName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = "Module.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = cInsertCode_SparkLines.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);


                //cInsertCode_SparkLines.DestinationIsNamedRange

                //cInsertCode_SparkLines.SourceIsNamedRange

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 7 }, "Details");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = "Direction";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                switch (cInsertCode_SparkLines.Direction)
                {
                    case ClsInsertCode_SparkLines.enumDirection.eSpkDir_Column:
                        objCell.sText = "Column";
                        objCell.sHiddenText = "A column chart sparkline.";
                        break;
                    case ClsInsertCode_SparkLines.enumDirection.eSpkDir_ColumnStacked100:
                        objCell.sText = "Column Stacked 100";
                        objCell.sHiddenText = "A win/loss chart sparkline.";
                        break;
                    case ClsInsertCode_SparkLines.enumDirection.eSpkDir_Line:
                        objCell.sText = "Line";
                        objCell.sHiddenText = "A line chart sparkline.";
                        break;
                    case ClsInsertCode_SparkLines.enumDirection.eSpkDir_Unknown:
                        objCell.sText = "Unknown";
                        objCell.sHiddenText = "";
                        break;
                    default:
                        objCell.sText = "Unknown";
                        objCell.sHiddenText = "";
                        break;
                }
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                /*
                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.sText = "";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                if (cInsertCode_SparkLines.SourceIsNamedRange)
                { objCell.sText = "Source is Named Range"; }
                else
                { objCell.sText = "Source is NOT Named Range"; }
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                */






                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.sText = "Source";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (cInsertCode_SparkLines.SourceIsNamedRange)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Named Range";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.SourceRange;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Sheet Name";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.SourceShtName;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Range";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.SourceRange;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.sText = "Destination";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (cInsertCode_SparkLines.DestinationIsNamedRange)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Named Range";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.DestinationRange;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Sheet Name";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.DestinationShtName;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.sText = "Range";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.sText = cInsertCode_SparkLines.DestinationRange;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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

        private void displayHtmlSummary()
        {
            try
            {
                string sHtml = cConfigReporter.getHtml();

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "SparkLines");

                frm.ShowDialog(this);

                frm = null;
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
