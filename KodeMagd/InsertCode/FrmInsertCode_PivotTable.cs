using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Misc;
using System.Diagnostics;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_PivotTable : Form
    {
        System.Windows.Forms.ToolTip ttMessage = new System.Windows.Forms.ToolTip();
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        private const string csWarningConnectionString = "Connection Strings for the Excel.Connections can have differences different from ADODB Connection strings\n" +
                                                        "Please try adding an item at the beginning to explain what type of string it is.\n " +
                                                        "For example add \"ODBC;\" to the beginning of a odbc type of connection string " +
                                                        "and it should work as a Excel connection string";

        private ClsInsertCode_PivotTable cInsertCode_PivotTable = new ClsInsertCode_PivotTable();

        public FrmInsertCode_PivotTable()
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

        private void FrmInsertCode_PivotTable_Load(object sender, EventArgs e)
        {
            try
            {
                ClsDefaults.FormatControl(ref ttMessage, true);
                ttMessage.SetToolTip(this, null);
                foreach (Control ctrl in this.Controls)
                { ttMessage.SetToolTip(ctrl, null); }

                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref grpSourceType);

                ClsDefaults.FormatControl(ref optSelectedRange);
                ClsDefaults.FormatControl(ref optNamedRange);
                ClsDefaults.FormatControl(ref optDatabase);

                ClsDefaults.FormatControl(ref lblCommandType);
                ClsDefaults.FormatControl(ref cmbCommandType);

                ClsDefaults.FormatControl(ref lblSource);
                ClsDefaults.FormatControl(ref txtSource);

                ClsDefaults.FormatControl(ref lblConnectionString);
                ClsDefaults.FormatControl(ref txtConnectionString);

                ClsDefaults.FormatControl(ref btnConnectionStringRecent);
                ClsDefaults.FormatControl(ref btnConnectionStringExpend);
                ClsDefaults.FormatControl(ref btnConnectionStringBuild);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnGenerate);

                ClsDefaults.FormatControl(ref lblDestination);
                ClsDefaults.FormatControl(ref chkNewSheet);

                ClsDefaults.FormatControl(ref lblSheetName);
                ClsDefaults.FormatControl(ref cmbSheetName);
                ClsDefaults.FormatControl(ref txtSheetName);

                ClsDefaults.FormatControl(ref lblAddress);
                ClsDefaults.FormatControl(ref txtAddress);
                ClsDefaults.FormatControl(ref dgFields);

                ClsDefaults.FormatControl(ref ssStatus);

                optDatabase.Checked = false;
                optSelectedRange.Checked = true;
                optNamedRange.Checked = false;

                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                chkNewSheet.Checked = true;
                newSheetCheckChange();
                fillDataGridCombo();
                fillComboBoxCommandType();
                fillCmbSheetName();

                txtAddress.Text = "A1";


                cControlPosition.setControl(grpSourceType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(optSelectedRange, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optNamedRange, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optDatabase, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblCommandType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(cmbCommandType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(cmbSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(txtConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(btnConnectionStringRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnConnectionStringExpend, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnConnectionStringBuild, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblDestination, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkNewSheet, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblSheetName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbSheetName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtSheetName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblAddress, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtAddress, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(dgFields, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

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

        private void fillDataGridCombo()
        {
            try
            {
                DataGridViewComboBoxColumn dgvc = (DataGridViewComboBoxColumn)dgFields.Columns[colOrientation.Index];
                string sTemp = "";

                sTemp = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_Aggregate);

                if (!dgvc.Items.Contains(sTemp))
                { dgvc.Items.Add(sTemp); }

                sTemp = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_ColumnTitle);

                if (!dgvc.Items.Contains(sTemp))
                { dgvc.Items.Add(sTemp); }

                sTemp = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_NotUsed);

                if (!dgvc.Items.Contains(sTemp))
                { dgvc.Items.Add(sTemp); }

                sTemp = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_Page);

                if (!dgvc.Items.Contains(sTemp))
                { dgvc.Items.Add(sTemp); }

                sTemp = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_RowTitle);

                if (!dgvc.Items.Contains(sTemp))
                { dgvc.Items.Add(sTemp); }

                foreach (ClsInsertCode_PivotTable.enumPivotTblColOrientation eOrien in Enum.GetValues(typeof(ClsInsertCode_PivotTable.enumPivotTblColOrientation)))
                {
                    sTemp = ClsInsertCode_PivotTable.getNormalName(eOrien);

                    if (!dgvc.Items.Contains(sTemp))
                    { dgvc.Items.Add(sTemp); }
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

        private void newSheetCheckChange() 
        {
            try
            {
                if (chkNewSheet.Checked)
                { 
                    txtSheetName.Visible = true;
                    cmbSheetName.Visible = false;
                }
                else
                {
                    txtSheetName.Visible = false;
                    cmbSheetName.Visible = true;
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

        private void fillCmbSheetName() 
        {
            try
            {
                //Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                List<string> lst = new List<string>();

                foreach (Excel.Worksheet sht in ClsMisc.ActiveWorkBook().Worksheets) 
                { lst.Add(sht.Name); }

                lst.Sort();

                cmbSheetName.Items.Clear();
                foreach (string sSheetName in lst)
                { cmbSheetName.Items.Add(sSheetName); }

                lst = null;
                //wrk = null;
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

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                generate();
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

        private void generate()
        {
            try
            {
                ClsInsertCode_PivotTable cInsertCode_PivotTable = new ClsInsertCode_PivotTable();
                bool bIsOk = true;
                string sMessage = "";

                if (chkNewSheet.Checked == true)
                {
                    if (txtSheetName.Text == null)
                    { cInsertCode_PivotTable.destinationSheetName = ""; }
                    else
                    { cInsertCode_PivotTable.destinationSheetName = txtSheetName.Text.Trim(); }
                }
                else
                {
                    if (cmbSheetName.Text == null)
                    { cInsertCode_PivotTable.destinationSheetName = ""; }
                    else
                    { cInsertCode_PivotTable.destinationSheetName = cmbSheetName.Text.Trim(); }
                }

                if (cInsertCode_PivotTable.destinationSheetName.Trim() == "")
                {
                    bIsOk = false;
                    sMessage = "Please enter a Sheet name for the destination";
                }
                else if (cInsertCode_PivotTable.destinationSheetName.Trim().Length > 31)
                {
                    bIsOk = false;
                    sMessage = "Destination Sheet Name is to long.  Sheet names can only be max 31 charactors.";
                }

                if (bIsOk)
                {
                    cInsertCode_PivotTable.commandType = ClsInsertCode_PivotTable.convertCommandType(cmbCommandType.Text);

                    for (int iRow = 0; iRow < dgFields.RowCount; iRow++)
                    {
                        ClsInsertCode_PivotTable.strPivotField objField = new ClsInsertCode_PivotTable.strPivotField();

                        string sTempName = (string)dgFields[colName.Index, iRow].Value;
                        string sTempOrientation = (string)dgFields[colOrientation.Index, iRow].Value;

                        objField.sName = sTempName;
                        objField.eOrientation = ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_NotUsed;

                        foreach (ClsInsertCode_PivotTable.enumPivotTblColOrientation eOrien in Enum.GetValues(typeof(ClsInsertCode_PivotTable.enumPivotTblColOrientation)))
                        {
                            if (sTempOrientation == ClsInsertCode_PivotTable.getNormalName(eOrien))
                            { objField.eOrientation = eOrien; }
                        }

                        cInsertCode_PivotTable.addPivotField(objField);
                    }
                }

                if (bIsOk) 
                {
                    if (optDatabase.Checked & !optNamedRange.Checked & !optSelectedRange.Checked)
                    { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eDatabase; }
                    else if (!optDatabase.Checked & optNamedRange.Checked & !optSelectedRange.Checked)
                    { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eNamedRange; }
                    else if (!optDatabase.Checked & !optNamedRange.Checked & optSelectedRange.Checked)
                    { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eSelectedRange; }
                    else
                    { 
                        bIsOk = false;
                        sMessage = "Unexpected source type";
                    }
                }

                if (bIsOk) 
                {
                    switch (cInsertCode_PivotTable.sourceType) 
                    {
                        case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                            if (string.IsNullOrEmpty(txtConnectionString.Text))
                            { cInsertCode_PivotTable.connectionString = ""; }
                            else
                            { cInsertCode_PivotTable.connectionString = txtConnectionString.Text; }

                            if (cmbCommandType.Text == ADODB.CommandTypeEnum.adCmdTable.ToString() || cmbCommandType.Text == ADODB.CommandTypeEnum.adCmdTableDirect.ToString())
                            { cInsertCode_PivotTable.sql = cmbSource.Text; }
                            else
                            { cInsertCode_PivotTable.sql = txtSource.Text; }
                            break;
                        case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                            cInsertCode_PivotTable.sql = cmbSource.Text;
                            break;
                        case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                            cInsertCode_PivotTable.sql = txtSource.Text;
                            break;
                    }
                }

                if (bIsOk)
                {
                    cInsertCode_PivotTable.destinationNewSheet = chkNewSheet.Checked;
                    cInsertCode_PivotTable.destinationAddress = txtAddress.Text;
                }

                if (bIsOk)
                {
                    //http://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.workbook.pivottablewizard%28v=vs.80%29.aspx

                    cInsertCode_PivotTable.generateCode(ref cCodeMapper);

                    configHtmlSummary(ref cInsertCode_PivotTable);
                    displayHtmlSummary();

                    cInsertCode_PivotTable = null;
                    this.Close();
                }
                else
                { 
                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    cInsertCode_PivotTable = null;
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

        private void optSelectedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                optChanged();
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

        private void optChanged()
        {
            try
            {
                cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eUnknown;
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                Excel.Worksheet sht = wrk.ActiveSheet;
                Excel.Range rngSelection = wrk.Application.Selection;

                if (optSelectedRange.Checked & !optNamedRange.Checked & !optDatabase.Checked)
                { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eSelectedRange; }

                if (!optSelectedRange.Checked & optNamedRange.Checked & !optDatabase.Checked)
                { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eNamedRange; }

                if (!optSelectedRange.Checked & !optNamedRange.Checked & optDatabase.Checked)
                { cInsertCode_PivotTable.sourceType = ClsInsertCode_PivotTable.enumSourceType.eDatabase; }

                switch (cInsertCode_PivotTable.sourceType)
                {
                    case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                        lblSource.Text = "Source (Address of select Range)";
                        cmbSource.Text = "";
                        txtSource.Text = rngSelection.Address;
                        txtConnectionString.Text = "";
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                        cmbSource.Items.Clear();
                        foreach (string sRange in ClsMisc.namedRanges())
                        { cmbSource.Items.Add(sRange); }
                        lblSource.Text = "Source (Name of Range)";
                        cmbSource.Text = "";
                        txtSource.Text = "";
                        txtConnectionString.Text = "";

                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                        lblSource.Text = "Source (SQL)";
                        cmbSource.Text = "";
                        txtSource.Text = "";
                        txtConnectionString.Text = "";
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eUnknown:
                        break;
                    default:
                        break;
                }

                visibleControlCheck();

                wrk = null;
                sht = null;
                rngSelection = null;
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

        private void visibleControlCheck()
        {
            try
            {
                switch (cInsertCode_PivotTable.sourceType)
                {
                    case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                        cmbSource.Visible = false;
                        txtSource.Visible = true;
                        btnSourceExpand.Visible = true;

                        lblCommandType.Visible = false;
                        cmbCommandType.Visible = false;

                        lblConnectionString.Visible = false;
                        txtConnectionString.Visible = false;
                        btnConnectionStringExpend.Visible = false;
                        btnConnectionStringRecent.Visible = false;
                        btnConnectionStringBuild.Visible = false;

                        txtSource.Enabled = false;
                        
                        ttMessage.SetToolTip(this, "");
                        foreach (Control ctrl in this.Controls)
                        { ttMessage.SetToolTip(ctrl, ""); }

                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                        cmbSource.Visible = true;
                        txtSource.Visible = false;
                        txtSource.Enabled = true;
                        btnSourceExpand.Visible = false;

                        lblCommandType.Visible = false;
                        cmbCommandType.Visible = false;

                        lblConnectionString.Visible = false;
                        txtConnectionString.Visible = false;
                        btnConnectionStringExpend.Visible = false;
                        btnConnectionStringRecent.Visible = false;
                        btnConnectionStringBuild.Visible = false;

                        ttMessage.SetToolTip(this, "");
                        foreach (Control ctrl in this.Controls)
                        { ttMessage.SetToolTip(ctrl, ""); }

                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                        if (cmbCommandType.Text == ADODB.CommandTypeEnum.adCmdTable.ToString()
                            || cmbCommandType.Text == ADODB.CommandTypeEnum.adCmdTableDirect.ToString())
                        {
                            cmbSource.Visible = true;
                            txtSource.Visible = false;
                            txtSource.Enabled = false;
                            btnSourceExpand.Visible = false;
                            fillCmbSource_Tables();
                        }
                        else
                        {
                            cmbSource.Visible = false;
                            txtSource.Visible = true;
                            txtSource.Enabled = true;
                            btnSourceExpand.Visible = true;
                        }

                        lblCommandType.Visible = true;
                        cmbCommandType.Visible = true;

                        lblConnectionString.Visible = true;
                        txtConnectionString.Visible = true;
                        btnConnectionStringExpend.Visible = true;
                        btnConnectionStringRecent.Visible = true;
                        btnConnectionStringBuild.Visible = true;

                        ttMessage.SetToolTip(this, csWarningConnectionString);
                        foreach (Control ctrl in this.Controls)
                        { ttMessage.SetToolTip(ctrl, csWarningConnectionString); }

                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eUnknown:
                        break;
                    default:
                        break;
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

        private void fillCmbSource_Tables()
        {
            try
            {
                ADOX.Catalog cat = new ADOX.Catalog();
                bool bIsOk = true;
                string sMessage = "";

                Debug.WriteLine("");
                Debug.WriteLine("Connection String: " + txtConnectionString.Text);
                ADODB.Connection con = new ADODB.Connection();

                try { con.Open(txtConnectionString.Text); }
                catch
                {
                    bIsOk = false;
                    sMessage = "Can't open connection, please check Connection String";
                }

                try { cat.ActiveConnection = con; }
                catch
                {
                    bIsOk = false;
                    sMessage = "Can't open connection, please check Connection String";
                }

                if (bIsOk)
                {
                    List<string> lst = new List<string>();

                    foreach (ADOX.Table tbl in cat.Tables)
                    { lst.Add(tbl.Name); }

                    lst.Sort();

                    cmbSource.Items.Clear();
                    foreach (string sItem in lst)
                    { cmbSource.Items.Add(sItem); }
                    lblWarning.Text = csWarningConnectionString;
                    
                    lst = null;
                }
                else
                { lblWarning.Text = sMessage; }

                cat = null;
                con = null;
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

        private void optNamedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                optChanged();
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

        private void optDatabase_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                optChanged();
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
                fillFields();
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

        private void fillFields() 
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";
                lblWarning.Text = "";
                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                while (dgFields.Rows.Count > 0)
                { dgFields.Rows.RemoveAt(0); }

                switch (cInsertCode_PivotTable.sourceType) 
                { 
                    case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                        string sName = cmbSource.Text;

                        dgFields.Rows.Clear();

                        if (sName != "")
                        {
                            Excel.Range rng = ClsMisc.getRange(sName); ;

                            if (rng != null)
                            { fillFields(rng); }
                        }
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                        Excel.Range rngSelected = ClsMisc.ActiveRange();

                        if (rngSelected != null)
                        { fillFields(rngSelected); }

                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                        ADODB.Connection con = new ADODB.Connection();
                        ADODB.Command cmd = new ADODB.Command();
                        ADODB.Recordset rst = new ADODB.Recordset();

                        con.ConnectionString = txtConnectionString.Text;
                        con.Mode = ADODB.ConnectModeEnum.adModeRead;

                        if (cInsertCode_PivotTable.commandType == ADODB.CommandTypeEnum.adCmdUnknown) 
                        {
                            bIsOk = false;
                            sMessage = "Please set the Command Type.";
                        }

                        try
                        { con.Open(); }
                        catch (Exception ex)
                        {
                            bIsOk = false;
                            sMessage = "Couldn't open connection with DB.";
                            Debug.Print(sMessage);
                            Debug.Print(DateTime.Now.ToString() + ":" + ex.Message);
                        }

                        if (cInsertCode_PivotTable.commandType == ADODB.CommandTypeEnum.adCmdUnknown)
                        { bIsOk = false; }

                        if (bIsOk)
                        {
                            try
                            {
                                cmd.ActiveConnection = con;
                                cmd.CommandType = cInsertCode_PivotTable.commandType;

                                if (cInsertCode_PivotTable.commandType == ADODB.CommandTypeEnum.adCmdTable || cInsertCode_PivotTable.commandType == ADODB.CommandTypeEnum.adCmdTableDirect)
                                { cmd.CommandText = cmbSource.Text; }
                                else
                                { cmd.CommandText = txtSource.Text; }
                            }
                            catch (Exception ex)
                            {
                                bIsOk = false;
                                sMessage = "Couldn't open connection with DB. (command)";
                                Debug.Print(sMessage);
                                Debug.Print(DateTime.Now.ToString() + ":" + ex.Message);
                            }
                        }

                        if (bIsOk)
                        {
                            try
                            {
                                rst.Open(cmd, Type.Missing, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, -1);

                                ClsSettings cSettings = new ClsSettings();
                                cSettings.addUsedConnectionString(txtConnectionString.Text.Trim());
                                cSettings = null;
                            }
                            catch (Exception ex)
                            {
                                bIsOk = false;
                                sMessage = "Couldn't open connection with DB. (command)";
                                Debug.Print(sMessage);
                                Debug.Print(DateTime.Now.ToString() + ":" + ex.Message);
                            }
                        }

                        if (bIsOk)
                        {
                            int iDataGridRow = 0;

                            while (dgFields.Rows.Count > 0)
                            { dgFields.Rows.RemoveAt(0); }

                            while (dgFields.Rows.Count < rst.Fields.Count)
                            { dgFields.Rows.Add(); }

                            for (int iField = 0; iField < rst.Fields.Count; iField++)
                            //foreach (ADODB.Colum objField in rst.Fields)
                            {
                                string sTemp = rst.Fields[iField].Name;

                                dgFields[colName.Index, iDataGridRow].Value = sTemp;
                                dgFields[colOrientation.Index, iDataGridRow].Value = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_NotUsed);
                                iDataGridRow++;
                            }
                            ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);
                        }
                        else
                        {
                            ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Warning);
                            lblWarning.Text = sMessage;
                        }

                        try
                        {
                            if (rst != null)
                            {
                                if (rst.State != (int)ADODB.ObjectStateEnum.adStateClosed)
                                { rst.Close(); }
                            }

                            rst = null;
                        }
                        catch (Exception ex)
                        {
                            sMessage = "Error trying to close RST.";
                            Debug.Print(sMessage);
                            Debug.Print(DateTime.Now.ToString() + ":" + ex.Message);
                        }

                        con = null;
                        cmd = null;
                        rst = null;
                        break;
                    default:
                        break;
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

        private void fillFields(Excel.Range rng)
        { 
            try
            {
                int iRow;

                while (dgFields.Rows.Count > 0)
                { dgFields.Rows.Remove(dgFields.Rows[0]); }

                for (int iColumnCounter = 1; iColumnCounter <= rng.Columns.Count; iColumnCounter++) 
                {
                    string sName;
                    string sOrientation;

                    if (rng[1, iColumnCounter].value == null)
                    { sName = ""; }
                    else
                    { sName = rng[1, iColumnCounter].value; }

                    dgFields.Rows.Add();
                    iRow = dgFields.Rows.Count - 1;
                    dgFields[colName.Index, iRow].Value = sName;

                    sOrientation = ClsInsertCode_PivotTable.getNormalName(ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_NotUsed);
                    
                    DataGridViewComboBoxColumn dgvc = (DataGridViewComboBoxColumn)dgFields.Columns[colOrientation.Index];

                    if (!dgvc.Items.Contains(sOrientation))
                    { dgvc.Items.Add(sOrientation); }
                    
                    dgFields[colOrientation.Index, iRow].Value = sOrientation;
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

        private void txtSource_TextChanged(object sender, EventArgs e)
        {
            try
            {
                fillFields();
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

        private void btnSourceExpand_Click(object sender, EventArgs e)
        {
            try
            {
                string sQuestion;
                bool bReadOnly;
                switch (cInsertCode_PivotTable.sourceType)
                {
                    case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                        bReadOnly = true;
                        sQuestion = "Selected Range";
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                        bReadOnly = false;
                        sQuestion = "Named Range";
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                        bReadOnly = false;
                        sQuestion = "Database Source";
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eUnknown:
                        bReadOnly = false;
                        sQuestion = "";
                        break;
                    default:
                        bReadOnly = false;
                        sQuestion = "";
                        break;
                }
                string sSource = FrmLargeTextBox.GetString(ClsDefaults.formTitle, sQuestion, txtSource.Text, bReadOnly);

                if (!string.IsNullOrEmpty(sSource))
                { txtSource.Text = sSource; }
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

        private void txtConnectionString_TextChanged(object sender, EventArgs e)
        {
            try
            {
                fillFields();

                visibleControlCheck();
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

        private void btnConnectionStringExpend_Click(object sender, EventArgs e)
        {
            try
            {
                string sQuestion = "Connection String";
                bool bReadOnly = false;

                string sConnectinString = FrmLargeTextBox.GetString(ClsDefaults.formTitle, sQuestion, txtConnectionString.Text, bReadOnly);

                if (!string.IsNullOrEmpty(sConnectinString))
                { txtConnectionString.Text = sConnectinString; }

                fillFields();
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

        private void cmbCommandType_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ADODB.CommandTypeEnum eCmdTypeTemp = ADODB.CommandTypeEnum.adCmdUnknown;
                bool bIsFound = false;

                foreach (ADODB.CommandTypeEnum eTemp in Enum.GetValues(typeof(ADODB.CommandTypeEnum)))
                {
                    if (eTemp.ToString () == cmbCommandType.Text)
                    {
                        eCmdTypeTemp = eTemp;
                        bIsFound = true;
                    }
                }

                if (bIsFound)
                { this.cInsertCode_PivotTable.commandType = eCmdTypeTemp; }
                else
                { this.cInsertCode_PivotTable.commandType = ADODB.CommandTypeEnum.adCmdUnknown; }

                visibleControlCheck();
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

        private void fillComboBoxCommandType() 
        {
            try
            {
                List<string> lstTemp = new List<string>();
                foreach(ADODB.CommandTypeEnum eCmdType in Enum.GetValues(typeof(ADODB.CommandTypeEnum)))
                { lstTemp.Add(eCmdType.ToString()); }

                lstTemp.Sort();

                cmbCommandType.Items.Clear();
                foreach (string sCmdType in lstTemp)
                { cmbCommandType.Items.Add(sCmdType); }

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

        private void dgFields_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                DataGridViewComboBoxCell cellComboBox = (DataGridViewComboBoxCell)dgFields[colOrientation.Index, e.RowIndex];

                List<string> lstTemp = new List<string>();
                foreach (ClsInsertCode_PivotTable.enumPivotTblColOrientation eOrientation in Enum.GetValues(typeof(ClsInsertCode_PivotTable.enumPivotTblColOrientation)))
                {
                    string sItem = ClsInsertCode_PivotTable.getNormalName(eOrientation);
                    lstTemp.Add(sItem);
                }

                lstTemp.Sort();

                cellComboBox.Items.Clear();
                foreach (string sCmdType in lstTemp)
                { cellComboBox.Items.Add(sCmdType); }

                lstTemp = null;
                cellComboBox = null;
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

        private void btnConnectionStringRecent_Click(object sender, EventArgs e)
        {
            try
            {
                recent();
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

        private void recent()
        {
            try
            {
                string sTemp = FrmRecentConnectionStrings.GetString();

                if (sTemp != "")
                { txtConnectionString.Text = sTemp; }
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

        private void chkNewSheet_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                newSheetCheckChange();
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

        private void btnBuildConnectionString_Click(object sender, EventArgs e)
        {
            try
            {
                build();
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

        private void build()
        {
            try
            {
                FrmConnectionString frm = new FrmConnectionString();

                frm.ShowDialog(this);

                string sResult = frm.Result;

                if (sResult != "")
                { txtConnectionString.Text = sResult; }

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

        private void FrmInsertCode_PivotTable_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref grpSourceType);

                cControlPosition.positionControl(ref optSelectedRange);
                cControlPosition.positionControl(ref optNamedRange);
                cControlPosition.positionControl(ref optDatabase);

                cControlPosition.positionControl(ref lblCommandType);
                cControlPosition.positionControl(ref cmbCommandType);

                cControlPosition.positionControl(ref lblSource);
                cControlPosition.positionControl(ref txtSource);
                cControlPosition.positionControl(ref cmbSource);

                cControlPosition.positionControl(ref lblConnectionString);
                cControlPosition.positionControl(ref txtConnectionString);

                cControlPosition.positionControl(ref btnConnectionStringRecent);
                cControlPosition.positionControl(ref btnConnectionStringExpend);
                cControlPosition.positionControl(ref btnConnectionStringBuild);
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnGenerate);

                cControlPosition.positionControl(ref lblDestination);
                cControlPosition.positionControl(ref chkNewSheet);

                cControlPosition.positionControl(ref lblSheetName);
                cControlPosition.positionControl(ref cmbSheetName);
                cControlPosition.positionControl(ref txtSheetName);

                cControlPosition.positionControl(ref lblAddress);
                cControlPosition.positionControl(ref txtAddress);

                cControlPosition.positionControl(ref lblWarning);

                cControlPosition.positionControl(ref dgFields);
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


        private void configHtmlSummary(ref ClsInsertCode_PivotTable cInsertCode_PivotTable)
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
                objCell.bPropHtml = true;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Description";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PivotTable.moduleName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PivotTable.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 7 }, "Details");

                switch (cInsertCode_PivotTable.sourceType)
                {
                    case ClsInsertCode_PivotTable.enumSourceType.eDatabase:
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Connection String";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PivotTable.connectionString.Trim();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Command Type";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PivotTable.commandType.ToString().Trim();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "SQL";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PivotTable.sql.Trim();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eNamedRange:
                    case ClsInsertCode_PivotTable.enumSourceType.eSelectedRange:
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = "Source Range";
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cInsertCode_PivotTable.sql.Trim();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                        break;
                    case ClsInsertCode_PivotTable.enumSourceType.eUnknown:
                        break;
                    default:
                        break;
                }

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                if (cInsertCode_PivotTable.destinationNewSheet == true)
                { objCell.sText = "Destination Sheet (Create new)"; }
                else
                { objCell.sText = "Destination Sheet (Use existing)"; }
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PivotTable.destinationSheetName.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Destination Address";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_PivotTable.destinationAddress.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);



                
                
                /********************
                 *   Fields table   *
                 ********************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 2, 2, 1 }, "Fields");

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "VBA Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Path";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                foreach (ClsInsertCode_PivotTable.strPivotField objField in cInsertCode_PivotTable.pivotFields.Distinct().OrderBy(x => x.sName))
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objField.sName;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    
                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = ClsMiscString.makeValidVarName(objField.sName, "fld");
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = ClsInsertCode_PivotTable.getNormalName(objField.eOrientation);
                    objCell.sHiddenText = "";

                    /*
                    switch (objField.eOrientation)
                    {
                        case ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_Aggregate:
                            objCell.sText = "Aggregate";
                            objCell.sHiddenText = "";
                            break;
                        case ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_ColumnTitle:
                            objCell.sText = "Column Title";
                            objCell.sHiddenText = "";
                            break;
                        case ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_NotUsed:
                            objCell.sText = "Not Used";
                            objCell.sHiddenText = "";
                            break;
                        case ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_Page:
                            objCell.sText = "Top of Page";
                            objCell.sHiddenText = "";
                            break;
                        case ClsInsertCode_PivotTable.enumPivotTblColOrientation.eColOri_RowTitle:
                            objCell.sText = "Row Title";
                            objCell.sHiddenText = "";
                            break;
                    }
                    */ 
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }

                objCell.lstFormatDetails = null;
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

        private void displayHtmlSummary()
        {
            try
            {
                string sHtml = cConfigReporter.getHtml();

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Pivot_Table");

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

        private void FrmInsertCode_PivotTable_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.B)
                    { build(); }

                    if (e.KeyCode == Keys.R)
                    { recent(); }

                    if (e.KeyCode == Keys.G)
                    { generate(); }

                    if (e.KeyCode == Keys.C)
                    { this.Close(); }
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
