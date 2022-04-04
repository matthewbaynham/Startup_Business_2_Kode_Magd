using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Misc;
using KodeMagd.InsertCode;
using KodeMagd.Settings;
using KodeMagd.Reporter;


namespace KodeMagd
{
    public partial class FrmInsertCode_Rst : Form
    {
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        string sCtrl_ControlName = "";
        string sCtrl_FieldName = "";
        FrmInsertCode_Rst_PopulateListboxCombobox.enumControlType eCtrl_ControlType = FrmInsertCode_Rst_PopulateListboxCombobox.enumControlType.enumCtrlType_Listbox;

        string sDestinationRangeName = "";
        string sDestinationRangeWrkName = "";
        string sDestinationRangeShtName = "";
        int iDestinationRangeRow = 1;
        int iDestinationRangeColumn = 1;
        ClsInsertCode_Rst.enumDestinationTypeRangeType eDestinationRangeType = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Unknown;

        bool bUseWithStatement = false;
        ADODB.CursorTypeEnum eRstCursorType = ADODB.CursorTypeEnum.adOpenForwardOnly;
        ADODB.LockTypeEnum eRstLockType = ADODB.LockTypeEnum.adLockReadOnly;

        private ClsControlPosition cControlPosition = new ClsControlPosition();

        public FrmInsertCode_Rst()
        {
            try
            {
                InitializeComponent();
                optSql.Checked = true;

                eRstCursorType = ADODB.CursorTypeEnum.adOpenForwardOnly;
                eRstLockType = ADODB.LockTypeEnum.adLockReadOnly;

                VBA.VBComponent VBComp = ClsMisc.ActiveVBComponent();

                if (VBComp != null)
                {
                    cCodeMapper = new ClsCodeMapper();
                    cCodeMapper.readCode(ClsMisc.ActiveVBComponent());
                }
                else
                { cCodeMapper = null; }

                VBComp = null;
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

        private void btnCancel_Click(object sender, EventArgs e)
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

        private bool isOk(ref List<string> lstMessage) 
        {
            try
            {
                bool bIsOk;
                string sMessage = "";

                bIsOk = true;
                string sName = ClsMiscString.makeValidVarName(txtName.Text.Trim());

                if (string.IsNullOrEmpty(sName))
                {
                    bIsOk = false;
                    lstMessage.Add("Please enter a Name.");
                }

                List<char> lstInvalidChar = new List<char>{'.', '!', '@', '&', '$', '#', '!', '"', '#', '$', '%', '&', 
                                                      '\u0027', '(', ')', '*', '+', ',', '-', '.', '/', ':', ';', '<', '=', 
                                                           '>', '?', '@', '[', '\\', ']', '^', '_', '`', '`', '{', '|', '}', 
                                                           '~', ' ', '¡', '¢', '£', '¤', '¥', '¦', '§', '¨', '©', 'ª', 'ª', 
                                                           '«', '¬', '®', ',', '¯', '°', '±', '²', '³', '´', 'µ', '¶', '·', 
                                                           '¸', '¹', 'º', '»', '¼', '½', '¾', '¿', 'À', 'Á', 'Â', 'ǁ', 'ǂ', 
                                                           'ǀ', '˛'};

                if (bIsOk) 
                {
                    sMessage = "";
                    foreach (char cTemp in lstInvalidChar)
                    {
                        if (sName.Contains(cTemp)) 
                        {
                            bIsOk = false;
                            sMessage += " " + cTemp.ToString();
                        }
                    }
                    if (!bIsOk) { lstMessage.Add("VBA variable are not allow to contain " + sMessage); }
                }

                /*
                 * Check the parameters
                 */
                //for (int iParRow = 0; iParRow < dgParameters.RowCount - 1; iParRow++)
                foreach (DataGridViewRow objRow in dgParameters.Rows)
                {
                    int iParRow = objRow.Index; 

                    //if (string.IsNullOrEmpty(dgParameters[ciParamCol_Name, iParRow].Value.ToString().Trim()))
                    if (dgParameters[colName.Index, iParRow].Value == null)
                        {
                        bIsOk = false;
                        lstMessage.Add("Each Parameter must have a name.");
                    }

                    if (dgParameters[colType.Index, iParRow].Value == null)
                    {
                        bIsOk = false;
                        lstMessage.Add("Each Parameter must have a type.");
                    }

                    if (dgParameters[colSize.Index, iParRow].Value == null)
                    {
                        bIsOk = false;
                        lstMessage.Add("Each Parameter must have a size.");
                    }
                    else
                    {
                        int iOutput = 0;

                        if (int.TryParse(dgParameters[colSize.Index, iParRow].Value.ToString(), out iOutput))
                        {
                            if (iOutput <= 0) 
                            {
                                bIsOk = false;
                                lstMessage.Add("Please make sure the size is a positive integer.");
                            }
                        }
                        else
                        {
                            bIsOk = false;
                            lstMessage.Add("Please make sure the size and an integer.");
                        }
                        
                    }

                    if (dgParameters[colValue.Index, iParRow].Value == null)
                    {
                        bIsOk = false;
                        lstMessage.Add("Each Parameter must have a Value.");
                    }
                }

                lstInvalidChar = null;

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
                ClsInsertCode_Rst cInsertRst = new ClsInsertCode_Rst();
                List<string> lstMessages = new List<string>();
                bool bIsOk = isOk(ref lstMessages);

                if (bIsOk)
                {
                    if (chkAddReference.Checked)
                    {
                        FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_ADO, ref ssStatus);

                        if (!frmReference.referenceAlreadySet)
                        { frmReference.ShowDialog(this); }

                        frmReference = null;
                    }
                }

                if (bIsOk)
                {
                    cInsertRst.Name = ClsMiscString.makeValidVarName(txtName.Text);
                    cInsertRst.SQL = txtSource.Text;
                    cInsertRst.ConnectionString = txtConnectionString.Text;
                    cInsertRst.CursorType = eRstCursorType;
                    cInsertRst.LockType = eRstLockType;

                    if (optStoreProcedure.Checked)
                    { cInsertRst.commandType = ADODB.CommandTypeEnum.adCmdStoredProc; }
                    else if (optSql.Checked)
                    { cInsertRst.commandType = ADODB.CommandTypeEnum.adCmdText; }
                    else
                    { cInsertRst.commandType = ADODB.CommandTypeEnum.adCmdUnknown; }

                    //for (int iParRow = 0; iParRow < dgParameters.RowCount; iParRow++)
                    foreach (DataGridViewRow objRow in dgParameters.Rows)
                    {
                        int iParRow = objRow.Index;

                        string sParName = dgParameters[colName.Index, iParRow].Value.ToString();
                        ADODB.DataTypeEnum eDataType = ClsConvert.DataTypeEnum(dgParameters[colType.Index, iParRow].Value.ToString());

                        int iSize = 0;
                        if (!int.TryParse(dgParameters[colSize.Index, iParRow].Value.ToString(), out iSize))
                        { iSize = 0; }
                        
                        string sValue = dgParameters[colValue.Index, iParRow].Value.ToString();

                        bool bAssignVariable;

                        if (dgParameters[ColAssignVariable.Index, iParRow].Value == null)
                        { bAssignVariable = false; }
                        else
                        { bAssignVariable = (bool)dgParameters[ColAssignVariable.Index, iParRow].Value; }

                        cInsertRst.AddParameter(sParName, eDataType, bAssignVariable, iSize, sValue);
                    }

                    ClsInsertCode_Rst.enumDestinationType eDestinationType = ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop;

                    if (!optListboxCombo.Checked & !optRange.Checked & optEmptyLoop.Checked)
                    { eDestinationType = ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop; }
                    else if (!optListboxCombo.Checked & optRange.Checked & !optEmptyLoop.Checked)
                    { eDestinationType = ClsInsertCode_Rst.enumDestinationType.eRstDest_Range; }
                    else if (optListboxCombo.Checked & !optRange.Checked & !optEmptyLoop.Checked)
                    { eDestinationType = ClsInsertCode_Rst.enumDestinationType.eRstDest_ListboxCombo; }
                    else
                    { 
                        bIsOk = false;
                        eDestinationType = ClsInsertCode_Rst.enumDestinationType.eRstDest_Unknown;
                        lstMessages.Add("Unexpected destination type");
                    }

                    if (eDestinationType == ClsInsertCode_Rst.enumDestinationType.eRstDest_ListboxCombo)
                    {
                        cInsertRst.ListboxComboboxName = sCtrl_ControlName;
                        cInsertRst.RstFieldName = sCtrl_FieldName;
                    }

                    cInsertRst.destinationType = eDestinationType;
                    cInsertRst.destinationTypeRangeType = eDestinationRangeType;

                    switch (eDestinationRangeType)
                    {
                        case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds:
                            cInsertRst.destinationRangeColumn = iDestinationRangeColumn;
                            cInsertRst.destinationRangeRow = iDestinationRangeRow;
                            cInsertRst.destinationRangeShtName = sDestinationRangeShtName;
                            cInsertRst.destinationRangeWrkName = sDestinationRangeWrkName;
                            break;
                        case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named:
                            cInsertRst.destinationRangeName = sDestinationRangeName;
                            cInsertRst.destinationRangeShtName = sDestinationRangeShtName;
                            cInsertRst.destinationRangeWrkName = sDestinationRangeWrkName;
                            break;
                    }
                }

                if (bIsOk)
                {
                    ClsSettings cSettings = new ClsSettings();
                    cSettings.addUsedConnectionString(txtConnectionString.Text.Trim());
                    cSettings = null;
                    bIsOk = cInsertRst.isOk(ref lstMessages);
                }

                if (bIsOk)
                {
                    cInsertRst.Insert_RstLoop(ref cCodeMapper);

                    configHtmlSummary(ref cInsertRst);
                    displayHtmlSummary();

                    cInsertRst = null;
                    lstMessages = null;

                    this.Close();
                }
                else
                {
                    lstMessages = lstMessages.Distinct().ToList();
                    string sMessage = "";
                    foreach (string sTemp in lstMessages)
                    { sMessage += sTemp + "\n\r"; }

                    lblWarning.Text = sMessage;
                    lblWarning.Visible = true;

                    MessageBox.Show(sMessage,
                                    ClsDefaults.messageBoxTitle(), 
                                    MessageBoxButtons.OK, 
                                    MessageBoxIcon.Exclamation);
                    
                    lstMessages = null;
                    cInsertRst = null;
                    lstMessages = null;
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

        private void btnOptions_Click(object sender, EventArgs e)
        {
            try
            {
                options();
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

        private void options()
        {
            try
            {

                FrmRstOpenLoopClose_Options frmRstOptions = new FrmRstOpenLoopClose_Options();

                frmRstOptions.LockType = eRstLockType;
                frmRstOptions.CursorType = eRstCursorType;

                frmRstOptions.ShowDialog(this);

                eRstLockType = frmRstOptions.LockType;
                eRstCursorType = frmRstOptions.CursorType;

                frmRstOptions = null;
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

        private void FrmRstOpenLoopClose_Load(object sender, EventArgs e)
        {
            try 
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;
                txtName.Text = ClsDefaults.defaultName;

                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref btnBuildConnectionString);
                ClsDefaults.FormatControl(ref btnCancel);
                ClsDefaults.FormatControl(ref btnDestinationDetails);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnOptions);
                ClsDefaults.FormatControl(ref btnRecent);

                ClsDefaults.FormatControl(ref chkAddReference);

                ClsDefaults.FormatControl(ref txtConnectionString);
                ClsDefaults.FormatControl(ref txtName);
                ClsDefaults.FormatControl(ref txtSource);

                ClsDefaults.FormatControl(ref grpSourceType);
                ClsDefaults.FormatControl(ref grpDestinatin);

                ClsDefaults.FormatControl(ref lblName, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref lblConnectionString, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref lblParameters, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref lblSource, ClsDefaults.enumLabelState.eLbl_normal);
                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                ClsDefaults.FormatControl(ref optSql);
                ClsDefaults.FormatControl(ref optStoreProcedure);
                ClsDefaults.FormatControl(ref optRange);
                ClsDefaults.FormatControl(ref optListboxCombo);
                ClsDefaults.FormatControl(ref optEmptyLoop);

                ClsDefaults.FormatControl(ref dgParameters);

                ClsDefaults.FormatControl(ref ssStatus);

                optStoreProcedure.Checked = false;
                optSql.Checked = true;

                lblSource.Text = "SQL Query";
                txtSource.Text = "select * from <Table Name>";
                chkAddReference.Checked = false;

                optListboxCombo.Checked = false;
                optRange.Checked = false;
                optEmptyLoop.Checked = true;
                checkVisibleDestinationBtn();

                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnBuildConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnCancel, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnDestinationDetails, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOptions, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(chkAddReference, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(txtConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(grpSourceType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(grpDestinatin, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblSource, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(optSql, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optStoreProcedure, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optRange, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optListboxCombo, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optEmptyLoop, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(dgParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
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

        private void dgParameters_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try 
            {
                dgParameters[ColAssignVariable.Index, e.RowIndex].Value = true;

                DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgParameters[colType.Index, e.RowIndex];

                Array arrDataType = Enum.GetValues(typeof(ADODB.DataTypeEnum));
                Array.Sort(arrDataType);

                List<string> lstDataType = new List<string>();

                foreach (ADODB.DataTypeEnum eTemp in arrDataType)
                { lstDataType.Add(eTemp.ToString()); }
                lstDataType.Sort();

                foreach (string sTemp in lstDataType)
                { cellDataType.Items.Add(sTemp); }

                arrDataType = null;
                lstDataType = null;
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

        private void dgParameters_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == colType.Index) 
                {
                    //if the datatype is change to something with a known size (e.g. integer) then we can set the size field as well
                    if (dgParameters[colType.Index, e.RowIndex].Value != null)
                    {
                        if (dgParameters[colType.Index, e.RowIndex].Value.ToString().Trim() != "")
                        {
                            ADODB.DataTypeEnum eDataType = ClsConvert.DataTypeEnum(dgParameters[colType.Index, e.RowIndex].Value.ToString());
                            int iSize = ClsMisc.getDefaultSize(eDataType);
                            dgParameters[colSize.Index, e.RowIndex].Value = iSize;
                        }
                    }
                
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
        
        private void sourseChanged() 
        {
            try
            {
                if (optSql.Checked)
                {
                    lblSource.Text = "select * from <Table Name>";
                    optStoreProcedure.Checked = false;
                }
                else if (optStoreProcedure.Checked)
                {
                    lblSource.Text = "<Store Procedure Name>";
                    optSql.Checked = false;
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

        private void btnRecent_Click(object sender, EventArgs e)
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

        private void dgParameters_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == ColAssignVariable.Index)
                {
                    bool bAssignVariable;
                    int iRow = e.RowIndex;

                    if (iRow >= 0)
                    {
                        string sAssignVariable = "";
                        if (dgParameters[ColAssignVariable.Index, iRow].Value == null)
                        { sAssignVariable = ""; }
                        else
                        { sAssignVariable = dgParameters[ColAssignVariable.Index, iRow].Value.ToString(); }

                        if (!bool.TryParse(sAssignVariable, out bAssignVariable))
                        { bAssignVariable = false; }

                        if (bAssignVariable)
                        {
                            DataGridViewComboBoxCell ComboCellString = new DataGridViewComboBoxCell();

                            List<string> lst = cCodeMapper.variableNames();

                            lst.Sort();
                            ComboCellString.Items.Clear();

                            foreach (string sVarName in lst)
                            { ComboCellString.Items.Add(sVarName); }

                            dgParameters[colValue.Index, iRow] = ComboCellString;
                        }
                        else
                        {
                            DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                            dgParameters[colValue.Index, iRow] = TextCellString;
                        }
                    }
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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                add();
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

        private void add()
        {
            try
            {
                string sText = FrmInputBox.GetString("Parameter", "Please enter a parameter name.");

                if (!string.IsNullOrEmpty(sText)) 
                {
                    int iRow = dgParameters.Rows.Add();

                    dgParameters[colName.Index, iRow].Value = sText;

                    this.ActiveControl = dgParameters;
                    dgParameters.CurrentCell = dgParameters[colName.Index, iRow];
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

        private void optStoreProcedure_Click(object sender, EventArgs e)
        {
            try
            {
                optSql.Checked = false;
                lblSource.Text = "Stored Procedure Name";
                txtSource.Text = "Stored Procedure Name";
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

        private void optSql_Click(object sender, EventArgs e)
        {
            try
            {
                optStoreProcedure.Checked = false;
                lblSource.Text = "SQL Query";
                txtSource.Text = "select * from <Table Name>";
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

        private void btnDestinationDetails_Click(object sender, EventArgs e)
        {
            try
            {
                destinationDetails();
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

        private void destinationDetails()
        {
            try
            {
                ClsInsertCode_Rst.enumDestinationType eDestination = ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop;
                bool bIsOk = true;
                string sMessage = "";

                if (optListboxCombo.Checked & !optEmptyLoop.Checked & !optRange.Checked)
                { eDestination = ClsInsertCode_Rst.enumDestinationType.eRstDest_ListboxCombo; }
                else if (!optListboxCombo.Checked & optEmptyLoop.Checked & !optRange.Checked)
                { eDestination = ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop; }
                else if (!optListboxCombo.Checked & !optEmptyLoop.Checked & optRange.Checked)
                { eDestination = ClsInsertCode_Rst.enumDestinationType.eRstDest_Range; }
                else
                { 
                    bIsOk = false; 
                    sMessage = "Unexpected internal variable";
                }

                if (bIsOk)
                {
                    switch (eDestination)
                    {
                        case ClsInsertCode_Rst.enumDestinationType.eRstDest_Range:
                            FrmInsertCode_Rst_Range frmRange = new FrmInsertCode_Rst_Range();

                            if (string.IsNullOrEmpty(sDestinationRangeWrkName))
                            { sDestinationRangeWrkName = ClsMisc.ActiveWorkBook().Name; }

                            if (string.IsNullOrEmpty(sDestinationRangeShtName))
                            { sDestinationRangeShtName = ClsMisc.ActiveRange().Worksheet.Name; }

                            frmRange.shtName = sDestinationRangeShtName;
                            frmRange.wrkName = sDestinationRangeWrkName;
                            frmRange.namedRange = sDestinationRangeName;
                            frmRange.rangeType = eDestinationRangeType;
                            frmRange.row = iDestinationRangeRow;
                            frmRange.column = iDestinationRangeColumn;

                            frmRange.ShowDialog(this);

                            sDestinationRangeShtName = frmRange.shtName;
                            sDestinationRangeWrkName = frmRange.wrkName;
                            sDestinationRangeName = frmRange.namedRange;
                            eDestinationRangeType = frmRange.rangeType;
                            iDestinationRangeRow = frmRange.row;
                            iDestinationRangeColumn = frmRange.column;

                            frmRange = null;
                            break;
                        case ClsInsertCode_Rst.enumDestinationType.eRstDest_ListboxCombo:
                            FrmInsertCode_Rst_PopulateListboxCombobox frmControl = new FrmInsertCode_Rst_PopulateListboxCombobox();

                            frmControl.controlName = sCtrl_ControlName;
                            //frmControl.fieldName = sCtrl_FieldName;
                            frmControl.ControlType = eCtrl_ControlType;

                            frmControl.ShowDialog(this);

                            sCtrl_ControlName = frmControl.controlName;
                            sCtrl_FieldName = frmControl.fieldName;
                            eCtrl_ControlType = frmControl.ControlType;

                            frmControl = null;
                            break;
                        case ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop:
                            MessageBox.Show("No details", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            break;
                    }
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

        private void FrmInsertCode_Rst_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnRemove);
                cControlPosition.positionControl(ref btnBuildConnectionString);
                cControlPosition.positionControl(ref btnCancel);
                cControlPosition.positionControl(ref btnDestinationDetails);
                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnOptions);
                cControlPosition.positionControl(ref btnRecent);

                cControlPosition.positionControl(ref chkAddReference);

                cControlPosition.positionControl(ref txtConnectionString);
                cControlPosition.positionControl(ref txtName);
                cControlPosition.positionControl(ref txtSource);

                cControlPosition.positionControl(ref grpSourceType);
                cControlPosition.positionControl(ref grpDestinatin);

                cControlPosition.positionControl(ref lblName);
                cControlPosition.positionControl(ref lblConnectionString);
                cControlPosition.positionControl(ref lblParameters);
                cControlPosition.positionControl(ref lblSource);
                cControlPosition.positionControl(ref lblWarning);

                cControlPosition.positionControl(ref optSql);
                cControlPosition.positionControl(ref optStoreProcedure);
                cControlPosition.positionControl(ref optRange);
                cControlPosition.positionControl(ref optListboxCombo);
                cControlPosition.positionControl(ref optEmptyLoop);

                cControlPosition.positionControl(ref dgParameters);
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

        private void configHtmlSummary(ref ClsInsertCode_Rst cInsertRst)
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
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 5, 1 }, "Auto generated code is located.");

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
                objCell.sText = "Module.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertRst.moduleName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertRst.functionName;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Connection String.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertRst.ConnectionString;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);


                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Destination Type.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                switch (cInsertRst.destinationType)
                {
                    case ClsInsertCode_Rst.enumDestinationType.eRstDest_ListboxCombo:
                        objCell.sText = "List box or Combo box";
                        break;
                    case ClsInsertCode_Rst.enumDestinationType.eRstDest_EmptyLoop:
                        objCell.sText = "Empty Loop";
                        break;
                    case ClsInsertCode_Rst.enumDestinationType.eRstDest_Range:
                        switch (cInsertRst.destinationRangeType)
                        {
                            case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds:
                                objCell.sText = "Range defined by coordinates";
                                break;
                            case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named:
                                objCell.sText = "Named Range";
                                break;
                            case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Unknown:
                                objCell.sText = "Range";
                                break;
                            default:
                                objCell.sText = "Range";
                                break;
                        }
                        break;
                    case ClsInsertCode_Rst.enumDestinationType.eRstDest_Unknown:
                        objCell.sText = "Unknown";
                        break;
                    default:
                        objCell.sText = "Unknown";
                        break;
                }
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);










                if (cInsertRst.parameters.Count == 0)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameters";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "No Parameters";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {

                    /***************
                     *   A table   *
                     ***************/
                    cConfigReporter.TableAddNew(out iTableId, new List<int> { 5, 1, 1, 1, 1, 1 }, "Parameters");

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
                    objCell.sText = "Using Variable";
                    objCell.sHiddenText = "The value is assigned to the parameter through assigning a variable rather than hardcoding the value in the code.";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Database Data Type";
                    objCell.sHiddenText = "Relate to the data type used by the database server";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Variable Data Type";
                    objCell.sHiddenText = "Relates to Datatypes in VBA.";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Size";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Value";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    foreach (ClsInsertCode_Rst.strParameter objParameter in cInsertRst.parameters.Distinct().OrderBy(x => x.sName.Trim().ToUpper()))
                    {
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.sName;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (objParameter.bAssignVariable)
                        { objCell.sText = "Variable"; }
                        else
                        { objCell.sText = "Hardcoded"; }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.eDataType.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = cDataTypes.getName(objParameter.eVarType);
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.Size.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.Value;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    }
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Create_Recordset");

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

        private void FrmInsertCode_Rst_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.B)
                    { build(); }

                    if (e.KeyCode == Keys.D)
                    { destinationDetails(); }

                    if (e.KeyCode == Keys.O)
                    { options(); }

                    if (e.KeyCode == Keys.R)
                    { remove(); }

                    if (e.KeyCode == Keys.E)
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

        private void btnRemove_Click(object sender, EventArgs e)
        {
            try
            {
                remove();
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

        private void remove()
        {
            try
            {
                //string sText = FrmInputBox.GetString("Parameter", "Please enter a parameter name.");

                //if (!string.IsNullOrEmpty(sText))
                //{
                //    int iRow = dgParameters.Rows.Add();

                //    dgParameters[colName.Index, iRow].Value = sText;

                //    this.ActiveControl = dgParameters;
                //    dgParameters.CurrentCell = dgParameters[colName.Index, iRow];
                //}
                bool bIsOk = true;
                string sMessage = "";

                if (dgParameters.CurrentCell == null)
                {
                    bIsOk = false;
                    sMessage = "No current cell selected";
                }

                if (bIsOk)
                {
                    int iRow = dgParameters.CurrentCell.RowIndex;
                    dgParameters.Rows.RemoveAt(iRow);
                }
                else
                { MessageBox.Show(text: sMessage, caption: ClsDefaults.messageBoxTitle(), buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Exclamation); }
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

        private void checkVisibleDestinationBtn()
        {
            try
            {
                bool bIsLstCmb;
                bool bIsRange;
                bool bIsLoop;

                if (optListboxCombo.Checked == null)
                { bIsLstCmb = false; }
                else
                {
                    if (optListboxCombo.Checked == true)
                    { bIsLstCmb = true; }
                    else
                    { bIsLstCmb = false; }
                }

                if (optRange.Checked == null)
                { bIsRange = false; }
                else
                {
                    if (optRange.Checked == true)
                    { bIsRange = true; }
                    else
                    { bIsRange = false; }
                }

                if (optEmptyLoop.Checked == null)
                { bIsLoop = false; }
                else
                {
                    if (optEmptyLoop.Checked == true)
                    { bIsLoop = true; }
                    else
                    { bIsLoop = false; }
                }

                if (bIsLstCmb && !bIsLoop && !bIsRange)
                { btnDestinationDetails.Visible = true; }
                else if (!bIsLstCmb && bIsLoop && !bIsRange)
                { btnDestinationDetails.Visible = false; }
                else if (!bIsLstCmb && !bIsLoop && bIsRange)
                { btnDestinationDetails.Visible = true; }
                else
                { btnDestinationDetails.Visible = false; }
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

        private void optListboxCombo_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkVisibleDestinationBtn();
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

        private void optRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkVisibleDestinationBtn();
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

        private void optEmptyLoop_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                checkVisibleDestinationBtn();
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

