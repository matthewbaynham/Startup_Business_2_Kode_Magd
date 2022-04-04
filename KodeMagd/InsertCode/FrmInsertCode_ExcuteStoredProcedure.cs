using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Misc;
using KodeMagd.Settings;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_ExcuteStoredProcedure : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();
        
        public FrmInsertCode_ExcuteStoredProcedure()
        {
            try
            {
                InitializeComponent();

                //cCode = new ClsCodeMapper();
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

        private void FrmInsertCode_ExcuteStoredProcedure_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnRecent);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnBuildConnectionString);
                ClsDefaults.FormatControl(ref btnAdd);

                ClsDefaults.FormatControl(ref chkExecuteAsynchronously);
                ClsDefaults.FormatControl(ref chkAddReferences);

                ClsDefaults.FormatControl(ref lblConnectionString);
                ClsDefaults.FormatControl(ref txtConnectionString);

                ClsDefaults.FormatControl(ref lblSPName);
                ClsDefaults.FormatControl(ref txtSPName);

                ClsDefaults.FormatControl(ref dgParameters);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnBuildConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(chkExecuteAsynchronously, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(chkAddReferences, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                
                cControlPosition.setControl(lblConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                
                cControlPosition.setControl(lblSPName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtSPName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                
                cControlPosition.setControl(dgParameters, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                txtSPName.Text = "<Store Procedure Name>";
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
                if (chkAddReferences.Checked)
                {
                    FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_ADO, ref ssStatus);

                    if (!frmReference.referenceAlreadySet)
                    { frmReference.ShowDialog(this); }

                    frmReference = null;
                }

                bool bIsOk = true;
                string sMessage = "";
                ClsInsertCode_ExcuteStoredProcedure cInsertCode_ExcuteStoredProcedure = new ClsInsertCode_ExcuteStoredProcedure();
                ClsDataTypes cDataTypes = new ClsDataTypes();

                if (bIsOk)
                {
                    for (int iRow = 0; iRow < dgParameters.RowCount; iRow++ )
                    {
                        ClsInsertCode_ExcuteStoredProcedure.strParameter objParameter = new ClsInsertCode_ExcuteStoredProcedure.strParameter();
                        ADODB.Parameter adoParameter = new ADODB.Parameter();

                        string sName = "";
                        if (dgParameters[ColName.Index, iRow].Value == null)
                        { sName = ""; }
                        else
                        { sName = dgParameters[ColName.Index, iRow].Value.ToString(); }
                        
                        int iSize = 0;
                        string sSize = "";

                        if (dgParameters[ColSize.Index, iRow].Value == null)
                        { sSize = ""; }
                        else
                        { sSize = dgParameters[ColSize.Index, iRow].Value.ToString(); }

                        if (!int.TryParse(sSize, out iSize))
                        {
                            bIsOk = false;
                            sMessage += "Size in Parameter " + sName + " can not be converted into a Integer.\n\r";
                        }

                        string sType = "";
                        if (dgParameters[ColType.Index, iRow].Value == null)
                        { sType = ""; }
                        else
                        { sType = dgParameters[ColType.Index, iRow].Value.ToString(); }
                        ADODB.DataTypeEnum eType = ClsDataTypes.getAdodbDataType(sType);

                        string sDirection = "";
                        if (dgParameters[ColDirection.Index, iRow].Value == null)
                        { sDirection = ""; }
                        else
                        { sDirection = dgParameters[ColDirection.Index, iRow].Value.ToString(); }
                        ADODB.ParameterDirectionEnum eDirection = ClsMisc.getAdodbDirection(sDirection);

                        string sValue = "";
                        if (dgParameters[ColValue.Index, iRow].Value == null)
                        { sValue = ""; }
                        else
                        { sValue = dgParameters[ColValue.Index, iRow].Value.ToString(); }
                        adoParameter.Name = sName;
                        adoParameter.Size = iSize;
                        if (eDirection != ADODB.ParameterDirectionEnum.adParamUnknown)
                        { adoParameter.Direction = eDirection; }
                        adoParameter.Type = eType;

                        bool bVariable;

                        if (dgParameters[ColVariable.Index, iRow].Value == null)
                        { bVariable = false; }
                        else
                        { 
                            string sVariable = dgParameters[ColVariable.Index, iRow].Value.ToString();

                            if (!bool.TryParse(sVariable, out bVariable))
                            { bVariable = false; }
                        }

                        if (bVariable == false)
                        {
                            if (!ClsDataTypes.typeCheck(eType, sValue))
                            {
                                bIsOk = false;
                                if (sValue == "")
                                { sMessage += "Paramater: " + sName + "\n\rCouldn't convert value <Empty String> into the correct datatype " + eType.ToString() + "\n\r"; }
                                else
                                { sMessage += "Paramater: " + sName + "\n\rCouldn't convert value '" + sValue + "' into the correct datatype " + eType.ToString() + "\n\r"; }
                            }
                        }

                        if (bIsOk)
                        {
                            if (bVariable)
                            {
                                //error here when the parameter is a variable name and the datatype doesn't allow strings 
                                objParameter.bIsVariable = true;
                                objParameter.sVbaVariableName = sValue;
                                adoParameter.Value = null;
                            }
                            else
                            {
                                objParameter.bIsVariable = false;
                                objParameter.sVbaVariableName = "";
                                try
                                { adoParameter.Value = sValue; }
                                catch
                                {
                                    bIsOk = false;
                                    sMessage = "Failed to assign value to Parameter please check value and datatype";
                                }
                            }

                            if (bIsOk)
                            {
                                ClsSettings cSettings = new ClsSettings();
                                cSettings.addUsedConnectionString(txtConnectionString.Text.Trim());
                                cSettings = null;
                            }

                            if (bIsOk)
                            {
                                objParameter.eVbaType = ClsDataTypes.vbVarType.vbUnknown;//gets assigned when genereating
                                objParameter.iOrder = iRow;
                                objParameter.objParameter = adoParameter;

                                cInsertCode_ExcuteStoredProcedure.addParameter(objParameter);
                            }
                        }
                    }
                }

                if (bIsOk)
                {
                    if (chkExecuteAsynchronously.Checked == null)
                    { cInsertCode_ExcuteStoredProcedure.asynchronously = false; }
                    else
                    { cInsertCode_ExcuteStoredProcedure.asynchronously = chkExecuteAsynchronously.Checked; }

                    if (string.IsNullOrEmpty(txtConnectionString.Text))
                    { cInsertCode_ExcuteStoredProcedure.connectionString = ""; }
                    else
                    { cInsertCode_ExcuteStoredProcedure.connectionString = txtConnectionString.Text; }

                    if (string.IsNullOrEmpty(txtSPName.Text))
                    { cInsertCode_ExcuteStoredProcedure.storedProcedure = ""; }
                    else
                    { cInsertCode_ExcuteStoredProcedure.storedProcedure = txtSPName.Text; }
                }

                if (bIsOk)
                {
                    configHtmlSummary(ref cInsertCode_ExcuteStoredProcedure);
                    displayHtmlSummary();

                    cInsertCode_ExcuteStoredProcedure.generateCode(ref cCodeMapper);
                    this.Close();
                }
                else
                {
                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                cInsertCode_ExcuteStoredProcedure = null; ;
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

        private void fillNewRow(int iRow) 
        {
            try
            {
                //ComboBoxes
                DataGridViewComboBoxCell cmbType = (DataGridViewComboBoxCell)dgParameters[ColType.Index, iRow];
                DataGridViewComboBoxCell cmbDirection = (DataGridViewComboBoxCell)dgParameters[ColDirection.Index, iRow];
                List<ADODB.DataTypeEnum> lstNotSupported = ClsDataTypes.NotSupportedAdoDataTypes();

                cmbType.Items.Clear();
                cmbDirection.Items.Clear();

                List<string> lstTypes = new List<string>();
                List<string> lstDirection = new List<string>();

                foreach (ADODB.DataTypeEnum eType in Enum.GetValues(typeof(ADODB.DataTypeEnum)))
                {
                    if (!lstNotSupported.Contains(eType))
                    { lstTypes.Add(eType.ToString()); }
                }

                foreach (ADODB.ParameterDirectionEnum eType in Enum.GetValues(typeof(ADODB.ParameterDirectionEnum)))
                { lstDirection.Add(eType.ToString()); }

                lstTypes.Sort();
                lstDirection.Sort();

                foreach (string sType in lstTypes)
                { cmbType.Items.Add(sType); }

                foreach (string sType in lstDirection)
                { cmbDirection.Items.Add(sType); }

                DataGridViewTextBoxCell cmbSize = (DataGridViewTextBoxCell)dgParameters[ColSize.Index, iRow];

                cmbSize.ValueType = typeof(int);

                fillCmbVariable(iRow);

                cmbType = null;
                cmbDirection = null;
                lstTypes = null;
                lstDirection = null;
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

        private void fillCmbVariable(int iRow)
        {
            try
            {
                string sVariable = "";
                if (dgParameters[ColVariable.Index, iRow].Value == null)
                { sVariable = ""; }
                else
                { sVariable = dgParameters[ColVariable.Index, iRow].Value.ToString(); }

                bool bVariable;
                if (bool.TryParse(sVariable, out bVariable))
                {
                    if (bVariable == true)
                    {
                        DataGridViewComboBoxCell cmbVarible = (DataGridViewComboBoxCell)dgParameters[ColValue.Index, iRow];

                        foreach (ClsCodeMapper.strVariables objVariable in cCodeMapper.lstVariablesInCurrentScope().OrderBy(x => x.sName))
                        { cmbVarible.Items.Add(objVariable.sName); }
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

        private void dgParameters_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                fillNewRow(e.RowIndex);
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

        private void dgParameters_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                string sValue;

                if (e.FormattedValue == null)
                { sValue = ""; }
                else
                { sValue = e.FormattedValue.ToString(); }

                if (e.ColumnIndex == ColSize.Index)
                {
                    decimal dDec;

                    if (!decimal.TryParse(sValue, out dDec))
                    {
                        MessageBox.Show("Size must be an integer", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        e.Cancel = true;
                    }
                    else 
                    {
                        int iInt;

                        if (!int.TryParse(sValue, out iInt))
                        {
                            MessageBox.Show("Size must be an integer", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            e.Cancel = true;
                        }
                    }
                }
                else if (e.ColumnIndex == ColVariable.Index) //ciParamCol_VariableChk
                {
                    bool bIsVariable = true;

                    if (!bool.TryParse(sValue, out bIsVariable))
                    {
                        MessageBox.Show("Error: Check box must be an boolean", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        e.Cancel = true;
                    }
                    cellChangeValueType(e.RowIndex, bIsVariable);
                }
                else if (e.ColumnIndex == ColType.Index) 
                {
                    ADODB.DataTypeEnum eType;

                    eType = ClsDataTypes.getAdodbDataType(sValue);

                    int iSize = ClsDataTypes.getDataTypeSize(eType);

                    dgParameters[ColSize.Index, e.RowIndex].Value = iSize;
                }
                else
                {
                    //do nothing
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

        private void cellChangeValueType(int iRow, bool bIsVariable)
        {
            /*
             * changes the value cell from a combobox that list all the variable to a textbox for typing in hardcoded values
             */
            try
            {
                if (bIsVariable)
                {
                    List<string> lst = cCodeMapper.variableNames();

                    DataGridViewComboBoxCell ComboCellString = new DataGridViewComboBoxCell();
                    
                    lst.Sort();
                    ComboCellString.Items.Clear();

                    foreach (string sVarName in lst)
                    { ComboCellString.Items.Add(sVarName); }

                    dgParameters[ColValue.Index, iRow] = ComboCellString;

                    this.fillCmbVariable(iRow);

                    lst = null;
                }
                else
                {
                    DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                    dgParameters[ColValue.Index, iRow] = TextCellString;
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
                string sParameterName = FrmInputBox.GetString("Parameter Name", "Please enter the name of the parameter you want to add.");

                if (sParameterName.Trim() != "")
                {
                    int iRow = dgParameters.Rows.Add();

                    dgParameters[ColName.Index, iRow].Value = sParameterName;
                    this.ActiveControl = dgParameters;
                    dgParameters.CurrentCell = dgParameters[ColName.Index, iRow];
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

        private void FrmInsertCode_ExcuteStoredProcedure_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnRecent);
                cControlPosition.positionControl(ref btnGenerate);
                cControlPosition.positionControl(ref btnBuildConnectionString);
                cControlPosition.positionControl(ref btnAdd);

                cControlPosition.positionControl(ref chkExecuteAsynchronously);
                cControlPosition.positionControl(ref chkAddReferences);

                cControlPosition.positionControl(ref lblConnectionString);
                cControlPosition.positionControl(ref txtConnectionString);

                cControlPosition.positionControl(ref lblSPName);
                cControlPosition.positionControl(ref txtSPName);

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

        private void configHtmlSummary(ref ClsInsertCode_ExcuteStoredProcedure cInsertCode_ExcuteStoredProcedure)
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
                objCell.sText = ClsMisc.ActiveVBComponent().Name; //cInsertCode_CommandBarClass.className;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module where VBA has been inserted.";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_ExcuteStoredProcedure.vbaFnName; //cInsertCode_CommandBarClass.SampleCodeModulePrefix.Trim() + cInsertCode_CommandBarClass.className.Trim();
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
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 4 }, "Settings");

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
                objCell.sText = "Value";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Stored Procedure Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_ExcuteStoredProcedure.storedProcedure;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Connection String";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cInsertCode_ExcuteStoredProcedure.connectionString;
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Asynchronously?";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                if (cInsertCode_ExcuteStoredProcedure.asynchronously)
                {
                    objCell.sText = "Is asynchronous";
                    objCell.sHiddenText = "When the command is sent to the database server to run the stored procedure, the VBA will not waiting for anything it'll just continue to the next line of VBA.";
                }
                else
                {
                    objCell.sText = "Is NOT asynchronous";
                    objCell.sHiddenText = "When the command is sent to the database server to run the stored procedure, the VBA will wait until the database server has finished before continuing to the next line of code.";
                }
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                if (cInsertCode_ExcuteStoredProcedure.parameters.Count == 0)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameters";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "No parameters used.";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                }
                else
                {
                    /************************
                     *   Parameters table   *
                     ************************/
                    cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 3, 1, 1 }, "Parameters");

                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "VBA Variable Name";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameter Name";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameter Type";
                    objCell.sHiddenText = "Using ADODB data types";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = "Parameter Size";
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    foreach (ClsInsertCode_ExcuteStoredProcedure.strParameter objParameter in cInsertCode_ExcuteStoredProcedure.parameters.OrderBy(x => x.iOrder))
                    {
                        //Add Row
                        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        if (objParameter.bIsVariable == true)
                        { objCell.sText = objParameter.sVbaVariableName; }
                        else
                        { objCell.sText = "<N/A>"; }
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);


                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.objParameter.Name;
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.objParameter.Type.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                        objCell.iOrder = 0;
                        objCell.bPropHtml = true;
                        objCell.sText = objParameter.objParameter.Size.ToString();
                        objCell.sHiddenText = "";
                        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                    }
                }

                cDataTypes = null;
                objCell.lstFormatDetails = null;
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Run_Store_Procedure");

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

        private void FrmInsertCode_ExcuteStoredProcedure_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

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
