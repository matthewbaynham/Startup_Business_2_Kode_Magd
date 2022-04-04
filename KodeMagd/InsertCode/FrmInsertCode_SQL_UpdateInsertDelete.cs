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
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using System.Diagnostics;
using KodeMagd.Settings;
using KodeMagd.Reporter;

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_SQL_UpdateInsertDelete : Form
    {
        public struct strTableDetails
        {
            public string sName;
            public string sCaption;
        }

        private List<strTableDetails> lstTableDetails = new List<strTableDetails>();
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsConfigReporter cConfigReporter = new ClsConfigReporter();
        private ClsCodeMapper cCode;
        private DateTime dLastKeyPress = DateTime.Now;
        private ClsCodeMapper cCodeMapper = new ClsCodeMapper();

        private string sFnNameCountRecord = "CountRecord";

        /*         :                                 :                                 :
         *=========:=================================:=================================:
         *         :  SQL                            : Via Recordset                   :
         *=========:=================================:=================================:
         * update  : List Fields that are required   : List Fields that are required   :
         *         : in the unique key.              : in the unique key.              :
         *         : List Fields that are going to   : List Fields that are going to   :
         *         : be updated.                     : be updated.                     :
         *         :                                 :                                 :
         *---------:---------------------------------:---------------------------------:
         * insert  : List all Fields.                : List all Fields.                :
         *         :                                 :                                 :
         *---------:---------------------------------:---------------------------------:
         * delete  : List Fields that are required   : List Fields that are required   :
         *         : in the unique key.              : in the unique key.              :
         *         :                                 :                                 :
         *=========:=================================:=================================:
         */

        /*=======================================================: 
         * Flags mean                                            :
         *=========:===========:===============:=================:
         *         : Selected  : Conditional   : Audit Condition :
         *=========:===========:===============:=================:
         * update  : to update : Filter on     : filter audit on :
         *---------:-----------:---------------:-----------------:
         * Insert  : to insert : N/A           : filter audit on :
         *---------:-----------:---------------:-----------------:
         * delete  : N/A       : Filter on     : filter audit on :
         *=========:===========:===============:=================:
         */

        public FrmInsertCode_SQL_UpdateInsertDelete()
        {
            try
            {
                InitializeComponent();

                if (ClsMisc.ActiveVBComponent() != null)
                {
                    cCode = new ClsCodeMapper();
                    cCode.readCode(ClsMisc.ActiveVBComponent());
                }
                else
                { cCode = null; }
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

        private void FrmInsertCode_SQL_UpdateInsertDelete_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;
                txtName.Text = ClsDefaults.defaultName;

                ClsDefaults.FormatControl(ref optInsert);
                ClsDefaults.FormatControl(ref optDelete);
                ClsDefaults.FormatControl(ref optUpdate);

                ClsDefaults.FormatControl(ref optViaRecordset);
                ClsDefaults.FormatControl(ref optSQL);

                ClsDefaults.FormatControl(ref grpAction);
                ClsDefaults.FormatControl(ref grpType);

                ClsDefaults.FormatControl(ref lblConnectionString);
                ClsDefaults.FormatControl(ref txtConnectionString);
                ClsDefaults.FormatControl(ref btnConnectionStringBuild);
                ClsDefaults.FormatControl(ref btnConnectionStringRecent);

                ClsDefaults.FormatControl(ref lblInstructions);

                ClsDefaults.FormatControl(ref lblName);
                ClsDefaults.FormatControl(ref txtName);

                ClsDefaults.FormatControl(ref lblTableName);
                ClsDefaults.FormatControl(ref txtTableName);
                ClsDefaults.FormatControl(ref cmbTableName);

                ClsDefaults.FormatControl(ref dgFields);

                ClsDefaults.FormatControl(ref btnConnectionStringExpand);
                ClsDefaults.FormatControl(ref btnAddFields);
                ClsDefaults.FormatControl(ref btnRemoveFields);

                ClsDefaults.FormatControl(ref chkAddReference);
                ClsDefaults.FormatControl(ref btnGenerate);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref chkAsynchronousWithAuditCheck);
                ClsDefaults.FormatControl(ref chkAdhocTableName);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(optInsert, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optDelete, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optUpdate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(optViaRecordset, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(optSQL, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(grpAction, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(grpType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnConnectionStringBuild, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnConnectionStringRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblInstructions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(lblTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(txtTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(dgFields, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(btnConnectionStringExpand, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(btnAddFields, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemoveFields, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(chkAddReference, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnGenerate, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(chkAsynchronousWithAuditCheck, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(chkAdhocTableName, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                optDelete.Checked = false;
                optUpdate.Checked = false;
                optInsert.Checked = true;

                optViaRecordset.Checked = false;
                optSQL.Checked = true;

                chkAsynchronousWithAuditCheck.Checked = true;

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

        private void FrmInsertCode_SQL_UpdateInsertDelete_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref optInsert);
                cControlPosition.positionControl(ref optDelete);
                cControlPosition.positionControl(ref optUpdate);

                cControlPosition.positionControl(ref optViaRecordset);
                cControlPosition.positionControl(ref optSQL);

                cControlPosition.positionControl(ref grpAction);
                cControlPosition.positionControl(ref grpType);

                cControlPosition.positionControl(ref lblConnectionString);
                cControlPosition.positionControl(ref txtConnectionString);
                cControlPosition.positionControl(ref btnConnectionStringBuild);
                cControlPosition.positionControl(ref btnConnectionStringRecent);
                cControlPosition.positionControl(ref lblInstructions);

                cControlPosition.positionControl(ref lblName);
                cControlPosition.positionControl(ref txtName);

                cControlPosition.positionControl(ref lblTableName);
                cControlPosition.positionControl(ref txtTableName);
                cControlPosition.positionControl(ref cmbTableName);

                cControlPosition.positionControl(ref dgFields);

                cControlPosition.positionControl(ref btnConnectionStringExpand);
                cControlPosition.positionControl(ref btnAddFields);
                cControlPosition.positionControl(ref btnRemoveFields);

                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref chkAddReference);
                cControlPosition.positionControl(ref btnGenerate);

                cControlPosition.positionControl(ref chkAsynchronousWithAuditCheck);
                cControlPosition.positionControl(ref chkAdhocTableName);
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

        private void btnConnectionStringBuild_Click(object sender, EventArgs e)
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

                fillCmbTableName(false);
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

        private void fillDataGridCombos(int iRow) 
        {
            try
            {
                DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgFields[ColDataType.Index, iRow];

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
                cellDataType = null;
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

        private void btnAddFields_Click(object sender, EventArgs e)
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
                //add field
                string sFieldName = FrmInputBox.GetString("Field Name", "Please enter the Field Name");
                bool bIsFound = false;
                //do we already have this field


                if (sFieldName.Trim() != "")
                {
                    if (bIsFound)
                    { MessageBox.Show("Field already exists", ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                    else
                    {
                        int iRow = dgFields.Rows.Add();

                        dgFields[ColName.Index, iRow].Value = sFieldName.Trim();

                        this.ActiveControl = dgFields;
                        dgFields.CurrentCell = dgFields[ColName.Index, iRow];
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

        private void switchVariableStatus(int iRow)
        {
            try
            {
                bool bAssignVariable;

                if (iRow >= 0)
                {
                    if (!bool.TryParse(dgFields[ColValueVariable.Index, iRow].Value.ToString(), out bAssignVariable))
                    { bAssignVariable = false; }

                    if (bAssignVariable)
                    {
                        DataGridViewComboBoxCell ComboCellString = new DataGridViewComboBoxCell();

                        //List<string> lst = cCode.variableNames();

                        //lst.Sort();
                        ComboCellString.Items.Clear();

                        //foreach (string sVarName in lst)
                        //{ ComboCellString.Items.Add(sVarName); }
                        
                        foreach (ClsCodeMapper.strVariables objTemp in cCodeMapper.lstVariablesInCurrentScope().OrderBy(x => x.sName))
                        { ComboCellString.Items.Add(objTemp.sName.Trim()); }

                        dgFields[ColValue.Index, iRow] = ComboCellString;
                        
                        ComboCellString = null;
                    }
                    else
                    {
                        DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                        dgFields[ColValue.Index, iRow] = TextCellString;
                        TextCellString = null;
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

        public void autoFillFields() 
        {
            try
            {
                /*
                 *  using the connection string do a select top 1 * from <tablename>
                 *  loop through the fields
                 */
                ADODB.Connection con = new ADODB.Connection();
                ADODB.Recordset rst = new ADODB.Recordset();
                bool bIsOk = true;
                string sMessage = "";
                string sSql = "SELECT TOP 1 * FROM [" + txtTableName.Text + "]";
                
                try
                { 
                    con.Open(ConnectionString: txtConnectionString.Text);
                }
                catch 
                {
                    bIsOk = false;
                    sMessage = "Can't open connection.\n\r\n\rPlease check Connection String.";
                }

                try
                {
                    rst.Open(sSql, con, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, -1);
                }
                catch 
                {
                    bIsOk = false;
                    sMessage = "Can't open Table.\n\r\n\rPlease check Table Name.";
                }

                ADOX.Table tbl = new ADOX.Table();

                //ADOX.Tables

                tbl = null;
                con = null;
                rst = null;
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

        public void fillCmbTableName(bool bPrompt) 
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
                    {
                        strTableDetails objTableDetails = new strTableDetails();

                        objTableDetails.sName = tbl.Name;
                        objTableDetails.sCaption = tbl.Name + " (" + tbl.Type.ToString() + ")";

                        lstTableDetails.Add(objTableDetails);

                        lst.Add(objTableDetails.sCaption); 
                    }

                    lst.Sort();

                    cmbTableName.Items.Clear();
                    foreach (string sItem in lst)
                    { cmbTableName.Items.Add(sItem); }

                    chkAdhocTableName.Checked = false;
                }
                else
                {
                    if (bPrompt)
                    { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                    else
                    { ssStatus.Text = sMessage; }
                }

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

        private string getTableNameFromCmbText(string sText)
        {
            try
            {
                string sResult = "";

                if (sText.Contains('(') & sText.Contains(')'))
                {
                    int iPos = sText.IndexOf('(');

                    if (iPos > 0)
                    { sResult = ClsMiscString.Left(ref sText, iPos - 1).Trim(); }
                    else
                    { sResult = sText; }
                }
                else
                { sResult = sText; }

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

        private void btnRemoveFields_Click(object sender, EventArgs e)
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
                if (dgFields.SelectedCells.Count > 0)
                {
                    List<int> lstSelectedRows = new List<int>();

                    foreach (DataGridViewCell objCell in dgFields.SelectedCells)
                    {
                        int iRow = objCell.RowIndex;
                        lstSelectedRows.Add(iRow); 
                    }

                    lstSelectedRows = lstSelectedRows.Distinct().ToList();
                    lstSelectedRows = lstSelectedRows.OrderByDescending(x => x).ToList();

                    foreach (int iRow in lstSelectedRows)
                    { dgFields.Rows.RemoveAt(iRow); }

                    lstSelectedRows = null;
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

        private void cmbTableName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string sTableName = getTableNameFromCmbText(cmbTableName.Text);

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

                List<string> lst = new List<string>();

                foreach (ADOX.Table tbl in cat.Tables)
                { lst.Add(tbl.Name); }

                if (!lst.Contains(sTableName))
                {
                    bIsOk = false;
                    sMessage = "Can't find Table Name";
                }

                if (bIsOk)
                {
                    while (dgFields.Rows.Count > 0)
                    { dgFields.Rows.RemoveAt(0); }

                    ADOX.Table tbl = cat.Tables[sTableName];

                    foreach (ADOX.Column col in tbl.Columns)
                    {
                        int iRow = dgFields.Rows.Add();

                        dgFields[ColName.Index, iRow].Value = col.Name;

                        DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgFields[ColDataType.Index, iRow];
                        List<string> lstItems = new List<string>();

                        foreach (string sItem in cellDataType.Items)
                        { lstItems.Add(sItem); }

                        string sDataType = col.Type.ToString();

                        if (!cellDataType.Items.Contains(sDataType))
                        { cellDataType.Items.Add(sDataType); }
                        cellDataType = null;

                        dgFields[ColDataType.Index, iRow].Value = sDataType;

                        int iSize = col.DefinedSize;

                        if (iSize == 0) 
                        { iSize = ClsDataTypes.getDataTypeSize((ADODB.DataTypeEnum)col.Type); }

                        dgFields[ColSize.Index, iRow].Value = iSize;

                        lstItems = null;
                        cellDataType = null;
                    }
                }
                else
                {
                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                cat = null;
                con = null;
                lst = null;
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

        /*
        private void btnGetTables_Click(object sender, EventArgs e)
        {
            try
            {
                fillCmbTableName(true);
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
        */

        private void dgFields_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            try
            {
                dgFields[ColSelection.Index, e.RowIndex].Value = true;
                dgFields[ColValueVariable.Index, e.RowIndex].Value = true;
                
                DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgFields[ColDataType.Index, e.RowIndex];

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
                cellDataType = null;
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

        private void chkAdhocTableName_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                bool bIsChecked;

                if (chkAdhocTableName.Checked == null)
                { bIsChecked = false; }
                else
                { bIsChecked = chkAdhocTableName.Checked; }

                if (bIsChecked)
                {
                    txtTableName.Visible = true;
                    cmbTableName.Visible = false;
                }
                else
                {
                    txtTableName.Visible = false;
                    cmbTableName.Visible = true;
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
                if (chkAddReference.Checked)
                {
                    FrmAddReference frmReference = new FrmAddReference(ClsReferences.enumFilterType.eFilt_ADO, ref ssStatus);

                    if (!frmReference.referenceAlreadySet)
                    { frmReference.ShowDialog(this); }

                    frmReference = null;
                }

                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsInsertCode_DBUpdateInsertDelete cInsertCode_DBUpdateInsertDelete = new ClsInsertCode_DBUpdateInsertDelete();
                bool bIsOk = true;
                string sMessage = "";

                cInsertCode_DBUpdateInsertDelete.fieldsEmpty();

                for (int iRow = 0; iRow < dgFields.RowCount; iRow++)
                {
                    ClsInsertCode_DBUpdateInsertDelete.strField objField = new ClsInsertCode_DBUpdateInsertDelete.strField();

                    if (bIsOk)
                    { objField.sName = dgFields[ColName.Index, iRow].Value.ToString(); }

                    string sAdoDataType = "";
                    if (dgFields[ColDataType.Index, iRow].Value == null)
                    { sAdoDataType = ""; }
                    else
                    { sAdoDataType = dgFields[ColDataType.Index, iRow].Value.ToString(); }

                    if (bIsOk)
                    {
                        objField.eDataType = ClsDataTypes.getAdodbDataType(sAdoDataType);

                        if (objField.eDataType == ADODB.DataTypeEnum.adIUnknown || objField.eDataType == ADODB.DataTypeEnum.adError)
                        {
                            bIsOk = false;
                            sMessage = "Please check the data types, for the field " + objField.sName;
                        }
                    }

                    if (bIsOk)
                    {
                        if (dgFields[ColValueVariable.Index, iRow].Value == null)
                        { objField.bIsVariable = false; }
                        else
                        {
                            string sTemp = dgFields[ColValueVariable.Index, iRow].Value.ToString();
                            bool bTemp;

                            if (bool.TryParse(sTemp, out bTemp))
                            { objField.bIsVariable = bTemp;}
                            else
                            { objField.bIsVariable = true;}
                        }
                    }

                    if (bIsOk)
                    {
                        string sSize = "";

                        if (dgFields[ColSize.Index, iRow].Value == null)
                        { sSize = "0"; }
                        else
                        { sSize = dgFields[ColSize.Index, iRow].Value.ToString(); }
                        
                        int iTemp = 0;

                        if (int.TryParse(sSize, out iTemp))
                        { objField.iSize = iTemp; }
                        else
                        { objField.iSize = 0; }
                    }

                    if (bIsOk)
                    {
                        if (dgFields[ColValue.Index, iRow].Value == null)
                        { objField.sVariableValue = ""; }
                        else
                        { objField.sVariableValue = dgFields[ColValue.Index, iRow].Value.ToString(); }

                        objField.sParameterName = objField.sName;

                        if (objField.bIsVariable == true)
                        { 
                            if (objField.sVariableValue.Trim() == "")
                            {
                                bIsOk = false;
                                sMessage = "Field '" + objField.sName + "' is suppose to be a variable however there is no variable assign to it.";
                            }
                        }
                        //else
                        //{ objField.sVariableValue = "l" + objField.sVariableValue; }
                    }

                    if (bIsOk)
                    {
                        if (dgFields[ColSelection.Index, iRow].Value == null)
                        { objField.bIsSelect = false; }
                        else
                        {
                            string sTemp = dgFields[ColSelection.Index, iRow].Value.ToString();
                            bool bTemp;

                            if (bool.TryParse(sTemp, out bTemp))
                            { objField.bIsSelect = bTemp; }
                            else
                            { objField.bIsSelect = false; }
                        }
                    }

                    if (bIsOk)
                    {
                        if (dgFields[ColAduitCondition.Index, iRow].Value == null)
                        { objField.bIsAuditCondition = false; }
                        else
                        {
                            string sTemp = dgFields[ColAduitCondition.Index, iRow].Value.ToString();
                            bool bTemp;

                            if (bool.TryParse(sTemp, out bTemp))
                            { objField.bIsAuditCondition = bTemp; }
                            else
                            { objField.bIsAuditCondition = false; }
                        }
                    }

                    if (bIsOk)
                    {
                        if (dgFields[ColWhere.Index, iRow].Value == null)
                        { objField.bIsConditional = false; }
                        else
                        {
                            string sTemp = dgFields[ColWhere.Index, iRow].Value.ToString();
                            bool bTemp;

                            if (bool.TryParse(sTemp, out bTemp))
                            { objField.bIsConditional = bTemp; }
                            else
                            { objField.bIsConditional = false; }
                        }
                    }

                    if (bIsOk)
                    {
                        if (objField.sVariableValue == "")
                        {
                            bIsOk = false;
                            if (objField.bIsVariable)
                            { sMessage = "Please check the Variable Name, for the field " + objField.sName; }
                            else
                            { sMessage = "Please check the Value, for the field " + objField.sName; }
                        }
                    }

                    if (bIsOk)
                    { cInsertCode_DBUpdateInsertDelete.fieldsAdd(objField); }
                }

                if (bIsOk)
                {
                    if (txtConnectionString.Text == null)
                    { cInsertCode_DBUpdateInsertDelete.connectionString = ""; }
                    else
                    { cInsertCode_DBUpdateInsertDelete.connectionString = txtConnectionString.Text; }

                    if (chkAdhocTableName.Checked)
                    {
                        if (txtTableName.Text == null)
                        { cInsertCode_DBUpdateInsertDelete.tableName = ""; }
                        else
                        { cInsertCode_DBUpdateInsertDelete.tableName = txtTableName.Text; }
                    }
                    else
                    {
                        if (cmbTableName.Text == null)
                        { cInsertCode_DBUpdateInsertDelete.tableName = ""; }
                        else
                        { cInsertCode_DBUpdateInsertDelete.tableName = getTableNameFromCation(cmbTableName.Text); }
                    }

                    if (txtName.Text == null)
                    { cInsertCode_DBUpdateInsertDelete.name = ""; }
                    else
                    { cInsertCode_DBUpdateInsertDelete.name = txtName.Text; }

                    cInsertCode_DBUpdateInsertDelete.doAuditCheck = chkAsynchronousWithAuditCheck.Checked;

                    cInsertCode_DBUpdateInsertDelete.fixAmbiguousFieldNames(ref bIsOk, ref sMessage);
                }

                if (bIsOk)
                {
                    if (!optDelete.Checked && !optInsert.Checked && optUpdate.Checked)
                    { cInsertCode_DBUpdateInsertDelete.type = ClsInsertCode_DBUpdateInsertDelete.enumType.eType_Update; }
                    else if (!optDelete.Checked && optInsert.Checked && !optUpdate.Checked)
                    { cInsertCode_DBUpdateInsertDelete.type = ClsInsertCode_DBUpdateInsertDelete.enumType.eType_Insert; }
                    else if (optDelete.Checked && !optInsert.Checked && !optUpdate.Checked)
                    { cInsertCode_DBUpdateInsertDelete.type = ClsInsertCode_DBUpdateInsertDelete.enumType.eType_Delete; }
                    else
                    { cInsertCode_DBUpdateInsertDelete.type = ClsInsertCode_DBUpdateInsertDelete.enumType.eType_Unknown; }

                    if (!optSQL.Checked && optViaRecordset.Checked)
                    { cInsertCode_DBUpdateInsertDelete.methodology = ClsInsertCode_DBUpdateInsertDelete.enumMethodology.eMeth_Recordset; }
                    else if (optSQL.Checked && !optViaRecordset.Checked)
                    { cInsertCode_DBUpdateInsertDelete.methodology = ClsInsertCode_DBUpdateInsertDelete.enumMethodology.eMeth_SQL; }
                    else
                    { cInsertCode_DBUpdateInsertDelete.methodology = ClsInsertCode_DBUpdateInsertDelete.enumMethodology.eMeth_Unknown; }

                    if (bIsOk)
                    {
                        ClsSettings cSettings = new ClsSettings();
                        cSettings.addUsedConnectionString(txtConnectionString.Text.Trim());
                        cSettings = null;
                    }

                    if (bIsOk)
                    { 
                        cInsertCode_DBUpdateInsertDelete.functionNameCount = ClsMiscString.nextFunctionName(ref cCodeMapper, sFnNameCountRecord); 

                        dupicateFieldCheck();
                        cInsertCode_DBUpdateInsertDelete.generateCode(ref cCodeMapper);

                        configHtmlSummary(ref cInsertCode_DBUpdateInsertDelete);

                        displayHtmlSummary();

                        this.Close();
                    }

                    cDataTypes = null;
                    cInsertCode_DBUpdateInsertDelete = null;
                }

                if (bIsOk)
                { this.Close(); }
                else
                { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
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

        private void configHtmlSummary(ref ClsInsertCode_DBUpdateInsertDelete cClassNew)
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
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

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
                objCell.sText = cClassNew.moduleName.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Module Name";
                objCell.sHiddenText = "Sample code is an example of working with the Class";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //Add Row
                cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = cClassNew.functionName.Trim();
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Function Name";
                objCell.sHiddenText = "Sample code is an example of working with the Class";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                /***************
                 *   A table   *
                 ***************/
                cConfigReporter.TableAddNew(out iTableId, new List<int> { 4, 4, 4, 1, 2, 2 }, "Fields");

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
                objCell.sText = "Parameter Name";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Variable / Value";
                objCell.sHiddenText = "";
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
                objCell.sText = "Data type";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                objCell.iOrder = 0;
                objCell.bPropHtml = true;
                objCell.sText = "Flags";
                objCell.sHiddenText = "";
                objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                foreach (ClsInsertCode_DBUpdateInsertDelete.strField objFields in cClassNew.fields)
                {
                    //Add Row
                    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objFields.sName;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objFields.sParameterName;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objFields.sVariableValue;
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objFields.iSize.ToString();
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = objFields.eDataType.ToString();
                    objCell.sHiddenText = "";
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                    string sFlags = "";
                    string sFlagsDescriptions = "";

                    if (objFields.bIsVariable)
                    {
                        sFlags += "V";
                        sFlagsDescriptions += "Is Variable\n";
                    }

                    if (objFields.bIsSelect)
                    {
                        sFlags += "S";
                        sFlagsDescriptions += "Is Selected\n";
                    }

                    if (objFields.bIsConditional)
                    {
                        sFlags += "C";
                        sFlagsDescriptions += "Is Conditional\n";
                    }

                    if (objFields.bIsAuditCondition)
                    {
                        sFlags += "A";
                        sFlagsDescriptions += "Is Audit Condition\n";
                    }

                    if (sFlagsDescriptions.Length > 1)
                    {
                        if (sFlagsDescriptions.EndsWith("\n"))
                        { sFlagsDescriptions = ClsMiscString.Left(sFlagsDescriptions, sFlagsDescriptions.Length - 1); }
                    }

                    objCell.iOrder = 0;
                    objCell.bPropHtml = true;
                    objCell.sText = sFlags;
                    objCell.sHiddenText = sFlagsDescriptions;
                    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "DB Write");

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

        private void dgFields_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == ColDataType.Index)
                {
                    string sDataType = "";
                    int iRow = e.RowIndex;

                    if (dgFields[ColDataType.Index, iRow].Value == null)
                    { sDataType = ""; }
                    else
                    { sDataType = dgFields[ColDataType.Index, iRow].Value.ToString(); }

                    int iSize = ClsDataTypes.getDataTypeSize(ClsDataTypes.getAdodbDataType(sDataType));

                    dgFields[ColSize.Index, iRow].Value = iSize;
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

        private string getTableNameFromCation(string sCaption)
        {
            try 
            {
                string sResult = "";

                if (lstTableDetails.Exists(x => x.sCaption.Trim().ToLower() == sCaption.Trim().ToLower()))
                {
                    strTableDetails objTableDetails = lstTableDetails.Find(x => x.sCaption.Trim().ToLower() == sCaption.Trim().ToLower());

                    sResult = objTableDetails.sName;
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

        /*=======================================================: 
         * Flags mean                                            :
         *=========:===========:===============:=================:
         *         : Selected  : Conditional   : Audit Condition :
         *=========:===========:===============:=================:
         * update  : to update : Filter on     : filter audit on :
         *---------:-----------:---------------:-----------------:
         * Insert  : to insert : N/A           : filter audit on :
         *---------:-----------:---------------:-----------------:
         * delete  : N/A       : Filter on     : filter audit on :
         *=========:===========:===============:=================:
         */
        private void setVisibleFields() 
        {
            try
            {
                string sInstructions = "";


                if (optInsert.Checked & !optDelete.Checked & !optUpdate.Checked)
                {
                    //insert
                    ColSelection.Visible = true;
                    ColWhere.Visible = false;

                    sInstructions = "Selected flag: These fields will be inserted into the database.";
                    sInstructions += "\n";
                    sInstructions += "Audit Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to insert.";
                }
                else if (!optInsert.Checked & optDelete.Checked & !optUpdate.Checked)
                {
                    //delete
                    ColSelection.Visible = false;
                    ColWhere.Visible = true;

                    sInstructions = "Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to delete.";
                    sInstructions += "\n";
                    sInstructions += "Audit Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to delete.";
                }
                else if (!optInsert.Checked & !optDelete.Checked & optUpdate.Checked)
                {
                    //update
                    ColSelection.Visible = true;
                    ColWhere.Visible = true;

                    sInstructions = "Selected flag: These fields will be inserted into the database.";
                    sInstructions += "\n";
                    sInstructions = "Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to update.";
                    sInstructions += "\n";
                    sInstructions += "Audit Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to update.";
                }
                else
                { 
                    //error
                }

                if (chkAsynchronousWithAuditCheck.Checked)
                { ColAduitCondition.Visible = true; }
                else
                { ColAduitCondition.Visible = false; }

                lblInstructions.Text = sInstructions;
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

        private void optInsert_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                setVisibleFields();
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

        private void optUpdate_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                setVisibleFields();
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

        private void optDelete_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                setVisibleFields();
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

        private void chkAsynchronousWithAuditCheck_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                setVisibleFields();
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

        private void dupicateFieldCheck() 
        {
            try
            {
                List<string> lstFieldName = new List<string>();
                
                for (int iRow = 0; iRow < dgFields.RowCount; iRow++)
                {
                    if (dgFields[ColName.Index, iRow].Value != null)
                    {
                        string sName = dgFields[ColName.Index, iRow].Value.ToString();

                        lstFieldName.Add(sName);
                    }
                }
                /*
                var varDuplicateItems = from x in lstFieldName
                                        group x.Trim().ToLower() by x.Trim().ToLower() into grouped
                                        where grouped.Count() > 1
                                        select grouped.Key.ToString();

                List<string> lstDuplicateItems = varDuplicateItems.ToList();
                */

                List<string> lstDuplicateItems = (from x in lstFieldName
                                                    group x.Trim().ToLower() by x.Trim().ToLower() into grouped
                                                    where grouped.Count() > 1
                                                    select grouped.Key.ToString()).ToList();
                
                if (lstDuplicateItems.Count > 0)
                {
                    string sMessage = "Some of the fields have been listed twice." + "\n\r\n\r";

                    foreach (string sTemp in lstDuplicateItems)
                    { sMessage += sTemp + "\n\r"; }

                    sMessage += "\n\rWhen you try to run your VBA please make sure each parameter has is unique.";

                    MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                lstDuplicateItems = null;
                lstFieldName = null;
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

        private void dgFields_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == ColValueVariable.Index)
                {
                    bool bAssignVariable;
                    int iRow = e.RowIndex;

                    if (iRow >= 0)
                    {
                        if (!bool.TryParse(dgFields[ColValueVariable.Index, iRow].Value.ToString(), out bAssignVariable))
                        { bAssignVariable = false; }

                        if (bAssignVariable)
                        {
                            DataGridViewComboBoxCell ComboCellString = new DataGridViewComboBoxCell();
                            
                            List<string> lst = cCode.variableNames();

                            lst.Sort();
                            ComboCellString.Items.Clear();

                            foreach (string sVarName in lst)
                            { ComboCellString.Items.Add(sVarName); }

                            dgFields[ColValue.Index, iRow] = ComboCellString;

                            ComboCellString = null;
                            lst = null;
                        }
                        else
                        {
                            DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                            dgFields[ColValue.Index, iRow] = TextCellString;
                            TextCellString = null;
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

        private void FrmInsertCode_SQL_UpdateInsertDelete_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.B)
                    { build(); }

                    if (e.KeyCode == Keys.V)
                    { remove(); }

                    if (e.KeyCode == Keys.E)
                    { recent(); }

                    if (e.KeyCode == Keys.T)
                    { fillCmbTableName(true); }

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

        //private void updateFlagMeaning()
        //{
        //    try
        //    {
        //        /*=======================================================: 
        //         * Flags mean                                            :
        //         *=========:===========:===============:=================:
        //         *         : Selected  : Conditional   : Audit Condition :
        //         *=========:===========:===============:=================:
        //         * update  : to update : Filter on     : filter audit on :
        //         *---------:-----------:---------------:-----------------:
        //         * Insert  : to insert : N/A           : filter audit on :
        //         *---------:-----------:---------------:-----------------:
        //         * delete  : N/A       : Filter on     : filter audit on :
        //         *=========:===========:===============:=================:
        //         */
        //        string sMessage = "";

        //        if (optDelete.Checked == true && optInsert.Checked != true && optUpdate.Checked != true)
        //        {
        //            //Delete
        //            sMessage = "Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to delete.";
        //            sMessage += "\n";
        //            sMessage += "Audit Contitional flag: the where condition in the SQL, i.e. the fields used to decide which records to delete.";


        //        }
        //        else if (optDelete.Checked != true && optInsert.Checked == true && optUpdate.Checked != true)
        //        {
        //            //Insert
        //            sMessage = "";


        //        }
        //        else if (optDelete.Checked != true && optInsert.Checked != true && optUpdate.Checked == true)
        //        {
        //            //Update
        //            sMessage = "";


        //        }


        //        //optDelete.Checked
        //        //optInsert.Checked
        //        //optUpdate.Checked

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

        private void btnConnectionStringExpand_Click(object sender, EventArgs e)
        {
            try
            {
                string sQuestion = "Connection String";
                bool bReadOnly = false;

                string sConnectinString = FrmLargeTextBox.GetString(ClsDefaults.formTitle, sQuestion, txtConnectionString.Text, bReadOnly);

                if (!string.IsNullOrEmpty(sConnectinString))
                { txtConnectionString.Text = sConnectinString; }

                //fillFields();
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
