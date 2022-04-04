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

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private string sLblQuestion;

        public FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters()
        {
            try
            {
                InitializeComponent();

                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref lblRememberQuestionMarks);
                ClsDefaults.FormatControl(ref btnCancel);
                ClsDefaults.FormatControl(ref btnOK);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref dgParameters);
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(lblRememberQuestionMarks, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOK, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnCancel, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
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

        //public string Question {
        //    get { 
        //        try 
        //        {
        //            return sLblQuestion; 
        //        }
        //        catch (Exception ex)
        //        {
        //            MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //            string sMessage = "";

        //            sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //            sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //            sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //            sMessage += ex.Message;

        //            MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //            return string.Empty;
        //        }
        //    }
        //    set { 
        //        try 
        //        {
        //            sLblQuestion = value;
        //            lblQuestion.Text = sLblQuestion;
        //        }
        //        catch (Exception ex)
        //        {
        //            MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //            string sMessage = "";

        //            sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //            sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //            sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //            sMessage += ex.Message;

        //            MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //        }
        //    }
        //}

        private List<ClsInsertCode_PopulateListboxCombobox.strParameter> ResultText { get; set; }

        private List<ClsInsertCode_PopulateListboxCombobox.strParameter> parameters { 
            get
            {
                try
                {
                    List<ClsInsertCode_PopulateListboxCombobox.strParameter> lstResult = new List<ClsInsertCode_PopulateListboxCombobox.strParameter>();

                    for (int iRow = 0; iRow < dgParameters.Rows.Count; iRow++)
                    {
                        //Name
                        string sName = "";

                        if (dgParameters[ColName.Index, iRow].Value == null)
                        { sName = ""; }
                        else
                        { sName = dgParameters[ColName.Index, iRow].Value.ToString(); }

                        //Data Type
                        string sDataType = "";

                        if (dgParameters[ColType.Index, iRow].Value == null)
                        { sDataType = ""; }
                        else
                        { sDataType = dgParameters[ColType.Index, iRow].Value.ToString(); }

                        ADODB.DataTypeEnum eDataType = ClsMisc.getAdodbDataType(sDataType);

                        //Direction
                        string sDirection = "";

                        if (dgParameters[ColDirection.Index, iRow].Value == null)
                        { sDataType = ""; }
                        else
                        { sDataType = dgParameters[ColDirection.Index, iRow].Value.ToString(); }

                        ADODB.ParameterDirectionEnum eDirection = ClsMisc.getAdodbDirection(sDataType);

                        //Size
                        string sSize = "";
                        int iSize = 0;

                        if (dgParameters[ColSize.Index, iRow].Value == null)
                        { sSize = "0"; }
                        else
                        { sSize = dgParameters[ColSize.Index, iRow].Value.ToString(); }

                        if (!int.TryParse(sSize, out iSize))
                        { iSize = 0; }

                        //Value
                        string sValue = "";

                        if (dgParameters[ColValue.Index, iRow].Value == null)
                        { sValue = ""; }
                        else
                        { sValue = dgParameters[ColValue.Index, iRow].Value.ToString(); }

                        ClsInsertCode_PopulateListboxCombobox.strParameter objParameter = new ClsInsertCode_PopulateListboxCombobox.strParameter();

                        objParameter.sName = sName;
                        objParameter.eDataType = eDataType;
                        objParameter.eDirection = eDirection;
                        objParameter.lSize = iSize;
                        objParameter.sValue = sValue;

                        lstResult.Add(objParameter);
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

                    return new List<ClsInsertCode_PopulateListboxCombobox.strParameter>();
                }
            }
            set 
            {
                try
                {
                    foreach (ClsInsertCode_PopulateListboxCombobox.strParameter objParameter in value)
                    {
                        int iRow = dgParameters.Rows.Add();

                        dgParameters[ColName.Index, iRow].Value = objParameter.sName;
                        dgParameters[ColType.Index, iRow].Value = objParameter.eDataType.ToString();
                        dgParameters[ColDirection.Index, iRow].Value = objParameter.eDirection.ToString();
                        dgParameters[ColSize.Index, iRow].Value = objParameter.lSize.ToString();
                        dgParameters[ColValue.Index, iRow].Value = objParameter.sValue;
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

        public static List<ClsInsertCode_PopulateListboxCombobox.strParameter> GetParameters(string title, List<ClsInsertCode_PopulateListboxCombobox.strParameter> lstParameters)
        {
            try
            {
                FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters box = new FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters { Text = title, parameters = lstParameters };

                if (box.ShowDialog() == DialogResult.OK)
                {
                    return box.ResultText;
                }

                return new List<ClsInsertCode_PopulateListboxCombobox.strParameter>();
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
                return new List<ClsInsertCode_PopulateListboxCombobox.strParameter>();
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                ok();
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
        
        private void ok()
        {
            try
            {
                this.ResultText = parameters;
                this.DialogResult = DialogResult.OK;
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                this.DialogResult = DialogResult.Cancel;
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
                string sName = FrmInputBox.GetString("Please enter parameter name", "Name");

                if (sName.Trim() != "")
                {
                    int iRow = dgParameters.Rows.Add();

                    dgParameters[ColName.Index, iRow].Value = sName;

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
                int iRow = dgParameters.CurrentRow.Index;
                string sName = "";

                if (dgParameters[ColName.Index, iRow].Value == null)
                { sName = ""; }
                else
                { sName = dgParameters[ColName.Index, iRow].Value.ToString(); }

                DialogResult dlg = MessageBox.Show("Are you sure you want to remove the '" + sName + "' parameter?", ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlg == System.Windows.Forms.DialogResult.Yes)
                { dgParameters.Rows.RemoveAt(iRow); }
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
                DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgParameters[ColType.Index, e.RowIndex];

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

                //Direction
                DataGridViewComboBoxCell cellDirection = (DataGridViewComboBoxCell)dgParameters[ColDirection.Index, e.RowIndex];

                Array arrDirection = Enum.GetValues(typeof(ADODB.ParameterDirectionEnum));
                Array.Sort(arrDirection);

                List<string> lstDirection = new List<string>();

                foreach (ADODB.ParameterDirectionEnum eTemp in arrDirection)
                { lstDirection.Add(eTemp.ToString()); }
                lstDirection.Sort();

                foreach (string sTemp in lstDirection)
                { cellDirection.Items.Add(sTemp); }

                arrDirection = null;
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

        private void FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblRememberQuestionMarks);
                cControlPosition.positionControl(ref btnOK);
                cControlPosition.positionControl(ref btnCancel);
                cControlPosition.positionControl(ref btnRemove);
                cControlPosition.positionControl(ref btnAdd);
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

        private void FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.R)
                    { remove(); }

                    if (e.KeyCode == Keys.O)
                    { ok(); }

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

        private void FrmInsertCode_Rst_PopulateListboxCombobox_RstParameters_Load(object sender, EventArgs e)
        {

        }
    }
}
