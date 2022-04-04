using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Reflection;
using KodeMagd.Settings;
using KodeMagd.Reporter;

/*
Am i going about everything the wrong way should i just use the objects?

•	SqlConnectionStringBuilder 
•	OleDbConnectionStringBuilder 
•	OdbcConnectionStringBuilder 
•	OracleConnectionStringBuilder 
instead? 
http://msdn.microsoft.com/en-us/library/vstudio/ms254500%28v=vs.100%29.aspx
http://msdn.microsoft.com/en-us/library/vstudio/system.data.oledb.oledbconnectionstringbuilder%28v=vs.100%29.aspx
http://msdn.microsoft.com/en-us/library/vstudio/system.data.odbc.odbcconnectionstringbuilder%28v=vs.100%29.aspx
http://msdn.microsoft.com/en-us/library/vstudio/system.data.oracleclient.oracleconnectionstringbuilder%28v=vs.100%29.aspx
*/


namespace KodeMagd.Misc
{
    public partial class FrmConnectionString : Form
    {
        private const int ciAtt_Type = 0;
        private const int ciAtt_Value = 1;

        SqlConnectionStringBuilder conStringBuilder;
        ClsReadConnectionString cReadConnectionString = new ClsReadConnectionString();
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsConfigReporter cConfigReporter = new ClsConfigReporter();

        string sResult = "";

        public string Result 
        {
            get
            {
                try
                {
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
        }

        public FrmConnectionString()
        {
            try
            {
                InitializeComponent();
                conStringBuilder = new SqlConnectionStringBuilder();
            }
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

        private void FrmConnectionString_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnDelete);
                ClsDefaults.FormatControl(ref btnOk);
                ClsDefaults.FormatControl(ref btnRecent);

                ClsDefaults.FormatControl(ref dgAttributes);

                ClsDefaults.FormatControl(ref lblBackend);
                ClsDefaults.FormatControl(ref lblType);

                ClsDefaults.FormatControl(ref cmbBackend);
                ClsDefaults.FormatControl(ref cmbType);

                ClsDefaults.FormatControl(ref txtConnectionString);
                ClsDefaults.FormatControl(ref txtNotes);

                ClsDefaults.FormatControl(ref ssStatus);

                cReadConnectionString.read();
                fillBackend();
                fillType();

                cControlPosition.setControl(btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnDelete, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnRecent, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnOk, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(dgAttributes, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblBackend, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbBackend, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lblType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(cmbType, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);

                cControlPosition.setControl(txtConnectionString, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(txtNotes, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
            }
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

        private void fillCmbType(ref DataGridViewComboBoxCell cmb) 
        {
            try
            {
                //ConnectionReset, ConnectionString
                List<string> lstPropNames = new List<string>();
                string sCurrentValue = "";
                lstPropNames.Clear();

                foreach (PropertyInfo propInfo in conStringBuilder.GetType().GetProperties())
                {
                    if (propInfo.CanWrite)
                    { lstPropNames.Add(propInfo.Name.Trim()); }

                }

                if (cmb.Value == null)
                { sCurrentValue = ""; }
                else if (string.IsNullOrEmpty(cmb.Value.ToString()))
                //if (string.IsNullOrEmpty(cmb.Value))
                { sCurrentValue = ""; }
                else
                {
                    sCurrentValue = cmb.Value.ToString();
                    lstPropNames.Add(sCurrentValue.Trim());
                    cmb.Value = null;
                }

                lstPropNames = lstPropNames.Distinct().ToList();
                lstPropNames.Sort();
                
                cmb.Items.Clear();
                foreach (string sTemp in lstPropNames)
                { cmb.Items.Add(sTemp); }

                cmb.Value = sCurrentValue;
            }
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


        private void dgAttributes_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                updateTxtConnectionString();
            }
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

        private void updateTxtConnectionString()
        {
            try
            {
                this.conStringBuilder = new SqlConnectionStringBuilder();

                List<string> lstConnectionString = new List<string>();

                foreach (DataGridViewRow objRow in dgAttributes.Rows)
                {
                    string sTypeName;
                    
                    if (dgAttributes[ColType.Index, objRow.Index].Value == null)
                    { sTypeName = string.Empty; }
                    else
                    { sTypeName = (string)dgAttributes[ColType.Index, objRow.Index].Value; }

                    string sValue;

                    if (dgAttributes[ColValue.Index, objRow.Index].Value == null)
                    { sValue = ""; }
                    else
                    { sValue = (string)dgAttributes[ColValue.Index, objRow.Index].Value; }

                    if (sValue == null | sValue.Trim() == "")
                    {
                        string sType = "";
                        List<string> lstPropNames = new List<string>();
                        lstPropNames.Clear();

                        foreach (PropertyInfo propInfo in conStringBuilder.GetType().GetProperties())
                        {
                            if (propInfo.CanWrite)
                            {
                                if (propInfo.Name == sTypeName)
                                { sType = propInfo.PropertyType.ToString(); }
                            }
                        }

                        lstPropNames.Sort();

                        switch (sType)
                        {
                            case "System.String":
                                DataGridViewTextBoxCell TextCellString = new DataGridViewTextBoxCell();
                                dgAttributes[ciAtt_Value, objRow.Index] = TextCellString;
                                break;
                            case "System.Int":
                            case "System.Int16":
                            case "System.Int32":
                            case "System.Int64":
                            case "System.IntPtr":
                                DataGridViewTextBoxCell TextCellInt = new DataGridViewTextBoxCell();
                                TextCellInt.Style.NullValue = "Integer value";
                                dgAttributes[ciAtt_Value, objRow.Index] = TextCellInt;
                                break;
                            case "System.Boolean":
                                DataGridViewCheckBoxCell CheckCell = new DataGridViewCheckBoxCell();
                                dgAttributes[ciAtt_Value, objRow.Index] = CheckCell;
                                dgAttributes[ciAtt_Value, objRow.Index].Value = true;
                                break;
                            case "System.Collections.ICollection":
                                DataGridViewButtonCell ButtonCell = new DataGridViewButtonCell();
                                ButtonCell.Value = "Press";
                                dgAttributes[ciAtt_Value, objRow.Index] = ButtonCell;
                                break;
                            default:
                                DataGridViewTextBoxCell TextCellDefault = new DataGridViewTextBoxCell();
                                dgAttributes[ciAtt_Value, objRow.Index] = TextCellDefault;
                                //if (sType == "")
                                //{ dgAttributes[ciAtt_Value, e.RowIndex].Value = "<not found>"; }
                                //else
                                //{ 
                                dgAttributes[ciAtt_Value, objRow.Index].Value = sType;
                                //}
                                break;
                        }
                    }

                    switch (sTypeName)
                    {
                        case "AsynchronousProcessing":
                        case "BrowsableConnectionString":
                        case "ContextConnection":
                        case "Encrypt":
                        case "Enlist":
                        case "IntegratedSecurity":
                        case "MultipleActiveResultSets":
                        case "PersistSecurityInfo":
                        case "Pooling":
                        case "TrustServerCertificate":
                        case "Replication":
                        case "UserInstance":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { lstConnectionString.Add(sTypeName + "="); }
                            else
                            {
                                bool bValue;

                                //switch (dgAttributes[ColValue.Index, objRow.Index].Value.ToString().ToLower())
                                //{
                                //    case "true":
                                //        lstConnectionString.Add(sTypeName + "=true");
                                //        break;
                                //    case "false":
                                //        lstConnectionString.Add(sTypeName + "=false");
                                //        break;
                                //    default:
                                        if (bool.TryParse((string)dgAttributes[ColValue.Index, objRow.Index].Value, out bValue))
                                        {
                                            if (bValue)
                                            { lstConnectionString.Add(sTypeName + "=true"); }
                                            else
                                            { lstConnectionString.Add(sTypeName + "=false"); }
                                        }
                                        else
                                        { lstConnectionString.Add(sTypeName + "=" + (string)dgAttributes[ColValue.Index, objRow.Index].Value); }
                                //        break;
                                //}
                            }
                            break;
                        case "AttachDBFilename":
                        case "CurrentLanguage":
                        case "DataSource":
                        case "FailoverPartner":
                        case "InitialCatalog":
                        case "MaxPoolSize":
                        case "MinPoolSize":
                        case "NetworkLibrary":
                        case "Password":
                        case "TransactionBinding":
                        case "TypeSystemVersion":
                        case "UserID":
                        case "WorkstationID":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { lstConnectionString.Add(sTypeName + "=\"\"");}
                            else
                            { lstConnectionString.Add(sTypeName + "=" + ClsMiscString.addQuotes((string)dgAttributes[ciAtt_Value, objRow.Index].Value));}
                            break;
                        case "ConnectTimeout":
                        case "LoadBalanceTimeout":
                        case "PacketSize":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { lstConnectionString.Add(sTypeName + "=0"); }
                            else
                            { lstConnectionString.Add(sTypeName + "=" + (string)dgAttributes[ciAtt_Value, objRow.Index].Value); }
                            break;
                        default:
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { lstConnectionString.Add(sTypeName + "="); }
                            else
                            {
                                int iNumber;
                                float fNumber;
                                DateTime dDate;
 
                                if (int.TryParse((string)dgAttributes[ciAtt_Value, objRow.Index].Value, out iNumber))
                                { lstConnectionString.Add(sTypeName + "=" + iNumber.ToString()); }
                                else if (float.TryParse((string)dgAttributes[ciAtt_Value, objRow.Index].Value, out fNumber))
                                { lstConnectionString.Add(sTypeName + "=" + fNumber.ToString()); }
                                else if (DateTime.TryParse((string)dgAttributes[ciAtt_Value, objRow.Index].Value, out dDate))
                                { lstConnectionString.Add(sTypeName + "=#" + dDate.ToString("mm/dd/yyyy") + "#"); }
                                else
                                { lstConnectionString.Add(sTypeName + "=" + (string)dgAttributes[ciAtt_Value, objRow.Index].Value);}
                            }
                        break;
                    }

                    /*
                    switch (sTypeName)
                    {
                        case "AsynchronousProcessing":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.AsynchronousProcessing = false; }
                            else
                            { this.conStringBuilder.AsynchronousProcessing = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "AttachDBFilename":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.AttachDBFilename = string.Empty; }
                            else
                            { this.conStringBuilder.AttachDBFilename = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "BrowsableConnectionString":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.BrowsableConnectionString = false; }
                            else
                            { this.conStringBuilder.BrowsableConnectionString = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "ConnectTimeout":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.ConnectTimeout = 0; }
                            else
                            { this.conStringBuilder.ConnectTimeout = (int)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "ContextConnection":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.ContextConnection = false; }
                            else
                            { this.conStringBuilder.ContextConnection = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "CurrentLanguage":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.CurrentLanguage = string.Empty; }
                            else
                            { this.conStringBuilder.CurrentLanguage = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "DataSource":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.DataSource = string.Empty; }
                            else
                            { this.conStringBuilder.DataSource = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "Encrypt":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.Encrypt = false; }
                            else
                            { this.conStringBuilder.Encrypt = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "Enlist":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.Enlist = false; }
                            else
                            { this.conStringBuilder.Enlist = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "FailoverPartner":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.FailoverPartner = string.Empty; }
                            else
                            { this.conStringBuilder.FailoverPartner = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "InitialCatalog":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.InitialCatalog = string.Empty; }
                            else
                            { this.conStringBuilder.InitialCatalog = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "IntegratedSecurity":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.IntegratedSecurity = false; }
                            else
                            { this.conStringBuilder.IntegratedSecurity = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "LoadBalanceTimeout":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.LoadBalanceTimeout = 0; }
                            else
                            { this.conStringBuilder.LoadBalanceTimeout = (int)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "MaxPoolSize":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.MaxPoolSize = 0; }
                            else
                            { this.conStringBuilder.MaxPoolSize = (int)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "MinPoolSize":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.MinPoolSize = 0; }
                            else
                            { this.conStringBuilder.MinPoolSize = (int)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "MultipleActiveResultSets":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.MultipleActiveResultSets = false; }
                            else
                            { this.conStringBuilder.MultipleActiveResultSets = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "NetworkLibrary":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.NetworkLibrary = string.Empty; }
                            else
                            { this.conStringBuilder.NetworkLibrary = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "PacketSize":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.PacketSize = 0; }
                            else
                            { this.conStringBuilder.PacketSize = (int)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "Password":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.Password = string.Empty; }
                            else
                            { this.conStringBuilder.Password = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "PersistSecurityInfo":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.PersistSecurityInfo = false; }
                            else
                            { this.conStringBuilder.PersistSecurityInfo = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "Pooling":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.Pooling = false; }
                            else
                            { this.conStringBuilder.Pooling = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "Replication":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.Replication = false; }
                            else
                            { this.conStringBuilder.Replication = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "TransactionBinding":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.TransactionBinding = string.Empty; }
                            else
                            { this.conStringBuilder.TransactionBinding = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "TrustServerCertificate":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.TrustServerCertificate = false; }
                            else
                            { this.conStringBuilder.TrustServerCertificate = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "TypeSystemVersion":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.TypeSystemVersion = string.Empty; }
                            else
                            { this.conStringBuilder.TypeSystemVersion = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "UserID":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.UserID = string.Empty; }
                            else
                            { this.conStringBuilder.UserID = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "UserInstance":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.UserInstance = false; }
                            else
                            { this.conStringBuilder.UserInstance = (bool)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        case "WorkstationID":
                            if (dgAttributes[ciAtt_Value, objRow.Index].Value == null)
                            { this.conStringBuilder.WorkstationID = string.Empty; }
                            else
                            { this.conStringBuilder.WorkstationID = (string)dgAttributes[ciAtt_Value, objRow.Index].Value; }
                            break;
                        default:
                            //this.conStringBuilder.Add(sTypeName, sValue);
                            //this.conStringBuilder.Keys. Add(sTypeName, sValue);
                            this.conStringBuilder.Add(sTypeName, sValue);
                            break;
                    }
                    */
                }
                //txtConnectionString.Text = this.conStringBuilder.ConnectionString;
                if (lstConnectionString.Count == 0)
                { txtConnectionString.Text = ""; }
                else
                { txtConnectionString.Text = string.Join(";", lstConnectionString.ToArray()); }
            }
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

        private void dgAttributes_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //try
            //{

            //    //DataGridViewComboBoxCell cellDataType = (DataGridViewComboBoxCell)dgParameters[ciParamCol_Type, e.RowIndex - 1];
            //    DataGridViewComboBoxCell cmd = (DataGridViewComboBoxCell)dgAttributes[ciAtt_Type, e.RowIndex];

            //    fillCmbType(ref cmd);
            //}
            //catch (Exception ex)
            //{
            //    MethodBase mbTemp = MethodBase.GetCurrentMethod();

            //    string sMessage = "";

            //    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
            //    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
            //    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
            //    sMessage += ex.Message;

            //    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            //}
        }

        private void dgAttributes_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewComboBoxCell cmd = (DataGridViewComboBoxCell)dgAttributes[ciAtt_Type, e.RowIndex];

                fillCmbType(ref cmd);
            }
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
                string sRecentString = FrmRecentConnectionStrings.GetString();

                //MessageBox.Show(sRecentString, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                List<string> lstElements = sRecentString.Split(';').ToList();

                foreach (string sElement in lstElements)
                {
                    if (sElement.Contains('='))
                    {
                        int iRow = dgAttributes.Rows.Add();
                        int iPosEquals = sElement.IndexOf('=');

                        string sType = ClsMiscString.Left(sElement , iPosEquals).Trim();
                        string sValue = ClsMiscString.Right(sElement, sElement.Length - iPosEquals - 1).Trim();

                        DataGridViewComboBoxCell cmb = (DataGridViewComboBoxCell)dgAttributes[ColType.Index, iRow];

                        if (!cmb.Items.Contains(sType))
                        {cmb.Items.Add(sType); }
                        
                        dgAttributes[ColType.Index, iRow].Value = sType;
                        dgAttributes[ColValue.Index, iRow].Value = sValue;
                    }
                }

                updateTxtConnectionString();
            }
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

        private void fillDgAttributes() 
        {
            try
            {
                dgAttributes.Rows.Clear();

                foreach (PropertyInfo propInfo in conStringBuilder.GetType().GetProperties())
                {
                    if (propInfo.CanWrite)
                    {
                        if (propInfo.GetValue(propInfo, null) != null)
                        {
                            string[] rowNew = new string[] { propInfo.Name, propInfo.GetValue(propInfo, null).ToString() };

                            dgAttributes.Rows.Add(rowNew);
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

        private void dgAttributes_CellContentClick(object sender, DataGridViewCellEventArgs e)
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

        private void cmbBackend_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dgAttributes.Rows.Clear();
                cmbType.Text = "";
                txtConnectionString.Text = "";
                txtNotes.Text = "";
                fillType();
            }
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

        private void fillBackend()
        {
            try
            {
                List<string> lstBackends = new List<string>();

                lstBackends = cReadConnectionString.ListBackend();

                lstBackends.Sort();

                cmbBackend.Items.Clear();
                foreach (string sBackend in lstBackends)
                { cmbBackend.Items.Add(sBackend); }
            }
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

        private void fillType()
        {
            try
            {
                List<string> lstTypes = new List<string>();

                if (string.IsNullOrEmpty(cmbBackend.Text.Trim()))
                { lstTypes = cReadConnectionString.ListType(); }
                else
                { lstTypes = cReadConnectionString.ListType(cmbBackend.Text.Trim()); }

                lstTypes.Sort();

                cmbType.Items.Clear();
                foreach (string sType in lstTypes)
                { cmbType.Items.Add(sType); }

            }
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

        private void cmbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillDetails();
            }
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

        private void fillDetails() 
        {
            try
            {
                ClsReadConnectionString.strConnectionString objConnStr = cReadConnectionString.getConnectionString(cmbBackend.Text, cmbType.Text);

                if (objConnStr.sBackend == "" | objConnStr.sType == "")
                {
                    while (dgAttributes.RowCount > 0)
                    { dgAttributes.Rows.RemoveAt(0); }

                    //txtConnectionString.Text = "";
                    txtNotes.Text = "";

                }
                else
                {
                    while (dgAttributes.RowCount > 0)
                    { dgAttributes.Rows.RemoveAt(0); }

                    foreach (ClsReadConnectionString.strConnStrElement objElement in objConnStr.lstElements)
                    {
                        int iRowNew = dgAttributes.Rows.Add();

                        dgAttributes[ColValue.Index, iRowNew].Value = objElement.sValue;
                        bool bIsFound = false;

                        DataGridViewComboBoxCell cmbCell = (DataGridViewComboBoxCell)dgAttributes[ColType.Index, iRowNew];

                        foreach (string sItem in cmbCell.Items)
                        {
                            if (sItem.ToLower().Trim() == objElement.sName.ToLower().Trim())
                            { bIsFound = true; }
                        }

                        if (!bIsFound)
                        { cmbCell.Items.Add(objElement.sName.Trim()); }

                        dgAttributes[ColType.Index, iRowNew].Value = objElement.sName.Trim();
                    }

                    //txtConnectionString.Text = objConnStr.;
                    txtNotes.Text = objConnStr.sNotes.Trim();
                    //txtConnectionString.Text = conStringBuilder.ConnectionString;
                }

                updateTxtConnectionString();
            }
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

        private void FrmConnectionString_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnDelete);
                cControlPosition.positionControl(ref btnRecent);
                cControlPosition.positionControl(ref btnOk);
                cControlPosition.positionControl(ref btnClose);

                cControlPosition.positionControl(ref dgAttributes);

                cControlPosition.positionControl(ref lblBackend);
                cControlPosition.positionControl(ref cmbBackend);
                cControlPosition.positionControl(ref lblType);
                cControlPosition.positionControl(ref cmbType);

                cControlPosition.positionControl(ref txtConnectionString);
                cControlPosition.positionControl(ref txtNotes);

            }
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
                List<string> lstAnswer = FrmDoubleInputBox.GetResult("Please enter Type and Value", "", "Type", "Value");

                if (lstAnswer.Count != null)
                {
                    if (lstAnswer.Count == 2)
                    {
                        string sType = lstAnswer[0].Trim();
                        string sValue = lstAnswer[1].Trim();
                        bool bIsOk = true;
                        int iRow = 0;

                        if (sType == string.Empty & sValue == string.Empty)
                        { bIsOk = false; }

                        if (bIsOk)
                        {
                            iRow = dgAttributes.Rows.Add();

                            DataGridViewComboBoxCell cmb = (DataGridViewComboBoxCell)dgAttributes[ColType.Index, iRow];

                            if (!cmb.Items.Contains(sType))
                            {
                                DialogResult dlgResult = MessageBox.Show("The Type: " + sType + " is not in the list of types do you want to add it?", ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (dlgResult == System.Windows.Forms.DialogResult.Yes)
                                { cmb.Items.Add(sType); }
                                else
                                { bIsOk = false; }
                            }
                        }

                        if (bIsOk)
                        {
                            dgAttributes[ColType.Index, iRow].Value = sType;
                            dgAttributes[ColValue.Index, iRow].Value = sValue;

                            updateTxtConnectionString();
                        }
                        else
                        { dgAttributes.Rows.RemoveAt(iRow); }
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

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                delete();
            }
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

        private void delete()
        {
            try
            {
                List<int> lstSelectedRows = new List<int>();

                foreach (DataGridViewCell cell in dgAttributes.SelectedCells)
                { lstSelectedRows.Add(cell.RowIndex); }
                lstSelectedRows = lstSelectedRows.Distinct().ToList();
                lstSelectedRows.Sort();

                if (lstSelectedRows.Count > 0) 
                {
                    string sMessage;

                    if (lstSelectedRows.Count == 1)
                    {
                        sMessage = "Are you sure you delete the ";
                        if (dgAttributes[ColType.Index, lstSelectedRows[0]].Value == null)
                        { sMessage += lstSelectedRows.ToString() + " down?"; }
                        else
                        { sMessage += "Item " + dgAttributes[ColType.Index, lstSelectedRows[0]].Value.ToString() + "?"; }
                    }
                    else
                    { sMessage = "Are you sure you delete the " + lstSelectedRows.Count.ToString() + " selected rows?"; }

                    DialogResult dlg = MessageBox.Show(sMessage,
                                                       ClsDefaults.messageBoxTitle(),
                                                       MessageBoxButtons.YesNo,
                                                       MessageBoxIcon.Question);

                    if (dlg == System.Windows.Forms.DialogResult.Yes)
                    {
                        for (int iRow = lstSelectedRows.Count; iRow > 0; iRow--)
                        { dgAttributes.Rows.RemoveAt(lstSelectedRows[iRow - 1]); }

                        updateTxtConnectionString();
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

        private void btnOk_Click(object sender, EventArgs e)
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
                if (txtConnectionString.Text == null)
                { sResult = ""; }
                else
                { sResult = txtConnectionString.Text.Trim(); }

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

        private void configHtmlSummary()
        {
            try
            {
                ClsDataTypes cDataTypes = new ClsDataTypes();
                ClsConfigReporter.strTableCell objCell = new ClsConfigReporter.strTableCell();
                int iTableId;
                int iRowId = 0;

                ///***************
                // *   A table   *
                // ***************/
                //cConfigReporter.TableAddNew(out iTableId, 2, "Auto generated code is located.");

                ////Add Row
                //cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                //objCell.iOrder = 0;
                //objCell.sText = "Name";
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //objCell.iOrder = 0;
                //objCell.sText = "Description";
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                ////Add Row
                //cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                //objCell.iOrder = 0;
                //objCell.sText = cInsertCode_OpenForm.FormName; //ClsMisc.ActiveVBComponent().Name; //cInsertCode_CommandBarClass.className;
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //objCell.iOrder = 0;
                //objCell.sText = "Form name.";
                //if (cInsertCode_OpenForm.isNewForm)
                //{ objCell.sHiddenText = "New Form being opened by this code."; }
                //else
                //{ objCell.sHiddenText = "Existing Form being opened by this code."; }
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                ////Add Row
                //cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                //objCell.iOrder = 0;
                //objCell.sText = cInsertCode_OpenForm.ModuleCallForm;
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //objCell.iOrder = 0;
                //objCell.sText = "Module where form is opened from.";
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                ////Add Row
                //cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                //objCell.iOrder = 0;
                //objCell.sText = cInsertCode_OpenForm.FunctionCallForm; // ClsMisc.ActiveVBComponent().Name; //cInsertCode_CommandBarClass.SampleCodeModulePrefix.Trim() + cInsertCode_CommandBarClass.className.Trim();
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //objCell.iOrder = 0;
                //objCell.sText = "Function, Sub or Property where VBA has been inserted.";
                //objCell.sHiddenText = "";
                //objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //if (cInsertCode_OpenForm.parameters.Count == 0)
                //{
                //    /***************
                //     *   A table   *
                //     ***************/
                //    cConfigReporter.TableAddNew(out iTableId, new List<int> { 1, 4 }, "Details");

                //    //Add Row
                //    cConfigReporter.TableAddNewRow(iTableId, out iRowId, false);

                //    objCell.iOrder = 0;
                //    objCell.sText = "Parameters";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //    objCell.iOrder = 0;
                //    objCell.sText = "No parameters used.";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                //}
                //else
                //{
                //    /************************
                //     *   Parameters table   *
                //     ************************/
                //    cConfigReporter.TableAddNew(out iTableId, new List<int> { 3, 1, 3, 1, 1 }, "Parameters");

                //    //Add Row
                //    cConfigReporter.TableAddNewRow(iTableId, out iRowId, true);

                //    objCell.iOrder = 0;
                //    objCell.sText = "Name internally in the form.";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //    objCell.iOrder = 0;
                //    objCell.sText = "Name outside form";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //    objCell.iOrder = 0;
                //    objCell.sText = "Value assigned to parameter";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //    objCell.iOrder = 0;
                //    objCell.sText = "Datatype";
                //    objCell.sHiddenText = "";
                //    objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //    cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //    foreach (ClsInsertCode_OpenForm.strParameter objParameter in cInsertCode_OpenForm.parameters.Distinct().OrderBy(x => x.sNamePublicOutsideForm))
                //    {
                //        //Add Row
                //        cConfigReporter.TableAddNewRow(iTableId, out iRowId);

                //        objCell.iOrder = 0;
                //        objCell.sText = objParameter.sNamePrivatelyInForm;
                //        objCell.sHiddenText = "";
                //        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //        objCell.iOrder = 0;
                //        objCell.sText = objParameter.sNamePublicOutsideForm;
                //        objCell.sHiddenText = "";
                //        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //        objCell.iOrder = 0;
                //        objCell.sText = objParameter.sValueGiveToParameter;
                //        objCell.sHiddenText = "";
                //        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);

                //        objCell.iOrder = 0;
                //        objCell.sText = cDataTypes.getName(objParameter.eDataType);
                //        objCell.sHiddenText = "";
                //        objCell.lstFormatDetails = new List<ClsConfigReporter.enumFormatDetails>();

                //        cConfigReporter.TableAddNewCell(iTableId, iRowId, objCell);
                //    }
                //}
            }
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

                FrmHtmlReportViewer frm = new FrmHtmlReportViewer(sHtml, "Open_Form");

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

        private void FrmConnectionString_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.D)
                    { delete(); }

                    if (e.KeyCode == Keys.R)
                    { recent(); }

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
    }
}
