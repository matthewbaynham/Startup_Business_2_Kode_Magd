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

namespace KodeMagd.Rename
{
    public partial class FrmRenameModuleOrForm : Form
    {
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        ClsControlPosition cControlPosition = new ClsControlPosition();
        private string sTextAll = "<All>";
        public FrmRenameModuleOrForm()
        {
            try
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref dgModule);
                ClsDefaults.FormatControl(ref lstModuleType);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                
                cControlPosition.setControl(dgModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(lstModuleType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
            }
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

        private void FrmRenameModuleOrForm_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModuleType();
                lstModuleType.SetItemChecked(lstModuleType.Items.IndexOf("<All>"), true);
                fillDgModules();
            }
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

        private void fillCmbModuleType()
        {
            try
            {
                List<string> lst = new List<string>();

                foreach (Microsoft.Vbe.Interop.vbext_ComponentType eTemp in Enum.GetValues(typeof(Microsoft.Vbe.Interop.vbext_ComponentType)))
                {
                    string sTemp = ClsDataTypes.convertModuleType(eTemp);
                    lst.Add(sTemp.Trim());
                }
                lst.Sort();

                lstModuleType.Items.Clear();
                lstModuleType.Items.Add(sTextAll);
                lstModuleType.SetSelected(lstModuleType.Items.IndexOf(sTextAll), true);
                foreach (string sTemp in lst)
                { 
                    lstModuleType.Items.Add(sTemp);
                    lstModuleType.SetSelected(lstModuleType.Items.IndexOf(sTemp), false);
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

        private void fillDgModules()
        {
            try
            {
                List<ClsCodeMapper.strModuleDetails> lst = cCodeMapperWrk.getLstModuleDetails().OrderBy(x => x.sName).ToList();

                bool bFilter;

                if (lstModuleType.CheckedItems.Count == 0)
                { bFilter = false; }
                else
                {
                    if (lstModuleType.CheckedItems.Contains(sTextAll) | lstModuleType.CheckedItems.Count == lstModuleType.Items.Count)
                    { bFilter = false; }
                    else
                    { bFilter = true; }
                }

                dgModule.Rows.Clear();
                foreach (ClsCodeMapper.strModuleDetails eItem in lst)
                {
                    bool bAdd;

                    if (bFilter)
                    {
                        if (lstModuleType.CheckedItems.Contains(ClsDataTypes.convertModuleType(eItem.eType)))
                        { bAdd = true; }
                        else
                        { bAdd = false; }
                    }
                    else
                    { bAdd = true; }
 
                    if (bAdd)
                    {
                        int iRow = dgModule.Rows.Add();

                        dgModule[ColName.Index, iRow].Value = eItem.sName;
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

        private void checkAllSelected()
        {
            try
            {
                if (lstModuleType.SelectedItem == lstModuleType.Items[lstModuleType.Items.IndexOf(sTextAll)]) 
                {
                    if (lstModuleType.CheckedItems.Contains(sTextAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < lstModuleType.Items.Count; iIndex++)
                        {
                            if (iIndex != lstModuleType.Items.IndexOf(sTextAll))
                            { lstModuleType.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (lstModuleType.CheckedItems.Count == 0 | (lstModuleType.CheckedItems.Count == 1 & lstModuleType.CheckedItems.Contains(sTextAll)))
                    { bAnySelected=false; }
                    else
                    { bAnySelected=true; }

                    if (bAnySelected)
                    { lstModuleType.SetItemChecked(lstModuleType.Items.IndexOf(sTextAll), false); }
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

        private void lstModuleType_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelected();
                fillDgModules();
            }
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

        private void lstModuleType_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelected();
                fillDgModules();
            }
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

        private void FrmRenameModuleOrForm_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                
                cControlPosition.positionControl(ref dgModule);
                cControlPosition.positionControl(ref lstModuleType);
            }
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

        private void dgModule_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";
                string sOldName;
                string sNewName;

                if (dgModule[e.ColumnIndex, e.RowIndex].Value == null)
                { sOldName = ""; }
                else
                { sOldName = dgModule[e.ColumnIndex, e.RowIndex].Value.ToString(); }

                if (dgModule[e.ColumnIndex, e.RowIndex].EditedFormattedValue == null)
                { sNewName = ""; }
                else
                { sNewName = dgModule[e.ColumnIndex, e.RowIndex].EditedFormattedValue.ToString(); }

                if (sNewName.Trim().ToLower() != sOldName.Trim().ToLower())
                {

                    string sNamingError = "";
                    if (!ClsMisc.validVariableNameCheck(sNewName, out sMessage))
                    {
                        bIsOk = false;
                        sMessage = sNamingError;
                    }

                    if (cCodeMapperWrk.getLstModuleDetails().Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower()))
                    {
                        bIsOk = false;
                        sMessage = "Module all ready exists with the name '" + sNewName + "'";
                    }

                    if (bIsOk)
                    {
                        cCodeMapperWrk.renameModule(sNewName, sOldName);

                        //update cCodeMapper
                        cCodeMapperWrk.renameModuleInList(sOldName, sNewName);
                        //cCodeMapperWrk.renameVariableInFunctionList(sOldName, sNewName);
                    }
                    else
                    { 
                        MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        e.Cancel = true;
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

        private void FrmRenameModuleOrForm_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
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
