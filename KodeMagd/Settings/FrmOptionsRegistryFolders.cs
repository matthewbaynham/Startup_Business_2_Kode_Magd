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

namespace KodeMagd.Settings
{
    public partial class FrmOptionsRegistryFolders : Form
    {
        private ClsControlPosition cControlPosition = new ClsControlPosition();

        public FrmOptionsRegistryFolders()
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

        /*
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                bool bIsOk = true;

                if (bIsOk)
                {
                    string sResult = "";
                    List<string> lstTemp = new List<string>();

                    foreach (string sTemp in lstRegistryFolders.Items)
                    { lstTemp.Add(sTemp); }

                    lstTemp = lstTemp.Distinct().ToList();
                    lstTemp.Sort();

                    sResult = string.Join("\t", lstTemp.ToArray());




                    ClsSettings cSettings = new ClsSettings();

                    lstRegistryFolders.Items.Clear();

                    cSettings.addRegistryDirForReferences();

                    this.Close();
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
        */
        public void fillLstRegistryFolders()
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();
                
                lstRegistryFolders.Items.Clear();

                List<string> lst = cSettings.registryDirForReferences;

                lst.Sort();

                foreach (string sTemp in lst)
                { lstRegistryFolders.Items.Add(sTemp); }
            }
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
                //FrmInputBox frm = new FrmInputBox();
                string sNew = FrmInputBox.GetString("Add Folder", "Please enter the full path of the folder within the registry");

                sNew = sNew.Trim();

                if (sNew != "")
                {
                    ClsSettings cSettings = new ClsSettings();

                    lstRegistryFolders.Items.Add(sNew);

                    cSettings.addRegistryDirForReferences(sNew);
                    cSettings.Save();

                    cSettings = null;
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
                removeItem();
            }
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

        private void removeItem() 
        {
            try
            {
                DialogResult drAreYouSure = MessageBox.Show("Remove registry directory?", "Remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (drAreYouSure == DialogResult.Yes)
                {
                    if (lstRegistryFolders.SelectedItems.Count == 1)
                    {
                        ClsSettings cSettings = new ClsSettings();
                        int iSelected = lstRegistryFolders.SelectedIndex;

                        string sDir = lstRegistryFolders.Items[iSelected].ToString();
                        //cSettings.addRegistryDirForReferences(sDir);
                        cSettings.removeRegistryDirForReferences(sDir);
                        cSettings.Save();

                        lstRegistryFolders.Items.RemoveAt(iSelected);

                        cSettings = null;
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

        private void FrmOptionsRegistryFolders_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;
                
                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnAdd);
                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref lstRegistryFolders);
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl((Control)btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl((Control)btnAdd, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl((Control)btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl((Control)lstRegistryFolders, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                fillLstRegistryFolders();
            }
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

        private void FrmOptionsRegistryFolders_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                cControlPosition.positionControl(ref btnAdd);
                cControlPosition.positionControl(ref btnRemove);
                cControlPosition.positionControl(ref lstRegistryFolders);
            }
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

        private void lstRegistryFolders_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                removeItem();
            }
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

        private void FrmOptionsRegistryFolders_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

                    if (e.KeyCode == Keys.R)
                    { removeItem(); }

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
