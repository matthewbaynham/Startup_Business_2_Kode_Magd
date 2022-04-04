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
using KodeMagd.InsertCode;
using KodeMagd.Settings;
using KodeMagd.WorkbookAnalysis;

namespace KodeMagd
{
    public partial class FrmOptions : Form
    {
        public FrmOptions()
        {
            try
            {
                InitializeComponent();

                //createMenu();
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

        private void FrmOptions_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref btnOK);

                ClsDefaults.FormatControl(ref btnCodeInColourDefaults);
                ClsDefaults.FormatControl(ref btnRegistryFolders);
                ClsDefaults.FormatControl(ref btnDefault);
                
                ClsDefaults.FormatControl(ref txtIndentSize);

                ClsDefaults.FormatControl(ref lblSplitLinesDescription, ClsDefaults.enumLabelState.eLbl_normal);

                ClsDefaults.FormatControl(ref chkFocusActivePane);
                ClsDefaults.FormatControl(ref chkIndentFirstLevel);
                ClsDefaults.FormatControl(ref chkIndentFirstLevel);
                ClsDefaults.FormatControl(ref chkSplitLines);
                ClsDefaults.FormatControl(ref chkUseWith);

                ClsDefaults.FormatControl(ref lblTxtFormattingLineBreakMethodology, ClsDefaults.enumLabelState.eLbl_normal);

                ClsDefaults.FormatControl(ref txtFormatCutTextChar);
                ClsDefaults.FormatControl(ref txtIndentSize);

                ClsDefaults.FormatControl(ref ssStatus);

                fillControls();
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

        private void fillControls() 
        {
            try
            {
                ClsSettings cSettings = new ClsSettings();

                txtIndentSize.Value = (decimal)cSettings.IndentSize;
                chkIncludeErrorHandler.Checked = cSettings.InsertErrorHandlers;
                chkIndentFirstLevel.Checked = cSettings.IndentFirstLevel;
                chkSplitLines.Checked = cSettings.SplitConcatinatedLines;
                chkFocusActivePane.Checked = cSettings.SetFocusActivePane;
                chkUseWith.Checked = cSettings.UseWith;
                chkTips.Checked = cSettings.UserTips;

                fillCmbFormatLineCutMethodology();
                fillCmbFormatOfDimStatement();
                cmbFormatLineCutMethodology.Text = ClsMisc.Convert_FormatLineCutMethodology(cSettings.FormatLineCutMethodology);
                cmbFormatOfDimStatement.Text = ClsMisc.Convert_FormatLineVarDim(cSettings.FormatVarDimType);

                txtFormatCutTextChar.Value = cSettings.InsertCode_Format_CharCutOffPoint;

                cSettings = null;
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
                ClsSettings cSettings = new ClsSettings();

                cSettings.IndentSize = (int)this.txtIndentSize.Value;
                cSettings.IndentFirstLevel = this.chkIndentFirstLevel.Checked;
                cSettings.SplitConcatinatedLines = this.chkSplitLines.Checked;
                cSettings.SetFocusActivePane = this.chkFocusActivePane.Checked;
                cSettings.InsertErrorHandlers = this.chkIncludeErrorHandler.Checked;
                cSettings.UseWith = this.chkUseWith.Checked;
                cSettings.FormatLineCutMethodology = ClsMisc.Convert_FormatLineCutMethodology(cmbFormatLineCutMethodology.Text);
                cSettings.FormatVarDimType = ClsMisc.Convert_FormatLineVarDim(cmbFormatOfDimStatement.Text);
                cSettings.InsertCode_Format_CharCutOffPoint = (int)txtFormatCutTextChar.Value;
                cSettings.UserTips = chkTips.Checked;

                cSettings.Save();

                cSettings = null;

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

        private void fillCmbFormatLineCutMethodology()
        { 
            try
            {
                List<string> lstItems = new List<string>();

                foreach (ClsInsertCode.enumFormatLineCutMethodology eTemp in Enum.GetValues(typeof(ClsInsertCode.enumFormatLineCutMethodology)))
                { lstItems.Add(ClsMisc.Convert_FormatLineCutMethodology(eTemp)); }
                lstItems.Sort();

                cmbFormatLineCutMethodology.Items.Clear();
                foreach(string sTemp in lstItems)
                { cmbFormatLineCutMethodology.Items.Add(sTemp); }
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

        private void fillCmbFormatOfDimStatement()
        {
            try
            {
                List<string> lstItems = new List<string>();

                foreach (ClsCodeMapper.enumVarDimType eTemp in Enum.GetValues(typeof(ClsCodeMapper.enumVarDimType)))
                { lstItems.Add(ClsMisc.Convert_FormatLineVarDim(eTemp)); }
                lstItems.Sort();

                cmbFormatOfDimStatement.Items.Clear();
                foreach (string sTemp in lstItems)
                { cmbFormatOfDimStatement.Items.Add(sTemp); }
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

        private void btnRegistryFolders_Click(object sender, EventArgs e)
        {
            try
            {
                registryFolders();
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

        private void registryFolders()
        {
            try
            {
                FrmOptionsRegistryFolders frm = new FrmOptionsRegistryFolders();

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

        private void btnDefault_Click(object sender, EventArgs e)
        {
            try
            {
                resetDefault();
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

        private void resetDefault()
        {
            try
            {
                DialogResult drAreYouSure = MessageBox.Show("Are you sure you want to reset all the setting to the default values?", 
                                                            "Reset", 
                                                            MessageBoxButtons.YesNo,
                                                            MessageBoxIcon.Question);

                if (drAreYouSure == DialogResult.Yes)
                {
                    ClsSettings cSettings = new ClsSettings();

                    cSettings.Reset();

                    cSettings = null;

                    fillControls();
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

        private void cmbFormatLineCutMethodology_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbFormatLineCutMethodology.Text == ClsMisc.Convert_FormatLineCutMethodology(ClsInsertCode.enumFormatLineCutMethodology.eFmtLineCut_AfterXChar))
                {
                    txtFormatCutTextChar.Visible = true;
                    lblTextFormattingLineBreakCutAfterNumberOfChar.Visible = true;
                }
                else
                {
                    txtFormatCutTextChar.Visible = false;
                    lblTextFormattingLineBreakCutAfterNumberOfChar.Visible = false;
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

        private void FrmOptions_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.D)
                    { resetDefault(); }

                    if (e.KeyCode == Keys.R)
                    { registryFolders(); }

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

        private void btnCodeInColourDefaults_Click(object sender, EventArgs e)
        {
            try
            {
                FrmCodeInColour_Options frm = new FrmCodeInColour_Options();

                frm.ShowDialog();

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
