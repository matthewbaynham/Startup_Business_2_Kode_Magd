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

namespace KodeMagd.Format
{
    public partial class FrmRemoveLineNo : Form
    {
        ClsControlPosition cControlPosition = new ClsControlPosition();
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();

        public FrmRemoveLineNo()
        {
            try
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnRemove);
                ClsDefaults.FormatControl(ref btnClose);
                
                ClsDefaults.FormatControl(ref lblModules);
                ClsDefaults.FormatControl(ref lstModules);
                
                ClsDefaults.FormatControl(ref lblFunctions);
                ClsDefaults.FormatControl(ref lstFunctions);
                
                ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnRemove, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstModules, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                
                cControlPosition.setControl(lblFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                
                cControlPosition.setControl(lblWarning, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
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

        private void FrmRemoveLineNo_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                
                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillLstModule();
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

        private void fillLstModule()
        {
            try
            {
                List<string> lst = new List<string>();

                foreach (ClsCodeMapper.strModuleDetails cModuleDetails in cCodeMapperWrk.getLstModuleDetails())
                { lst.Add(cModuleDetails.sName.Trim()); }
                lst.Sort();

                lstModules.Items.Clear();
                lstModules.Items.Add(ClsDefaults.textAll);
                lstModules.SetSelected(lstModules.Items.IndexOf(ClsDefaults.textAll), true);
                foreach (string sTemp in lst)
                {
                    lstModules.Items.Add(sTemp);
                    lstModules.SetSelected(lstModules.Items.IndexOf(sTemp), false);
                }

                lstModules.SetItemChecked(lstModules.Items.IndexOf(ClsDefaults.textAll), true);
                fillLstFunctions();
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

        private void lstModules_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelectedModules();
                fillLstFunctions();
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

        private void lstModules_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelectedModules();
                fillLstFunctions();
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

        private void checkAllSelectedModules()
        {
            try
            {
                if (lstModules.SelectedItem == lstModules.Items[lstModules.Items.IndexOf(ClsDefaults.textAll)])
                {
                    if (lstModules.CheckedItems.Contains(ClsDefaults.textAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < lstModules.Items.Count; iIndex++)
                        {
                            if (iIndex != lstModules.Items.IndexOf(ClsDefaults.textAll))
                            { lstModules.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (lstModules.CheckedItems.Count == 0 | (lstModules.CheckedItems.Count == 1 & lstModules.CheckedItems.Contains(ClsDefaults.textAll)))
                    { bAnySelected = false; }
                    else
                    { bAnySelected = true; }

                    if (bAnySelected)
                    { lstModules.SetItemChecked(lstModules.Items.IndexOf(ClsDefaults.textAll), false); }
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

        private void checkAllSelectedFunction()
        {
            try
            {
                if (lstFunctions.SelectedItem == lstFunctions.Items[lstModules.Items.IndexOf(ClsDefaults.textAll)])
                {
                    if (lstFunctions.CheckedItems.Contains(ClsDefaults.textAll))
                    {
                        //unselect all other items
                        for (int iIndex = 0; iIndex < lstFunctions.Items.Count; iIndex++)
                        {
                            if (iIndex != lstFunctions.Items.IndexOf(ClsDefaults.textAll))
                            { lstFunctions.SetItemChecked(iIndex, false); }
                        }
                    }
                }
                else
                {
                    //if any of the other items are selected deselect <All>
                    bool bAnySelected;

                    if (lstFunctions.CheckedItems.Count == 0 | (lstFunctions.CheckedItems.Count == 1 & lstFunctions.CheckedItems.Contains(ClsDefaults.textAll)))
                    { bAnySelected = false; }
                    else
                    { bAnySelected = true; }

                    if (bAnySelected)
                    { lstFunctions.SetItemChecked(lstFunctions.Items.IndexOf(ClsDefaults.textAll), false); }
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

        public void fillLstFunctions() 
        {
            try
            {
                //string sPreviousValue;
                bool bIncludePrefixModName;
                bool bWarning = false;

                if (lstModules.CheckedItems.Count != 1 | lstModules.CheckedItems.Contains(ClsDefaults.textAll))
                { bIncludePrefixModName = true; }
                else
                { bIncludePrefixModName = false; }

                lstFunctions.Text = null;

                List<ClsCodeMapper.strModuleDetails> lst = cCodeMapperWrk.getLstModuleDetails();

                lst = lst.OrderBy(x => x.sName).ToList();

                lstFunctions.Items.Clear();

                bool bFilter;

                if (lstModules.CheckedItems.Count == 0)
                { bFilter = false; }
                else
                {
                    if (lstModules.CheckedItems.Contains(ClsDefaults.textAll) | lstModules.CheckedItems.Count == lstModules.Items.Count)
                    { bFilter = false; }
                    else
                    { bFilter = true; }
                }

                int iAllIndex = lstFunctions.Items.Add(ClsDefaults.textAll);

                foreach (ClsCodeMapper.strModuleDetails objMod in lst)
                {
                    bool bAddModule = false;

                    if (bFilter)
                    {
                        if (lstModules.CheckedItems.Contains(objMod.sName))
                        { bAddModule = true; }
                    }
                    else
                    { bAddModule = true; }

                    if (bAddModule)
                    {
                        foreach (string sFunctionName in cCodeMapperWrk.getLstFunctionNames(objMod.sName, bIncludePrefixModName).Distinct().OrderBy(x => x))
                        {
                            string sItem = sFunctionName;

                            if (bIncludePrefixModName == true && sFunctionName.Contains('.'))
                            {
                                int iPosDot = sFunctionName.IndexOf('.');
                                string sBeginning = ClsMiscString.Left(sFunctionName, iPosDot);
                                string sEnding = ClsMiscString.Right(sFunctionName, sFunctionName.Length - iPosDot - 1);

                                if (cCodeMapperWrk.existsGoToWithLineNo(sEnding , sBeginning ))
                                {
                                    sItem += " (*)";
                                    bWarning = true;
                                }
                            }
                            else
                            {
                                if (cCodeMapperWrk.existsGoToWithLineNo(sFunctionName, objMod.sName))
                                {
                                    sItem += " (*)";
                                    bWarning = true;
                                }
                            }
                            lstFunctions.Items.Add(sItem);
                        }
                    }
                }

                lstFunctions.SetItemChecked(iAllIndex, true);
                
                if (bWarning)
                {
                    lblWarning.Text = "(*) = Function, Sub or Property contains a GOTO which requires line numbers";
                    ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_normal);
                }
                else
                { ClsDefaults.FormatControl(ref lblWarning, ClsDefaults.enumLabelState.eLbl_Invisible); }
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

        private void FrmRemoveLineNo_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnRemove);
                cControlPosition.positionControl(ref btnClose);
                
                cControlPosition.positionControl(ref lblModules);
                cControlPosition.positionControl(ref lstModules);
                
                cControlPosition.positionControl(ref lblFunctions);
                cControlPosition.positionControl(ref lstFunctions);
                
                cControlPosition.positionControl(ref lblWarning);
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

        private void btnRemove_Click(object sender, EventArgs e)
        {
            try
            {
                RemoveNumber();
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

        private void RemoveNumber() 
        {
            try
            {
                foreach (string sFunction in lstFunctions.CheckedItems)
                {
                    if (sFunction.Contains("."))
                    { 
                        int iPos = sFunction.IndexOf('.');

                        string sModuleName = ClsMiscString.Left(sFunction, iPos);
                        string sFunctionName = ClsMiscString.Right(sFunction, sFunction.Length - iPos - 1);

                        if (ClsMiscString.Right(ref sFunctionName, 4) == " (*)")
                        { sFunctionName = ClsMiscString.Left(ref sFunctionName, sFunctionName.Length - 4); }

                        cCodeMapperWrk.removeLineNo(sFunctionName, sModuleName);
                    }
                    else
                    {
                        string sModuleName = "";
                        string sFunctionName = sFunction;

                        if (lstModules.CheckedItems.Count == 1)
                        { sModuleName = lstModules.CheckedItems[0].ToString(); }

                        if (ClsMiscString.Right(ref sFunctionName, 4) == " (*)")
                        { sFunctionName = ClsMiscString.Left(ref sFunctionName, sFunctionName.Length - 4); }

                        cCodeMapperWrk.removeLineNo(sFunctionName, sModuleName);
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

        private void lstFunctions_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                checkAllSelectedFunction();
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

        private void lstFunctions_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                checkAllSelectedFunction();
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

        private void FrmRemoveLineNo_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.R)
                    { RemoveNumber(); }

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
