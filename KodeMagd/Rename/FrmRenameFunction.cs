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
    public partial class FrmRenameFunction : Form
    {
        ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();
        ClsControlPosition cControlPosition = new ClsControlPosition();
        private string sTextAll = "<All>";
        
        public FrmRenameFunction()
        {
            try
            {
                InitializeComponent();

                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);
                ClsDefaults.FormatControl(ref lstModule);
                ClsDefaults.FormatControl(ref lstModuleType);
                ClsDefaults.FormatControl(ref dgFunctions);
                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(lstModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                cControlPosition.setControl(lstModuleType, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(dgFunctions, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
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

        private void FrmRenameFunction_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModuleType();
                lstModuleType.SetItemChecked(lstModuleType.Items.IndexOf(sTextAll), true);
                fillLstModules();
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

        private void fillLstModules()
        {
            try
            {
                string sPreviousValue;

                if (lstModule.Text == null)
                { sPreviousValue = ""; }
                else
                { sPreviousValue = lstModule.Text; }
                lstModule.Text = null;

                List<ClsCodeMapper.strModuleDetails> lst = cCodeMapperWrk.getLstModuleDetails();

                lst = lst.OrderBy(x => x.sName).ToList();

                lstModule.Items.Clear();

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

                foreach (ClsCodeMapper.strModuleDetails eItem in lst)
                {
                    if (bFilter)
                    {
                        if (lstModuleType.CheckedItems.Contains(ClsDataTypes.convertModuleType(eItem.eType)))
                        { lstModule.Items.Add(eItem.sName); }
                    }
                    else
                    { lstModule.Items.Add(eItem.sName); }
                }

                lstModule.Text = sPreviousValue;
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
                fillLstModules();
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
                fillLstModules();
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

        private void FrmRenameFunction_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref btnClose);
                
                cControlPosition.positionControl(ref lstModule);
                cControlPosition.positionControl(ref lstModuleType);

                cControlPosition.positionControl(ref dgFunctions);
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

        private void lstModule_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillDgFunctions();
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

        private void fillDgFunctions()
        {
            try
            {
                dgFunctions.Rows.Clear();

                if (lstModule.SelectedIndex == null)
                {

                }
                else
                {
                    string sModuleName = lstModule.Items[lstModule.SelectedIndex].ToString();

                    List<ClsCodeMapper.strFunctions> lst = cCodeMapperWrk.getLstFunctions(sModuleName);

                    lst = lst.OrderBy(x => x.sName).ToList();

                    foreach (ClsCodeMapper.strFunctions objTemp in lst)
                    {
                        int iRow = dgFunctions.Rows.Add();

                        dgFunctions[ColName.Index, iRow].Value = objTemp.sName;
                        dgFunctions[ColScope.Index, iRow].Value = ClsCodeMapper.convertToText(objTemp.eScope);
                        if (objTemp.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property)
                        { dgFunctions[ColType.Index, iRow].Value = ClsCodeMapper.convertToText(objTemp.eFunctionType) + " (" + ClsCodeMapper.convertToText(objTemp.ePropertyType) + ")"; }
                        else
                        { dgFunctions[ColType.Index, iRow].Value = ClsCodeMapper.convertToText(objTemp.eFunctionType); }

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

        //private void btnRename_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        bool bIsOk = checkOK();

        //        if (bIsOk)
        //        {
        //            renameFunction();
        //        }

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

        private bool checkOK(string sNewName)
        {
            try
            {

                bool bIsOK = true;
                bool bIsWarning = false;
                string sMessage = "";
                //string sNewName = "";

                if (dgFunctions.CurrentRow == null)
                {
                    bIsOK = false;
                    sMessage = "Please make sure you select a Function, Sub or Property.";
                }

                if (sNewName == null)
                { sNewName = ""; }

                if (sNewName == "")
                {
                    bIsOK = false;
                    sMessage = "Please make sure you enter a new name";
                }

                if (bIsOK)
                {
                    List<ClsCodeMapper.strVariables> lstVar = cCodeMapperWrk.getLstVariableDetails();

                    if (lstVar.Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower()))
                    {
                        bIsWarning = true;
                        sMessage = "There are Variables with the same name \"" + sNewName + "\"";
                    }

                    List<ClsCodeMapper.strFunctions> lstFn = cCodeMapperWrk.getLstFunctions();

                    bool bFoundFunctions = false;
                    bool bFoundSubroutines = false;
                    bool bFoundProperties = false;

                    if (lstFn.Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower() && x.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Function))
                    { bFoundFunctions = true; }

                    if (lstFn.Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower() && x.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Sub))
                    { bFoundSubroutines = true; }

                    if (lstFn.Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower() && x.eFunctionType == ClsCodeMapper.enumFunctionType.eFnType_Property))
                    { bFoundProperties = true; }

                    if (bFoundFunctions || bFoundProperties || bFoundSubroutines)
                    {
                        bIsWarning = true;
                        sMessage = "There are ";

                        if (bFoundFunctions)
                        { sMessage += "Functions, "; }

                        if (bFoundSubroutines)
                        { sMessage += "Subs, "; }

                        if (bFoundProperties)
                        { sMessage += "Properties, "; }

                        sMessage += "with the same name \"" + sNewName + "\"";
                    }

                    List<ClsCodeMapper.strModuleDetails> lstMod = cCodeMapperWrk.getLstModuleDetails();

                    if (lstMod.Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower()))
                    {
                        bIsOK = false;
                        sMessage = "The name \"" + sNewName + "\" is already used for a code module.";
                    }
                }

                if (!bIsOK)
                { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
                else
                {
                    if (bIsWarning)
                    {
                        sMessage += "\n\r\n\rAre you sure you want to proceed?";

                        DialogResult dlg = MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dlg != System.Windows.Forms.DialogResult.Yes)
                        { bIsOK = false; }
                    }
                }

                return bIsOK;
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

        //private void renameFunction() 
        //{
        //    try
        //    {
        //        bool bIsOk = true;
        //        string sMessage = "";
        //        string sOldName = "";
        //        string sNewName = "";
        //        string sModuleName = "";

        //        int iRow = 0;

        //        if (dgFunctions.CurrentRow == null)
        //        {
        //            bIsOk = false;
        //            sMessage = "No function or Sub is selected";
        //        }
        //        else
        //        { iRow = dgFunctions.CurrentRow.Index;}

        //        if (dgFunctions[ColName.Index, iRow].Value == null)
        //        {
        //            bIsOk = false;
        //            sMessage = "No function or Sub is selected";
        //        }
        //        else if (dgFunctions[ColName.Index, iRow].Value == "")
        //        {
        //            bIsOk = false;
        //            sMessage = "No function or Sub is selected";
        //        }
        //        else
        //        { sOldName = dgFunctions[ColName.Index, iRow].Value.ToString(); }

        //        if (txtNewName.Text == "")
        //        { sNewName = ""; }
        //        else
        //        { sNewName = txtNewName.Text; }

        //        if (lstModule.Text == null)
        //        { sModuleName = ""; }
        //        else
        //        { sModuleName = lstModule.Text; }

        //        sNewName = sNewName.Trim();
        //        sOldName = sOldName.Trim();
        //        sModuleName = sModuleName.Trim();

        //        if (sNewName == "")
        //        {
        //            bIsOk = false;
        //            sMessage = "Please enter a New Name.";
        //        }
        //        else 
        //        { bIsOk = ClsMisc.validVariableNameCheck(sNewName, out sMessage); }

        //        if (bIsOk)
        //        { 
        //            cCodeMapperWrk.renameFunction(sNewName, sOldName, sModuleName);
        //            this.Close();
        //        }
        //        else
        //        { MessageBox.Show(sMessage, ClsDefaults.messageBoxTitle(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
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

        //private void dgFunctions_SelectionChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        string sNewValue = "";

        //        if (dgFunctions.CurrentRow != null) 
        //        {
        //            int iRow = dgFunctions.CurrentRow.Index;

        //            if (dgFunctions[ColName.Index, iRow].Value == null)
        //            { sNewValue = ""; }
        //            else
        //            { sNewValue = dgFunctions[ColName.Index, iRow].Value.ToString(); }

        //            //txtNewName.Text = sNewValue;
        //        }
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

        private void dgFunctions_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                bool bIsOk = true;
                string sMessage = "";
                string sOldName = "";
                string sNewName = "";
                string sModuleName = "";

                if (dgFunctions[ColName.Index, e.RowIndex].Value == null)
                {
                    bIsOk = false;
                    sMessage = "No function or Sub is selected";
                }
                else if (dgFunctions[ColName.Index, e.RowIndex].Value.ToString().Trim() == "")
                {
                    bIsOk = false;
                    sMessage = "No function or Sub is selected";
                }
                else
                { sOldName = dgFunctions[ColName.Index, e.RowIndex].Value.ToString(); }

                if (dgFunctions[ColName.Index, e.RowIndex].EditedFormattedValue == null)
                {
                    bIsOk = false;
                    sMessage = "No function or Sub is selected";
                }
                else if (dgFunctions[ColName.Index, e.RowIndex].EditedFormattedValue.ToString().Trim() == "")
                {
                    bIsOk = false;
                    sMessage = "No function or Sub is selected";
                }
                else
                { sNewName = dgFunctions[ColName.Index, e.RowIndex].EditedFormattedValue.ToString(); }

                if (lstModule.Text == null)
                { sModuleName = ""; }
                else
                { sModuleName = lstModule.Text; }

                sNewName = sNewName.Trim();
                sOldName = sOldName.Trim();
                sModuleName = sModuleName.Trim();

                if (sOldName.Trim().ToLower() != sNewName.Trim().ToLower())
                {

                    if (sNewName == "")
                    {
                        bIsOk = false;
                        sMessage = "Please enter a New Name.";
                    }
                    else
                    { bIsOk = ClsMisc.validVariableNameCheck(sNewName, out sMessage); }

                    if (cCodeMapperWrk.getLstFunctions(sModuleName).Exists(x => x.sName.Trim().ToLower() == sNewName.Trim().ToLower()))
                    {
                        bIsOk = false;
                        sMessage = "There is already a Function or Sub  or Property with the name '" + sNewName + "'";
                    }

                    if (bIsOk)
                    {
                        cCodeMapperWrk.renameFunction(sNewName, sOldName, sModuleName);
                        cCodeMapperWrk.renameFunctionInList(sNewName, sOldName, sModuleName);
                        
                        //if it's a get/set/let in a class then check if there are anyother identical names on the GUI that need to be changed
                        for (int iRow = 0; iRow < dgFunctions.Rows.Count; iRow++) 
                        {
                            string sTempName;
                            if (dgFunctions[ColName.Index, iRow].Value == null)
                            { sTempName = ""; }
                            else
                            { sTempName = dgFunctions[ColName.Index, iRow].Value.ToString(); }

                            if (sTempName.Trim().ToLower() == sOldName.Trim().ToLower())
                            { dgFunctions[ColName.Index, iRow].Value = sNewName; }
                        }
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

        private void FrmRenameFunction_KeyDown(object sender, KeyEventArgs e)
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
