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
    public partial class FrmRenameVariable : Form
    {
        List<ClsCodeMapper.strVariables> lstVariables = new List<ClsCodeMapper.strVariables>();
        private ClsControlPosition cControlPosition = new ClsControlPosition();
        private ClsCodeMapperWrk cCodeMapperWrk = new ClsCodeMapperWrk();

        public FrmRenameVariable()
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

        private void FrmRenameVariable_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref lblModule);
                ClsDefaults.FormatControl(ref lblVariables);

                ClsDefaults.FormatControl(ref lstModule);
                ClsDefaults.FormatControl(ref dgVariables);

                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref ssStatus);


                cControlPosition.setControl(lblModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(lstModule, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                cControlPosition.setControl(lblVariables, ClsControlPosition.enumAnchorHorizontal.eAnchor_Left, ClsControlPosition.enumAnchorVertical.eAnchor_Top);
                cControlPosition.setControl(dgVariables, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);
                
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cCodeMapperWrk.Wrk = ClsMisc.ActiveWorkBook();

                fillCmbModuleType();
                fillLstModules();
                fillDgVariables();
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

        private void FrmRenameVariable_Resize(object sender, EventArgs e)
        {
            try
            {
                cControlPosition.positionControl(ref lblModule);
                cControlPosition.positionControl(ref lstModule);

                cControlPosition.positionControl(ref lblVariables);
                cControlPosition.positionControl(ref dgVariables);

                cControlPosition.positionControl(ref btnClose);
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

                cmbModuleType.Items.Clear();
                cmbModuleType.Items.Add("<All>");
                foreach(string sTemp in lst)
                { cmbModuleType.Items.Add(sTemp); }
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

                if (cmbModuleType.Text == null)
                { bFilter = false; }
                else
                {
                    if (cmbModuleType.Text == "<All>" | cmbModuleType.Text == "")
                    { bFilter = false; }
                    else
                    { bFilter = true; }
                }

                foreach (ClsCodeMapper.strModuleDetails eItem in lst)
                {
                    if (bFilter)
                    {
                        if (cmbModuleType.Text == ClsDataTypes.convertModuleType(eItem.eType))
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

        private void fillDgVariables()
        {
            try
            {
                lstVariables = cCodeMapperWrk.getLstVariableDetails();

                while (dgVariables.RowCount > 0)
                { dgVariables.Rows.RemoveAt(0); }

                if (lstModule.Text == null | lstModule.Text.Trim() == "")
                { lstVariables = new List<ClsCodeMapper.strVariables>(); }
                else
                { lstVariables = cCodeMapperWrk.getLstVariableDetails(lstModule.Text); }

                lstVariables = lstVariables.OrderBy(x => x.sName).ToList();

                foreach (ClsCodeMapper.strVariables objVariable in lstVariables)
                {
                    int iRow = dgVariables.Rows.Add();

                    dgVariables[ColIndex.Index, iRow].Value = lstVariables.IndexOf(objVariable);
                    dgVariables[ColName.Index, iRow].Value = objVariable.sName;
                    dgVariables[ColType.Index, iRow].Value = objVariable.sDatatype;
                    dgVariables[ColFunction.Index, iRow].Value = objVariable.sFunctionName;
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

        private void cmbModule_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillDgVariables();
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

        private void cmbModuleType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fillLstModules();

                lstModule.Text = "";

                while (dgVariables.RowCount > 0)
                { dgVariables.Rows.RemoveAt(0); }
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
                fillDgVariables();
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

        private void dgVariables_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                string sCellsValueOnEntering = "";
                string sCellsValueOnLeaving = "";
                bool bIsChanged = false;
                bool bIsOk = true;
                string sMessage = "";

                if (dgVariables[e.ColumnIndex, e.RowIndex].Value == null)
                { sCellsValueOnEntering = ""; }
                else
                { sCellsValueOnEntering = dgVariables[e.ColumnIndex, e.RowIndex].Value.ToString(); }

                if (dgVariables[e.ColumnIndex, e.RowIndex].EditedFormattedValue == null)
                { sCellsValueOnLeaving = ""; }
                else
                { sCellsValueOnLeaving = dgVariables[e.ColumnIndex, e.RowIndex].EditedFormattedValue.ToString(); }

                if (e.ColumnIndex == ColName.Index)
                {
                    string sNameError;
                    if (!ClsMisc.validVariableNameCheck(sCellsValueOnLeaving, out sNameError))
                    {
                        bIsOk = false;
                        sMessage = sNameError;
                    }
                }

                if (sCellsValueOnEntering != sCellsValueOnLeaving)
                {
                    int iVarIndex = -1;

                    if (dgVariables[ColName.Index, e.RowIndex].Value != null)
                    {
                        string sVarIndex = dgVariables[ColIndex.Index, e.RowIndex].Value.ToString();

                        if (!int.TryParse(sVarIndex, out iVarIndex))
                        { iVarIndex = -1; }
                    }

                    if (iVarIndex == -1)
                    {
                        bIsOk = false; 
                        sMessage = "Ooops there is a problem selecting the variable."; 
                    }
                    else if (iVarIndex < 0 && iVarIndex < lstVariables.Count - 1)
                    {
                        bIsOk = false;
                        sMessage = "Error: internal variable is wrong.";
                    }
                    else
                    {
                        ClsCodeMapper.strVariables varTemp = lstVariables[iVarIndex];

                        bool bIsNameUsed = false;

                        switch (varTemp.eScope)
                        {
                            case ClsCodeMapper.enumScopeVar.eScope_Function:
                                bIsNameUsed = cCodeMapperWrk.variableNameExistInFunction(sCellsValueOnLeaving, varTemp.sFunctionName, varTemp.sModuleName);

                                if (!bIsNameUsed)
                                { bIsNameUsed = cCodeMapperWrk.variableNameExistsGlobally(sCellsValueOnLeaving, true); }
                                break;
                            case ClsCodeMapper.enumScopeVar.eScope_Global:
                                //check no module has this variable at all
                                bIsNameUsed = cCodeMapperWrk.variableNameExistsGlobally(sCellsValueOnLeaving, false);
                                break;
                            case ClsCodeMapper.enumScopeVar.eScope_Module:
                                //need to check locals in same module
                                bIsNameUsed = cCodeMapperWrk.variableNameExistsInModule(sCellsValueOnLeaving, varTemp.sModuleName);
                                if (!bIsNameUsed)
                                {
                                    //only checks globals in all modules
                                    bIsNameUsed = cCodeMapperWrk.variableNameExistsGlobally(sCellsValueOnLeaving, true);
                                }
                                break;
                        }

                        if (ClsMiscString.ingoreNull(sCellsValueOnLeaving).Trim() == "")
                        {
                            bIsOk = false;
                            sMessage = "Please enter a new value name.";
                        }

                        if (bIsOk)
                        {
                            if (bIsNameUsed)
                            {
                                bIsOk = false;
                                sMessage = "Variable Name you have entered will have conflict, please choose another name.";
                            }
                            else
                            {
                                switch (varTemp.eScope)
                                {
                                    case ClsCodeMapper.enumScopeVar.eScope_Function:
                                        cCodeMapperWrk.renameVariable(sCellsValueOnLeaving, varTemp.sName, varTemp.sModuleName, varTemp.sFunctionName);
                                        break;
                                    case ClsCodeMapper.enumScopeVar.eScope_Global:
                                        cCodeMapperWrk.renameVariable(sCellsValueOnLeaving, varTemp.sName);
                                        break;
                                    case ClsCodeMapper.enumScopeVar.eScope_Module:
                                        cCodeMapperWrk.renameVariable(sCellsValueOnLeaving, varTemp.sName, varTemp.sModuleName);
                                        break;
                                }

                                cCodeMapperWrk.renameVariableInFunctionList(varTemp.sName, sCellsValueOnLeaving, varTemp.sFunctionName, varTemp.sModuleName);

                                varTemp.sName = sCellsValueOnLeaving;
                                lstVariables[iVarIndex] = varTemp;
                                bIsChanged = true;
                            }
                        }

                        if (!bIsChanged)
                        { 
                            e.Cancel = true;
                        }
                    }
                }

                if (!bIsOk)
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

        private void FrmRenameVariable_KeyDown(object sender, KeyEventArgs e)
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
