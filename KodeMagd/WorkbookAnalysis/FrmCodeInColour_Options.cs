using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using KodeMagd.Settings;
using KodeMagd.Misc;
using KodeMagd.Reporter;

namespace KodeMagd.WorkbookAnalysis
{
    public partial class FrmCodeInColour_Options : Form
    {
        private enum enumMode
        {
            eSetDefaults,
            eOneReport
        }

        private enumMode eFormMode;
        private bool bOK = false;

        private List<ClsConfigReporter.strCss> lstResults = new List<ClsConfigReporter.strCss>();

        public FrmCodeInColour_Options()
        {
            try
            {
                eFormMode = enumMode.eSetDefaults;
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

        public FrmCodeInColour_Options(List<ClsConfigReporter.strCss> lstSettings)
        {
            try
            {
                eFormMode = enumMode.eOneReport;
                this.lstResults = lstSettings;
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


        public List<ClsConfigReporter.strCss> result
        {
            get
            {
                try
                {
                    return lstResults;
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

                    return new List<ClsConfigReporter.strCss>();
                }
            }
        }

        private void fillColourButtons()
        {
            try
            {
                switch (this.eFormMode)
                {
                    case enumMode.eSetDefaults:
                        ClsSettings_CodeInColour cSettings_CodeInColour = new ClsSettings_CodeInColour();

                        btnDeclaringVariable.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_DeclareVariables); 
                        btnAssigningValues.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_AssignVariables);
                        btnIfStatements.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_If);
                        btnLoops.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Loops);
                        btnFunctions.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_DeclareFunctions);
                        btnComments.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Comments);
                        btnErrorCode.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Errors);
                        btnWith.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_With);

                        cSettings_CodeInColour = null;
                        break;
                    case enumMode.eOneReport:
                        
                        setButtonColour(ref btnDeclaringVariable, ClsCodeInColour.sCssName_DeclareVariables);
                        setButtonColour(ref btnAssigningValues, ClsCodeInColour.sCssName_AssignVariables);
                        setButtonColour(ref btnIfStatements, ClsCodeInColour.sCssName_IfStatements);
                        setButtonColour(ref btnLoops, ClsCodeInColour.sCssName_Loops);
                        setButtonColour(ref btnFunctions, ClsCodeInColour.sCssName_DeclareFunctions);
                        setButtonColour(ref btnComments, ClsCodeInColour.sCssName_Comments);
                        setButtonColour(ref btnErrorCode, ClsCodeInColour.sCssName_Errors);
                        setButtonColour(ref btnWith, ClsCodeInColour.sCssName_With);
                        break;
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

        private void setButtonColour(ref Button btn, string sColourName)
        {
            try
            {
                int iIndex = this.lstResults.FindIndex(x => x.sName.Trim().ToUpper() == sColourName.Trim().ToUpper());

                if (iIndex != -1)
                {
                    ClsConfigReporter.strCss objCssStyle = this.lstResults[iIndex];

                    int iIndexItem = objCssStyle.lstCssStyles.FindIndex (x => x.sName.Trim().ToUpper() == "color".Trim().ToUpper());

                    if (iIndexItem != -1)
                    { btn.BackColor = ClsMisc.convertRGBColour(objCssStyle.lstCssStyles[iIndexItem].sValue); }
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

        private void FrmCodeInColour_Options_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref lblTitle);
                ClsDefaults.FormatControl(ref btnOK);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref lblAssigningValues);
                ClsDefaults.FormatControl(ref btnAssigningValues);
                ClsDefaults.FormatControl(ref lblComments);
                ClsDefaults.FormatControl(ref btnComments);
                ClsDefaults.FormatControl(ref lblDeclaringVariable);
                ClsDefaults.FormatControl(ref btnDeclaringVariable);
                ClsDefaults.FormatControl(ref lblErrorCode);
                ClsDefaults.FormatControl(ref btnErrorCode);
                ClsDefaults.FormatControl(ref lblFunctions);
                ClsDefaults.FormatControl(ref btnFunctions);
                ClsDefaults.FormatControl(ref lblIfStatements);
                ClsDefaults.FormatControl(ref btnIfStatements);
                ClsDefaults.FormatControl(ref lblLoops);
                ClsDefaults.FormatControl(ref btnLoops);
                ClsDefaults.FormatControl(ref lblWith);
                ClsDefaults.FormatControl(ref btnWith);

                ClsDefaults.FormatControl(ref ssStatus);

                fillColourButtons();

                switch (eFormMode)
                {
                    case enumMode.eOneReport:
                        lblTitle.Text = "Please select text colours for this report.";
                        btnResetDefault.Visible = false;
                        break;
                    case enumMode.eSetDefaults:
                        lblTitle.Text = "Please select text colours for default settings.";
                        btnResetDefault.Visible = true;
                        break;
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

        private void btnDeclaringVariable_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnDeclaringVariable.BackColor;

                dlgColour.ShowDialog(this);

                btnDeclaringVariable.BackColor = dlgColour.Color;
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

        private void btnAssigningValues_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnAssigningValues.BackColor;

                dlgColour.ShowDialog(this);

                btnAssigningValues.BackColor = dlgColour.Color;
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

        private void btnIfStatements_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnIfStatements.BackColor;

                dlgColour.ShowDialog(this);

                btnIfStatements.BackColor = dlgColour.Color;
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

        private void btnLoops_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnLoops.BackColor;

                dlgColour.ShowDialog(this);

                btnLoops.BackColor = dlgColour.Color;
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

        private void btnFunctions_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnFunctions.BackColor;

                dlgColour.ShowDialog(this);

                btnFunctions.BackColor = dlgColour.Color;
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

        private void btnComments_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnComments.BackColor;

                dlgColour.ShowDialog(this);

                btnComments.BackColor = dlgColour.Color;
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

        private void btnErrorCode_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnErrorCode.BackColor;

                dlgColour.ShowDialog(this);

                btnErrorCode.BackColor = dlgColour.Color;
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

        private void btnWith_Click(object sender, EventArgs e)
        {
            try
            {
                dlgColour = new ColorDialog();

                dlgColour.AnyColor = true;
                dlgColour.SolidColorOnly = true;
                dlgColour.AllowFullOpen = true;
                dlgColour.Color = btnWith.BackColor;

                dlgColour.ShowDialog(this);

                btnWith.BackColor = dlgColour.Color;
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
                switch (eFormMode) 
                {
                    case enumMode.eOneReport:
                        this.bOK = true;
                        storeResults();
                        break;
                    case enumMode.eSetDefaults:
                        saveSettings();
                        break;
                    default:
                        break;
                }

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

        private void storeResults()
        {
            try
            {
                storeResults(ref btnAssigningValues, ClsCodeInColour.sCssName_AssignVariables);
                storeResults(ref btnComments, ClsCodeInColour.sCssName_Comments);
                storeResults(ref btnFunctions, ClsCodeInColour.sCssName_DeclareFunctions);
                storeResults(ref btnDeclaringVariable, ClsCodeInColour.sCssName_DeclareVariables);
                storeResults(ref btnErrorCode, ClsCodeInColour.sCssName_Errors);
                storeResults(ref btnIfStatements, ClsCodeInColour.sCssName_IfStatements);
                storeResults(ref btnLoops, ClsCodeInColour.sCssName_Loops);
                storeResults(ref btnWith, ClsCodeInColour.sCssName_With);

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

        private void storeResults(ref Button btn, string sColourName)
        {
            try
            {
                string sColour = ClsMisc.convertColourRGB(btn.BackColor);
                int iIndex = lstResults.FindIndex(x => x.sName == sColourName);
                ClsConfigReporter.strCss objCssStyle;

                if (iIndex == -1)
                {
                    objCssStyle = new ClsConfigReporter.strCss();
                    objCssStyle.sName = sColourName;
                    objCssStyle.lstCssStyles = new List<ClsConfigReporter.strCssStyle>();
                }
                else
                {
                    objCssStyle = lstResults[iIndex];
                }

                if (objCssStyle.lstCssStyles.Exists(x => x.sName.Trim().ToUpper() == "color".Trim().ToUpper()))
                {
                    int iIndexColour = objCssStyle.lstCssStyles.FindIndex(x => x.sName.Trim().ToUpper() == "color".Trim().ToUpper());

                    ClsConfigReporter.strCssStyle objCssStyleItem = objCssStyle.lstCssStyles[iIndexColour];

                    objCssStyleItem.sValue = sColour;

                    objCssStyle.lstCssStyles[iIndexColour] = objCssStyleItem;
                }
                else
                {
                    ClsConfigReporter.strCssStyle objCssStyleItem = new ClsConfigReporter.strCssStyle();

                    objCssStyleItem.sName = "color";
                    objCssStyleItem.sValue = sColour;

                    objCssStyle.lstCssStyles.Add(objCssStyleItem);
                }

                if (iIndex == -1)
                { lstResults.Add(objCssStyle); }
                else
                { lstResults[iIndex] = objCssStyle; }
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

        private void saveSettings()
        { 
            try
            {
                ClsSettings_CodeInColour cSettings_CodeInColour = new ClsSettings_CodeInColour();

                cSettings_CodeInColour.lineColour_DeclareVariables = ClsMisc.convertColourRGB(btnDeclaringVariable.BackColor);
                cSettings_CodeInColour.lineColour_AssignVariables = ClsMisc.convertColourRGB(btnAssigningValues.BackColor);
                cSettings_CodeInColour.lineColour_If = ClsMisc.convertColourRGB(btnIfStatements.BackColor);
                cSettings_CodeInColour.lineColour_Loops = ClsMisc.convertColourRGB(btnLoops.BackColor);
                cSettings_CodeInColour.lineColour_DeclareFunctions = ClsMisc.convertColourRGB(btnFunctions.BackColor);
                cSettings_CodeInColour.lineColour_Comments= ClsMisc.convertColourRGB(btnComments.BackColor);
                cSettings_CodeInColour.lineColour_Errors = ClsMisc.convertColourRGB(btnErrorCode.BackColor);
                cSettings_CodeInColour.lineColour_With = ClsMisc.convertColourRGB(btnWith.BackColor);

                cSettings_CodeInColour.Save();
                cSettings_CodeInColour = null;
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

        public bool OK
        {
            get
            {
                try
                {
                    return this.bOK;
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
            set 
            {
                try
                {
                    this.bOK = value;
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

        private void btnResetDefault_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dlg = MessageBox.Show("Are you sure you want to reset back to the default settings?", "Reset", MessageBoxButtons.YesNo);

                if (dlg == System.Windows.Forms.DialogResult.Yes)
                {
                    ClsSettings_CodeInColour cSettings_CodeInColour = new ClsSettings_CodeInColour();

                    cSettings_CodeInColour.Reset();
                    cSettings_CodeInColour.Save();

                    btnDeclaringVariable.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_DeclareVariables);
                    btnAssigningValues.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_AssignVariables);
                    btnIfStatements.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_If);
                    btnLoops.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Loops);
                    btnFunctions.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_DeclareFunctions);
                    btnComments.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Comments);
                    btnErrorCode.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_Errors);
                    btnWith.BackColor = ClsMisc.convertRGBColour(cSettings_CodeInColour.lineColour_With);

                    cSettings_CodeInColour = null;

                    fillColourButtons();
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
