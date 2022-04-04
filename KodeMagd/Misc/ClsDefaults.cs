using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace KodeMagd.Misc
{
    class ClsDefaults
    {
        static DateTime dtePreviousProgressBarUpdate;

        private const string sProgressBarName = "Progress Bar";
        private const string sStatusLabelName = "Status Label";
        private const string sDefaultName = "Name";
        private const string sTextAll = "<All>";
        private const string sTextCodeOutsideFunctions = "<Code Outside functions>";
        private const string csSampleCodeModulePrefix = "SampleCode_";
        private const string csWebsite = "http://www.kodeMagd.de";

        public enum enumStyle 
        {
            eStyle1,
            eStyle2,
            eStyle3,
            eStyle4,
            eStyle5,
            eStyle6
        }

        public enum enumFontSize
        {
            eFontSize_Small,
            eFontSize_Normal,
            eFontSize_Large,
            eFontSize_XLarge
        }

        public enum enumFontType
        {
            //eFont_Title,
            //eFont_SubTitle,
            eFont_Bold,
            eFont_DosLook,
            eFont_Normal
        }

        public enum enumLabelState
        {
            eLbl_normal,
            eLbl_Invisible,
            eLbl_Warning
        }

        public enum enumSpecialEffect
        {
            eEff_DosLook,
            eEff_Normal,
            eEff_Grey, 
            eEff_Invisible
        }

        public static Color FormColour
        { get { return Color.WhiteSmoke; } }

        public static Color ControlColour
        { get { return Color.White; } }

        public static Color ControlColourGrey
        { get { return Color.LightGray; } }

        public static Color FontColour
        { get { return Color.Black; } }

        public static Color FontWarningColour
        { get { return Color.Red; } }

        public static string messageBoxTitle()
        {
            try
            {
                string sResult = ClsCodeEditorGUI.csCommandBarName + " - " + ClsMisc.ActiveWorkBook().Name;

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

        public static string website
        {
            get
            {
                try
                {
                    return csWebsite;
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

        public static string formTitle
        {
            get
            {
                try
                {
                    Excel.Workbook wrk = ClsMisc.ActiveWorkBook();

                    string sName = wrk.Name;

                    return sName;
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

        public static Font GetFont(enumFontType eType)
        {
            try {
                return GetFont(eType, enumFontSize.eFontSize_Normal);
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

                return null;
            }
        }

        public static Font GetFont(enumFontType eType, enumFontSize eSize)
        {
            try {
                Font fnt;
                float fFontSize = (float)8.25;

                switch (eSize)
                {
                    case enumFontSize.eFontSize_Small:
                        fFontSize = (float)6;
                        break;
                    case enumFontSize.eFontSize_Normal:
                        fFontSize = (float)8.25;
                        break;
                    case enumFontSize.eFontSize_Large:
                        fFontSize = (float)10;
                        break;
                    case enumFontSize.eFontSize_XLarge:
                        fFontSize = (float)16;
                        break;
                }

                switch(eType)
                {
                    case enumFontType.eFont_Normal:
                        fnt = new Font("Arial", fFontSize, FontStyle.Regular);
                        break;
                    case enumFontType.eFont_Bold:
                        fnt = new Font("Arial", fFontSize, FontStyle.Bold);
                        break;
                    case enumFontType.eFont_DosLook:
                        fnt = new Font("Raster Fonts", fFontSize, FontStyle.Bold);
                        break;
                    //case enumFontType.eFont_Normal:
                    //    fnt = new Font("Arial", (float)8.25);
                    //    break;
                    default:
                        fnt = new Font("Arial", fFontSize);
                        break;
                }

                return fnt; 
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
                
                return null;
            }
        }

        public static void FormatForm(ref Form frm) 
        {
            try
            {
                //foreach (Control con in frm.Controls)
                //{
                //    if (con.GetType().ToString() == "Button") { FormatButton(ref (Button)con); }
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

        public static void FormatControl(ref Button con)
        {
            try
            {

                con.BackColor = ControlColour;
                con.ForeColor = FontColour;
                con.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref NumericUpDown txt)
        {
            try
            {
                txt.BackColor = ControlColour;
                txt.ForeColor = FontColour;
                txt.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref RichTextBox txt)
        {
            try
            {
                txt.TabStop = true;
                txt.AcceptsTab = false;
                txt.AllowDrop = true;
                txt.BackColor = ControlColour;
                txt.ForeColor = FontColour;
                txt.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TextBox txt)
        {
            try
            {
                txt.TabStop = true;
                txt.AcceptsTab = false;
                txt.AllowDrop = true;
                txt.BackColor = ControlColour;
                txt.ForeColor = FontColour;
                txt.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TextBox txt, enumSpecialEffect eSpecialEffect)
        {
            try
            {
                switch (eSpecialEffect)
                {
                    case enumSpecialEffect.eEff_Normal:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Normal);
                        txt.BackColor = ControlColour;
                        txt.Visible = true;
                        txt.ForeColor = FontColour;
                        break;
                    case enumSpecialEffect.eEff_Grey:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Bold);
                        txt.BackColor = ControlColourGrey;
                        txt.Visible = true;
                        txt.ForeColor = FontColour;
                        break;
                    case enumSpecialEffect.eEff_Invisible:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Normal);
                        txt.BackColor = Color.Transparent;
                        txt.Visible = false;
                        txt.ForeColor = FontColour;
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

        public static void FormatControl(ref TextBox txt, bool bReadOnly)
        {
            try
            {
                txt.Enabled = !bReadOnly;
                txt.AcceptsTab = true;
                txt.AllowDrop = true;
                txt.BackColor = ControlColour;
                txt.ForeColor = FontColour;
                txt.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TextBox txt, bool bReadOnly, enumSpecialEffect eSpecialEffect)
        {
            try
            {
                txt.Enabled = !bReadOnly;

                switch (eSpecialEffect)
                {
                    case enumSpecialEffect.eEff_Normal:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Normal);
                        txt.BackColor = ControlColour;
                        txt.Visible = true;
                        txt.ForeColor = FontColour;
                        break;
                    case enumSpecialEffect.eEff_DosLook:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_DosLook);
                        txt.BackColor = Color.Black;
                        txt.Visible = true;
                        txt.ForeColor = Color.White;
                        break;
                    case enumSpecialEffect.eEff_Grey:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Bold);
                        txt.BackColor = ControlColourGrey;
                        txt.Visible = true;
                        txt.ForeColor = FontColour;
                        break;
                    case enumSpecialEffect.eEff_Invisible:
                        txt.AcceptsTab = true;
                        txt.AllowDrop = true;
                        txt.Font = GetFont(enumFontType.eFont_Normal);
                        txt.BackColor = Color.Transparent;
                        txt.Visible = false;
                        txt.ForeColor = FontColour;
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

        public static void FormatControl(ref ComboBox cmb)
        {
            try
            {
                cmb.BackColor = ControlColour;
                cmb.ForeColor = FontColour;
                cmb.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref ListBox lst)
        {
            try
            {
                lst.BackColor = ControlColour;
                lst.ForeColor = FontColour;
                lst.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref CheckedListBox lst)
        {
            try
            {
                lst.BackColor = ControlColour;
                lst.ForeColor = FontColour;
                lst.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref Label lbl)
        {
            try
            {
                FormatControl(ref lbl, enumLabelState.eLbl_normal);
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

        public static void FormatControl(ref Label lbl, enumFontSize eSize)
        {
            try
            {
                FormatControl(ref lbl, enumLabelState.eLbl_normal, eSize);
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

        public static void FormatControl(ref LinkLabel lbl)
        {
            try
            {
                FormatControl(ref lbl, enumLabelState.eLbl_normal);
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

        public static void FormatControl(ref System.Windows.Forms.ToolTip tt, bool bIsVisible)
        {
            try
            {
                tt.ToolTipIcon = ToolTipIcon.Info;
                tt.UseFading = true;
                tt.UseAnimation = true;
                //tt.ShowAlways = true;
                tt.InitialDelay = 500;
                tt.AutoPopDelay = 5000;
                tt.ReshowDelay = 0;
                tt.BackColor = Color.AliceBlue;
                //tt.ToolTipIcon.Font = GetFont(enumFontType.eFont_Normal);
                tt.Active = bIsVisible;
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

        public static void FormatControl(ref TreeView tree)
        {
            try
            {
                tree.BackColor = ControlColour;
                tree.ForeColor = FontColour;
                tree.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TreeNode trNd)
        {
            try
            {
                FormatControl(ref trNd, enumStyle.eStyle1);
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

        public static void FormatControl(ref TreeNode trNd, enumStyle eStyle)
        {
            try
            {
                switch (eStyle) 
                {
                    case enumStyle.eStyle1:
                        trNd.BackColor = Color.CadetBlue;
                        trNd.ForeColor = Color.Black;
                        break;
                    case enumStyle.eStyle2:
                        trNd.BackColor = Color.DarkViolet;
                        trNd.ForeColor = Color.Black;
                        break;
                    case enumStyle.eStyle3:
                        trNd.BackColor = Color.Orange;
                        trNd.ForeColor = Color.Black;
                        break;
                    case enumStyle.eStyle4:
                        trNd.BackColor = Color.LimeGreen;
                        trNd.ForeColor = Color.Black;
                        break;
                    case enumStyle.eStyle5:
                        trNd.BackColor = Color.DarkRed;
                        trNd.ForeColor = Color.White;
                        break;
                    case enumStyle.eStyle6:
                        trNd.BackColor = Color.Wheat;
                        trNd.ForeColor = Color.Black;
                        break;
                    default:
                        trNd.BackColor = Color.Black;
                        trNd.ForeColor = Color.White;
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

        public static void FormatControl(ref Label lbl, enumLabelState eLblState)
        {
            try
            {
                FormatControl(ref lbl, eLblState, enumFontSize.eFontSize_Normal);
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


        public static void FormatControl(ref Label lbl, enumLabelState eLblState, enumFontSize eSize)
        {
            try
            {
                switch (eLblState)
                {
                    case enumLabelState.eLbl_Invisible:
                        lbl.Visible = false;
                        lbl.ForeColor = FormColour;
                        break;
                    case enumLabelState.eLbl_normal:
                        lbl.Visible = true;
                        lbl.ForeColor = FontColour;
                        break;
                    case enumLabelState.eLbl_Warning:
                        lbl.Visible = true;
                        lbl.ForeColor = FontWarningColour;
                        break;
                    default:
                        lbl.Visible = true;
                        lbl.ForeColor = FontColour;
                        break;
                }

                lbl.BackColor = System.Drawing.Color.Transparent;
                lbl.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref LinkLabel lbl, enumLabelState eLblState)
        {
            try
            {
                switch (eLblState)
                {
                    case enumLabelState.eLbl_Invisible:
                        lbl.Visible = false;
                        lbl.ForeColor = FormColour;
                        break;
                    case enumLabelState.eLbl_normal:
                        lbl.Visible = true;
                        lbl.ForeColor = FontColour;
                        break;
                    case enumLabelState.eLbl_Warning:
                        lbl.Visible = true;
                        lbl.ForeColor = FontWarningColour;
                        break;
                    default:
                        lbl.Visible = true;
                        lbl.ForeColor = FontColour;
                        break;
                }

                lbl.BackColor = System.Drawing.Color.Transparent;
                lbl.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref CheckBox chk)
        {
            try
            {
                chk.BackColor = FormColour;
                chk.ForeColor = FontColour;
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

        public static void FormatControl(ref Panel pnl)
        {
            try
            {
                pnl.BackColor = FormColour;
                pnl.ForeColor = FontColour;
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

        public static void FormatControl(ref StatusStrip ss)
        {
            try
            {
                ss.LayoutStyle = ToolStripLayoutStyle.Flow;

                ToolStripLabel tsLabel = new ToolStripLabel();

                tsLabel.Width = 2 * ss.Width / 3;

                FormatControl(ref tsLabel);

                ss.Items.Add(tsLabel);

                ToolStripProgressBar progBar = new ToolStripProgressBar();


                progBar.Width = ss.Width / 3;

                FormatControl(ref progBar);

                ss.Items.Add(progBar);

                ss.BackColor = FormColour;
                ss.ForeColor = FontColour;

                changeStatusStrip_ProgressBar(ref ss, false);
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

        public static void FormatControl(ref ToolStripLabel tsLabel)
        {
            try
            {
                tsLabel.Name = sStatusLabelName;
                tsLabel.ForeColor = ClsDefaults.FontColour;
                tsLabel.BackColor = ClsDefaults.FormColour;
                tsLabel.DisplayStyle = ToolStripItemDisplayStyle.Text;
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

        public static void FormatControl(ref ToolStripProgressBar progBar)
        {
            try
            {
                progBar.BackColor = ClsDefaults.FormColour;
                progBar.Style = ProgressBarStyle.Marquee;
                progBar.AutoSize = true;
                progBar.ForeColor = ClsDefaults.FontColour;
                progBar.Name = sProgressBarName;
                progBar.Minimum = 0;
                progBar.Maximum = 22;
                progBar.Value = 0;
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

        public static void FormatControl(ref RadioButton opt)
        {
            try
            {
                opt.BackColor = FormColour;
                opt.ForeColor = FontColour;
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

        public static void FormatControl(ref GroupBox grp)
        {
            try
            {
                grp.BackColor = FormColour;
                grp.ForeColor = FontColour;
                grp.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TabControl tbCtrl)
        {
            try
            {
                tbCtrl.BackColor = FormColour;
                tbCtrl.ForeColor = FontColour;
                tbCtrl.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref TabPage tab)
        {
            try
            {
                tab.BackColor = FormColour;
                tab.ForeColor = FontColour;
                tab.Font = GetFont(enumFontType.eFont_Normal);
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

        public static void FormatControl(ref DataGridView dg)
        {
            try
            {
                dg.RowHeadersDefaultCellStyle.BackColor = FormColour;
                dg.ColumnHeadersDefaultCellStyle.BackColor = FormColour;
                dg.BackgroundColor = FormColour;
                dg.BackColor = FormColour;
                dg.ForeColor = FontColour;
                dg.EditMode = DataGridViewEditMode.EditOnEnter;
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

        public static void changeStatusStrip_ProgressBar(ref StatusStrip ss)
        {
            try
            {
                if (dtePreviousProgressBarUpdate.AddSeconds(0.1) < DateTime.Now)
                {
                    ToolStripItem[] objTSI = ss.Items.Find(sProgressBarName, true);

                    if (objTSI.Count() != 0)
                    {
                        if (objTSI[0].GetType() == typeof(ToolStripProgressBar))
                        {
                            ToolStripProgressBar progBar = (ToolStripProgressBar)objTSI[0];

                            if (progBar.Value == progBar.Maximum)
                            { progBar.Value = progBar.Minimum; }
                            else
                            { progBar.Value++; }

                            ss.Refresh();
                        }
                    }

                    dtePreviousProgressBarUpdate = DateTime.Now;
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

        public static void changeStatusStrip_ProgressBar(ref StatusStrip ss, bool bIsVisible)
        {
            try
            {
                ToolStripItem[] objTSI = ss.Items.Find(sProgressBarName, true);

                if (objTSI.Count() != 0)
                {
                    if (objTSI[0].GetType() == typeof(ToolStripProgressBar))
                    {
                        ToolStripProgressBar progBar = (ToolStripProgressBar)objTSI[0];

                        progBar.Visible = bIsVisible;
                        ss.Refresh();
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

        public static string defaultName
        {
            get 
            {
                try
                {
                    return sDefaultName;
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

        public static string textAll
        {
            get
            {
                try
                {
                    return sTextAll;
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

        public static string textCodeOutsideFunctions
        {
            get
            {
                try
                {
                    return sTextCodeOutsideFunctions;
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
        
        public static string sampleCodeModulePrefix
        {
            get 
            {
                try
                {
                    return csSampleCodeModulePrefix;
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
    }
}
