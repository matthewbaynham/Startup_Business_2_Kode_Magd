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

namespace KodeMagd.InsertCode
{
    public partial class FrmInsertCode_Rst_Range : Form
    {
        public FrmInsertCode_Rst_Range()
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

        private void FrmInsertCode_Rst_Range_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref lblColumn);
                ClsDefaults.FormatControl(ref lblNamedRange);
                ClsDefaults.FormatControl(ref lblRow);
                ClsDefaults.FormatControl(ref lblShtName);
                ClsDefaults.FormatControl(ref lblWrkName);

                ClsDefaults.FormatControl(ref txtColumn);
                ClsDefaults.FormatControl(ref txtNamedRange);
                ClsDefaults.FormatControl(ref txtRow);
                ClsDefaults.FormatControl(ref txtShtName);
                ClsDefaults.FormatControl(ref txtWrkName);

                ClsDefaults.FormatControl(ref grpType);

                ClsDefaults.FormatControl(ref optCoordinates);
                ClsDefaults.FormatControl(ref optNamedRange);

                ClsDefaults.FormatControl(ref ssStatus);

                optCoordinates.Checked = true;
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

        public string wrkName
        {
            get
            {
                try
                {
                    string sTemp;

                    if (string.IsNullOrEmpty(txtWrkName.Text))
                    { sTemp = ""; }
                    else
                    { sTemp = txtWrkName.Text; }

                    return sTemp;
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
            set
            {
                try
                {
                    txtWrkName.Text = value;
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

        public string shtName
        {
            get
            {
                try
                {
                    string sTemp;

                    if (string.IsNullOrEmpty(txtShtName.Text))
                    { sTemp = ""; }
                    else
                    { sTemp = txtShtName.Text; }

                    return sTemp;
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
            set
            {
                try
                {
                    txtShtName.Text = value;
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

        public string namedRange
        {
            get
            {
                try
                {
                    string sTemp;

                    if (string.IsNullOrEmpty(txtNamedRange.Text))
                    { sTemp = ""; }
                    else
                    { sTemp = txtNamedRange.Text; }

                    return sTemp;
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
            set
            {
                try
                {
                    txtNamedRange.Text = value;
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

        public int row
        {
            get
            {
                try
                {
                    int iTemp;

                    if (string.IsNullOrEmpty(txtRow.Text))
                    { iTemp = 1; }
                    else
                    {
                        if (!int.TryParse(txtRow.Text, out iTemp))
                        { iTemp = 1; }
                    }

                    return iTemp;
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

                    return 1;
                }
            }
            set
            {
                try
                {
                    txtRow.Text = value.ToString();
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

        public int column
        {
            get
            {
                try
                {
                    int iTemp;

                    if (string.IsNullOrEmpty(txtColumn.Text))
                    { iTemp = 1; }
                    else
                    {
                        if (!int.TryParse(txtColumn.Text, out iTemp))
                        { iTemp = 1; }
                    }

                    return iTemp;
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

                    return 1;
                }
            }
            set
            {
                try
                {
                    txtColumn.Text = value.ToString();
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

        public ClsInsertCode_Rst.enumDestinationTypeRangeType rangeType
        {
            get
            {
                try
                {
                    ClsInsertCode_Rst.enumDestinationTypeRangeType eTemp;

                    if (optCoordinates.Checked & !optNamedRange.Checked)
                    { eTemp = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds; }
                    else if (!optCoordinates.Checked & optNamedRange.Checked)
                    { eTemp = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named; }
                    else
                    { eTemp = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds; }

                    return eTemp;
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

                    return ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds;
                }
            }
            set
            {
                try
                {
                    ClsInsertCode_Rst.enumDestinationTypeRangeType eTemp;
                    eTemp = value;

                    switch (eTemp) 
                    {
                        case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds:
                            optCoordinates.Checked = true;
                            optNamedRange.Checked = false;
                            break;
                        case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named:
                            optCoordinates.Checked = false;
                            optNamedRange.Checked = true;
                            break;
                        default:
                            optCoordinates.Checked = false;
                            optNamedRange.Checked = false;
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
        }

        private void optNamedRange_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                assignType();
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

        private void optCoordinates_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                assignType();
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

        private void assignType() 
        {
            try
            {
                ClsInsertCode_Rst.enumDestinationTypeRangeType eType = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Unknown;

                if (!optNamedRange.Checked & optCoordinates.Checked)
                { eType = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds;}
                else if (optNamedRange.Checked & !optCoordinates.Checked)
                { eType = ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named;}

                switch (eType)
                {
                    case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Coordinateds:
                        lblColumn.Visible = true;
                        txtColumn.Visible = true;
                        lblRow.Visible = true;
                        txtRow.Visible = true;

                        lblNamedRange.Visible = false;
                        txtNamedRange.Visible = false;
                        break;
                    case ClsInsertCode_Rst.enumDestinationTypeRangeType.eRng_Named:
                        lblColumn.Visible = false;
                        txtColumn.Visible = false;
                        lblRow.Visible = false;
                        txtRow.Visible = false;

                        lblNamedRange.Visible = true;
                        txtNamedRange.Visible = true;
                        break;
                    default:
                        lblColumn.Visible = false;
                        txtColumn.Visible = false;
                        lblRow.Visible = false;
                        txtRow.Visible = false;

                        lblNamedRange.Visible = false;
                        txtNamedRange.Visible = false;
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

        private void FrmInsertCode_Rst_Range_KeyDown(object sender, KeyEventArgs e)
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
