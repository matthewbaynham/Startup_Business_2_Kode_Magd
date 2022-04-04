using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using Microsoft.Vbe.Interop.Forms;
using System.Reflection;
using KodeMagd.Misc;
using Microsoft.Win32;

namespace KodeMagd.Settings
{
    public partial class FrmAddReference : Form
    {
//        private StatusStrip ssParent; not really working
        ClsControlPosition cControlPosition = new ClsControlPosition();

        private ClsListViewColumnSorter lvwColumnSorter; 
        private string sRefName = "";
        private string sRefGUID = "";
        private ClsReferences.enumFilterType eFilterType;

        /* http://stackoverflow.com/questions/742851/how-add-a-com-exposed-net-project-to-the-vb6-or-vba-references-dialog
         * The entries in the References dialog box come from the HKCR\TypeLib registry key, 
         * not from HKCR\CLSID. If your assembly does not show up in the References dialog 
         * but compiled DLL's can still use your COM assembly, it means the classes and 
         * interfaces were correctly registered for your assembly, but that the type library 
         * itself was not. This would appear to be a bug with VS2008 in Vista, as mentioned 
         * in the forum thread you linked to. Running regasm explicitly from your MSI should 
         * fix it
         */
        //http://en.wikipedia.org/wiki/Windows_Registry (good one)
        //http://en.wikipedia.org/wiki/Globally_Unique_Identifier
        //http://stackoverflow.com/questions/12723821/list-all-available-progid
        //http://msdn.microsoft.com/en-us/library/k3677y81.aspx
        //http://msdn.microsoft.com/en-us/library/ms724072%28v=vs.85%29.aspx
        //http://support.microsoft.com/kb/256986
        //http://msdn.microsoft.com/en-us/library/b2hs0tae%28v=vs.71%29.aspx (might be important or might not)
        //http://msdn.microsoft.com/en-us/library/microsoft.win32.registrykey%28v=vs.80%29.aspx

        public ClsReferences.enumFilterType filterType
        {
            get 
            {
                try
                {
                    return eFilterType;
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

                    return ClsReferences.enumFilterType.eFilt_None;
                }
            }
            set 
            {
                try
                {
                    eFilterType = value;
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

        public FrmAddReference(ClsReferences.enumFilterType filterType, ref StatusStrip ss)
        {
            try
            {
                InitializeComponent();

                //ss.Refresh();
                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                eFilterType = filterType;

                // Create an instance of a ListView column sorter and assign it 
                // to the ListView control.
                lvwColumnSorter = new ClsListViewColumnSorter();

                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                lstVwAssembliesTypeLib.ListViewItemSorter = lvwColumnSorter;

                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                lvwColumnSorter.SortColumn = colAssName.Index;

                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                setLabel();

                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                chkFiltered.Checked = true;

                ClsDefaults.changeStatusStrip_ProgressBar(ref ss);

                fillLstVw(eFilterType, ref ss);
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

        private void FrmAddReference_Load(object sender, EventArgs e)
        {
            try
            {
                this.Text = ClsDefaults.formTitle;
                this.BackColor = ClsDefaults.FormColour;

                ClsDefaults.FormatControl(ref btnAddReference);
                ClsDefaults.FormatControl(ref btnClose);

                ClsDefaults.FormatControl(ref chkFiltered);

                ClsDefaults.FormatControl(ref lblComments);

                ClsDefaults.FormatControl(ref ssStatus);

                cControlPosition.setControl(btnAddReference, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(btnClose, ClsControlPosition.enumAnchorHorizontal.eAnchor_Right, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lblComments, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);
                cControlPosition.setControl(chkFiltered, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_Bottom);

                cControlPosition.setControl(lstVwAssembliesTypeLib, ClsControlPosition.enumAnchorHorizontal.eAnchor_StretchHor, ClsControlPosition.enumAnchorVertical.eAnchor_StretchVert);

                //move these to the initialise event so that we can update the status bar on the previous form
                /*
                chkFiltered.Checked = true;

                fillLstVw(eFilterType);
                */
                reLayout();
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

        public void setLabel()
        {
            try
            {
                string sReferenceName = "";

                switch (filterType) 
                {
                    case ClsReferences.enumFilterType.eFilt_Access:
                        sReferenceName = "Access";
                        break;
                    case ClsReferences.enumFilterType.eFilt_ADO:
                        sReferenceName = "ADO";
                        break;
                    case ClsReferences.enumFilterType.eFilt_None:
                        sReferenceName = "";

                        break;
                    case ClsReferences.enumFilterType.eFilt_Outlook:
                        sReferenceName = "Outlook";
                        break;
                    case ClsReferences.enumFilterType.eFilt_Scripting:
                        sReferenceName = "Scripting";
                        break;
                    default:
                        break;
                }

                string sText;

                if (filterType == ClsReferences.enumFilterType.eFilt_None)
                { 
                    sText = "No Filtering";
                    chkFiltered.Enabled = false;
                }
                else
                {
                    sText = "Please select the version of the " + sReferenceName + " Reference you wish to use.";
                    chkFiltered.Enabled = true;
                }

                lblComments.Text = sText;

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

        private void addReference()
        {
            try
            {
                Scripting.FileSystemObject fso = new Scripting.FileSystemObject();
                Excel.Workbook wrk = ClsMisc.ActiveWorkBook();
                
                foreach (ListViewItem objItem in lstVwAssembliesTypeLib.SelectedItems)
                {
                    string sPath = objItem.SubItems[colAssPath.Index].Text;

                    if (fso.FileExists(sPath))
                    {
                        try
                        {
                            wrk.VBProject.References.AddFromFile(sPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(text: "Add Reference failed:\n\rPlease check you don't already have the refence or any reference that would clash.", caption: "Add Reference Failed", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        try
                        {
                            string sGUID = objItem.SubItems[colAssGUID.Index].Text;
                            string sVersion = objItem.SubItems[colAssVersion.Index].Text;
                            
                            int iVersionMajor = ClsMisc.getVersionMajor(sVersion);
                            int iVersionMinor = ClsMisc.getVersionMinor(sVersion);

                            if (!(iVersionMajor == -1 || iVersionMinor == -1))
                            { wrk.VBProject.References.AddFromGuid(sGUID, iVersionMajor, iVersionMinor); }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(text: "Add Reference failed:\n\rPlease check you don't already have the refence or any reference that would clash.", caption: "Add Reference Failed", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                        }
                    }
                }

                wrk = null;
                fso = null;
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

        private void btnAddReference_Click(object sender, EventArgs e)
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
                addReference();

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

        public bool referenceAlreadySet
        {
            get 
            {
                try
                {
                    bool bResult = false;

                    return bResult;
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
        }

        private void fillLstVw(ClsReferences.enumFilterType eFilterType, ref StatusStrip ss)
        {
            try
            {
                //empty listview
                lstVwAssembliesTypeLib.Items.Clear();

                ClsReferences cReferences;
                if (chkFiltered.Checked)
                { cReferences = new ClsReferences(eFilterType, ref ss); }
                else
                { cReferences = new ClsReferences(ClsReferences.enumFilterType.eFilt_None, ref ss);}

                List<ClsReferences.strAsssembly> lstAss = cReferences.assembliesTypeLib;

                if (lstAss != null)
                {
                    foreach (ClsReferences.strAsssembly objAss in lstAss)
                    {
                        ClsDefaults.changeStatusStrip_ProgressBar(ref ss);
                        ListViewItem objValue = new ListViewItem(new string[] { objAss.sName, objAss.sVersion, objAss.sPath, objAss.sGUID, objAss.sWinXX });

                        lstVwAssembliesTypeLib.Items.Add(objValue);
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

        public void reLayout()
        {
            try
            {
                int iFormWidth = 0;
                int iFormHeight = 0;

                if (this.Width < 480)
                { iFormWidth = 480; }
                else
                { iFormWidth = this.Width; }

                if (this.Height < 360)
                { iFormHeight = 360; }
                else
                { iFormHeight = this.Height; }

                btnClose.Top = iFormHeight - (btnClose.Height + 50);
                btnClose.Left = iFormWidth - (btnClose.Width + 20);

                btnAddReference.Top = iFormHeight - (btnAddReference.Height + 50);
                btnAddReference.Left = iFormWidth - (btnClose.Width + btnAddReference.Width + 40);

                lstVwAssembliesTypeLib.Height = iFormHeight - 90;
                lstVwAssembliesTypeLib.Width = iFormWidth - 40;

                chkFiltered.Top = iFormHeight - (btnClose.Height + 50);

                lblComments.Top = iFormHeight - (btnClose.Height + 50 - lblComments.Height);
            
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

        private void FrmAddReference_Resize(object sender, EventArgs e)
        {
            try
            {
                //reLayout();
                cControlPosition.positionControl(ref btnAddReference);
                cControlPosition.positionControl(ref btnClose);

                cControlPosition.positionControl(ref lblComments);
                cControlPosition.positionControl(ref chkFiltered);

                cControlPosition.positionControl(ref lstVwAssembliesTypeLib);

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

        private void chkFiltered_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                ClsDefaults.changeStatusStrip_ProgressBar(ref ssStatus);
                lstVwAssembliesTypeLib.Enabled = false;
                ClsDefaults.changeStatusStrip_ProgressBar(ref ssStatus);
                chkFiltered.Enabled = false;
                ClsDefaults.changeStatusStrip_ProgressBar(ref ssStatus);

                fillLstVw(this.filterType, ref ssStatus);
                
                lstVwAssembliesTypeLib.Enabled = true;
                chkFiltered.Enabled = true;
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

        private void lstVwAssembliesTypeLib_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            try
            {
                // Determine if clicked column is already the column that is being sorted.
                if (e.Column == lvwColumnSorter.SortColumn)
                {
                    // Reverse the current sort direction for this column.
                    if (lvwColumnSorter.Order == SortOrder.Ascending)
                    {
                        lvwColumnSorter.Order = SortOrder.Descending;
                    }
                    else
                    {
                        lvwColumnSorter.Order = SortOrder.Ascending;
                    }
                }
                else
                {
                    // Set the column number that is to be sorted; default to ascending.
                    lvwColumnSorter.SortColumn = e.Column;
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }

                // Perform the sort with these new sort options.
                this.lstVwAssembliesTypeLib.Sort();
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

        private void lstVwAssembliesTypeLib_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                addReference();

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

        private void FrmAddReference_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.A)
                    { add(); }

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
