using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    class ClsControlPosition
    {
        public enum enumAnchorHorizontal
        {
            eAnchor_Left,  //all of a control gets anchored to the left edge of the form
            eAnchor_Right, //all of a control gets anchored to the right edge of the form
            eAnchor_StretchHor //the left side of a control gets anchored to the left and the right side gets anchored to the right
        };

        public enum enumAnchorVertical
        {
            eAnchor_Top,
            eAnchor_Bottom,
            eAnchor_StretchVert
        };

        public struct strControl
        {
            public enumAnchorHorizontal eAnchHor;
            public enumAnchorVertical eAnchVert;
            public string sName;
            //public int iId;
            public int iLeft;
            public int iTop;
            public int iWidth;
            public int iHeight;
            public int iFormWidth;
            public int iFormHeight;
        };

        List<strControl> lstCtrl = new List<strControl>();

        public ClsControlPosition()
        {
            try
            {
                lstCtrl.Clear();
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

        ~ClsControlPosition()
        {
            try
            {
                lstCtrl = null;
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

        public void setControl(System.Windows.Forms.Control cntl, enumAnchorHorizontal eAnchHor, enumAnchorVertical eAnchVert)
        {
            try
            {
                int iFormWidth = cntl.FindForm().Width;
                int iFormHeight = cntl.FindForm().Height;

                strControl objCntrl = new strControl();

                objCntrl.sName = cntl.Name;
                objCntrl.iLeft = cntl.Left;
                objCntrl.iTop = cntl.Top;
                objCntrl.iWidth = cntl.Width;
                objCntrl.iHeight = cntl.Height;
                //objCntrl.iId = cntl.i;
                objCntrl.eAnchHor = eAnchHor;
                objCntrl.eAnchVert = eAnchVert;
                objCntrl.iFormWidth = iFormWidth;
                objCntrl.iFormHeight = iFormHeight;

                lstCtrl.Add(objCntrl);
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

        public void positionControl(ref System.Windows.Forms.Control cntl)
        {
            try
            {
                if (cntl.FindForm() == cntl.Parent)
                {
                    string sCtrlName = cntl.Name;
                    int iFormWidth = cntl.FindForm().Width;
                    int iFormHeight = cntl.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex > 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    cntl.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    cntl.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    cntl.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    cntl.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref Button btn)
        {
            try
            {
                if (btn.FindForm() == btn.Parent)
                {
                    string sCtrlName = btn.Name;
                    int iFormWidth = btn.FindForm().Width;
                    int iFormHeight = btn.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    btn.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    btn.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    btn.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    btn.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref ListBox lst)
        {
            try
            {
                if (lst.FindForm() == lst.Parent)
                {
                    string sCtrlName = lst.Name;
                    int iFormWidth = lst.FindForm().Width;
                    int iFormHeight = lst.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    lst.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    lst.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    lst.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    lst.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref CheckedListBox lst)
        {
            try
            {
                if (lst.FindForm() == lst.Parent)
                {
                    string sCtrlName = lst.Name;
                    int iFormWidth = lst.FindForm().Width;
                    int iFormHeight = lst.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    lst.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    lst.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    lst.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    lst.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref TextBox txt)
        {
            try
            {
                if (txt.FindForm() == txt.Parent)
                {
                    string sCtrlName = txt.Name;
                    int iFormWidth = txt.FindForm().Width;
                    int iFormHeight = txt.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    txt.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    txt.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    txt.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    txt.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref RichTextBox rtf)
        {
            try
            {
                if (rtf.FindForm() == rtf.Parent)
                {
                    string sCtrlName = rtf.Name;
                    int iFormWidth = rtf.FindForm().Width;
                    int iFormHeight = rtf.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    rtf.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    rtf.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    rtf.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    rtf.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref Label lbl)
        {
            try
            {
                if (lbl.FindForm() == lbl.Parent)
                {
                    string sCtrlName = lbl.Name;
                    int iFormWidth = lbl.FindForm().Width;
                    int iFormHeight = lbl.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    lbl.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    lbl.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    lbl.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    lbl.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref GroupBox grp)
        {
            try
            {
                if (grp.FindForm() == grp.Parent)
                {
                    string sCtrlName = grp.Name;
                    int iFormWidth = grp.FindForm().Width;
                    int iFormHeight = grp.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    grp.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    grp.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    grp.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    grp.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref RadioButton opt)
        {
            try
            {
                if (opt.FindForm() == opt.Parent)
                {
                    string sCtrlName = opt.Name;
                    int iFormWidth = opt.FindForm().Width;
                    int iFormHeight = opt.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    opt.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    opt.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    opt.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    opt.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref TreeView tv)
        {
            try
            {
                if (tv.FindForm() == tv.Parent)
                {
                    string sCtrlName = tv.Name;
                    int iFormWidth = tv.FindForm().Width;
                    int iFormHeight = tv.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    tv.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    tv.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    tv.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    tv.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref CheckBox chk)
        {
            try
            {
                if (chk.FindForm() == chk.Parent)
                {
                    string sCtrlName = chk.Name;
                    int iFormWidth = chk.FindForm().Width;
                    int iFormHeight = chk.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    chk.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    chk.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    chk.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    chk.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref ComboBox cmb)
        {
            try
            {
                if (cmb.FindForm() == cmb.Parent)
                {
                    string sCtrlName = cmb.Name;
                    int iFormWidth = cmb.FindForm().Width;
                    int iFormHeight = cmb.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    cmb.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    cmb.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    cmb.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    cmb.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref ListView lv)
        {
            try
            {
                if (lv.FindForm() == lv.Parent)
                {
                    string sCtrlName = lv.Name;
                    int iFormWidth = lv.FindForm().Width;
                    int iFormHeight = lv.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex != null)
                    {
                        if (iIndex >= 0 & iIndex < lstCtrl.Count)
                        {
                            strControl objCntrl = lstCtrl[iIndex];

                            switch (objCntrl.eAnchHor)
                            {
                                case enumAnchorHorizontal.eAnchor_Left:
                                    break;
                                case enumAnchorHorizontal.eAnchor_Right:
                                    lv.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                                case enumAnchorHorizontal.eAnchor_StretchHor:
                                    lv.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                    break;
                            }

                            switch (objCntrl.eAnchVert)
                            {
                                case enumAnchorVertical.eAnchor_Top:
                                    break;
                                case enumAnchorVertical.eAnchor_Bottom:
                                    lv.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                                case enumAnchorVertical.eAnchor_StretchVert:
                                    lv.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                    break;
                            }
                        }
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

        public void positionControl(ref DataGridView dg)
        {
            try
            {
                if (dg.FindForm() == dg.Parent)
                {
                    string sCtrlName = dg.Name;
                    int iFormWidth = dg.FindForm().Width;
                    int iFormHeight = dg.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex >= 0 & iIndex < lstCtrl.Count)
                    {
                        strControl objCntrl = lstCtrl[iIndex];

                        switch (objCntrl.eAnchHor)
                        {
                            case enumAnchorHorizontal.eAnchor_Left:
                                break;
                            case enumAnchorHorizontal.eAnchor_Right:
                                dg.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                break;
                            case enumAnchorHorizontal.eAnchor_StretchHor:
                                dg.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                break;
                        }

                        switch (objCntrl.eAnchVert)
                        {
                            case enumAnchorVertical.eAnchor_Top:
                                break;
                            case enumAnchorVertical.eAnchor_Bottom:
                                dg.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                break;
                            case enumAnchorVertical.eAnchor_StretchVert:
                                dg.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                break;
                        }
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

        public void positionControl(ref Panel pnl)
        {
            try
            {
                if (pnl.FindForm() == pnl.Parent)
                {
                    string sCtrlName = pnl.Name;
                    int iFormWidth = pnl.FindForm().Width;
                    int iFormHeight = pnl.FindForm().Height;

                    int iIndex = lstCtrl.FindIndex(x => x.sName == sCtrlName);

                    if (iIndex >= 0 & iIndex < lstCtrl.Count)
                    {
                        strControl objCntrl = lstCtrl[iIndex];

                        switch (objCntrl.eAnchHor)
                        {
                            case enumAnchorHorizontal.eAnchor_Left:
                                break;
                            case enumAnchorHorizontal.eAnchor_Right:
                                pnl.Left = (objCntrl.iLeft - objCntrl.iFormWidth) + iFormWidth;
                                break;
                            case enumAnchorHorizontal.eAnchor_StretchHor:
                                pnl.Width = (objCntrl.iWidth - objCntrl.iFormWidth) + iFormWidth;
                                break;
                        }

                        switch (objCntrl.eAnchVert)
                        {
                            case enumAnchorVertical.eAnchor_Top:
                                break;
                            case enumAnchorVertical.eAnchor_Bottom:
                                pnl.Top = (objCntrl.iTop - objCntrl.iFormHeight) + iFormHeight;
                                break;
                            case enumAnchorVertical.eAnchor_StretchVert:
                                pnl.Height = (objCntrl.iHeight - objCntrl.iFormHeight) + iFormHeight;
                                break;
                        }
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
    }
}
