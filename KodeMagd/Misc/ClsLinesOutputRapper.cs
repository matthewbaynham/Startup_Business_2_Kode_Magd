using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    public class ClsLinesOutputRapper
    {
        public struct strLineOut
        {
            public int iOrder;
            public string sLine;
        }

        private List<strLineOut> lstLines = new List<strLineOut>();
        private int iOrder = 0;

        public ClsLinesOutputRapper() 
        {
            try
            {
                lstLines = new List<strLineOut>();
                iOrder = 0;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        ~ClsLinesOutputRapper() 
        {
            try
            {
                lstLines = null;
                iOrder = 0;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void Add(ref string sLine)
        {
            try
            {
                strLineOut objLineOut = new strLineOut();

                iOrder++;
                objLineOut.iOrder = iOrder;
                objLineOut.sLine = sLine;

                lstLines.Add(objLineOut);
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void Add(string sLine)
        {
            try
            {
                strLineOut objLineOut = new strLineOut();

                iOrder++;
                objLineOut.iOrder = iOrder;
                objLineOut.sLine = sLine;

                lstLines.Add(objLineOut);
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public int Count
        {
            get 
            {
                try 
                {
                    int iResult = lstLines.Count;

                    return iResult;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = string.Empty;

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return 0;
                }
            }
        }

        public List<strLineOut> lines
        {
            get 
            {
                try
                {
                    List<strLineOut> lstResult = new List<strLineOut>();

                    lstResult = lstLines.OrderBy(x => x.iOrder).ToList<strLineOut>();

                    return lstResult;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = string.Empty;

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return new List<strLineOut>();
                }
            }
        }

        public void incrementOrderAboveX(int iOrder)
        {
            try
            {
                for (int iIndex = 0; iIndex < lstLines.Count;iIndex++)
                {
                    if (lstLines[iIndex].iOrder > iOrder)
                    {
                        strLineOut objLines = lstLines[iIndex];

                        objLines.iOrder++;

                        lstLines[iIndex] = objLines;
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = string.Empty;

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        public void removeReturnChar()
        {
            try
            {
                Predicate<strLineOut> predReturnChar = x => x.sLine.Contains('\n') || x.sLine.Contains('\r');
                
                int iIndex = lstLines.FindIndex(predReturnChar);

                while (iIndex != -1)
                {
                    if (lstLines[iIndex].sLine.Length == 1)
                    {
                        strLineOut objTemp = lstLines[iIndex];
                        objTemp.sLine = "";
                        lstLines[iIndex] = objTemp;

                        incrementOrderAboveX(iIndex);

                        lstLines.Insert(iIndex, objTemp);
                    }
                    else
                    {
                        strLineOut objTemp = lstLines[iIndex];

                        incrementOrderAboveX(objTemp.iOrder);

                        int iPos = objTemp.sLine.ToList().FindIndex(y => y == '\n' || y == '\r');

                        strLineOut objBefore = new strLineOut();
                        strLineOut objAfter = new strLineOut();

                        objBefore.sLine = ClsMiscString.Left(ref objTemp.sLine, iPos);
                        objBefore.iOrder = objTemp.iOrder;

                        objAfter.sLine = ClsMiscString.Right(ref objTemp.sLine, lstLines[iIndex].sLine.Length - iPos - 1);
                        objAfter.iOrder = objTemp.iOrder + 1;

                        //string sBefore = ClsMiscString.Left(ref objTemp.sLine, iPos);
                        //string sAfter = ClsMiscString.Right(ref objTemp.sLine, lstLines[iIndex].sLine.Length - iPos - 1);

                        lstLines.Insert(iIndex, objBefore);
                        lstLines[iIndex + 1] = objAfter;
                    }

                    iIndex = lstLines.FindIndex(predReturnChar);
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
