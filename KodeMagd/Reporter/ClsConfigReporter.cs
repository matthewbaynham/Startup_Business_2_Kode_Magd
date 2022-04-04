using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Misc;

namespace KodeMagd.Reporter
{
    public class ClsConfigReporter
    {
        /*****************************************************
         *                                                   *
         *   When the user has pressed loads of buttons      *
         *   to do something then hits Generate he will      * 
         *   forget what he has pressed.                     *
         *   This Class takes reports to him what he has     *
         *   entered, either via MessageBox or HTML report   *
         *   displayed on the screen.                        *
         *                                                   *
         *****************************************************/

        public struct strLine
        {
            public int iOrder;
            public string sLine;
        }

        public enum enumFormatDetails
        {
            eFmt_Italic,
            eFmt_Bold,
            eFmt_Oblique,
            eFmt_VeryLarge,
            eFmt_Large,
            eFmt_Small,
            eFmt_VerySmall,

            eFmt_Maroon,
            eFmt_Red,
            eFmt_Orange,
            eFmt_Yellow,
            eFmt_Olive,

            eFmt_Purple,
            eFmt_Fuchsia,
            eFmt_White,
            eFmt_Lime,
            eFmt_Green,

            eFmt_Navy,
            eFmt_Blue,
            eFmt_Aqua,
            eFmt_Teal,

            eFmt_Black,
            eFmt_Silver,
            eFmt_Gray
        }

        public struct strTableCell
        {
            public int iOrder; //horizontally in the row
            public int iHtmlId;
            public bool bPropHtml;
            public string sText;
            public string sHiddenText;
            public List<enumFormatDetails> lstFormatDetails;
        }

        public struct strTableRow
        {
            public int iOrder;
            public int iHtmlId;
            public bool bIsHeader;
            public List<strTableCell> lstText;
        }

        public struct strTable
        {
            public int iOrder;
            public int iHtmlId;
            public int columns;
            public string sCaption;
            public List<strTableRow> lstText;
            public List<int> lstColumnSize;
        }

        public struct strHtmlId
        {
            public int iHtmlId;
            public string sType;
        }

        public struct strCssStyle
        {
            public string sName;
            public string sValue;
        }

        public struct strCss
        {
            public string sName;
            public List<strCssStyle> lstCssStyles;
        }

        public List<strHtmlId> lstHtmlId = new List<strHtmlId>();
        public ClsSettings cSettings = new ClsSettings();
        public List<strTable> lstTables = new List<strTable>();
        public List<strCss> lstCssExtra = new List<strCss>();

        public void addExtraCss(strCss objCss)
        {
            try
            {
                Predicate<strCss> prepCss = x => x.sName.Trim().ToLower() == objCss.sName.Trim().ToLower();

                if (lstCssExtra.Exists(prepCss))
                { lstCssExtra.RemoveAll(prepCss); }
                lstCssExtra.Add(objCss);
                
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

        public ClsConfigReporter() 
        {
            try
            {
                lstTables = new List<strTable>();
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

        ~ClsConfigReporter() 
        {
            try
            {
                lstTables = null;
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

        public void clear(ref List<strLine> lstHtml)
        {
            try
            {
                lstHtml = new List<strLine>();
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
        public void addHeader(ref List<strLine> lstHtml)
        {
            try
            {

                int iIndent = 0;

                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<!DOCTYPE html>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<html>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<head>");
                iIndent++;
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<title>" + ClsCodeEditorGUI.csCommandBarName + " - " + ClsMisc.ActiveWorkBook().Name + "</title>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<style type=\"text/css\">");
                iIndent++;
                //addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".bluetitle { color: blue; }");
                //addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".cell_format { color: black; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".process_details { font-family:\"Arial\"; font-size:20px; border: 1px solid #EFCFCF; text-align:left; Width:67%; margin-top:3px; margin-bottom:3px; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".table_caption { font-family:\"Arial\"; font-size:16px; color: black; margin-top:20px; margin-bottom:0px; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".table_format { font-family:\"Arial\"; font-size:16px; border: 1px solid #EFEFEF; text-align:left; width:100%; margin-top:3px; margin-bottom:30px; }");
                //addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".row_format { font-family:\"Arial\"; font-size:16px; border: 1px solid #EFEFEF; text-align:left; }");
                //addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".cellformatone { font-family:\"Arial\"; font-size:16px; border: 1px solid #EFEFEF; text-align:left;  margin-top:3px; margin-bottom:3px; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".row_format { font-family:\"Arial\"; font-size:16px; border: 1px solid #EFEFEF; text-align:left; font-weight:normal;  margin-top:3px; margin-bottom:3px; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".title_row_format { font-family:\"Arial\"; font-size:18px; border: 1px solid #EFEFEF; text-align:left; font-weight:bold; margin-top:3px; margin-bottom:3px; }");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + ".hiddenText { font-family:\"Arial\"; font-size:14px; color: gray; border: 1px solid #EFEFEF; text-align:left; padding-left:15px; font-style:italic; }");

                if (lstCssExtra.FindAll(x => x.lstCssStyles.Count != 0).Count != 0)
                {
                    foreach (strCss objCss in lstCssExtra.FindAll(x => x.lstCssStyles.Count != 0))
                    {
                        string sCssLine = "." + objCss.sName + " {";
                        foreach (strCssStyle objCssStyle in objCss.lstCssStyles)
                        { sCssLine += objCssStyle.sName + ":" + objCssStyle.sValue + ";"; }
                        sCssLine += "}";

                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + sCssLine);
                    }
                }

                iIndent--;
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "</style>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<meta http-equiv='X-UA-Compatible' content='IE=9' >");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<meta name=\"Application\" content=\"" + prepHtmlText(ClsCodeEditorGUI.csCommandBarName) + "\">");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<meta name=\"Excel File\" content=\"" + prepHtmlText(ClsMisc.ActiveWorkBook().Name) + "\">");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<meta name=\"User\" content=\"" + prepHtmlText(Environment.UserName) + "\">");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<meta name=\"Created at\" content=\"" + prepHtmlText(DateTime.Now.ToString()) + "\">");
                iIndent--;
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "</head>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<body>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<h1 class=\"process_details\">" + prepHtmlText(ClsCodeEditorGUI.csCommandBarName) + "</h1>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"process_details\">Excel Doc Name: " + prepHtmlText(ClsMisc.ActiveWorkBook().Name) + "</div>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"process_details\">Excel Doc Path: " + prepHtmlText(ClsMisc.ActiveWorkBook().Path) + "</div>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"process_details\">User: " + prepHtmlText(Environment.UserName) + "</div>");
                addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"process_details\">" + prepHtmlText(DateTime.Now.ToString()) + "</div>");

                if (lstHtmlId.Count > 0)
                {
                    addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<script type=\"text/javascript\">");
                    iIndent++;
                    addNextLine(ref lstHtml, cSettings.Indent(iIndent));

                    foreach (strHtmlId objHtmlId in lstHtmlId)
                    {
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "function switchMenu(obj) {");
                        iIndent++;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "var el = document.getElementById(obj);");
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "if ( el.style.display != \"none\" ) {");
                        iIndent++;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "el.style.display = \"none\";");
                        iIndent--;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "} else {");
                        iIndent++;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "el.style.display = \"\";");
                        iIndent--;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "}");
                        iIndent--;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "}");

                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "if (document.getElementById(\"" + prepHtmlText(getHtmlIdText(objHtmlId.iHtmlId)) + "\") != null) {");
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "document.getElementById(\"" + prepHtmlText(getHtmlIdText(objHtmlId.iHtmlId)) + "\").style.display = \"none\"; }");
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent));
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent));
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent));
                    }

                    iIndent--;
                    addNextLine(ref lstHtml, "</script>");
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

        public void addFooter(ref List<strLine> lstHtml)
        {
            try
            {
                addNextLine(ref lstHtml, "</body>");
                addNextLine(ref lstHtml, "</html>");
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

        private void addTitle(ref List<strLine> lstHtml)
        {
            try
            {
                addNextLine(ref lstHtml, "<h1 class=\"bluetitle\">" + ClsDefaults.formTitle + "</h1>");
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

        public void addNextLine(ref List<strLine> lstHtml, string sLine)
        {
            try
            {
                int iMaxLineNo;

                if (lstHtml.Count == 0)
                { iMaxLineNo = 0; }
                else
                { iMaxLineNo = lstHtml.Max(x => x.iOrder); }

                strLine objLine = new strLine();
                objLine.iOrder = iMaxLineNo + 1;
                objLine.sLine = sLine;
                lstHtml.Add(objLine);
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

        public string getHtml()
        {
            try
            {
                List<strLine> lstHtml = new List<strLine>();
                string sDoc = "";
                int iIndent = 0;

                buildListHtmlId();

                addHeader(ref lstHtml);

                foreach (strTable objTable in lstTables.OrderBy(x => x.iOrder))
                {
                    if (objTable.sCaption.Trim() != "")
                    { addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"table_caption\">" + prepHtmlText(objTable.sCaption) + "</div>"); }

                    if (objTable.columns == 1)
                    {
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<div class=\"table_format\">");

                        foreach (strTableRow objTableRow in objTable.lstText.OrderBy(x => x.iOrder))
                        {
                            iIndent++;
                            string sRowStyle = "";
                            if (objTableRow.bIsHeader == true)
                            { sRowStyle = "title_row_format"; }
                            else
                            { sRowStyle = "row_format"; }

                            if (objTableRow.lstText.Count > 0)
                            {
                                strTableCell objTableCell = objTableRow.lstText[0];
                                {
                                    iIndent++;
                                    string sLine = cSettings.Indent(iIndent);

                                    if (objTableCell.sHiddenText.Trim() == "")
                                    {
                                        if (objTableCell.sText.Length == objTableCell.sText.Trim().TrimStart().Length)
                                        { sLine += "<div class=\"" + sRowStyle + "\" " + extraFormatting(objTableCell.lstFormatDetails) + " >"; }
                                        else
                                        {
                                            int iCellIndent = objTableCell.sText.Length - objTableCell.sText.Trim().TrimStart().Length;
                                            sLine += "<div class=\"" + sRowStyle + "\" style=\" padding-left:" + (10 * iCellIndent).ToString() + "px;\" " + extraFormatting(objTableCell.lstFormatDetails) + " >";
                                        }

                                        if (objTableCell.bPropHtml)
                                        { sLine += prepHtmlText(objTableCell.sText); }
                                        else
                                        { sLine += objTableCell.sText; }
                                        sLine += "<br>";
                                        sLine += "</div>";
                                    }
                                    else
                                    {
                                        sLine += "<div class=\"cell_format\">";

                                        if (objTableCell.sText.Length == objTableCell.sText.Trim().TrimStart().Length)
                                        { sLine += "<a href=\"#\" onclick=\"switchMenu('" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "'); \" " + extraFormatting(objTableCell.lstFormatDetails) + " >" + prepHtmlText(objTableCell.sText.Trim()) + "</a>"; }
                                        else
                                        {
                                            int iCellIndent = objTableCell.sText.Length - objTableCell.sText.Trim().TrimStart().Length;
                                            sLine += "<a href=\"#\" onclick=\"switchMenu('" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "');\" style=\"padding-left:" + (10 * iCellIndent).ToString() + "px;\" " + extraFormatting(objTableCell.lstFormatDetails) + ">" + prepHtmlText(objTableCell.sText.Trim()) + "</a>";
                                        }

                                        sLine += "<div id=\"" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "\" class=\"hiddenText\" style=\"display:none\">";

                                        if (objTableCell.bPropHtml)
                                        { sLine += prepHtmlText(objTableCell.sHiddenText); }
                                        else
                                        { sLine += objTableCell.sHiddenText; }
                                        
                                        sLine += "</div>";
                                        sLine += "</div>";
                                    }

                                    addNextLine(ref lstHtml, sLine);
                                    iIndent--;
                                }
                            }
                            iIndent--;
                        }

                        addNextLine(ref lstHtml, "</div>");
                    }
                    else
                    {
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<table class=\"table_format\">");
                        iIndent++;

                        int iTotalWidth = objTable.lstColumnSize.Sum();

                        foreach (int iWidth in objTable.lstColumnSize)
                        {
                            double dSize = 100 * (double)iWidth / (double)iTotalWidth;
                            addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<col width=\"" + dSize.ToString() + "%\">");
                        }

                        foreach (strTableRow objTableRow in objTable.lstText.OrderBy(x => x.iOrder).ToList())
                        {
                            string sRowStyle = "";
                            if (objTableRow.bIsHeader == true)
                            { sRowStyle = "title_row_format"; }
                            else
                            { sRowStyle = "row_format"; }

                            string sLine = cSettings.Indent(iIndent);

                            addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "<tr class=\"" + sRowStyle + "\">");
                            iIndent++;

                            foreach (strTableCell objTableCell in objTableRow.lstText.OrderBy(x => x.iOrder))
                            {
                                sLine = cSettings.Indent(iIndent);

                                if (objTableCell.sHiddenText.Trim() == "")
                                {
                                    if (objTableCell.sText.Length == objTableCell.sText.Trim().TrimStart().Length)
                                    { sLine += "<td class=\"" + sRowStyle + "\" " + extraFormatting(objTableCell.lstFormatDetails) + " >"; }
                                    else
                                    {
                                        int iCellIndent = objTableCell.sText.Length - objTableCell.sText.Trim().TrimStart().Length;
                                        sLine += "<td class=\"" + sRowStyle + "\" style=\"padding-left:" + (10 * iCellIndent).ToString() + "px;\" " + extraFormatting(objTableCell.lstFormatDetails) + " >";
                                    }
                                    sLine += prepHtmlText(objTableCell.sText);
                                    sLine += "</td>";
                                }
                                else
                                {
                                    sLine += "<td class=\"" + sRowStyle + "\">";

                                    if (objTableCell.sText.Length == objTableCell.sText.Trim().TrimStart().Length)
                                    { sLine += "<a href=\"#\" onclick=\"switchMenu('" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "');\" " + extraFormatting(objTableCell.lstFormatDetails) + " >" + prepHtmlText(objTableCell.sText) + "</a>"; }
                                    else
                                    {
                                        int iCellIndent = objTableCell.sText.Length - objTableCell.sText.Trim().TrimStart().Length;
                                        sLine += "<a href=\"#\" onclick=\"switchMenu('" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "');\" style=\"padding-left:" + (10 * iCellIndent).ToString() + "px;\" " + extraFormatting(objTableCell.lstFormatDetails) + " >" + prepHtmlText(objTableCell.sText) + "</a>";
                                    }

                                    sLine += "<div id=\"" + prepHtmlText(getHtmlIdText(objTableCell.iHtmlId)) + "\" class=\"hiddenText\" style=\"display:none\">";
                                    sLine += prepHtmlText(objTableCell.sHiddenText);
                                    sLine += "</div>";
                                    sLine += "</td>";
                                }

                                addNextLine(ref lstHtml, sLine);
                            }

                            iIndent--;
                            addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "</tr>");
                        }

                        iIndent--;
                        addNextLine(ref lstHtml, cSettings.Indent(iIndent) + "</table>");
                    }
                }

                addFooter(ref lstHtml);

                lstHtml = lstHtml.OrderBy(x => x.iOrder).ToList();

                foreach (strLine objLine in lstHtml)
                { sDoc += objLine.sLine + "\n"; }

                return sDoc;
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

                return sMessage;
            }
        }

        public int addParagraph()
        {
            try
            {
                int iId = 0;

                return iId;
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

                return 0;
            }
        }

        public void TableAddNew(out int iTableId, int iNoOfColumns)
        {
            try
            {
                TableAddNew(out iTableId, iNoOfColumns, "");
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

                iTableId = 0;
            }
        }

        public void TableAddNew(out int iTableId, int iNoOfColumns, string sCaption)
        {
            try
            {
                List<int> lstColSize = new List<int>();

                for (int iCounter = 0; iCounter < iNoOfColumns; iCounter++)
                { lstColSize.Add(1); }

                TableAddNew(out iTableId, lstColSize, sCaption);
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

                iTableId = 0;
            }
        }

        public void TableAddNew(out int iTableId, List<int> lstColSize)
        {
            try
            {
                TableAddNew(out iTableId, lstColSize, "");
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

                iTableId = 0;
            }
        }

        public void TableAddNew(out int iTableId, List<int> lstColSize, string sCaption)
        {
            try
            {
                int iId = 0;

                if (lstTables.Count == 0)
                { iId = 0; }
                else
                { iId = lstTables.Max(x => x.iOrder); }

                iId++;

                strTable objTable = new strTable();

                objTable.iOrder = iId;
                objTable.iHtmlId = 0;
                objTable.sCaption = sCaption;
                objTable.lstColumnSize = new List<int>();

                objTable.lstColumnSize = lstColSize;

                objTable.lstText = new List<strTableRow>();
                objTable.columns = lstColSize.Count;

                lstTables.Add(objTable);

                iTableId = iId;
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

                iTableId = 0;
            }
        }

        public void TableAddNewRow(int iTableId, out int iRowId)
        {
            try
            {
                TableAddNewRow(iTableId, out iRowId, false);
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

                iRowId = 0;
            }
        }

        public void TableAddNewRow(int iTableId, out int iRowId, bool bIsHeader)
        {
            try
            {
                if (lstTables.Exists(x => x.iOrder == iTableId))
                {
                    int iTblId = lstTables.FindIndex(x => x.iOrder == iTableId);

                    strTableRow objRow = new strTableRow();

                    if (lstTables[iTblId].lstText.Count == 0)
                    {
                        objRow.iOrder = 1;
                    }
                    else
                    {
                        int iId = lstTables[iTblId].lstText.Max(x => x.iOrder) + 1;

                        objRow.iOrder = iId;
                    }

                    iRowId = objRow.iOrder;

                    objRow.bIsHeader = bIsHeader;
                    objRow.lstText = new List<strTableCell>();

                    lstTables[iTblId].lstText.Add(objRow);
                }
                else
                { iRowId = 0; }
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

                iRowId = 0;
            }
        }


        public void TableAddNewCell(int iTableId, int iRowId, strTableCell objCell)
        {
            try
            {
                //bool bIsOk = true;  if this bit goes wrong don't bother with anything

                if (objCell.sText == null)
                { objCell.sText = ""; }

                if (objCell.sHiddenText == null)
                { objCell.sHiddenText = ""; }

                if (lstTables.Exists(x => x.iOrder == iTableId))
                {
                    int iIdTable = lstTables.FindIndex(x => x.iOrder == iTableId);

                    if (lstTables[iIdTable].lstText.Exists(x => x.iOrder == iRowId))
                    {
                        int iIdRow = lstTables[iIdTable].lstText.FindIndex(x => x.iOrder == iRowId);

                        int iCellId = 0;

                        if (lstTables[iIdTable].lstText[iIdRow].lstText.Count == 0)
                        { iCellId = 1; }
                        else
                        { iCellId = lstTables[iIdTable].lstText[iIdRow].lstText.Max(x => x.iOrder) + 1; }

                        objCell.iOrder = iCellId;
                        if (objCell.sHiddenText.Trim() == "")
                        { objCell.iHtmlId = 0; }
                        else
                        { objCell.iHtmlId = getNextHtmlID(); }

                        lstTables[iIdTable].lstText[iIdRow].lstText.Add(objCell);
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

                iRowId = 0;
            }
        }

        private int getNextHtmlID()
        {
            try
            {
                int iMaxHtmlId = 0;

                if (lstTables.Count > 0)
                {
                    if (iMaxHtmlId < lstTables.Max(x => x.iHtmlId))
                    { iMaxHtmlId = lstTables.Max(x => x.iHtmlId); }

                    foreach (strTable objTable in lstTables)
                    {
                        if (objTable.lstText.Count > 0)
                        {
                            if (iMaxHtmlId < objTable.lstText.Max(x => x.iHtmlId))
                            { iMaxHtmlId = objTable.lstText.Max(x => x.iHtmlId); }

                            foreach (strTableRow objRow in objTable.lstText)
                            {
                                if (objRow.lstText.Count > 0)
                                {
                                    if (iMaxHtmlId < objRow.lstText.Max(x => x.iHtmlId))
                                    { iMaxHtmlId = objRow.lstText.Max(x => x.iHtmlId); }
                                }
                            }
                        }
                    }
                }

                return iMaxHtmlId + 1;
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

                return 0;
            }
        }

        public string getHtmlIdText(int iID)
        {
            try
            {
                return "ID" + ClsMiscString.Right("0000000000" + iID.ToString(), 10);
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

        private void buildListHtmlId()
        {
            try
            {
                lstHtmlId.Clear();

                foreach (strTable objTable in lstTables)
                {
                    if (objTable.iHtmlId > 0)
                    {
                        strHtmlId objHtmlId = new strHtmlId();

                        objHtmlId.iHtmlId = objTable.iHtmlId;
                        objHtmlId.sType = "Table";

                        lstHtmlId.Add(objHtmlId);
                    }

                    foreach (strTableRow objRow in objTable.lstText)
                    {
                        if (objRow.iHtmlId > 0)
                        {
                            strHtmlId objHtmlId = new strHtmlId();

                            objHtmlId.iHtmlId = objRow.iHtmlId;
                            objHtmlId.sType = "Row";

                            lstHtmlId.Add(objHtmlId);
                        }

                        foreach (strTableCell objCell in objRow.lstText)
                        {
                            if (objCell.iHtmlId > 0)
                            {
                                strHtmlId objHtmlId = new strHtmlId();

                                objHtmlId.iHtmlId = objCell.iHtmlId;
                                objHtmlId.sType = "Cell";

                                lstHtmlId.Add(objHtmlId);
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

        public static string prepHtmlText(string sText)
        {
            try
            {
                sText = sText.Replace(">", "&gt;").Replace("<", "&lt;").Replace("\"", "&quot;").Replace("\n\r", "<br>").Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");

                return sText;
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

        private static string extraFormatting(List<enumFormatDetails> lstHtmlFormats)
        {
            try
            {
                string sResult = "";
                string sStyle = "";

                foreach (enumFormatDetails eFormat in lstHtmlFormats.Distinct().OrderBy(x => x.ToString()))
                {
                    switch (eFormat)
                    {
                        case enumFormatDetails.eFmt_Bold:
                            sStyle += "font-weight:bold;";
                            break;
                        case enumFormatDetails.eFmt_Italic:
                            sStyle += "font-style:italic;";
                            break;
                        case enumFormatDetails.eFmt_VerySmall:
                            sStyle += "font-size:8px;";
                            break;
                        case enumFormatDetails.eFmt_Small:
                            sStyle += "font-size:12px;";
                            break;
                        case enumFormatDetails.eFmt_Large:
                            sStyle += "font-size:20px;";
                            break;
                        case enumFormatDetails.eFmt_VeryLarge:
                            sStyle += "font-size:30px;";
                            break;
                        case enumFormatDetails.eFmt_Oblique:
                            sStyle += "font-style:oblique;";
                            break;
                        case enumFormatDetails.eFmt_Maroon:
                            sStyle += "color:maroon;";
                            break;
                        case enumFormatDetails.eFmt_Red:
                            sStyle += "color:red;";
                            break;
                        case enumFormatDetails.eFmt_Orange:
                            sStyle += "color:orange;";
                            break;
                        case enumFormatDetails.eFmt_Yellow:
                            sStyle += "color:yellow;";
                            break;
                        case enumFormatDetails.eFmt_Olive:
                            sStyle += "color:olive;";
                            break;
                        case enumFormatDetails.eFmt_Purple:
                            sStyle += "color:purple;";
                            break;
                        case enumFormatDetails.eFmt_Fuchsia:
                            sStyle += "color:fuchsia;";
                            break;
                        case enumFormatDetails.eFmt_White:
                            sStyle += "color:white;";
                            break;
                        case enumFormatDetails.eFmt_Lime:
                            sStyle += "color:lime;";
                            break;
                        case enumFormatDetails.eFmt_Green:
                            sStyle += "color:green;";
                            break;
                        case enumFormatDetails.eFmt_Navy:
                            sStyle += "color:navy;";
                            break;
                        case enumFormatDetails.eFmt_Blue:
                            sStyle += "color:blue;";
                            break;
                        case enumFormatDetails.eFmt_Aqua:
                            sStyle += "color:aqua;";
                            break;
                        case enumFormatDetails.eFmt_Teal:
                            sStyle += "color:teal;";
                            break;
                        case enumFormatDetails.eFmt_Black:
                            sStyle += "color:black;";
                            break;
                        case enumFormatDetails.eFmt_Silver:
                            sStyle += "color:silvey;";
                            break;
                        case enumFormatDetails.eFmt_Gray:
                            sStyle += "color:gray;";
                            break;
                    }
                }

                if (sStyle != "")
                { sResult += " style=\"" + sStyle + "\""; }

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
    }
}
