using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Globalization;
using System.Data;
using System.IO;
using System.Reflection;

namespace KodeMagd.Settings
{
    class ClsReadConnectionString
    {
        private const string csNodeNameConnectionStrings = "ConnectionStrings";
        private const string csNodeNameConnectionString = "ConnectionString";
        private const string csNodeNameBackend = "Backend";
        private const string csNodeNameType = "Type";
        private const string csNodeNameNotes = "Notes";
        private const string csNodeNameElements = "Elements";
        private const string csNodeNameElement = "Element";
        private const string csNodeNameElementName = "Name";
        private const string csNodeNameElementValue = "Value";

        //private const string csPath = "connectionStrings.xml"; <-- <-- this one was working before
        //private const string csPath = "Settings.connectionStrings.xml";
        private const string csPath = "KodeMagd.Settings.connectionStrings.xml";
        //private const string csPath = "KodeMagd.connectionStrings.xml";
        //http://www.codeproject.com/Articles/9159/Working-with-Embedded-Data


        public struct strConnStrElement
        {
            public string sName;
            public string sValue;
        }

        public struct strConnectionString
        {
            public string sBackend;
            public string sType;
            public string sNotes;
            public List<strConnStrElement> lstElements;// = new List<strConnStrElement>();
        }

        public List<strConnectionString> lstConnectionStrings = new List<strConnectionString>();

        //Reading Connection String file and returning the right data stuctures
        public ClsReadConnectionString()
        {
            try
            {
                //read();
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

        ~ClsReadConnectionString()
        {
            try
            {
                lstConnectionStrings = null;
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

        public void read() 
        {
            XmlTextReader reader;

            try
            {
                lstConnectionStrings = new List<strConnectionString>();

                //string sPath = Assembly.GetExecutingAssembly().Location + "." + csPath;

                System.IO.Stream strm = Assembly.GetExecutingAssembly().GetManifestResourceStream(csPath);

                XmlDocument doc = new XmlDocument();
                //doc.Load(sPath);
                doc.Load(strm);

                XmlNodeList lstNodesConnectionStrings = doc.GetElementsByTagName(csNodeNameConnectionString);

                foreach (XmlNode nodConnectionString in lstNodesConnectionStrings)
                {
                    strConnectionString objConnStr = new strConnectionString();

                    objConnStr.sType = string.Empty;
                    objConnStr.sNotes = string.Empty;
                    objConnStr.sBackend = string.Empty;
                    objConnStr.lstElements = new List<strConnStrElement>();

                    XmlElement elmtConnectionString = (XmlElement)nodConnectionString;

                    foreach (XmlNode nodChild in nodConnectionString.ChildNodes)
                    {
                        
                        switch (nodChild.Name) 
                        {
                            case csNodeNameType:
                                objConnStr.sType = nodChild.InnerText;
                                break;
                            case csNodeNameNotes:
                                objConnStr.sNotes = nodChild.InnerText;
                                break;
                            case csNodeNameBackend:
                                objConnStr.sBackend= nodChild.InnerText;
                                break;
                            case csNodeNameElements:
                                foreach (XmlNode nodElement in nodChild.ChildNodes)
                                {
                                    strConnStrElement objElement = new strConnStrElement();

                                    foreach(XmlNode nodElementChild in nodElement.ChildNodes)
                                    {
                                        switch (nodElementChild.Name)
                                        {
                                            case csNodeNameElementName:
                                                objElement.sName = nodElementChild.InnerText;
                                                break;
                                            case csNodeNameElementValue:
                                                objElement.sValue = nodElementChild.InnerText;
                                                break;
                                        }
                                    }

                                    objConnStr.lstElements.Add(objElement);
                                }
                                break;
                            default:
                                //Don't know
                                break;
                        }
                    }
                    lstConnectionStrings.Add(objConnStr);
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

        public void fillLstWithDefaultConnectionStrings() 
        {
            try
            {
                //List<string> lstTemp = new List<string>();
                //lstTemp.Add("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Test Stuff\\Database2.accdb;Mode=Share Deny None;Jet OLEDB:System database=C:\\Users\\Matthew\\AppData\\Roaming\\Microsoft\\Access\\System.mdw;Jet OLEDB:Registry Path=Software\\Microsoft\\Office\\14.0\\Access\\Access Connectivity Engine;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=True;Jet OLEDB:Bypass UserInfo Validation=False");

                lstConnectionStrings = new List<strConnectionString>();

                strConnectionString objTemp = new strConnectionString();
                strConnStrElement objElement = new strConnStrElement();

                /* One */
                objTemp.sBackend = "MS Access";
                objTemp.sType = "Standard security";
                objTemp.sNotes = "";
                objTemp.lstElements = new List<strConnStrElement>();

                objElement.sName = "Provider";
                objElement.sValue = "Microsoft.Jet.OLEDB.4.0";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Data Source";
                objElement.sValue = "C:\\*.mdb";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "User Id";
                objElement.sValue = "admin";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Password";
                objElement.sValue = "";
                objTemp.lstElements.Add(objElement);

                lstConnectionStrings.Add(objTemp);

                /* Two */
                objTemp.sBackend = "MS Access";
                objTemp.sType = "With database password";
                objTemp.sNotes = "This is the connection string to use when you have an access database protected with a password using the Set Database Password function in Access.";
                objTemp.lstElements = new List<strConnStrElement>();

                objElement.sName = "Provider";
                objElement.sValue = "Microsoft.Jet.OLEDB.4.0";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Data Source";
                objElement.sValue = "C:\\*.mdb";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Database Password";
                objElement.sValue = "MyDbPassword";
                objTemp.lstElements.Add(objElement);

                lstConnectionStrings.Add(objTemp);

                /* My string */
                objTemp.sBackend = "MS Access";
                objTemp.sType = "My string";
                objTemp.sNotes = "Created by me...";
                objTemp.lstElements = new List<strConnStrElement>();

                objElement.sName = "Provider";
                objElement.sValue = "Microsoft.ACE.OLEDB.12.0";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Data Source";
                objElement.sValue = "C:\\*.accdb";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Mode";
                objElement.sValue = "Share Deny None";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:System database";
                objElement.sValue = "C:\\Users\\Matthew\\AppData\\Roaming\\Microsoft\\Access\\System.mdw";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Registry Path";
                objElement.sValue = "Software\\Microsoft\\Office\\14.0\\Access\\Access Connectivity Engine";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Engine Type";
                objElement.sValue = "6";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Database Locking Mode";
                objElement.sValue = "1";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Global Partial Bulk Ops";
                objElement.sValue = "2";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Global Bulk Transactions";
                objElement.sValue = "1";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Create System Database";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Encrypt Database";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Don't Copy Locale on Compact";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Compact Without Replica Repair";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:SFP";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Support Complex Data";
                objElement.sValue = "True";
                objTemp.lstElements.Add(objElement);

                objElement.sName = "Jet OLEDB:Bypass UserInfo Validation";
                objElement.sValue = "False";
                objTemp.lstElements.Add(objElement);

                lstConnectionStrings.Add(objTemp);
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

        public void write() 
        {
            XmlWriter writer;
            
            try
            {

                
                string sPath = Assembly.GetExecutingAssembly().Location + "." + csPath;
                //System.IO.Stream strm = Assembly.GetExecutingAssembly().GetManifestResourceStream(csPath);
                //DataSet ds;
                
                //ds.

                //http://msdn.microsoft.com/en-us/library/system.io.stream%28v=vs.71%29.aspx

                /*
                 * fixed this here
                */
                writer = XmlWriter.Create(sPath);
                //writer = XmlWriter.Create(strm);
                writer.WriteStartDocument();
                writer.WriteStartElement(csNodeNameConnectionStrings);

                foreach (strConnectionString objConnStr in lstConnectionStrings)
                {
                    writer.WriteStartElement(csNodeNameConnectionString);
                    writer.WriteElementString(csNodeNameBackend, objConnStr.sBackend);
                    writer.WriteElementString(csNodeNameType, objConnStr.sType);
                    writer.WriteElementString(csNodeNameNotes, objConnStr.sNotes);
                    writer.WriteStartElement(csNodeNameElements);

                    foreach (strConnStrElement objElement in objConnStr.lstElements)
                    {
                        writer.WriteStartElement(csNodeNameElement);
                        writer.WriteElementString(csNodeNameElementName, objElement.sName);
                        writer.WriteElementString(csNodeNameElementValue, objElement.sValue);
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();

                writer = null;
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
                writer = null;
            }
        }

        public void add(string sType, string sBackend, string sNotes, string sConnectionString) 
        {
            try
            {
                strConnectionString objConnStr = new strConnectionString();

                objConnStr.sBackend = sBackend;
                objConnStr.sType = sType;
                objConnStr.sNotes = sNotes;
                objConnStr.lstElements = new List<strConnStrElement>();

                List<string> lstElementsString = sConnectionString.Split(';').ToList();

                foreach (string sElement in lstElementsString) 
                {
                    strConnStrElement objElement = new strConnStrElement();

                    List<string> lstElement = sElement.Split('=').ToList();

                    if (lstElement.Count == 2) 
                    {
                        objElement.sName = lstElement[0].ToLower().Trim();
                        objElement.sValue = lstElement[1].ToLower().Trim();

                        objConnStr.lstElements.Add(objElement);
                    }
                }

                bool bIsFound = false;

                foreach (strConnectionString objTemp in lstConnectionStrings)
                {
                    if (isMatch(objTemp.lstElements, objConnStr.lstElements)) 
                    { bIsFound = true; }
                }

                if (!bIsFound)
                { lstConnectionStrings.Add(objConnStr); }
                else
                { MessageBox.Show("Connection String already in list", "Not added",  MessageBoxButtons.OK, MessageBoxIcon.Exclamation); }
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

        private bool isMatch(List<strConnStrElement> lstElements1, List<strConnStrElement> lstElements2)
        {
            try
            {
                bool bResult;
                bool bAllInOneAreInTwo = true;
                bool bAllInTwoAreInOne = true;

                foreach (strConnStrElement objElement1 in lstElements1)
                {
                    bool bIsFound = false;

                    foreach (strConnStrElement objElement2 in lstElements2)
                    {
                        if (objElement1.sName.ToLower().Trim() == objElement2.sName.ToLower().Trim())
                        {
                            int iOne;

                            if (int.TryParse(objElement1.sValue, out iOne))
                            {
                                int iTwo;

                                if (int.TryParse(objElement2.sValue, out iTwo))
                                {
                                    if (iOne == iTwo)
                                    { bIsFound = true; }
                                }
                            }

                            if (!bIsFound)
                            {
                                float fOne;

                                if (float.TryParse(objElement1.sValue, out fOne))
                                {
                                    float fTwo;

                                    if (float.TryParse(objElement2.sValue, out fTwo))
                                    {
                                        if (fOne == fTwo)
                                        { bIsFound = true; }
                                    }
                                }
                            }

                            if (!bIsFound)
                            {
                                if (objElement1.sValue.Trim().ToLower() == objElement2.sValue.Trim().ToLower())
                                { bIsFound = true; }
                            }
                        }
                    }

                    if (!bIsFound)
                    { bAllInOneAreInTwo = false; }
                }

                foreach (strConnStrElement objElement2 in lstElements2)
                {
                    bool bIsFound = false;
                    foreach (strConnStrElement objElement1 in lstElements1)
                    {
                        if (objElement1.sName.ToLower().Trim() == objElement2.sName.ToLower().Trim())
                        {
                            int iOne;

                            if (int.TryParse(objElement1.sValue, out iOne))
                            {
                                int iTwo;

                                if (int.TryParse(objElement2.sValue, out iTwo))
                                {
                                    if (iOne == iTwo)
                                    { bIsFound = true; }
                                }
                            }

                            if (!bIsFound)
                            {
                                float fOne;

                                if (float.TryParse(objElement1.sValue, out fOne))
                                {
                                    float fTwo;

                                    if (float.TryParse(objElement2.sValue, out fTwo))
                                    {
                                        if (fOne == fTwo)
                                        { bIsFound = true; }
                                    }
                                }
                            }

                            if (!bIsFound)
                            {
                                if (objElement1.sValue.Trim().ToLower() == objElement2.sValue.Trim().ToLower())
                                { bIsFound = true; }
                            }
                        }
                    }

                    if (!bIsFound)
                    { bAllInTwoAreInOne = false; }
                }

                if (bAllInOneAreInTwo & bAllInTwoAreInOne)
                { bResult = true; }
                else
                { bResult = false; }

                return bResult;
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

                return false;
            }
        }

        public List<string> ListBackend()
        {
            try
            {
                List<string> lstResult = new List<string>();
                TextInfo tiTemp = new CultureInfo("").TextInfo;

                foreach (strConnectionString objConnStr in lstConnectionStrings)
                {
                    string sTemp = tiTemp.ToTitleCase(objConnStr.sBackend.Trim()).Trim();

                    lstResult.Add(sTemp);
                }

                lstResult = lstResult.Distinct().ToList();
                lstResult.Sort();

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

                return new List<string>(); ;
            }
        }

        public List<string> ListType()
        {
            try
            {
                List<string> lstResult = new List<string>();
                TextInfo tiTemp = new CultureInfo("").TextInfo;

                foreach (strConnectionString objConnStr in lstConnectionStrings)
                { lstResult.Add(tiTemp.ToTitleCase(objConnStr.sType.Trim()).Trim()); }

                lstResult = lstResult.Distinct().ToList();
                lstResult.Sort();

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

                return new List<string>(); ;
            }
        }

        public List<string> ListType(string sBackend)
        {
            try
            {
                List<string> lstResult = new List<string>();
                TextInfo tiTemp = new CultureInfo("").TextInfo;

                foreach (strConnectionString objConnStr in lstConnectionStrings)
                {
                    if (sBackend.ToLower().Trim() == objConnStr.sBackend.ToLower().Trim())
                    {
                        string sTemp = tiTemp.ToTitleCase(objConnStr.sType.Trim()).Trim();

                        lstResult.Add(sTemp);
                    }
                }

                lstResult = lstResult.Distinct().ToList();
                lstResult.Sort();

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

                return new List<string>(); 
            }
        }

        public strConnectionString getConnectionString(string sBackEnd, string sType) 
        {
            try
            {
                strConnectionString objResult = new strConnectionString();
                //int iPos = lstConnectionStrings.FindIndex(x => x.sBackend.ToLower().Trim() == sBackEnd.ToLower().Trim());
                
                if (lstConnectionStrings.Exists(x => x.sBackend.ToLower().Trim() == sBackEnd.ToLower().Trim() && x.sType.ToLower().Trim() == sType.ToLower().Trim()))
                {
                    int iPos = lstConnectionStrings.FindIndex(x => x.sBackend.ToLower().Trim() == sBackEnd.ToLower().Trim() && x.sType.ToLower().Trim() == sType.ToLower().Trim());

                    objResult = lstConnectionStrings[iPos];
                }
                else
                {
                    objResult = new strConnectionString();
                    objResult.sType = "";
                    objResult.sBackend = "";
                    objResult.sNotes = "";
                    objResult.lstElements = new List<strConnStrElement>();
                    objResult.lstElements.Clear();
                }

                return objResult;
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

                return new strConnectionString();
            }
        }
    }
}