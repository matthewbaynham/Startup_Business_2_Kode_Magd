using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace KodeMagd.Misc
{
    class ClsWebsiteDetails
    {
        private const string csDB = "dblic";
        private const string csUser = "wp_user";
        //private const string csAddress = "mysql.webhosting13.1blu.de";
        private const string csAddress = "localhost";
        private const string csPwd = "Me_wp*!123_Stuff";
        private const string csDriver = "{MySQL ODBC 5.2 ANSI Driver}";
        private const string csUrl = "https://www.KodeMagd.de";

        public string url
        {
            get
            {
                try
                {
                    return csUrl;
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

                    return "";
                }
            }
        }

        public string db
        {
            get
            {
                try
                {
                    return csDB;
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

                    return "";
                }
            }
        }

        public string user
        {
            get 
            {
                try 
                {
                    return csUser;
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

                    return "";
                }
            }
        }

        public string address
        {
            get 
            {
                try 
                {
                    return csAddress;
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

                    return "";
                }
            }
        }

        public string password
        {
            get 
            {
                try 
                {
                    return csPwd;
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

                    return "";
                }
            }
        }

        public string driver
        {
            get 
            {
                try 
                {
                    return csDriver;
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

                    return "";
                }
            }
        }

        //public string connectionString 
        //{
        //    get 
        //    {
        //        try 
        //        {
        //            string sConnectionString = "DRIVER=" + csDriver + ";SERVER=" + csAddress + ";DATABASE=" + csDB + ";USER=" + csUser + ";PASSWORD=" + csPwd + ";OPTION=3";
        //            return sConnectionString;
        //        }
        //        catch (Exception ex)
        //        {
        //            MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //            string sMessage = "";

        //            sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //            sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //            sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //            sMessage += ex.Message;

        //            MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

        //            return "";
        //        }
        //    }
        //}

        public string mysql_connectionString
        {
            get
            {
                try
                {
                    const string csDB = "dblic";
                    const string csUser = "wp_user";
                    //const string csAddress = "localhost";
                    //const string csAddress = "127.0.0.1";
                    //const string csPort = "80";


                    const string csAddress = "192.168.2.121";
                    const string csPort = "3306";
                    const string csPwd = "Me_wp*!123_Stuff";

                    //string sResult = "server=" + csAddress + ":" + csPort + ";"
                    //               + "database=" + csDB + ";"
                    //               + "persistsecurityinfo=True;"
                    //               + "user id=" + csUser + ";"
                    //               + "password=" + csPwd; // +";option=3";

                    MySqlConnectionStringBuilder conn_string = new MySqlConnectionStringBuilder();
                    conn_string.Server = csAddress;
                    conn_string.Port = 3306;
                    conn_string.UserID = csUser;
                    conn_string.Password = csPwd;
                    conn_string.Database = csDB;
                    conn_string.PersistSecurityInfo = true;

                    string sResult = conn_string.ToString();

                    //string sResult = "server=" + csAddress + ";"
                    //               + "port=" + csPort + ";"
                    //               + "database=" + csDB + ";"
                    //               + "persistsecurityinfo=True;"
                    //               + "user id=" + csUser + ";"
                    //               + "password=" + csPwd; // +";option=3";

                    //string sResult = "server=" + csAddress + ";"
                    //               + "database=" + csDB + ";"
                    //               + "persistsecurityinfo=True;"
                    //               + "user id=" + csUser + ";"
                    //               + "password=" + csPwd; // +";option=3";

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

                    return "";
                }
            }
        }
    }
}
