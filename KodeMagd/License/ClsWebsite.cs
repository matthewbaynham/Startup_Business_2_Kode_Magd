using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.License
{
    class ClsWebsite
    {
        private const int ciDBTimeout = 60; //in seconds
        private const string csDB = "db197564x2008664";
        private const string csUser = "s197564_2008664";
        private const string csAddress = "mysql.webhosting13.1blu.de";
        private const string csPwd = "me_wp*!123";
        private const string csDriver = "{MySQL ODBC 5.2 ANSI Driver}";

        private ADODB.Connection connWebsite;

        ClsWebsite()
        {
            try
            {
                connect();
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

        /*
Dim con As ADODB.Connection
Dim rst As ADODB.Recordset
Dim lFieldCounter As Long

Set con = New ADODB.Connection
Set rst = New ADODB.Recordset

Const csDB As String = "db197564x2008664"
Const csUser As String = "s197564_2008664"
Const csAddress As String = "mysql.webhosting13.1blu.de"
Const csPwd As String = "me_wp*!123"
Const csDriver As String = "{MySQL ODBC 5.2 ANSI Driver}"

Const csSql As String = "select * from xyz123dave_posts;"

With con
    .ConnectionString = "DRIVER=" & csDriver & ";" & _
                        "SERVER=" & csAddress & ";" & _
                        "DATABASE=" & csDB & ";" & _
                        "USER=" & csUser & ";" & _
                        "PASSWORD=" & csPwd & ";" & _
                        "OPTION=3"
    .Open
End With
         */

        private void connect()
        {
            try
            {
                string sConnectionString = "DRIVER=" + csDriver + ";SERVER=" + csAddress + ";DATABASE=" + csDB + ";USER=" + csUser + ";PASSWORD=" + csPwd + ";OPTION=3";
                

                switch ((ADODB.ObjectStateEnum)this.connWebsite.State) 
                {
                    case ADODB.ObjectStateEnum.adStateClosed:
                        connWebsite.Open(sConnectionString); 
                        break;
                    case ADODB.ObjectStateEnum.adStateConnecting:
                        break;
                    case ADODB.ObjectStateEnum.adStateExecuting:
                        break;
                    case ADODB.ObjectStateEnum.adStateFetching:
                        break;
                    case ADODB.ObjectStateEnum.adStateOpen:
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

        public bool isUserNameTaken(string sUserName)
        {
            try
            {
                ADODB.Command cmd = new ADODB.Command();
                ADODB.Recordset rst = new ADODB.Recordset();
                ADODB.Parameter parUserName = new ADODB.Parameter();
                string sSql = "select  from  where  = ? ";

                cmd.CommandText = sSql;
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdText;
                cmd.CommandTimeout = ciDBTimeout;

                cmd.CommandText = sSql;
                parUserName = cmd.CreateParameter("CheckUser", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255, sUserName);
                cmd.Parameters.Append(parUserName);

                rst.Open(cmd, Type.Missing, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, -1);

                if (rst.BOF && rst.EOF)
                { }
                else
                { }

                rst.Close();

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

                return true;
            }
        }

        public bool isUserEMailAddressTaken(string sEMailAddress)
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

                return true;
            }
        }

        public bool isMachineIDTaken(string sMachineID)
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

                return true;
            }
        }

        public bool addUser(string sUserName, string sEMailAddress, string sPassword, string sMachineID)
        {
            try
            {


                string sSql = "insert into tbl () values ()";

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

                return true;
            }
        }
    }
}
