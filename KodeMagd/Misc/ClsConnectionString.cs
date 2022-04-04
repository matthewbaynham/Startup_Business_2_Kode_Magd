using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    class ClsConnectionString
    {
        public enum enumAttributeType
        {
            eAttType_Bool,
            eAttType_String,
            eAttType_Int,
            eAttType_Collection,
            eAttType_Unknown
        }

        public struct strConnStrAttribute 
        {
            public enumAttributeType eType;
            public string sName;
            public bool bReadOnly;
        }

        List<strConnStrAttribute> stConnStrAttributes;
        SqlConnectionStringBuilder builder;

        public ClsConnectionString() 
        { 
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
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

        ~ClsConnectionString() 
        { 
            try
            {
                builder = null;
                stConnStrAttributes = null;
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

        public void fillLstConnStrAttributes() 
        {
            stConnStrAttributes = new List<strConnStrAttribute>();

            strConnStrAttribute objTemp = new strConnStrAttribute();

            objTemp.sName = "ApplicationName";
            objTemp.eType = enumAttributeType.eAttType_String;




            foreach (PropertyInfo propInfo in this.builder.GetType().GetProperties()) 
            {
                //propInfo.Name;
            }


            //conStringBuilder.ApplicationName string 
            //conStringBuilder.AsynchronousProcessing bool
            //conStringBuilder.AttachDBFilename string
            //conStringBuilder.ConnectionString string
            //conStringBuilder.ConnectTimeout int
            //conStringBuilder.ContextConnection bool
            //conStringBuilder.Count int (ReadOnly)
            //conStringBuilder.CurrentLanguage string
            //conStringBuilder.DataSource string
            //conStringBuilder.Encrypt bool
            //conStringBuilder.Enlist bool
            //conStringBuilder.FailoverPartner string
            //conStringBuilder.InitialCatalog string
            //conStringBuilder.IntegratedSecurity bool
            //conStringBuilder.IsFixedSize bool
            //conStringBuilder.IsReadOnly bool
            //conStringBuilder.Keys (collection of keys)
            //conStringBuilder.LoadBalanceTimeout int
            //conStringBuilder.MaxPoolSize int
            //conStringBuilder.MinPoolSize int
            //conStringBuilder.MultipleActiveResultSets bool
            //conStringBuilder.NetworkLibrary string
            //conStringBuilder.PacketSize int
            //conStringBuilder.Password string
            //conStringBuilder.PersistSecurityInfo bool
            //conStringBuilder.Pooling bool
            //conStringBuilder.Replication bool
            //conStringBuilder.TransactionBinding string
            //conStringBuilder.TrustServerCertificate bool
            //conStringBuilder.TypeSystemVersion string 
            //conStringBuilder.UserID string
            //conStringBuilder.UserInstance bool
            //conStringBuilder.Values (collection of values)
            //conStringBuilder.WorkstationID string

        
        }

        //public enumAttributeType attributeType(string sName) 
        //{
        //}


        public string ApplicationName 
        {
            get 
            {
                try
                {
                    return builder.ApplicationName;
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
                    builder.ApplicationName = value;
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
}
