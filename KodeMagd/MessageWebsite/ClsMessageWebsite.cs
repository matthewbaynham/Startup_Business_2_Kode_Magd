using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Sockets;
using System.Threading;
using System.Net;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.MessageWebsite
{
    public static class ClsMessageWebsite
    {
        //public static void Connect(String sServer, String sMessage, ref TextBox txtResults)
        public static void sendInfo(ref bool bIsOk, ref string sErrorMessage, string sMachineId, string sSynchronisationKey, string sVersionNo)
        {
            try
            {
                //http://msdn.microsoft.com/en-us/library/system.net.sockets.tcpclient.aspx

                ClsSettings cSettings = new ClsSettings(); 

                // Create a TcpClient. 
                // Note, for this client to work you need to have a TcpServer  
                // connected to the same address as specified by the server, port 
                // combination.
                //Int32 iPort = 13000;
                TcpClient client = new TcpClient(cSettings.WebsiteAddress, cSettings.WebsitePortNo);

                string sMessage = "synchronous|" + sMachineId.Trim() + "|" + sSynchronisationKey.Trim() + "|" + sVersionNo;

                // Translate the passed message into ASCII and store it as a Byte array.
                Byte[] data = System.Text.Encoding.ASCII.GetBytes(sMessage);

                // Get a client stream for reading and writing. 
                //  Stream stream = client.GetStream();

                NetworkStream NsStream = client.GetStream();

                // Send the message to the connected TcpServer. 
                NsStream.Write(data, 0, data.Length);

                //txtResults.Text += "Sent: " + sMessage + "\n";
                //Console.WriteLine("Sent: {0}", sMessage);

                // Receive the TcpServer.response. 

                // Buffer to store the response bytes.
                data = new Byte[256];

                // String to store the response ASCII representation.
                String responseData = String.Empty;

                // Read the first batch of the TcpServer response bytes.
                Int32 bytes = NsStream.Read(data, 0, data.Length);
                responseData = System.Text.Encoding.ASCII.GetString(data, 0, bytes);

                if (responseData != cSettings.WebsiteConfirmationReply)
                {
                    bIsOk = false;
                    sErrorMessage = "Wrong Confirmation Message from website";
                    foreach (string sTemp in responseData.Split('|'))
                    { sErrorMessage += "\n" + sTemp; }
                }
                
                //txtResults.Text += "Received: " + responseData + "\n";
                //Console.WriteLine("Received: {0}", responseData);

                // Close everything.
                NsStream.Close();
                client.Close();

                cSettings = null;
            }
            catch (ArgumentNullException e)
            {
                bIsOk = false;
                sErrorMessage = "ArgumentNullException: " + e.Message;
                //Console.WriteLine("ArgumentNullException: {0}", e);
            }
            catch (SocketException e)
            {
                bIsOk = false;
                sErrorMessage = "SocketException: " + e.Message;
                //Console.WriteLine("SocketException: {0}", e);
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                bIsOk = false;
                sErrorMessage = ex.Message;
                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
            //txtResults.Text += "\n Press Enter to continue...";
            //Console.WriteLine("\n Press Enter to continue...");
            //Console.Read();
        }

        
        //private string UserName = "Unknown";
        //private StreamWriter swSender;
        //private StreamReader srReceiver;
        //private TcpClient tcpServer;
        // Needed to update the form with messages from another thread
        //private delegate void UpdateLogCallback(string strMessage);
        // Needed to set the form to a "disconnected" state from another thread
        //private delegate void CloseConnectionCallback(string strReason);
        //private Thread thrMessaging;
        //private bool Connected;
        //IPAddress ipAddr;
        //int iPort;

        public static void sendInfo_old(ref bool bIsOk, ref string sErrorMessage, string sMachineId, string sSynchronisationKey)
        {
            try
            {
                StreamWriter swSender;
                TcpClient tcpServer;
                IPAddress ipAddr = IPAddress.Parse("192.168.2.121");
                int iPort = 1986;
                string sMessage = "synchronous|" + sMachineId.Trim() + "|" + sSynchronisationKey.Trim();

                // Start a new TCP connections to the chat server
                tcpServer = new TcpClient();
                try
                {
                    tcpServer.Connect(ipAddr, iPort);
                }
                catch (Exception ex)
                {
                    bIsOk = false;
                    sErrorMessage = "Failed to connect to Website.\n\r\n\r";
                    sErrorMessage += ex.Message;
                }

                try
                {
                    // Send the desired username to the server
                    swSender = new StreamWriter(tcpServer.GetStream());
                    swSender.WriteLine(sMessage);
                    swSender.Flush();
                    swSender.Close();
                }
                catch (Exception ex)
                {
                    bIsOk = false;
                    sErrorMessage = "Failed to send information to Website.\n\r\n\r";
                    sErrorMessage += ex.Message;
                }

                // Close the objects
                tcpServer.Close();
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

        //private void ReceiveMessages()
        //{
        //    try
        //    {
        //        bool bIsOk = true;
        //        string sErrorMessage = "";


        //        // Receive the response from the server
        //        srReceiver = new StreamReader(tcpServer.GetStream());
        //        // If the first character of the response is 1, connection was successful
        //        string ConResponse = srReceiver.ReadLine();
        //        // If the first character is a 1, connection was successful
        //        if (ConResponse[0] == '1')
        //        {
        //            // Update the form to tell it we are now connected
        //            //this.Invoke(new UpdateLogCallback(this.UpdateLog), new object[] { "Connected Successfully!" });
        //        }
        //        else // If the first character is not a 1 (probably a 0), the connection was unsuccessful
        //        {
        //            bIsOk = false;
        //            sErrorMessage = "Could not connect to website.";
        //        }
        //        /*
        //        // While we are successfully connected, read incoming lines from the server
        //        while (Connected)
        //        {
        //            // Show the messages in the log TextBox
        //            this.Invoke(new UpdateLogCallback(this.UpdateLog), new object[] { srReceiver.ReadLine() });
        //        }
        //        */ 
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        /*
        // This method is called from a different thread in order to update the log TextBox
        private void UpdateLog(string strMessage)
        {
            try
            {
                // Append text also scrolls the TextBox to the bottom each time
                txtLog.AppendText(strMessage + "\r\n");
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
        */
    }
}
