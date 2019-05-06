/* 
 * MJM 06/26/2009
This program monitors SMDR output from a Vodavi XTS and exports the data to SQL
 */

#region Namespace Inclusions
using System;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Text;
using System.Drawing;
using System.IO.Ports;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data.SqlClient;
//
using SerialPortTerminal.Properties;
#endregion

namespace SerialPortTerminal
{
    #region Public Enumerations
    public enum DataMode { Text, Hex }
    public enum LogMsgType { Incoming, Outgoing, Normal, Warning, Error };
    #endregion

    public partial class frmTerminal : Form
    {
        SqlConnection myConnection = new SqlConnection("Your connection string");
       


        #region Local Variables

        // The main control for communicating through the RS-232 port
        private SerialPort comport = new SerialPort();

        // Various colors for logging info
        private Color[] LogMsgTypeColor = { Color.Blue, Color.Green, Color.Black, Color.Orange, Color.Red };



        #endregion

        #region Constructor
        public frmTerminal()
        {
            // Build the form
            InitializeComponent();

            // Restore the users settings
            InitializeControlValues();

            // Enable/disable controls based on the current state
            EnableControls();

            OpenSerPort();

            // When data is recieved through the port, call this method
            comport.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
        }
        #endregion

        #region Local Methods

        /// <summary> Save the user's settings. </summary>
        private void SaveSettings()
        {
            Settings.Default.BaudRate = int.Parse(cmbBaudRate.Text);
            Settings.Default.DataBits = int.Parse(cmbDataBits.Text);
            Settings.Default.DataMode = CurrentDataMode;
            Settings.Default.Parity = (Parity)Enum.Parse(typeof(Parity), cmbParity.Text);
            Settings.Default.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cmbStopBits.Text);
            Settings.Default.PortName = cmbPortName.Text;

            Settings.Default.Save();
        }

        /// <summary> Populate the form's controls with default settings. </summary>
        private void InitializeControlValues()
        {
            cmbParity.Items.Clear(); cmbParity.Items.AddRange(Enum.GetNames(typeof(Parity)));
            cmbStopBits.Items.Clear(); cmbStopBits.Items.AddRange(Enum.GetNames(typeof(StopBits)));

            cmbParity.Text = Settings.Default.Parity.ToString();
            cmbStopBits.Text = Settings.Default.StopBits.ToString();
            cmbDataBits.Text = Settings.Default.DataBits.ToString();
            cmbParity.Text = Settings.Default.Parity.ToString();
            cmbBaudRate.Text = Settings.Default.BaudRate.ToString();
            CurrentDataMode = Settings.Default.DataMode;

            cmbPortName.Items.Clear();
            foreach (string s in SerialPort.GetPortNames())
                cmbPortName.Items.Add(s);

            if (cmbPortName.Items.Contains(Settings.Default.PortName)) cmbPortName.Text = Settings.Default.PortName;
            else if (cmbPortName.Items.Count > 0) cmbPortName.SelectedIndex = 0;
            else
            {
                MessageBox.Show(this, "There are no COM Ports detected on this computer.\nPlease install a COM Port and restart this app.", "No COM Ports Installed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        /// <summary> Enable/disable controls based on the app's current state. </summary>
        private void EnableControls()
        {
            // Enable/disable controls based on whether the port is open or not
            gbPortSettings.Enabled = !comport.IsOpen;


            if (comport.IsOpen) btnOpenPort.Text = "&Close Port";
            else btnOpenPort.Text = "&Open Port";
        }



        /// <summary> Log data to the terminal window. </summary>
        /// <param name="msgtype"> The type of message to be written. </param>
        /// <param name="msg"> The string containing the message to be shown. </param>
        private void Log(LogMsgType msgtype, string msg)
        {
            rtfTerminal.Invoke(new EventHandler(delegate
            {
                rtfTerminal.SelectedText = string.Empty;
                rtfTerminal.SelectionFont = new Font(rtfTerminal.SelectionFont, FontStyle.Bold);
                rtfTerminal.SelectionColor = LogMsgTypeColor[(int)msgtype];
                rtfTerminal.AppendText(msg);
                rtfTerminal.ScrollToCaret();
            }));
            
        }

       

       
        #endregion

        #region Local Properties
        private DataMode CurrentDataMode = DataMode.Text;
        
        #endregion

        #region Event Handlers


        private void frmTerminal_Shown(object sender, EventArgs e)
        {
            Log(LogMsgType.Normal, String.Format("Application Started at {0}\n", DateTime.Now));
        }
        private void frmTerminal_FormClosing(object sender, FormClosingEventArgs e)
        {
            // The form is closing, save the user's preferences
            SaveSettings();
            if (comport.IsOpen) comport.Close();
        }



        private void cmbBaudRate_Validating(object sender, CancelEventArgs e)
        { int x; e.Cancel = !int.TryParse(cmbBaudRate.Text, out x); }
        private void cmbDataBits_Validating(object sender, CancelEventArgs e)
        { int x; e.Cancel = !int.TryParse(cmbDataBits.Text, out x); }
        private void btnOpenPort_Click(object sender, EventArgs e)
        {
            OpenSerPort();
        }
        private void OpenSerPort()
        {
            // If the port is open, close it.
            if (comport.IsOpen) comport.Close();
            else
            {
                // Set the port's settings
                comport.BaudRate = int.Parse(cmbBaudRate.Text);
                comport.DataBits = int.Parse(cmbDataBits.Text);
                comport.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cmbStopBits.Text);
                comport.Parity = (Parity)Enum.Parse(typeof(Parity), cmbParity.Text);
                comport.PortName = cmbPortName.Text;
                //comport.ReceivedBytesThreshold = 80;


                // Open the port
                comport.Open();
            }

            // Change the state of the form's controls
            EnableControls();



        }


        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // This method will be called when there is data waiting in the port's buffer


            string strData = String.Empty;


            //strData = comport.ReadExisting();
            strData = comport.ReadTo("  \r\n");


            // Display the text to the user in the terminal
            Log(LogMsgType.Incoming, strData + "  \r\n");


            // Append 2 line Inbounds and Transfers to 1 line
         ProcessString(ref strData);
            

        }
       

        private void WriteFile(ref string wristr)
        {
            if (chkDebug.Checked)
            {
                DateTime thisDate = DateTime.Today;
                string correctDate = thisDate.ToString("d");
                string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;
                bool exists = System.IO.Directory.Exists(AppPath + @"\LOGS");

                correctDate = correctDate.Replace("/", "-");
                if (exists == false)
                {
                    System.IO.Directory.CreateDirectory(AppPath + @"\LOGS");
                    exists = true;
                }

                if (exists == true)
                {
                    System.IO.StreamWriter objFile = new System.IO.StreamWriter(AppPath + @"\LOGS" + @"\" + "PhoneLog " + correctDate + ".txt", true);

                    wristr = System.Text.RegularExpressions.Regex.Replace(wristr, "\\s+", " ");

                    objFile.Write(wristr + System.Environment.NewLine);
                    objFile.Close();
                    objFile.Dispose();
                }
            }
        }
        private void InsertSQL(ref string s)
        {
            // keep a text file around just in case.
            WriteFile(ref s);
            
            // Parse the string into segments for SQL Insert Statement
            int spapos = 0;
           

            s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ");


            string tempstr = String.Empty;
            string Extension = String.Empty;
            string CircuitID = String.Empty;
            string CallDuration = String.Empty;
            string CallStartTime = String.Empty;
            string CallDate = String.Empty;
            string CallType = String.Empty;
            string NumberDialed = String.Empty;
            string Qualifier = String.Empty;
            string InternalExt = String.Empty;
            string InboundNumber = String.Empty;

            tempstr = s;

            spapos = s.IndexOf(" ");
            Extension = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            spapos = s.IndexOf(" ");
            CircuitID = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            spapos = s.IndexOf(" ");
            CallDuration = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            spapos = s.IndexOf(" ");
            CallStartTime = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            spapos = s.IndexOf(" ");
            CallDate = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            CallType = s.Substring(0, 1);
            s = s.Substring(1);

            spapos = s.IndexOf(" ");
            NumberDialed = s.Substring(0, spapos);
            s = s.Substring(spapos + 1);

            if (s.Contains("**")) //We have more parsing to do!
            {
                


                spapos = s.IndexOf(" ");
                if (spapos == -1)
                {
                    s = s + " ";
                    spapos = s.IndexOf(" ");
                }
                Qualifier = s.Substring(0, spapos);
                s = s.Substring(spapos + 1);

                spapos = s.IndexOf(" ");
                if (spapos == -1)
                {
                    s = s + " ";
                    spapos = s.IndexOf(" ");
                }
                InternalExt = s.Substring(0, spapos);
                s = s.Substring(spapos + 1);

                
                InboundNumber = s;
               
            }

           try
            {
                myConnection.Open();
                SqlCommand myCommand = new SqlCommand("INSERT INTO tblPRPCallAccounting (Extension, CircuitID, CallDuration, CallStartTime, CallDate, CallType, NumberDialed, Qualifier, InternalExt, InboundNumber) " +
                "Values ('"+Extension+"', '"+CircuitID+"','"+CallDuration+"','"+CallStartTime+"','"+CallDate+"','"+CallType+"','"+NumberDialed+"','"+Qualifier+"','"+InternalExt+"','"+InboundNumber+"')", myConnection);
                myCommand.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (System.InvalidOperationException ex)
            {
                if (!System.Diagnostics.EventLog.SourceExists("SMDR Logger"))
                    System.Diagnostics.EventLog.CreateEventSource(
                       "SMDR Logger", "Application");

                eventLog1.Source = "SMDR Logger";

                string str;
                str = "Source:" + ex.Source;
                str += "\n" + "Message:" + ex.Message;
                str += "\n" + "\n";
                str += "\n" + "Stack Trace :" + ex.StackTrace;
                //Console.WriteLine(str, "Specific Exception");
                eventLog1.WriteEntry(str);

            }
            catch (System.Data.SqlClient.SqlException ex)
            {
                if (!System.Diagnostics.EventLog.SourceExists("SMDR Logger"))
                    System.Diagnostics.EventLog.CreateEventSource(
                       "SMDR Logger", "Application");

                eventLog1.Source = "SMDR Logger";

                string str;
                str = "Source:" + ex.Source;
                str += "\n" + "Message:" + ex.Message;
                //Console.WriteLine(str, "Database Exception");
                eventLog1.WriteEntry(str);
            }
            catch (System.Exception ex)
            {
                if (!System.Diagnostics.EventLog.SourceExists("SMDR Logger"))
                    System.Diagnostics.EventLog.CreateEventSource(
                       "SMDR Logger", "Application");
                
                eventLog1.Source = "SMDR Logger";

                string str;
                str = "Source:" + ex.Source;
                str += "\n" + "Message:" + ex.Message;
                //Console.WriteLine(str, "Generic Exception");
                eventLog1.WriteEntry(str);
            }
            finally
            {
               if (myConnection.State == ConnectionState.Open)
                { 
                    myConnection.Close();
                }
            }



        }
        private void ProcessString(ref string str)
        {

        
            char[] delimiterChars = { '\r', '\n' };
           
            string[] lines = str.Split(delimiterChars);
            

            int len = lines.Length;
            int i = 0;



      
            while (i < len)
            {
                
                string mystr = String.Empty;
                string bufferstr = String.Empty;
                int arrlen = 0;
                arrlen = lines.Length;
                bufferstr = lines.GetValue(i).ToString();

               

                if ((bufferstr.Contains("**")) && (arrlen > 1))
                {
                    // Append the 2 line inbound str
                   
                    mystr = lines.GetValue(i).ToString();
                    if (lines.GetValue(i + 1).ToString().Length == 0)
                    {
                        i = i + 2;
                        mystr = mystr + lines.GetValue(i).ToString();
                    }
                }
                else
                {
                    mystr = bufferstr;
                    if ((mystr.Length == 0) && (i != arrlen - 1))
                    {
                        mystr = lines.GetValue(i + 1).ToString();
                    }
                }

                                
                if ((mystr.Contains("STA")) || (mystr.Contains("CO")) || (mystr.Contains("TOTAL")) || (mystr.Contains("START")) || (mystr.Contains("DATE")) || (mystr.Contains("DIALED")) || (mystr.Contains("ACCOUNT")) || (mystr.Contains("CODE")) || mystr.Contains("COST"))
                {
                    mystr = ""; 
                }

               

                if (mystr.Length != 0) 
                {
                    InsertSQL(ref mystr);
                    
                    
                }
                
                i++;
            }

          

        }


        #endregion

        
    }
}
