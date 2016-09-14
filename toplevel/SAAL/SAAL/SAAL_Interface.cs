/****************************** Module Header ******************************\
Module Name:  SAAL_Interface.cs
Project:      SAAL API
Copyright (c) Funai SGP and Funai MY R&D.

Funai Software Automation Adaptation Layer (SAAL)

This source is subject to the Funai Software Licensing Agreement.
Kindly contact Funai Co. Ltd. for further details.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 * 
 * [Revision History]
 * v0.0.1.0 - Initial draft
 * v0.0.2.0 - Added QD882, UART APIs
 * v0.0.3.0 - Revised UART APIs to respective devices (TG19, TG45 and PlugUnPlug)
 *            to retain the open and close API synonymous for user to handle instead
 *            of using comm port declaration specific
 * v0.0.4.0 - Revised CAM_CaptureAndSave() to add string append for customization
 *            Revised MIC_Amplitude() with ProcessSample() handling
 * v0.0.5.0 - Edit QD882_format command
 *            Remove QD882_SendCMD is not needed
 * v0.0.6.0 - Fix LG3803_SendCMD() to send cmd in proper string
 * v0.0.7.0 - Add UART APIs
 *          - Add TG45_SetFormat API to set the resolution format on TG45
 *          - Add TG19_Settings API to set the frequency and signal on TG19
 * v0.0.8.0 - Updated API documentation
 *          - Updated TG19_Settings() to include frequency tuning (API_SAAL_v0.1.0.doc)
 *          - updated QD882_GetEDID() handler (by Chandru)
 * v0.0.9.0 - Added IRToy lib
 *			- Added 1080i for QD882 control
 * v0.0.9.1 - Bug fix foor QD882 commands string incorrect (did not follow spec)
 *			- Add TG19_SendCmd API to send cmd for TG19
 * v0.0.9.2 - Add TG59A equipment
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IRTrans.NET;
using MinimalisticTelnet;
using AForge.Video;
using AForge.Video.DirectShow;
using AForge.Imaging;
using System.IO;
using System.Drawing.Imaging;
using NAudio.Wave;
using System.Diagnostics;
using System.IO.Ports;
using System.Threading;
using PluginContracts;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Management;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.Util;
using IrToyLibrary;
using DirectShowLib;

namespace SAAL
{
    public class SAAL_Interface
    {
        /* Global definition */

        // Global variables & Objects        
        private SerialPort PUP_PORT = new SerialPort();     // Declare a port for Plug Unplug Jig
        private SerialPort TG19_PORT = new SerialPort();    // Declare a port for TG19
        private SerialPort TG45_PORT = new SerialPort();    // Declare a port for TG45
        private SerialPort TG45_59_PORT = new SerialPort();    // Declare a port for TG59
        private SerialPort QD882_PORT = new SerialPort();   // Declare a port for QD882
        private SerialPort UART_PORT = new SerialPort();    // Declare a port for UART

        private System.Net.Sockets.TcpClient TG59_TCP;

        private String SerialDataInput;
        private Int32 IndexStr;

        private String SerialDataInputUART;
        private Int32 IndexStrUART;
        public Boolean NoReboot = true;
        public Boolean RemoteSent = false;
        public Boolean NoFreezed = true;
        public String QuantumDataLabel = "QDModel";
        public Boolean SMPTEBarDisplay = false;

        //constructor
        public SAAL_Interface()
        {
            QD882_PORT.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(QD882_DataReceivedhandler);
            UART_PORT.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(QD882_DataReceivedFromUART);
        }

        // Creates object for Plug Unplug
        const string PLUGUNPLUG_PORT = "C,0,0,0,0";        

        // Creates object for Camera and Microphone
        public Boolean IsVideoConnected = false;
        public Boolean IsAudioConnected = false;

        // Video camera declaration
        public FilterInfoCollection videoDevice;
        public VideoCaptureDevice videoDisplay = new VideoCaptureDevice();
        private static Bitmap currentImage = null;
        private Int32 mic_amplitude = 0;

        // Video capture emgu library
        public Emgu.CV.Capture videoDisplay_cv = null;

        // Create a telnet connection to LG3803
        TelnetConnection tc = null;
        public Boolean bIsTelnetConnected = false;
        public Boolean bIsTG59Connected = false;

        System.Net.Sockets.TcpClient m_tcp; 
        System.Net.Sockets.NetworkStream m_ns;
        public String m_IPADDR;
        public Boolean bIsTG59pingpass = false;

        // Create objects and definition for IRTRANS module
        
        /* Made the IRTRANS variables to be public so that other forms can use them - to keep a single declaration */
        public IRTransServer irtObj = null; //new IR
        public Boolean bIsIRTransConnected = false;
        public const Int32 NUMRC = 7;  // number of supported RC      

        // IRTOYS
        public IrToySettings IRToySettings = null;
        public IrToyLib IRToyObj = null;
        private Boolean IRToy_semkey = true;
        public struct IRTOYRC
        {
            public string KEYLABEL;
            public string RAWDATA;
        }
        public IRTOYRC[] PHILIPSRC5 = new IRTOYRC[255];
        public IRTOYRC[] PHILIPSRC6 = new IRTOYRC[255];
        public IRTOYRC[] MATSUSHITA = new IRTOYRC[255];
        public IRTOYRC[] NECTHAI = new IRTOYRC[255];
        public IRTOYRC[] NECINDIA = new IRTOYRC[255];
        public IRTOYRC[] NECNEW = new IRTOYRC[255];
        public IRTOYRC[] SANYO = new IRTOYRC[255];

        public enum EN_MODEL_TYPE
        {
            EN_MODEL_JUSTTV,
            EN_MODEL_SMARTTV,
            EN_MODEL_CASTTV,
            EN_MODEL_MAX
        }

        #region LG3803
        /* SAAL for LG3803 */

        /*  Function:       LG3803_Login()
         *  Parameters:     None
         *  Description:    checks connection to IRTrans server application availability
         */
        public Boolean LG3803_Login(String IPADDR)
        {
            tc = new TelnetConnection(IPADDR, 23);

            if (tc != null)
            {
                //login with user "lg3803",password "lg3803", using a timeout of 1s, and show server output
                string s = tc.Login("lg3803", "lg3803", 500);
                //string s = tc.Login("admin", "Ds@dm1n", 1000);
                Console.Write(s);

                // server output should end with "$" or ">", otherwise the connection failed
                string prompt = s.TrimEnd();
                prompt = s.Substring(prompt.Length - 1, 1);
                if (prompt != "$" && prompt != ">")
                {
                    // Revert button back to original status
                    throw new Exception("Connection failed");
                }
                else
                    bIsTelnetConnected = true;
            }
            else
                Console.WriteLine("Connection failed");

            return bIsTelnetConnected;
        }

        /*  Function:       LG3803_Logout()
         *  Parameters:     None
         *  Description:    logs out from the logged in connection
         */
        public Boolean LG3803_Logout()
        {
            if (tc != null)
                tc.WriteLine("exit");

            // display server output
            Console.Write(tc.Read());
            Console.WriteLine("***DISCONNECTED");

            // return flag to disconnect mode
            bIsTelnetConnected = false;

            return true;
        }

        /*  Function:       LG3803_SendCMD()
         *  Parameters:     Input - cmd (string of command to send to LG3803. Example, "rem RF 500")
         *  Description:    Sends the string format based on LG3803 specification to control the modulator
         */
        public Boolean LG3803_SendCMD(EnumListUser.EN_LG3803_TSSELECT tstype, EnumListUser.EN_LG3803_TSPATTERN pattern, EnumListUser.EN_LG3803_SIZE size, EnumListUser.EN_LF3803_MOD mod, String freq)
        {
            if (tstype < EnumListUser.EN_LG3803_TSSELECT.LAST &&
                pattern < EnumListUser.EN_LG3803_TSPATTERN.LAST &&
                size < EnumListUser.EN_LG3803_SIZE.LAST &&
                mod < EnumListUser.EN_LF3803_MOD.LAST)
            {
                // prepare string for TS setting
                string cmd;

                if (tc != null)
                {
                    // prepare string for transport stream setting
                    switch (tstype)
                    {
                        case EnumListUser.EN_LG3803_TSSELECT.PN:
                            cmd = "rem TS " + (Int32)tstype;
                            break;

                        case EnumListUser.EN_LG3803_TSSELECT.EXT_ASI:
                        case EnumListUser.EN_LG3803_TSSELECT.EXT_SPI:
                            cmd = "rem TS " + (Int32)tstype + ",0"; // always MPEG2TS
                            break;

                        case EnumListUser.EN_LG3803_TSSELECT.ROM:
                            cmd = "rem TS " + (Int32)tstype + "," + (Int32)pattern + "," + (Int32)size + ",2";
                            break;

                        default:
                            cmd = "rem TS ?";   // not supported, just display a get TS settings
                            break;
                    }
                    
                    tc.WriteLine(cmd);
                    Console.Write(tc.Read());   // display server output

                    // prepare string for modulation type
                    cmd = "rem SY " + (Int32)mod;
                    tc.WriteLine(cmd);
                    Console.Write(tc.Read());   // display server output

                    // prepare string for broadcast frequency
                    cmd = "rem RF " + freq;
                    tc.WriteLine(cmd);
                    Console.Write(tc.Read());   // display server output

                    return true;
                }
                else
                    return false;
            }
            else
            {
                return false;
            }
        }

        /*  Function:       LG3803_IsConnected()
         *  Parameters:     None
         *  Description:    Returns if the TELNET connection still running
         */
        public Boolean LG3803_IsConnected()
        {
            return bIsTelnetConnected;
        }
        #endregion

        #region IRTRANS
        /* SAAL for IRTRANS */

        // Macro function to return IR device is connected or not
        public bool pIsIRTransConnected
        {
            get { return bIsIRTransConnected; } //NEW IR
        }

        /*  Function:       IRTRANS_Init()
         *  Parameters:     None
         *  Description:    checks connection to IRTrans server application availability
         */
        public Boolean IRTRANS_Init()// this portion not pass field test.when running overnight
        {
            //Exception thrown: 'System.UnauthorizedAccessException' in System.dll
            try
            {
                irtObj = null; // Reset the connect whenever user refresh it
                irtObj = new IRTransServer("localhost"); // connect to IR trans server

                if (irtObj != null)
                {
                    irtObj.StartAsnycReceiver();
                    bIsIRTransConnected = true;
                }
                else
                    bIsIRTransConnected = false;
            }
            catch (Exception)
            {
                System.Diagnostics.Process[] processName = System.Diagnostics.Process.GetProcessesByName("IRServer");
                bIsIRTransConnected = false;
            }

            return bIsIRTransConnected;
        }

        /*  Function:       IRTRANS_IsAlive()
         *  Parameters:     None
         *  Description:    Returns if the IRTRANS server is still running
         */
        public Boolean IRTRANS_IsAlive()
        {
            return bIsIRTransConnected;
        }

        /*  Function:       IRTRANS_IsAlive()
         *  Parameters:     None
         *  Description:    Returns if the IRTRANS server is still running
         */
        public Boolean IRTRANS_SendCMD(string rctype, string key)//The thread 0x1db0 has exited with code 0 (0x0).
                                                                 //The thread 0x28b8 has exited with code 0 (0x0).
                                                                 //The thread 0x1e70 has exited with code 0 (0x0).
                                                                 //The thread 0x3298 has exited with code 0 (0x0).
        {
            Boolean bSend = false;

            // Checks to ensure RC Type is supported
            for (Int32 index = 0; index < NUMRC; index++)
            {
                if (rctype == EnumListUser.RCFORMAT[index])
                {
                    bSend = true;
                    break;
                }
            }

            irtObj.IRSend(rctype, key);

            RemoteSent = true;

            return bSend;
        }

        /*  Function:       IRTRANS_Close()
         *  Parameters:     None
         *  Description:    Call this for application exit
         */
        public void IRTRANS_Close()
        {
            if (irtObj != null)
                irtObj.Close();
        }


        public String IRTRANS_Test()
        {
            String Manufacturer;
            String DeviveID;
            ManagementObjectSearcher searcher =
                   new ManagementObjectSearcher("root\\CIMV2",
                   "SELECT * FROM Win32_USBHub");

            int i = 0;
            foreach (ManagementObject queryObj in searcher.Get())
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("USBHub WMI " + i);
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Caption: {0}", queryObj["Caption"]);
                Console.WriteLine("Description: {0}", queryObj["Description"]);
                //Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                //Console.WriteLine("Manufacturer: {0}", queryObj["Manufacturer"]); not manufacturer it's Caption
                //Console.WriteLine("Name: {0}", queryObj["Name"]);
                //Console.WriteLine("PNPDeviceID: {0}", queryObj["PNPDeviceID"]);

                Console.WriteLine();
                Console.WriteLine();
                Manufacturer = queryObj["Caption"].ToString();
                DeviveID = queryObj["DeviceID"].ToString();
                Console.WriteLine(Manufacturer);//IRTrans USB

                //Caption: IRTrans USB
                //Description: IRTrans USB
                Regex regex = new Regex(@"IRTrans USB");
                Match match = regex.Match(Manufacturer);
                i++;
                if (match.Success)
                {
                    return DeviveID;
                }

            }
            return null;
        }

        /*  Function:       Check IRTrans connection
         *  Parameters:     
         *  Description:    
         */
        public Boolean IRTRANS_Setup()
        {
            String Manufacturer = "";
            String PortName = "";
            String Name = "";
            bool ret = false;
            ManagementObjectSearcher searcher =
                   new ManagementObjectSearcher("Select * from Win32_SerialPort");

            foreach (ManagementObject queryObj in searcher.Get())
            {
                Manufacturer = queryObj["Caption"].ToString();
                PortName = queryObj["DeviceID"].ToString();
                Name = queryObj["Name"].ToString();
                Console.WriteLine(Manufacturer + " : " + PortName + " : " + Name);

                Regex regex = new Regex(@"IRTrans USB");
                Match match = regex.Match(Manufacturer);

                if (match.Success)
                {
                    ret = true;
                    break;
                }
            }

            return ret;
        }


        /* End Of - SAAL for IRTRANS */
        #endregion

        #region IRTOYS
        public Boolean IRTOYS_Setup(bool auto, string portman)
        {
            bool ret = false;
            // Checks for IRToy

            ManagementObjectCollection ManObjReturn;
            ManagementObjectSearcher ManObjSearch;
            ManObjSearch = new ManagementObjectSearcher("Select * from WIN32_PnPEntity");
            ManObjReturn = ManObjSearch.Get();

            if (auto)
            {
                // Automatically find COM port
                foreach (ManagementObject ManObj in ManObjReturn)
                {
                    if (ManObj["Manufacturer"] != null)
                    {
                        if (ManObj["Manufacturer"].ToString().Contains("DangerousPrototypes.com"))
                        {
                            if (ManObj["Caption"] != null)
                            {
                                string[] portnames = SerialPort.GetPortNames();

                                foreach (string port in portnames)
                                {
                                    if (ManObj["Caption"].ToString().Contains(port))
                                    {
                                        String resultport = Regex.Match(port, @"\d+").Value;
                                        Byte portnum = Byte.Parse(resultport);

                                        IRToySettings = new IrToySettings { ComPort = portnum, UseHandshake = true, UseNotifyOnComplete = true, UseTransmitCount = true };

                                        ret = true;

                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                // Manually find COM port
                String resultport = Regex.Match(portman, @"\d+").Value;
                Byte portnum = Byte.Parse(resultport);

                IRToySettings = new IrToySettings { ComPort = portnum, UseHandshake = true, UseNotifyOnComplete = true, UseTransmitCount = true };

                ret = true;
            }

            return ret;
        }      

        public void IRTOYS_init()
        {
            IRTOYS_PHILIPS_RC5_Init();
            IRTOYS_PHILIPS_RC6_Init();
            IRTOYS_MATSUSHITA_Init();
            IRTOYS_NECTHAI_Init();
            IRTOYS_NECINDIA_Init();
            IRTOYS_NECNEW_Init();
            IRTOYS_SANYO_Init();
        }

        private void IRTOYS_PHILIPS_RC5_Init()
        {
            String tempstr = Properties.Resources.IRTOY_PHILIPS_RC5.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        PHILIPSRC5[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            PHILIPSRC5[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_PHILIPS_RC6_Init()
        {
            String tempstr = Properties.Resources.IRTOY_PHILIPS_RC6.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        PHILIPSRC6[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            PHILIPSRC6[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_MATSUSHITA_Init()
        {
            String tempstr = Properties.Resources.IRTOY_MATSUSHITA.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        MATSUSHITA[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            MATSUSHITA[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_NECTHAI_Init()
        {
            String tempstr = Properties.Resources.IRTOY_NEC_THAI.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        NECTHAI[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            NECTHAI[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_NECINDIA_Init()
        {
            String tempstr = Properties.Resources.IRTOY_NEC_INDIA.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        NECINDIA[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            NECINDIA[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_NECNEW_Init()
        {
            String tempstr = Properties.Resources.IRTOY_NEC.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        NECNEW[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            NECNEW[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;

                        dbindex++;
                    }
                }
            }
        }

        private void IRTOYS_SANYO_Init()
        {
            String tempstr = Properties.Resources.IRTOY_SANYO.ToString();
            Int32 dbindex = 0;  //must not be more than IRTOYRC declared array structure

            if (tempstr != null)
            {
                Int32 first = tempstr.IndexOf("[COMMANDS]") + "[COMMANDS]".Length;
                string startstr = tempstr.Substring(first, tempstr.Length - first);

                if (startstr != null)
                {
                    Int32 startchar = startstr.IndexOf("[") + "[".Length;
                    Int32 endchar = startstr.IndexOf("]") + "]".Length;

                    while (startchar != 0 && endchar != 0 && dbindex < 255)
                    {
                        //Get the next found label
                        SANYO[dbindex].KEYLABEL = startstr.Substring(startchar, endchar - startchar - 1);

                        //Move to next raw data
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        endchar = startstr.IndexOf("FF FF") + ("FF FF").Length;

                        //Get the next found raw data
                        if (endchar != 0)
                            SANYO[dbindex].RAWDATA = startstr.Substring(0, endchar);

                        //Move to the next label
                        startstr = startstr.Substring(endchar, startstr.Length - endchar);
                        startchar = startstr.IndexOf("[") + "[".Length;
                        endchar = startstr.IndexOf("]") + "]".Length;
                        
                        dbindex++;
                    }
                }
            }
        }

        public void IRTOYS_OpenCon()
        {
            IRTOYS_CloseCon();

            try
            {
                IRToyObj = IrToyLib.Connect(IRToySettings);
            }
            catch (IrToyException ex)
            {
                //MessageBox.Show(ex.Message);
                
            }
            
        }

        public void IRTOYS_CloseCon()
        {
            if (IRToyObj != null)
                IRToyObj.Close();
        }

        public Boolean IRTOYS_SendCMD(string IRformat, string IRlabel)
        {
            bool ret = false;
            int loop = 0;

            while(!ret)
            {
                Console.WriteLine("Send loop : " + loop);
                ret = IRTOYS_SendCMD(IRformat, IRlabel, loop);
                loop++;
            }

            return ret;
        }

        private Boolean IRTOYS_SendCMD(string IRformat, string IRlabel, int loop)
        {
            bool ret = false;
            string IRToyRAWKey = "";

            if ((IRlabel != "") && (IRToy_semkey == true))
            {
                IRToy_semkey = false;

                IRTOYS_OpenCon();

                if (IRformat == "PHILIPS_RC5")
                {
                    for (Int32 index = 0; index < PHILIPSRC5.Length - 1; index++)
                    {
                        if (PHILIPSRC5[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = PHILIPSRC5[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "PHILIPS_RC6")
                {
                    for (Int32 index = 0; index < PHILIPSRC6.Length - 1; index++)
                    {
                        if (PHILIPSRC6[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = PHILIPSRC6[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "Matsushita")
                {
                    for (Int32 index = 0; index < MATSUSHITA.Length - 1; index++)
                    {
                        if (MATSUSHITA[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = MATSUSHITA[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "NEC_THAI")
                {
                    for (Int32 index = 0; index < NECTHAI.Length - 1; index++)
                    {
                        if (NECTHAI[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = NECTHAI[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "NEC_INDIA")
                {
                    for (Int32 index = 0; index < NECINDIA.Length - 1; index++)
                    {
                        if (NECINDIA[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = NECINDIA[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "NEC")
                {
                    for (Int32 index = 0; index < NECNEW.Length - 1; index++)
                    {
                        if (NECNEW[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = NECNEW[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                else if (IRformat == "SANYO")
                {
                    for (Int32 index = 0; index < SANYO.Length - 1; index++)
                    {
                        if (SANYO[index].KEYLABEL == IRlabel)
                        {
                            IRToyRAWKey = SANYO[index].RAWDATA;
                            ret = true;
                            break;
                        }
                    }
                }
                
                //try
                //{
                IRToyObj.Send(IRToyRAWKey.ToLower());
                IRTOYS_CloseCon();
                Thread.Sleep(3000);
                //IRToyObj.Send("00 00 ff ff"); // Reset     
                
                //}
                //catch
                //{
                //    Console.WriteLine("Error count : " + loop);
                //    IRTOYS_OpenCon();
                //    ret = false;
                //    throw;
                //}

                //Task.Delay(500);

                IRToy_semkey = true;
                
            }
            return ret;
        }

        private void IRTOYS_usb_reboot(string port)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = "DEVCON";
            proc.StartInfo.Arguments = "Remove *usb" + port;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.UseShellExecute = false;
            proc.Start();
        }
        #endregion

        #region Camera
        /* SAAL for Camera */
        /*  Function:       CAM_GetCAMList()
         *  Parameters:     Output CamList[] - string array of available camera names
         *  Description:    To get the list of available camera(s) stored in a string array
         */
        public Boolean CAM_GetCAMList(out string[] CamList)
        {
            videoDevice = new FilterInfoCollection(AForge.Video.DirectShow.FilterCategory.VideoInputDevice);
            string[] mycamlist = new string[videoDevice.Count];
            Int32 index = 0;

            foreach (AForge.Video.DirectShow.FilterInfo device in videoDevice)
            {
                mycamlist[index] = device.Name;
                index++;
            }

            CamList = mycamlist;

            return true;
        }

        /*  Function:       CAM_Init()
         *  Parameters:     Input cam_name
         *  Description:    Set the camera device that corresponds to the cam_name. This will also create a new thread
         *                  to capture image when CAM_CaptureAndSave() is called
         */
        public Boolean CAM_Init(string cam_name)
        {
            Int32 index = 0;

            videoDevice = new FilterInfoCollection(AForge.Video.DirectShow.FilterCategory.VideoInputDevice);
            foreach (AForge.Video.DirectShow.FilterInfo device in videoDevice)
            {
                if (device.Name == cam_name)
                {
                    IsVideoConnected = true;
                    videoDisplay = new VideoCaptureDevice(videoDevice[index].MonikerString);
                    break;
                }
                index++;
            }

            return IsVideoConnected;
        }

        /*  Function:       CAM_IsAlive()
         *  Parameters:     None
         *  Description:    Returns if camera is active (TRUE - active, FALSE - inactive)
         */
        public Boolean CAM_IsAlive()
        {
            return IsVideoConnected;
        }

        /*  Function:       CAM_Init()
         *  Parameters:     Input cam_name
         *  Description:    Set the camera device that corresponds to the cam_name. This will also create a new thread
         *                  to capture image when CAM_CaptureAndSave() is called
         */
        public Boolean CAM_Init_cv(string cam_name)
        {
            Int32 index = 0;
            IsVideoConnected = false;
            CvInvoke.UseOpenCL = false;
            //videoDevice = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            DsDevice[] _SystemCamereas = DsDevice.GetDevicesOfCat(DirectShowLib.FilterCategory.VideoInputDevice);
            //CaptureCollection videoDevice_cv = new CaptureCollection();

            foreach (DsDevice device in _SystemCamereas)
            {
                if (device.Name == cam_name)
                {
                    try
                    {
                        IsVideoConnected = true;
                        if (videoDisplay_cv != null)
                            videoDisplay_cv.Dispose();
                        videoDisplay_cv = new Emgu.CV.Capture(index);
                    }
                    catch (NullReferenceException)
                    {
                        IsVideoConnected = false;
                    }
                }
                index++;
            }

            return IsVideoConnected;
        }

        /* End of - SAAL for Camera */
        #endregion

        #region Mic
        /* Start SAAL for Microphone */
        /*  Function:       MIC_GetMICList()
         *  Parameters:     Output MICList[] - string array of available camera names
         *  Description:    To get the list of available microphone(s) stored in a string array
         */
        public Boolean MIC_GetMICList(out string[] MICList)
        {
            int MICInDevices = WaveIn.DeviceCount;
            Int32 index = 0;
            string[] myMICList = new string[WaveIn.DeviceCount];

            for (int MICInDevice = 0; MICInDevice < MICInDevices; MICInDevice++)
            {
                WaveInCapabilities deviceInfo = WaveIn.GetCapabilities(MICInDevice);
                myMICList[index] = deviceInfo.ProductName;
                index++;
            }

            MICList = myMICList;

            return true;
        }

        /*  Function:       MIC_Init()
         *  Parameters:     Input mic_name
         *  Description:    Set the microphone device that corresponds to the mic_name.
         */
        public Boolean MIC_Init(string mic_name)
        {
            Int32 index = 0;
            int MICInDevices = WaveIn.DeviceCount;

            for (int MICInDevice = 0; MICInDevice < MICInDevices; MICInDevice++)
            {
                WaveInCapabilities deviceInfo = WaveIn.GetCapabilities(MICInDevice);
                if (deviceInfo.ProductName == mic_name)
                {
                    int sampleRate = 8000; // 8 kHz
                    int channels = 1; // mono
                    WaveIn soundIn = new WaveIn();
                    soundIn.WaveFormat = new WaveFormat(sampleRate, channels);
                    soundIn.DeviceNumber = MICInDevice;
                    soundIn.DataAvailable += new EventHandler<WaveInEventArgs>(soundIn_DataAvailable);
                    soundIn.StartRecording();
                    //Console.WriteLine("[WAV]Input sampling at " + soundIn.WaveFormat.BitsPerSample);
                    IsAudioConnected = true;

                    break;
                }
                index++;
            }
            return IsAudioConnected;
        }

        /*  Event:          soundIn_DataAvailable()
         *  Parameters:     Input -> sender - callback from who, e -> passed in parameter
         *  Description:    convert sample data and convert to 32-bit data for ProcessSample()
         */
        private void soundIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            for (int index = 0; index < e.BytesRecorded; index += 2)
            {
                short sample = (short)((e.Buffer[index + 1] << 8) | e.Buffer[index + 0]);
                float sample32 = sample / 32768f;

                ProcessSample(sample32);
            }

            Task.Delay(10); // give up task to reduce CPU loading
        }

        /*  Function:       ProcessSample()
         *  Parameters:     Input-> sample1 (32-bit data to process)
         *  Description:    average the sampled data and provide the amplitude level in db
         */
        private double samplenumber;
        private double maxval;
        private double minval;
        private void ProcessSample(float sample1)
        {
            samplenumber += 1;

            if (sample1 > maxval)
            {
                maxval = sample1;
            }

            if (sample1 < minval)
            {
                minval = sample1;
            }

            if (samplenumber >= ((double)8000 / 5))    // calculate average amplitude per cycle (8000 / 5 = 1600 bytes)
            {
                samplenumber = 0;

                double tmpvolume = (maxval - minval) * 100;

                // This is a none standardized calculation in order to amplify the min RMS value to return in %
                if (tmpvolume < 0.8)    //rms value at 0.707
                    tmpvolume = 0;
                else if (tmpvolume > 0.8 && tmpvolume < 10)
                    tmpvolume *= 10;
                else if (tmpvolume > 10 && tmpvolume < 50)
                    tmpvolume *= 2;
                else if (tmpvolume > 100)
                    tmpvolume = 100;

                mic_amplitude = (Int32)decimal.Round((decimal)tmpvolume, 2);

                maxval = 0;
                minval = 0;
                
                //Console.WriteLine("Audio decibel is " + mic_amplitude + "db");
            }
        }

        /*  Function:       MIC_IsAlive()
         *  Parameters:     None
         *  Description:    Returns if microphone is active (TRUE - active, FALSE - inactive)
         */
        public Boolean MIC_IsAlive()
        {
            return IsAudioConnected;
        }

        /*  Function:       MIC_Amplitude()
         *  Parameters:     Output MICList[] - string array of available camera names
         *  Description:    Returns the value of amplitude (0 - no sound, 100 - max sound)
         */
        public Int32 MIC_Amplitude()
        {
            return mic_amplitude;
        }
        #endregion

        #region QD882
        /* SAAL for QD882 */

        /*  Function:       QD882_Setup()
         *  Parameters:     Input commport (serial port #), baudrate (speed of connection)
         *  Description:    Setup the serial port # and baurate and store into QD882_PORT
         */
        public Boolean QD882_Setup(string commport, Int32 baudrate)
        {
            if (QD882_PORT.IsOpen == true)
                QD882_PORT.Close();

            QD882_PORT.PortName = commport;
            QD882_PORT.BaudRate = baudrate;
            QD882_PORT.ReadTimeout = 1000;
            QD882_PORT.WriteTimeout = 100;

            

            QD882_PORT.Open();
            
            if (!QD882_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /*  Function:       QD882_GetEDID()
         *  Parameters:     Output edid_data
         *  Description:    Returns the stored EDID data in edid_data that is read from QD882
         */
        string myedidstr;
        public Boolean QD882_GetEDID(out string[] edid_data)
        {
            string[] temp_edid = null;
            //Int32 indexofedid = 0;

            if (QD882_PORT.IsOpen)
            {
                string str = "EDID?" + "ALLU\r";

                myedidstr = String.Empty;

                QD882_PORT.Write(str);
                //QD882_PORT.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(QD882_DataReceivedhandler);

                Thread.Sleep(3000);


                var start = myedidstr.Substring(myedidstr.LastIndexOf("EDID?ALLU\r") + 11);
                var end = myedidstr.Substring(myedidstr.LastIndexOf("\r\n"));


                int startindex = myedidstr.LastIndexOf("EDID?ALLU\r") + 11;
                int endindex = myedidstr.LastIndexOf("\r\n") - 2;
                int edid_length = (endindex - startindex) / 2;
                temp_edid = new String[edid_length];


                if (myedidstr != null)
                {
                    for (int index = 0; index < edid_length; index++)
                    {
                        temp_edid[index] = start.Substring(0, 2);
                        start = start.Substring(2);
                    }
                }
            }
            else
            {
                //edid_data = false;
            }
            edid_data = temp_edid;

            return true;
        }

        /*  Event:          DataReceivedhandler()
         *  Parameters:     Input -> sender - callback from who, e -> passed in parameter
         *  Description:    Triggered from the serial port when a data buffer (i.e. EDID packet) is received
         */
        public void QD882_DataReceivedhandler(object sender, SerialDataReceivedEventArgs e)
        {
            //An unhandled exception of type 'System.InvalidOperationException' occurred in System.dll
            //Additional information: The port is closed.

            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();

            Console.Write(IndexStr +"->"+indata);
            SerialDataInput += indata;

            IndexStr++;
        }

        public void QD882_DataReceivedFromUART(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();

            Console.Write(IndexStrUART +"->"+indata);
            SerialDataInputUART += indata;

            Regex PowerON = new Regex(@"Pon Finished");
            Match match_PowerON = PowerON.Match(indata);

            if (match_PowerON.Success)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("                       REBOOT DETECTED    ");
                NoReboot = false;
            }
            else
            {
                Console.WriteLine("                       NOREBOOT     ");
            }

            Regex NotFreezed = new Regex(@"[IdleFrame]::OnKey");
            Match match_NotFreezed = NotFreezed.Match(indata);

            if (RemoteSent)
            {
                if (match_NotFreezed.Success)
                {
                    Console.WriteLine();
                    Console.WriteLine();
                    Console.WriteLine("                       NOT FREEZED    ");
                    NoFreezed = true;
                }
                else
                {
                    Console.WriteLine("                       FREEZED     ");
                    NoFreezed = false;
                }

                RemoteSent = false;
            }

            IndexStrUART++;
        }

        

        /*  Function:       QD882_SetSource()
         *  Parameters:     Input source, format
         *  Description:    Setup callback fro the QD882
         */
        public Boolean QD882_SetSource(EnumListUser.EN_QD882_SOURCE source, EnumListUser.EN_QD882_FORMAT format)
        {
            string fullcmd = null;
            string cmd1 = null, cmd2 = null;

            switch (source)
            {
                case EnumListUser.EN_QD882_SOURCE.HDMI_D:
                    cmd1 = "XVSI 3\r\n";
                    break;
                case EnumListUser.EN_QD882_SOURCE.HDMI_H:
                    cmd1 = "XVSI 0\r\n";
                    break;
                case EnumListUser.EN_QD882_SOURCE.VGA:
                    cmd1 = "XVSI 9\r\n";
                    break;
            }

            switch (format)
            {
                case EnumListUser.EN_QD882_FORMAT.Q1360_768_60hZ:
                    cmd2 = "FMTL CVR1360H\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1280_768_60hZ:
                    cmd2 = "FMTL CVT1250E\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1400_1050_60hZ:
                    cmd2 = "FMTL CVT1460\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1680_1050_60Hz:
                    cmd2 = "FMTL CVT1660D\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1280_1024_85Hz:
                    cmd2 = "FMTL CVT1285G\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1280_960_60Hz:
                    cmd2 = "FMTL CVR1260\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1152_864_75Hz:
                    cmd2 = "FMTL DMT1175\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q640_480_72Hz:
                    cmd2 = "FMTL DMT0672\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q640_480_60Hz:
                    cmd2 = "FMTL DMT0660\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1024_768_60Hz:
                    cmd2 = "FMTL DMT1060\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q640_480_75Hz:
                    cmd2 = "FMTL DMT0675\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1280_960_60Hz_VGA:
                    cmd2 = "FMTL DMT1260A\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1280_1024_75Hz:
                    cmd2 = "FMTL DMT1275G\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q1920_1080_30Hz:
                    cmd2 = "FMTL 1080i30\r\n";
                    break;
                case EnumListUser.EN_QD882_FORMAT.Q_empty:
                    cmd2 = "ALLU\r\n";
                    break;
            }
            fullcmd = cmd1 + cmd2;
            QD882_PORT.Write(fullcmd + "ALLU\r\n");

            return true;
        }
        /*
         *  Function:       QD882_Test()
         *  Parameters:     void
         *  Description:    Iterate all COM available and return the valid connected QD882 COM port
         * */
        public String QD882_Test(BackgroundWorker worker)
        {
/*
known issue to note:
for 804 there is 2 physical serial port
1) RS 232
2) USB-COM

tested (1) have HW issue cannot used
tested (2) go to preference of the instrument to change either storage or COM once change MUST reboot devive to active the option
*/
            string activeCOM = null;
            string[] portnames = SerialPort.GetPortNames();
            int numopOfPort = portnames.Length;
            Console.WriteLine("num of port available " + numopOfPort);
            int num = 0;

            //UART_Test();
            foreach (string port in portnames)
            {
                //Console.WriteLine(port);
                try
                {
                    QD882_Setup(port, 9600);
                    SerialDataInput = String.Empty;
                    IndexStr = 0;
                    QD882_PORT.Write("VERF?\r\n"); //Returns the Firmware version number for the programmable devices

                    Thread.Sleep(100);// automatically close connection after 100 ms
                    QD882_ClosePort();
                    Console.WriteLine();
                    Console.WriteLine(SerialDataInput);

                    Regex regexQD882 = new Regex(@"20.18");
                    Match match = regexQD882.Match(SerialDataInput);
                    if (match.Success)  //20.1888001,01.04.11 (for QD882)
                                        //20.1887800,01.04.11 (for QD882)
                    {
                        activeCOM = port;
                        QuantumDataLabel = "QD882";
                        break; //once found just exit. faster.
                    }
                    Regex regexQD804 = new Regex(@"1207");
                    match = regexQD804.Match(SerialDataInput);
                    if (match.Success)  //12072742 (for QD804)
                    {
                        activeCOM = port;
                        QuantumDataLabel = "QD804";
                        break; //once found just exit. faster.
                    }

                    worker.ReportProgress(num * 100 / numopOfPort);
                    num++;

                }
                catch
                {

                }
                
            }
            Console.WriteLine("active COM is :" + activeCOM);
            return activeCOM;
        }

        public String UART_Test()
        {
            String Manufacturer;
            String DeviveID;
            ManagementObjectSearcher searcher =
                   new ManagementObjectSearcher("root\\CIMV2",
                   "SELECT * FROM WIN32_SerialPort");

            int i = 0;
            foreach (ManagementObject queryObj in searcher.Get())
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("SerialPort WMI " + i);
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Caption: {0}", queryObj["Caption"]);
                Console.WriteLine("Description: {0}", queryObj["Description"]);
                //Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                //Console.WriteLine("Manufacturer: {0}", queryObj["Manufacturer"]); not manufacturer it's Caption
                //Console.WriteLine("Name: {0}", queryObj["Name"]);
                //Console.WriteLine("PNPDeviceID: {0}", queryObj["PNPDeviceID"]);

                Console.WriteLine();
                Console.WriteLine();
                Manufacturer = queryObj["Caption"].ToString();
                DeviveID = queryObj["DeviceID"].ToString();
                Console.WriteLine(Manufacturer);//Silicon Labs CP210x USB to UART Bridge (COM1)

                //Caption: Silicon Labs CP210x USB to UART Bridge (COM1)
                //Description: Silicon Labs CP210x USB to UART Bridge
                Regex regex = new Regex(@"Silicon Labs");
                Match match = regex.Match(Manufacturer);
                i++;
                if (match.Success)
                {
                    return DeviveID;
                }
                
            }
            return null;
        }


        public Boolean QD882_SetSource(String COMport, String sourceType, String SourceFormat)
        {
            string fullcmd = null;
            string cmd1 = null, cmd2 = null;
            
            switch (sourceType)
            {
                case "HDMI_D":
                    cmd1 = "XVSI 3\r\n";
                    break;
                case "HDMI_H":
                    cmd1 = "XVSI 4\r\n";
                    break;
                case "VGA":
                    cmd1 = "XVSI 9\r\n";
                    break;
            }

            cmd2 = "FMTL " + SourceFormat + "\r\n";
            fullcmd = cmd1 + cmd2;

            Console.WriteLine("QD882_SetSource active COM is :" + COMport + " src type: " + sourceType + " format :" + SourceFormat);



            SerialDataInput = String.Empty;
            IndexStr = 0;
            QD882_PORT.Write(fullcmd + "ALLU\r\n"); 

            Thread.Sleep(100);// automatically close connection after 100 ms

            Console.WriteLine();
            Console.WriteLine(SerialDataInput);

            return true;
        }


        public String QD882_SetSource(String COMport, String sourceType)
        {
            string fullcmd = null;
            string cmd1 = null;

            switch (sourceType)
            {
                case "HDMI_D":
                    cmd1 = "XVSI 3\r\n";
                    break;
                case "HDMI_H":
                    cmd1 = "XVSI 4\r\n";
                    break;
                case "VGA":
                    cmd1 = "XVSI 9\r\n";
                    break;
            }

            fullcmd = cmd1;

            Console.WriteLine("QD882_SetSource active COM is :" + COMport + " src type: " + sourceType);



            SerialDataInput = String.Empty;
            IndexStr = 0;
            QD882_PORT.Write(fullcmd + "ALLU\r\n");

            Thread.Sleep(100);// automatically close connection after 100 ms

            Console.WriteLine();
            Console.WriteLine(SerialDataInput);

            return SerialDataInput;
        }

        public String QD882_SetSourceFormat(String COMport,String SourceFormat)
        {
            string fullcmd = null;
            string cmd2 = null;

            cmd2 = "FMTL " + SourceFormat + "\r\n";
            fullcmd = cmd2;

            Console.WriteLine("QD882_SetSource active COM is :" + COMport  + " format :" + SourceFormat);



            SerialDataInput = String.Empty;
            IndexStr = 0;
            QD882_PORT.Write(cmd2 + "\r\n");
            

            Thread.Sleep(100);// automatically close connection after 100 ms

            QD882_PORT.Write("ALLU \r\n");// same ALLU=FMTU

            Console.WriteLine();
            Console.WriteLine(SerialDataInput);

            return SerialDataInput;
        }

        public String QD882_SetSourcePattern(String COMport, String SourceFormat)
        {
            string fullcmd = null;
            string cmd2 = null;

            cmd2 = "IMGL " + SourceFormat + "\r\n";
            fullcmd = cmd2;

            Console.WriteLine("QD882_SetSourcePattern active COM is :" + COMport + " format :" + SourceFormat);



            SerialDataInput = String.Empty;
            IndexStr = 0;
            QD882_PORT.Write(cmd2 + "\r\n");


            Thread.Sleep(100);// automatically close connection after 100 ms

            QD882_PORT.Write("IMGU \r\n");

            Console.WriteLine();
            Console.WriteLine(SerialDataInput);

            return SerialDataInput;
        }


        public String QD882_SendCmd(String COMport, String cmd)
        {

            SerialDataInput = String.Empty;
            IndexStr = 0;
            QD882_PORT.Write(cmd + "\r\n");

            Thread.Sleep(100);// automatically close connection after 100 ms

            Console.WriteLine();
            Console.WriteLine(SerialDataInput);

            return SerialDataInput; 
        }

        /*  Function:       QD882_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        public void QD882_ClosePort()
        {
            if (QD882_PORT.IsOpen)
            {
                QD882_PORT.Close();
            }
        }

        /* End Of - SAAL for QD882 */
        #endregion

        #region TG45 & TG59
        /* SAAL for TG45 & TG59 */

        /*  Function:       LAN Setup
         *  Parameters:     Input - IP addr
         *  Description:    Establish LAN connection
         */
        public Boolean TG45_59_Lan_Setup(string host)
        {
            int port = 9000;
            bool ret = false;

            try
            {
                TG59_TCP = new System.Net.Sockets.TcpClient(host, port);
                if (TG59_TCP != null && TG59_TCP.Client != null && TG59_TCP.Client.Connected)
                    ret = true;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
                return ret;
            }
            return ret;
        }

        /*  Function:       LAN port close
         *  Parameters:     
         *  Description:    Close LAN connection
         */
        public void TG45_59_Lan_Close()
        {
            if (TG59_TCP != null && TG59_TCP.Client != null && TG59_TCP.Client.Connected)
                TG59_TCP.Close();
        }

        /*  Function:       LAN test connection
         *  Parameters:     Input - IP addr
         *  Description:    Check LAN connection
         */
        public Boolean TG45_59_Lan_TestPerCon(string host)
        {
            System.Text.Encoding enc = System.Text.Encoding.UTF8;

            if (!TG45_59_Lan_Setup(host))
                return false;

            System.Net.Sockets.NetworkStream ns = TG59_TCP.GetStream();

            string checkStatus = "sysls\n";
            byte[] sendBytes = enc.GetBytes(checkStatus);
            ns.Write(sendBytes, 0, sendBytes.Length);

            byte[] resBytes = new byte[256];
            int resSize;
            resSize = ns.Read(resBytes, 0, resBytes.Length);
            /*System.IO.MemoryStream ms = new System.IO.MemoryStream();
            do
            {
                resSize = ns.Read(resBytes, 0, resBytes.Length);
                if (resSize == 0)
                {
                    return false;
                }
                ms.Write(resBytes, 0, resSize);
                Thread.Sleep(100);
            } while (ns.DataAvailable);            
            string resMsg = enc.GetString(ms.ToArray());
            Console.WriteLine(resMsg);
            ms.Close();
            */
            TG45_59_Lan_Close();

            if (resSize == 0)
                return false;
            else
                return true;
        }

        /*  Function:       LAN send command
         *  Parameters:     Command string
         *  Description:    Send command
         */
        public String TG45_59_Lan_SendCmd(string str)
        {
            Encoding enc = Encoding.UTF8;
            System.Net.Sockets.NetworkStream ns = TG59_TCP.GetStream();
            string readData = "";

            if (TG59_TCP != null && TG59_TCP.Client != null && TG59_TCP.Client.Connected)
            {
                byte[] sendBytes = enc.GetBytes(str);
                ns.Write(sendBytes, 0, sendBytes.Length);
                MemoryStream ms = new MemoryStream();
                byte[] resBytes = new byte[256];
                int resSize;
                do
                {
                    resSize = ns.Read(resBytes, 0, resBytes.Length);
                    if(resSize == 0)
                    {
                        return ""; // return empty string upon fail
                    }
                    ms.Write(resBytes, 0, resSize);
                    Thread.Sleep(100);
                } while (ns.DataAvailable);

                readData = enc.GetString(ms.ToArray());
                ms.Close();
                Console.WriteLine("Check 1 -> " + readData);
            }
            return readData;
        }

        /*  Function:       TG45_59_Setup()
         *  Parameters:     Input - commport (string of command to send on TG59. Example, "COM1")
         *                  Input - baudrate (int of baudrate value. Example, "115200")
         *  Description:    Create the connection of plug unplug jig
         */
        public Boolean TG45_59_Setup(string commport, Int32 baudrate)
        {
            bool ret = false;
            TG45_59_PORT.PortName = commport;
            TG45_59_PORT.BaudRate = baudrate;
            TG45_59_PORT.ReadTimeout = 1000;
            TG45_59_PORT.WriteTimeout = 100;

            try
            {
                TG45_59_PORT.Open();
                if (TG45_59_PORT.IsOpen)
                    ret = true;
            }
            catch
            {
                return ret;
            }
            return ret;                   
        }

        /*  Function:       TG45_59_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        public void TG45_59_ClosePort()
        {
            if (TG45_59_PORT.IsOpen)
            {
                TG45_59_PORT.Close();
            }
        }

        /*  Function:       TG45_59_SendCmd()
         *  Parameters:     Command string
         *  Description:    Send command
         */
        public String TG45_59_SendCmd(string str)
        {
            SerialDataInput = String.Empty;
            string SerialData = "";
            try
            {
                if (TG45_59_PORT.IsOpen)
                {
                    TG45_59_PORT.Write(str);
                    SerialData = TG45_59_PORT.ReadLine();
                    Console.WriteLine("Check 1 -> " + SerialData);
                    Console.WriteLine("Check 2 -> " + TG45_59_PORT.ReadExisting()); // Flush out remaining data
                }
            }
            catch
            {
                return ""; // return empty string upon fail
            }            
            return SerialData;
        }

        /*  Function:       TG45_59_TestCon()
         *  Parameters:     Command string
         *  Description:    Send command
         */
        public Dictionary<string, string> TG45_59_TestCon()
        {
            string compStrA = "";
            string compStrB = "";
            string[] portnames = SerialPort.GetPortNames();

            Dictionary<string, string> EquipmentList = new Dictionary<string,string>();
            EquipmentList.Add("TG45", "");
            EquipmentList.Add("TG59", "");

            foreach (string port in portnames)
            {
                Console.WriteLine(port);
                try
                {
                    TG45_59_Setup(port, 115200);
                    compStrA = TG45_59_SendCmd("sysls\n");

                    Console.WriteLine("Output Data 1 : " + compStrA);
                    Regex regex = new Regex(@"\d+x\d+");
                    Match match = regex.Match(compStrA);

                    if (match.Success)
                    {
                        compStrB = TG45_59_SendCmd("page 1\n");

                        Console.WriteLine("Output Data  2: " + compStrB);
                        Regex regex1 = new Regex(@"OK");
                        Match match1 = regex1.Match(compStrB);

                        if (match1.Success)
                            EquipmentList["TG45"] = port;
                        else
                            EquipmentList["TG59"] = port;
                    }
                }
                catch
                {
                }
                finally
                {
                    TG45_59_ClosePort();
                }
            }
            /*
            foreach (KeyValuePair<string, string> pair in EquipmentList)
            {
                Console.WriteLine("SAAL -> " + pair.Key + " : " + pair.Value);
            }*/

            return EquipmentList;
        }

        /*  Function:       TG45_59_TestCon()
         *  Parameters:     string TG45 or TG59
         *  Description:    Send command
         */
        public String TG45_59_TestPerCon(string devID)
        {
            string compStrA = "";
            string compStrB = "";
            string portStr = "";
            string[] portnames = SerialPort.GetPortNames();

            foreach (string port in portnames)
            {
                try
                {
                    TG45_59_Setup(port, 115200);

                    // Checking TG45, TG59 availability
                    compStrA = TG45_59_SendCmd("sysls\n");                 
                    Regex regex = new Regex(@"\d+x\d+");
                    Match match = regex.Match(compStrA);

                    if (match.Success)
                    {
                        // Checking TG45 availability
                        compStrB = TG45_59_SendCmd("page 1\n");
                        Regex regex1 = new Regex(@"OK");
                        Match match1 = regex1.Match(compStrB);

                        if (match1.Success && devID == "TG45")
                        {
                            portStr = port;
                            break;
                        }
                        else if(!match1.Success && devID == "TG59")
                        {
                            portStr = port;
                            break;
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    TG45_59_ClosePort();
                }
            }

            return portStr;
        }

        /*  Function:       TG45_59_SetFormat()
         *  Parameters:     Input - format (string of command to send on TG45/59. Example, "chgsys 43")
         *  Description:    Create the connection of plug unplug jig
         */
        public Boolean TG45_59_SetFormat(string inpStr, string commport, bool isUART)
        {
            Console.WriteLine(inpStr);
            string retStr = "";
            bool retVal = false;
            string cmdStr = "chgsys ";            

            // Special case
            if (inpStr.Contains("NTSC-M"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.SDTV_PAL_M;

            else if (inpStr.Contains("525p"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.SDTV_525P;

            else if (inpStr.Contains("HDTV4"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDTV_720_60p;

            else if (inpStr.Contains("HDTV2"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDTV_1080_60i;

            else if (inpStr.Contains("1080 60P"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDTV_1080_60p;

            else if (inpStr.Contains("HDMI001"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_640x480p_60_4_3;

            else if (inpStr.Contains("HDMI002"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_640x480p_60_4_3;

            else if (inpStr.Contains("HDMI003"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_720x480p_60_16_9;

            else if (inpStr.Contains("HDMI004"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_1280x720p_60_16_9;

            else if (inpStr.Contains("HDMI005"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_1920x1080i_60_16_9;

            else if (inpStr.Contains("HDMI006"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_2880x480i_60_4_3;

            else if (inpStr.Contains("HDMI007"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_720_1440_x480i_60_16_9;

            else if (inpStr.Contains("HDMI016"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_1920x1080p_60_16_9;

            else if (inpStr.Contains("HDMI032"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_1920x1080p_23_97_24_16_9;

            else if (inpStr.Contains("HDMI034"))
                cmdStr += (int)EnumListUser.EN_TG45_59_FORMAT.HDMI_1920x1080p_29_97_30_16_9;
            
            cmdStr += "\n";
            Console.WriteLine(cmdStr);

            if(isUART)
            {
                TG45_59_Setup(commport, 115200);
                retStr = TG45_59_SendCmd(cmdStr);
                if (retStr.Contains("OK"))
                    retVal = true;
                TG45_59_ClosePort();
            }
            else
            {
                TG45_59_Lan_Setup(commport);
                retStr = TG45_59_Lan_SendCmd(cmdStr);
                if (retStr.Contains("OK"))
                    retVal = true;
                TG45_59_Lan_Close();
            }            

            return retVal;
        }
        #endregion

        #region TG19
        /* SAAL for TG19 */

        /*  Function:       Tg19_Setup()
         *  Parameters:     Input - commport (string of command to send on TG19. Example, "COM1")
         *                  Input - baudrate (int of baudrate value. Example, "9600")
         *  Description:    Create the connection of TG19
         */
        public Boolean Tg19_Setup(string commport, Int32 baudrate)
        {
            if (commport == "" || baudrate <= 0)
            {
                return false;
            }
            else
            {
                TG19_PORT.PortName = commport;
                TG19_PORT.BaudRate = baudrate;
                TG19_PORT.Open();

                if (!TG19_PORT.IsOpen)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        /*  Function:       TG19_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        public void TG19_ClosePort()
        {
            if(TG19_PORT.IsOpen)
            {
                TG19_PORT.Close();
            }
        }

        /*  Function:       TG19_Settings()
         *  Parameters:     Input settings, freq
         *  Description:    To set the TG19 modulator to modulate according to video format and frequency
         */
        public Boolean TG19_Settings(EnumListUser.EN_TG19_FORMAT settings, string freq)
        {
            string /*freq, */str = "";

            // send format setting
            switch (settings)
            {
                //case EN_TG19_FORMAT.FREQ_5590_MHz:
                //    freq = "55.90";
                //    Double double_freq = Convert.ToDouble(freq);
                //    str = "F/" + freq + "\r";
                //    break;

                case EnumListUser.EN_TG19_FORMAT.SIG_COLORBAR:
                    str = "VSIGCO/1" + "\r";
                    break;
            }
            if (str != null)
                TG19_PORT.Write(str);

            // send frequency setting
            if (freq != null)
            {
                Thread.Sleep(500);  //wait 500ms for previous message to send thru

                Double double_freq = Convert.ToDouble(freq);
                str = "F/" + freq + "\r";
                TG19_PORT.Write(str);
            }

            return true;
        }

        /*  Function:       TG19_SendCmd()
         *  Parameters:     Input - cmd (string of command to send on Plug Unplug Jig. Example, "PO,B,5,1")
         *  Description:    Sends the string format based on Plug Unplug Spec to control the jig
         */
        public Boolean TG19_SendCmd(string cmd)
        {
            if(!TG19_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                TG19_PORT.Write(cmd + "\r");
                Thread.Sleep(200);
                return true;
            }
        }

        /* End Of - SAAL for TG19 */
        #endregion

        #region PUP
        /* SAAL for Plug Unplug */

        /*  Function:       PUP_Setup()
         *  Parameters:     Input - commport (string of command to send on Plug Unplug Jig. Example, "PO,B,5,1")
         *                  Input - baudrate (int of baudrate value. Example, "57600")
         *  Description:    Create the connection of plug unplug jig
         */
        public Boolean PUP_Setup(string commport, Int32 baudrate)
        {
            PUP_PORT.PortName = commport;
            PUP_PORT.BaudRate = baudrate;
            PUP_PORT.Open();
            PUP_PORT.Write(PLUGUNPLUG_PORT + '\r');
            System.Threading.Thread.Sleep(100);

            if(!PUP_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /*  Function:       PUP_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        public void PUP_ClosePort()
        {
            if(PUP_PORT.IsOpen)
            {
                PUP_PORT.Close();
            }
        }

        /*  Function:       PUP_SendCmd()
         *  Parameters:     Input - cmd (string of command to send on Plug Unplug Jig. Example, "PO,B,5,1")
         *  Description:    Sends the string format based on Plug Unplug Spec to control the jig
         */
        public Boolean PUP_SendCmd(string cmd)
        {
            if(!PUP_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                PUP_PORT.Write(cmd + '\r');
                return true;
            }
        }

        /*  Function:       PUP_Test()
         *  Parameters:     Input - None
         *  Description:    Test either the PUP is connected or not if connected will return it's COM ID.
         *                  if not connected wil return null
         */
        public String PUP_Test()
        {
            String Manufacturer;
            String DeviceID;
            ManagementObjectSearcher searcher =
                   new ManagementObjectSearcher("root\\CIMV2",
                   "SELECT * FROM WIN32_SerialPort");

            int i = 0;
            foreach (ManagementObject queryObj in searcher.Get())
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("SerialPort WMI " + i);
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Caption: {0}", queryObj["Caption"]);
                Console.WriteLine("Description: {0}", queryObj["Description"]);
                //Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                //Console.WriteLine("Manufacturer: {0}", queryObj["Manufacturer"]); not manufacturer it's Caption
                //Console.WriteLine("Name: {0}", queryObj["Name"]);
                //Console.WriteLine("PNPDeviceID: {0}", queryObj["PNPDeviceID"]);

                Console.WriteLine();
                Console.WriteLine();
                Manufacturer = queryObj["Caption"].ToString();
                DeviceID = queryObj["DeviceID"].ToString();
                Console.WriteLine(Manufacturer);

                Regex regex = new Regex(@"Microchip Technology, Inc.");
                Match match = regex.Match(Manufacturer);
                i++;
                if (match.Success)
                {
                    return DeviceID;
                }

            }
            return null;
        }

        /* End Of - SAAL for Plug Unplug */
        #endregion

        #region UART
        /* SAAL for UART */

        /*  Function:       UART_Setup()
         *  Parameters:     Input - commport (string of command to send to UART. Example, "COM1")
         *                  Input - baudrate (int of baudrate value. Example, "115200")
         *  Description:    Create the connection of UART jig
         */
        public Boolean UART_Setup(string setup, Int32 baudrate)
        {
            UART_PORT.PortName = setup;
            UART_PORT.BaudRate = baudrate;

            try
            {
                UART_PORT.Open();// sometime when hotplug the device the name still there but not refesh.

                if (!UART_PORT.IsOpen)
                {
                    return false;
                }
                else
                {
                    string weather_spoon = "WEATHER_SPOON 1\r\n";
                    UART_PORT.Write(weather_spoon);
                    return true;
                }
            }
            catch(Exception)
            {
                return false;
            }
        }

        /*  Function:       UART_SendCmd()
         *  Parameters:     Input - cmd (string of command to send on UART. Example, "FM FACTORY")
         *  Description:    Sends the string format based on UART command spec to control the jig
         */
        public Boolean UART_SendCmd(string cmd)
        {
            if (!UART_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                UART_PORT.Write(cmd + "\r\n");
                return true;
            }
        }

        /*  Function:       UART_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        public void UART_ClosePort()
        {
            if (UART_PORT.IsOpen)
            {
                UART_PORT.Close();
            }
        }

        /* End Of - SAAL for UART */
        #endregion

        #region Obsolete

        #region TG59
        /* SAAL for TG59A */

        /*  Function:       TG59A_control()
         *  Parameters:     None
         *  Description:    control TG59A equipment
         */
        [Obsolete("Not use anymore.")]
        public String TG59A_control(String IPADDR, String sendMsg)
        {            
            System.Text.Encoding enc = System.Text.Encoding.UTF8;
            int port = 9000;
            System.Net.Sockets.TcpClient tcp = new System.Net.Sockets.TcpClient(IPADDR, port);
            System.Net.Sockets.NetworkStream ns = tcp.GetStream();

            Console.WriteLine(sendMsg);
            byte[] sendBytes = enc.GetBytes(sendMsg);
            ns.Write(sendBytes, 0, sendBytes.Length);

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            byte[] resBytes = new byte[256];
            int resSize;
            do
            {
                resSize = ns.Read(resBytes, 0, resBytes.Length);
                if (resSize == 0)
                {
                    return string.Empty;
                }
                ms.Write(resBytes, 0, resSize);
                Thread.Sleep(100);
            } while (ns.DataAvailable);

            string resMsg = enc.GetString(ms.ToArray());
            ms.Close();
            ns.Close();
            tcp.Close();
            return resMsg;
        }

        [Obsolete("Not use anymore.")]
        public Boolean TG59A_connect(String IPADDR)
        {
            int port = 9000;
            try
            {
                m_tcp = new System.Net.Sockets.TcpClient(IPADDR, port);
                m_ns = m_tcp.GetStream();
                bIsTG59Connected = true;
            }
            catch
            {
                bIsTG59Connected = false;
            }

            return bIsTG59Connected;
        }

        [Obsolete("Not use anymore.")]
        public Boolean TG59A_disconnect()
        {
            m_ns.Close();
            m_tcp.Close();
            bIsTG59Connected = false;

            return bIsTG59Connected;
        }

        [Obsolete("Not use anymore.")]
        public String TG59A_SendCMD(String sendMsg)
        {
            System.Text.Encoding enc = System.Text.Encoding.UTF8;
            string resMsg;

            if (bIsTG59Connected)
            {
                //try
                //{
                    Console.WriteLine(sendMsg);
                    byte[] sendBytes = enc.GetBytes(sendMsg);
                    m_ns.Write(sendBytes, 0, sendBytes.Length);// here got some bug if use backgoundworker

                    System.IO.MemoryStream ms = new System.IO.MemoryStream();
                    byte[] resBytes = new byte[256];
                    int resSize;

                    do
                    {
                        resSize = m_ns.Read(resBytes, 0, resBytes.Length);
                        if (resSize == 0)
                        {
                            return string.Empty;
                        }
                        ms.Write(resBytes, 0, resSize);
                        Thread.Sleep(100);
                    } while (m_ns.DataAvailable);

                    resMsg = enc.GetString(ms.ToArray());

                //}
                //catch
                //{
                //    return string.Empty;
                //}

                return resMsg;
            }

            return string.Empty;
        }

        /*  Function:       TG59_SetFormat()
         *  Parameters:     Input - format (string of command to send on TG59. Example, "chgsys 43")
         *  Description:    Create the connection of plug unplug jig
         */
        [Obsolete("Not use anymore.")]
        public Boolean TG59_SetFormat(EnumListUser.EN_TG45_FORMAT format)
        {
            string str = "";

            switch (format)
            {
                case EnumListUser.EN_TG45_FORMAT.HDMI_1080i_24Hz:
                    str = "chgsys 43" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_1080p_60Hz:
                    str = "chgsys 19" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480i_169:
                    str = "chgsys 6" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480i_43:
                    str = "chgsys 5" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480p_169:
                    str = "chgsys 2" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480p_43:
                    str = "chgsys 1" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_720p:
                    str = "chgsys 3" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_1080i_60Hz:
                    str = "chgsys 80" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_1080p_60Hz:
                    str = "chgsys 82" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_720p_60Hz:
                    str = "chgsys 90" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_1080i:
                    str = "chgsys 4" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_480i:
                    str = "chgsys 6" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_480p:
                    str = "chgsys 2" + '\n';
                    break;
            }
            if (str != null)
            {
                TG59A_control(m_IPADDR, str);
            }
            return true;
        }
        #endregion

        #region TG45
        /* SAAL for TG45 */

        /*  Function:       Tg45_Setup()
         *  Parameters:     Input - commport (string of command to send on TG45. Example, "COM1")
         *                  Input - baudrate (int of baudrate value. Example, "115200")
         *  Description:    Create the connection of plug unplug jig
         */
        [Obsolete("Not use anymore.")]
        public Boolean TG45_Setup(string commport, Int32 baudrate)
        {
            TG45_PORT.PortName = commport;
            TG45_PORT.BaudRate = baudrate;
            TG45_PORT.Open();

            if (!TG45_PORT.IsOpen)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /*  Function:       TG45_ClosePort()
         *  Parameters:     None
         *  Description:    Close the open port
         */
        [Obsolete("Not use anymore.")]
        public void TG45_ClosePort()
        {
            if (TG45_PORT.IsOpen)
            {
                TG45_PORT.Close();
            }
        }

        /*  Function:       TG45_SetFormat()
         *  Parameters:     Input - format (string of command to send on TG45. Example, "chgsys 43")
         *  Description:    Create the connection of plug unplug jig
         */
        [Obsolete("Not use anymore.")]
        public Boolean TG45_SetFormat(EnumListUser.EN_TG45_FORMAT format)
        {
            string str = "";

            switch (format)
            {
                case EnumListUser.EN_TG45_FORMAT.HDMI_1080i_24Hz:
                    str = "chgsys 43" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_1080p_60Hz:
                    str = "chgsys 19" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480i_169:
                    str = "chgsys 6" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480i_43:
                    str = "chgsys 5" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480p_169:
                    str = "chgsys 2" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_480p_43:
                    str = "chgsys 1" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_720p:
                    str = "chgsys 3" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_1080i_60Hz:
                    str = "chgsys 80" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_1080p_60Hz:
                    str = "chgsys 82" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_720p_60Hz:
                    str = "chgsys 90" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDMI_1080i:
                    str = "chgsys 4" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_480i:
                    str = "chgsys 6" + '\n';
                    break;

                case EnumListUser.EN_TG45_FORMAT.HDTV_480p:
                    str = "chgsys 2" + '\n';
                    break;
            }

            if (str != null)
                TG45_PORT.Write(str);

            return true;
        }

        /* End Of - SAAL for TG45 */
        #endregion

        #region Camera
        /*  Task:           video_NewFrame()
         *  Parameters:     Input -> sender - callback from who, e -> passed in parameter
         *  Description:    thread to update currentImage bitmap
         */

        [Obsolete("Not use anymore.")]
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            currentImage = (Bitmap)eventArgs.Frame.Clone();
            Task.Delay(10); // give up task to reduce CPU loading
        }

        /*  Function:       CAM_CaptureAndSave()
         *  Parameters:     Output String ImageName - file location of the image captured
         *  Description:    screen shot based on the videoDisplay resource
         */
        [Obsolete("Not use anymore.")]
        public Boolean CAM_CaptureAndSave(out string ImageName, string append)
        {
            string myfolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\SAAL_CAM\";
            string curdate = DateTime.Now.ToString("yyyyMMdd_");
            string curtime = DateTime.Now.ToString("hhmm.fff_");

            if (videoDisplay.IsRunning == false)
            {
                ImageName = "Empty";
                return false;
            }

            ImageName = myfolder + curdate + curtime + append + ".jpg";

            if (!Directory.Exists(myfolder))
                System.IO.Directory.CreateDirectory(myfolder);

            Bitmap captureSample = (Bitmap)currentImage.Clone();

            captureSample.Save(ImageName, ImageFormat.Jpeg);

            captureSample.Dispose();

            return true;
        }

        /*  Function:       CAM_CompareImage()
         *  Parameters:     
         *                  Input string Actual_Image
         *                  Input string Captured_Image
         *  Description:    compares Actual_Image and Captured_Image and returns the similarity rate in % value
         *                  0 - none similar image
         *                  100 - accurately similar image
         */
        [Obsolete("Not use anymore.")]
        public Double CAM_CompareImage(string Actual_Image, string Captured_Image)
        {
            double dSimilarityResult = 0;
            Bitmap myactual = new Bitmap(Actual_Image);
            Bitmap mycapture = new Bitmap(Captured_Image);

            ExhaustiveTemplateMatching tm = new ExhaustiveTemplateMatching(0);
            TemplateMatch[] matchings = tm.ProcessImage(myactual, mycapture);
            // check similarity level

            myactual.Dispose();
            mycapture.Dispose();

            if (matchings.Length > 0)
                dSimilarityResult = matchings[0].Similarity;
            else if (matchings.Length <= 0)
                dSimilarityResult = 0;

            double rounded = Math.Round(dSimilarityResult, 2) * 100;

            return rounded;
        }
        #endregion

        #endregion

    }
}
