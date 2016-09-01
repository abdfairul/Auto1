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

namespace SAAL
{
    public class SAAL_Interface
    {
        /* Global definition */

        // Global variables & Objects        
        private SerialPort PUP_PORT = new SerialPort();     // Declare a port for Plug Unplug Jig
        private SerialPort TG19_PORT = new SerialPort();    // Declare a port for TG19
        private SerialPort TG45_PORT = new SerialPort();    // Declare a port for TG45
        private SerialPort QD882_PORT = new SerialPort();   // Declare a port for QD882
        private SerialPort UART_PORT = new SerialPort();    // Declare a port for UART

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

        // Create a telnet connection to LG3803
        TelnetConnection tc = null;
        public Boolean bIsTelnetConnected = false;

        // Create objects and definition for IRTRANS module
        
        /* Made the IRTRANS variables to be public so that other forms can use them - to keep a single declaration */
        public IRTransServer irtObj = null; //new IR
        public Boolean bIsIRTransConnected = false;
        public const Int32 NUMRC = 7;  // number of supported RC
        public string[] RCFORMAT = new string[NUMRC] { "PHILIPS_RC5",
                                                    "PHILIPS_RC6",
                                                    "Matsushita",
                                                    "NEC_THAI",
                                                    "NEC_INDIA",
                                                    "NEC",
                                                    "SANYO"};

        public string[] RCKEY = new string[] {
                                                    "POWER",
                                                    "CHUP",
                                                    "CHDN",
                                                    "PREVCH",
                                                    "0",
                                                    "1",
                                                    "2",
                                                    "3",
                                                    "4",
                                                    "5",
                                                    "6",
                                                    "7",
                                                    "8",
                                                    "9",
                                                    ".",
                                                    "INFO",
                                                    "SAP",
                                                    "VOLUP",
                                                    "VOLDN",
                                                    "MUTE",
                                                    "SLEEP",
                                                    "CC",
                                                    "FREEZE",
                                                    "SOURCE",
                                                    "FORMAT",
                                                    "AUTOMODE",
                                                    "AUTOPIC",
                                                    "AUTOSND",
                                                    "OPTIONS",
                                                    "UP",
                                                    "DOWN",
                                                    "RIGHT",
                                                    "LEFT",
                                                    "OK",
                                                    "BACK",
                                                    "MENU",
                                                    "PLAY",
                                                    "STOP",
                                                    "SKPUP",
                                                    "SKPDWN",
                                                    "PAUSE",
                                                    "REV",
                                                    "FWD",
                                                    "RED",
                                                    "GREEN",
                                                    "BLUE",
                                                    "YELLOW",
                                                    "HOME",
                                                    "EJECT",
                                                    "DISCMENU",
                                                    "TITLE",
                                                    "MODE",
                                                    "CLEAR",
                                                    "10",
                                                    "SUBTITLE",
                                                    "ANGLE",
                                                    "REPEAT",
                                                    "A2B",
                                                    "SRCHMODE",
                                                    "PIP",
                                                    "WHITEB",
                                                    "FACT"
        };


        // Structures


        // Enumerations
        
        // Telnet command format:   TS_n, p, q
        public enum EN_LG3803_TSSELECT   //n - TS input selection
        {
            PN,
            EXT_ASI,
            EXT_SPI,
            INTERNAL,
            RESERVE,
            ROM,
            LAST
        }

        public enum EN_LG3803_TSPATTERN   //p - for ROM only. Others for future implementation if required
        {
            COLORBAR,
            RAMP,
            MONOSCOPE,
            LAST
        }

        public enum EN_LG3803_SIZE         //q
        {
            R1080I_16_9,
            R720P_16_9,
            R480I_16_9,
            R480I_4_3,
            LAST
        }

        // n    - Telnet command:  SY_n
        public enum EN_LF3803_MOD  
        {
            DIGITAL_8VSB,   // antenna tuning mode
            DIGITAL_64QAM,    // cable tuning mode
            DIGITAL_256QAM,   // cable tuning mode
            ANALOG_MOD,
            LAST
        }

        public enum EN_QD882_SOURCE
        {
            HDMI_D,
            HDMI_H,
            VGA
        }

        public enum EN_QD882_FORMAT
        {
            // For HMDI output format
            Q1920_1080_30Hz,
            Q1360_768_60hZ,
            Q1280_768_60hZ,
            Q1400_1050_60hZ,
            Q1680_1050_60Hz,
            Q1280_1024_85Hz,
            Q1280_960_60Hz,
            Q1152_864_75Hz,
            Q640_480_72Hz,
            Q640_480_60Hz,
            // For PC output format
            Q1024_768_60Hz,
            Q640_480_75Hz,
            Q1280_960_60Hz_VGA,
            Q1280_1024_75Hz,
            Q_empty
        }

        public enum EN_TG45_FORMAT
        {
            HDMI_480i_169,
            HDMI_480i_43,
            HDMI_480p_169,
            HDMI_480p_43,
            HDMI_720p,
            HDMI_1080i_24Hz,
            HDMI_1080p_60Hz,
            HDMI_1080i,
            HDTV_480i,
            HDTV_480p,
            HDTV_720p_60Hz,
            HDTV_1080i_60Hz,
            HDTV_1080p_60Hz
        }

        public enum EN_TG19_FORMAT
        {
            //FREQ_5590_MHz,
            SIG_COLORBAR
        }

        /* End Of - Global definition */


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
        public Boolean LG3803_SendCMD(EN_LG3803_TSSELECT tstype, EN_LG3803_TSPATTERN pattern, EN_LG3803_SIZE size, EN_LF3803_MOD mod, String freq)
        {
            if (tstype < EN_LG3803_TSSELECT.LAST &&
                pattern < EN_LG3803_TSPATTERN.LAST &&
                size < EN_LG3803_SIZE.LAST &&
                mod < EN_LF3803_MOD.LAST)
            {
                // prepare string for TS setting
                string cmd;

                if (tc != null)
                {
                    // prepare string for transport stream setting
                    switch (tstype)
                    {
                        case EN_LG3803_TSSELECT.PN:
                            cmd = "rem TS " + (Int32)tstype;
                            break;

                        case EN_LG3803_TSSELECT.EXT_ASI:
                        case EN_LG3803_TSSELECT.EXT_SPI:
                            cmd = "rem TS " + (Int32)tstype + ",0"; // always MPEG2TS
                            break;

                        case EN_LG3803_TSSELECT.ROM:
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
        public Boolean IRTRANS_Init()
        {
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
        public Boolean IRTRANS_SendCMD(string rctype, string key)
        {
            Boolean bSend = false;

            // Checks to ensure RC Type is supported
            for (Int32 index = 0; index < NUMRC; index++)
            {
                if (rctype == RCFORMAT[index])
                {
                    bSend = true;
                    break;
                }
            }

            irtObj.IRSend(rctype, key);

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

        /* End Of - SAAL for IRTRANS */


        /* SAAL for Camera */
        /*  Function:       CAM_GetCAMList()
         *  Parameters:     Output CamList[] - string array of available camera names
         *  Description:    To get the list of available camera(s) stored in a string array
         */
        public Boolean CAM_GetCAMList(out string[] CamList)
        {
            videoDevice = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            string[] mycamlist = new string[videoDevice.Count];
            Int32 index = 0;

            foreach (FilterInfo device in videoDevice)
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

            videoDevice = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo device in videoDevice)
            {
                if (device.Name == cam_name)
                {
                    IsVideoConnected = true;
                    videoDisplay = new VideoCaptureDevice(videoDevice[index].MonikerString);
                    videoDisplay.NewFrame += new NewFrameEventHandler(video_NewFrame);
                    break;
                }
                index++;
            }

            return IsVideoConnected;
        }

        /*  Task:           video_NewFrame()
         *  Parameters:     Input -> sender - callback from who, e -> passed in parameter
         *  Description:    thread to update currentImage bitmap
         */
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            currentImage = (Bitmap)eventArgs.Frame.Clone();
            Task.Delay(10); // give up task to reduce CPU loading
        }

        /*  Function:       CAM_IsAlive()
         *  Parameters:     None
         *  Description:    Returns if camera is active (TRUE - active, FALSE - inactive)
         */
        public Boolean CAM_IsAlive()
        {
            return IsVideoConnected;
        }

        /*  Function:       CAM_CaptureAndSave()
         *  Parameters:     Output String ImageName - file location of the image captured
         *  Description:    screen shot based on the videoDisplay resource
         */
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

        /* End of - SAAL for Camera */


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
                QD882_PORT.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(DataReceivedhandler);

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
        public void DataReceivedhandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();

            Console.WriteLine("Data Received:");
            Console.Write(indata);
            myedidstr += indata;

            //string[] myedid_data = new string[indata.Length];
            //for (Int32 index = 0; index < indata.Length; index++)
            //{
            //    myedid_data[index] = indata.Substring(index, 2);
            //}
        }

        /*  Function:       QD882_SetSource()
         *  Parameters:     Input source, format
         *  Description:    Setup callback fro the QD882
         */
        public Boolean QD882_SetSource(EN_QD882_SOURCE source, EN_QD882_FORMAT format)
        {
            string fullcmd = null;
            string cmd1 = null, cmd2 = null;

            switch (source)
            {
                case EN_QD882_SOURCE.HDMI_D:
                    cmd1 = "XVSI 3" + "ALLU\r\n";
                    break;
                case EN_QD882_SOURCE.HDMI_H:
                    cmd1 = "XVSI 0" + "ALLU\r\n";
                    break;
                case EN_QD882_SOURCE.VGA:
                    cmd1 = "XVSI 9" + "ALLU\r\n";
                    break;
            }

            switch (format)
            {
                case EN_QD882_FORMAT.Q1360_768_60hZ:
                    cmd2 = "FMTL CVR1360H" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1280_768_60hZ:
                    cmd2 = "FMTL CVT1250E" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1400_1050_60hZ:
                    cmd2 = "FMTL CVT1460" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1680_1050_60Hz:
                    cmd2 = "FMTL CVT1660D" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1280_1024_85Hz:
                    cmd2 = "FMTL CVT1285G" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1280_960_60Hz:
                    cmd2 = "FMTL CVR1260" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1152_864_75Hz:
                    cmd2 = "FMTL DMT1175" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q640_480_72Hz:
                    cmd2 = "FMTL DMT0672" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q640_480_60Hz:
                    cmd2 = "FMTL DMT0660" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1024_768_60Hz:
                    cmd2 = "FMTL DMT1060" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q640_480_75Hz:
                    cmd2 = "FMTL DMT0675" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1280_960_60Hz_VGA:
                    cmd2 = "FMTL DMT1260A" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1280_1024_75Hz:
                    cmd2 = "FMTL DMT1275G" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q1920_1080_30Hz:
                    cmd2 = "FMTL 1080i30" + "ALLU\r\n";
                    break;
                case EN_QD882_FORMAT.Q_empty:
                    cmd2 = "ALLU\r\n";
                    break;
            }
            fullcmd = cmd1 + cmd2;
            QD882_PORT.Write(fullcmd + "ALLU\r\n");

            return true;
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
        
        /* SAAL for TG45 */

        /*  Function:       Tg45_Setup()
         *  Parameters:     Input - commport (string of command to send on TG45. Example, "COM1")
         *                  Input - baudrate (int of baudrate value. Example, "115200")
         *  Description:    Create the connection of plug unplug jig
         */
        public Boolean Tg45_Setup(string commport, Int32 baudrate)
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
        public void TG45_ClosePort()
        {
            if(TG45_PORT.IsOpen)
            {
                TG45_PORT.Close();
            }
        }

        /*  Function:       TG45_SetFormat()
         *  Parameters:     Input - format (string of command to send on TG45. Example, "chgsys 43")
         *  Description:    Create the connection of plug unplug jig
         */
        public Boolean TG45_SetFormat(EN_TG45_FORMAT format)
        {
            string str = "";

            switch (format)
            {
                case EN_TG45_FORMAT.HDMI_1080i_24Hz:
                    str = "chgsys 43" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_1080p_60Hz:
                    str = "chgsys 19" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_480i_169:
                    str = "chgsys 6" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_480i_43:
                    str = "chgsys 5" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_480p_169:
                    str = "chgsys 2" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_480p_43:
                    str = "chgsys 1" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_720p:
                    str = "chgsys 3" + '\n';
                    break;

                case EN_TG45_FORMAT.HDTV_1080i_60Hz:
                    str = "chgsys 80" + '\n';
                    break;

                case EN_TG45_FORMAT.HDTV_1080p_60Hz:
                    str = "chgsys 82" + '\n';
                    break;

                case EN_TG45_FORMAT.HDTV_720p_60Hz:
                    str = "chgsys 90" + '\n';
                    break;

                case EN_TG45_FORMAT.HDMI_1080i:
                    str = "chgsys 4" + '\n';
                    break;

                case EN_TG45_FORMAT.HDTV_480i:
                    str = "chgsys 6" + '\n';
                    break;

                case EN_TG45_FORMAT.HDTV_480p:
                    str = "chgsys 2" + '\n';
                    break;
            }

            if (str != null)
                TG45_PORT.Write(str);

            return true;
        }

        /* End Of - SAAL for TG45 */

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
        public Boolean TG19_Settings(EN_TG19_FORMAT settings, string freq)
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

                case EN_TG19_FORMAT.SIG_COLORBAR:
                    str = "VSIGCO/1" + "\r";
                    break;
            }
            if (str != null)
                TG19_PORT.Write(str);

            // send frequency setting
            if (freq != null)
            {
                Thread.Sleep(500);  //wait 100ms for previous message to send thru

                Double double_freq = Convert.ToDouble(freq);
                str = "F/" + freq + "\r";
                TG19_PORT.Write(str);
            }

            return true;
        }

        /* End Of - SAAL for TG19 */

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

        /* End Of - SAAL for Plug Unplug */


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
            UART_PORT.Open();

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
                string str;

                str = cmd;
                UART_PORT.Write(str + "\r\n");
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
    }
}
