﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginContracts;
using System.Windows.Forms;
using System.Management;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Data;
using SAAL;
//using AForge.Video;
//using AForge.Video.DirectShow;
//using AForge.Imaging;
using System.Drawing;
using System.Drawing.Imaging;
using System.ComponentModel;
using System.Threading;

using OpenCvSharp;
using mainUI;

namespace Test
{
    public class QuantumAllFormat:IPlugin
    {
        #region Member: Extra
        private TextBox tbTester;
        private TextBox tbVer;
        #endregion

        IplImage imgOri, imgPreProcess;
        Bitmap bm,bm2;

        const double ScaleFactor = 2.5;
        const int MinNeighbors = 1;
        CvSize MinSize;

        CvHaarClassifierCascade cascade;

        int testx = 0;
        int testy = 0;
        CvPoint testpoint1 = new CvPoint(0, 0);
        CvPoint testpoint2 = new CvPoint(0, 0);
        CvPoint testpoint3 = new CvPoint(0, 0);
        CvPoint testpoint4 = new CvPoint(0, 0);
        CvPoint testpoint = new CvPoint(0, 0);
        CvFont font;
        DateTime timeStamp;
        CvFont fontTimeStamp;


        private Thread _cameraThread;

        SAAL_Interface mysaal = new SAAL_Interface();
        private SerialPort connection = new SerialPort();
        private RibbonTab ribbonEquipmentSetting;
        private DataGridViewRow toExecute;
        private bool BusyFlag;
        private bool DoFlag = true;
        private BeforeExecute BeforeFlag;
        private AfterExecute AfterFlag;
        private CellClickEvent ClickEventFlag;

        private RibbonComboBox cameraList; 
        private RibbonComboBox cameraResolution;
        private RibbonComboBox microfon;
        private PictureBox pictureBox;
        private PictureBox pctCvWindow;
        private PictureBox pictureBoxPopup;
        private ComboBox cbCamList;
        private int cbCamListIndex;
        private Button btOnOffCam;
        private Form popupBox;
        private int ResWidth = 1280;
        private int ResHeight = 720;

        private String connectedCOM;
        private String connectedUART;
        private String IR_RemoteConnected;
        private BackgroundWorker backgroundWorker;

        private RibbonHost ConnectionToolsHost;
        private Panel ConnectionCheckTools;
        private Button buttonTestEquipment;
        private ProgressBar progressBar;
        private Int32 inProgress;
        private Label QD882Label;
        private Label UARTLabel;
        private Label RemoteIRLabel;
        private Button buttonResetEquipment;
        private Button buttonCheckEquipment;
        private Button buttonUARTWeatherSpooning;
        private Button buttonIRSend;

        private RibbonHost signalTypeHost;
        private ToolTip toolTipInfo;
        private ListBox listBoxSignalType;//INTERFACE
        private String currentInterface;

        private RibbonHost signalFormatHost;
        private ListBox listBoxSignalFormat;//SOURCE
        private String currentSourceFormat = String.Empty;
        private String currentSpecInfo = String.Empty;

        private ErrorProvider ErrorInfoForInterface;
        private ErrorProvider ErrorInfoForFormat;
        private ErrorProvider ErrorInfoForPattern;

        private RibbonHost signalPatternHost;
        private ListBox listBoxSignalPattern;//CONTENT

        private int DEFINE_SOURCE = 1;
        private int DEFINE_SPEC_INFO = 2;
        private int DEFINE_TIMES = 6;
        private int DEFINE_RESULT = 8;
        private int DEFINE_VER = 9;
        private int DEFINE_TESTER = 10;
        private int DEFINE_DATE = 11;
        private int DEFINE_REMARK = 12;
        private int DEFINE_PICTURE_REF = 14;

        public DataGridViewRow ToExecute
        {
            get
            {
                return toExecute;
            }
            set
            {
                toExecute = value;
                DataRowProces();
                
            }

        }

        public PictureBox Picture
        {
            get
            {
                return pictureBox;
            }
            set
            {
                pictureBox = value;
            }

        }

        public bool Busy
        {
            get
            {
                return BusyFlag;
            }
        }

        public bool DoExecute
        {
            get
            {
                return DoFlag;
            }
        }

        public BeforeExecute BeforeExecute
        {
            get
            {
                return BeforeFlag;
            }
        }

        public AfterExecute AfterExecute
        {
            get
            {
                return AfterFlag;
            }
        }

        public CellClickEvent ClickEvent
        {
            get
            {
                return ClickEventFlag;
            }
        }

        public BeforeExecuteButAfterSelection BeforeExecuteButAfterSelection
        {
            get;
        }

        public ProgressBarChangedEvent ProgressBarChangedEvent
        {
            get;
        }

        private void DataRowProces()
        {
            BusyFlag = true;
            EnumListUser.ResultQuantumDataTest Result = EnumListUser.ResultQuantumDataTest.OK_NG;//assign first

            if ((connectedCOM == null) || (connectedUART==null))
            {
                MessageBox.Show("Please check the equipment connection");
                toExecute.Cells[DEFINE_REMARK].Value += " (#ERROR)";
                toExecute.Cells[DEFINE_RESULT].Value = "NT";
            }
            else
            {
                Int32 loop;

                int j;
                if (Int32.TryParse(toExecute.Cells[DEFINE_TIMES].Value.ToString(), out j))
                {
                    loop = j;
                }
                else
                {
                    loop = 1; // if fault set the loop only one loop
                }

                Console.WriteLine("+++==loop " + loop.ToString());

                Result = DataRowProcesExecuter(loop);
            }


            toExecute.Cells[DEFINE_TESTER].Value = tbTester.Text;
            toExecute.Cells[DEFINE_VER].Value = tbVer.Text;
            toExecute.Cells[DEFINE_DATE].Value = timeStamp;// common timestamp shared

            if (Result == EnumListUser.ResultQuantumDataTest.OK)
                toExecute.Cells[DEFINE_RESULT].Value = "OK";
            else
            { 
                if (Result == EnumListUser.ResultQuantumDataTest.NG1)
                    toExecute.Cells[DEFINE_RESULT].Value = "NG1";
                else if (Result == EnumListUser.ResultQuantumDataTest.NG2)
                    toExecute.Cells[DEFINE_RESULT].Value = "NG2";
                else if (Result == EnumListUser.ResultQuantumDataTest.NG3)
                    toExecute.Cells[DEFINE_RESULT].Value = "NG3";
                else if (Result == EnumListUser.ResultQuantumDataTest.NG4)
                    toExecute.Cells[DEFINE_RESULT].Value = "NG4";
                else if (Result == EnumListUser.ResultQuantumDataTest.NT)
                    toExecute.Cells[DEFINE_RESULT].Value = "NT";
                else if (Result == EnumListUser.ResultQuantumDataTest.OK_NG)
                    toExecute.Cells[DEFINE_RESULT].Value = "OK/NG";
                else if (Result == EnumListUser.ResultQuantumDataTest.PEND)
                    toExecute.Cells[DEFINE_RESULT].Value = "PEND";
                else if (Result == EnumListUser.ResultQuantumDataTest.dash)
                    toExecute.Cells[DEFINE_RESULT].Value = "-";

                SavePictureForReference(currentSourceFormat,"");

            }

            Thread.Sleep(2000);
            BusyFlag = false; 
        }

        private EnumListUser.ResultQuantumDataTest DataRowProcesExecuter(int loop)
        {
            String instrumentResponse = String.Empty;
            String formatCommand = String.Empty;
            String formatSpec = String.Empty;

            Match match;
            String[] format = EnumListUser.SourceFormat_QD882_FORMAT;

            formatCommand = toExecute.Cells[DEFINE_SOURCE].Value.ToString();// to check

            foreach (string src in format)
            {
                Regex formatRegex = new Regex(src);
                match = formatRegex.Match(formatCommand);

                if (match.Success)// emptty etc.
                {
                    // update the current if in database.
                    currentSourceFormat = toExecute.Cells[DEFINE_SOURCE].Value.ToString();
                    currentSpecInfo = toExecute.Cells[DEFINE_SPEC_INFO].Value.ToString();

                    Console.WriteLine("1) Curr:" + currentSourceFormat + " " + currentSpecInfo);
                }
            }

            Console.WriteLine("2) Curr:" + currentSourceFormat + " " + currentSpecInfo);

            var paths = new[] { "QuantumData", currentSourceFormat };
            toExecute.Cells[DEFINE_PICTURE_REF].Tag = paths;

            EnumListUser.ResultQuantumDataTest Result = EnumListUser.ResultQuantumDataTest.OK_NG;//assign first

            for (int i = 0; i < loop; i++)//any of loop failed will indicate failed and break the loop
            {
                #region Non-correspondence espected: TV doesn't reboot or freeze, and normal performance is possible.

                bool contains = Regex.IsMatch(currentSpecInfo, "Non-correspondence", RegexOptions.IgnoreCase);
                if (contains)// espected: TV doesn't reboot or freeze, and normal performance is possible.
                {
                    Console.WriteLine("NORMAL: loop " + i + " Performing command:" + currentSourceFormat);
                    instrumentResponse = mysaal.QD882_SetSourceFormat(connectedCOM, currentSourceFormat);

                    // UI updating...
                    UpdateCurrentSettingUI(currentSourceFormat, instrumentResponse);
                    bool containsError = Regex.IsMatch(instrumentResponse, "ERROR", RegexOptions.IgnoreCase);
                    if (containsError)
                    {
                        Result = EnumListUser.ResultQuantumDataTest.NT;
                        toExecute.Cells[DEFINE_REMARK].Value += instrumentResponse;
                        break;
                    }

                    Thread.Sleep(2000);// enough time to see the picture
                                       // send IR command to tv
                    if (IR_RemoteConnected == null)
                    {
                        Result = EnumListUser.ResultQuantumDataTest.NT;
                        MessageBox.Show("Please check the IR remote connection");
                        toExecute.Cells[DEFINE_REMARK].Value += "IR not connected ";
                        break;
                    }
                    else
                    {
                        for (int irsend = 0; irsend < 2; irsend++) // to check the TV not freezed
                        {
                            mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "MENU");//MENU //FORMAT //SOURCE //INFO
                            Thread.Sleep(2000);// enough time to see the picture changing
                                               // save the bitmap?
                            SavePictureForReference(currentSourceFormat,"Menu");
                        }
                    }

                    if (mysaal.NoFreezed && mysaal.NoReboot && mysaal.SMPTEBarDisplay)
                    {
                        Result = EnumListUser.ResultQuantumDataTest.OK;
                    }
                    else
                    {
                        Result = EnumListUser.ResultQuantumDataTest.NG1;
                        if (!mysaal.NoFreezed)
                            toExecute.Cells[DEFINE_REMARK].Value += "TV freezed ";
                        else if (!mysaal.NoReboot)
                            toExecute.Cells[DEFINE_REMARK].Value += "TV reboot ";
                        else if (!mysaal.SMPTEBarDisplay)
                            toExecute.Cells[DEFINE_REMARK].Value += "Pattern Not display ";
                        break;
                    }

                }
                #endregion //
                #region Correspondence espected: TV doesn't reboot or freeze, and normal performance is possible.
                else // Correspondence resolution
                // send remote 5 times to check [Automatic][4:3][Widescreen]
                //                              [Automatic][4:3][Widescreen][Full] for wxga model
                //                              [Automatic][4:3][Widescreen][Unscaled] for FHD model
                //                              [Automatic][4:3][Widescreen][Movie expand 16.9,14,9],[Super Zoom]
                //                              [4:3][Widescreen][Unscaled]
                //                              [Automatic][4:3][Widescreen]
                //                              [Full][Unscaled]
                {// expected: TV doesn't reboot or freeze, and normal performance is possible.
                 //           The image is normally displayed.
                 //           Confirm the all picture format is same as Picture Format Info 

                    // expected: 1) TV doesn't reboot or freeze, and normal performance is possible.
                    Console.WriteLine("SPECIAL: loop " + i + " Performing command:" + currentSourceFormat);
                    instrumentResponse = mysaal.QD882_SetSourceFormat(connectedCOM, currentSourceFormat);

                    // UI updating...
                    UpdateCurrentSettingUI(currentSourceFormat, instrumentResponse);
                    bool containsError = Regex.IsMatch(instrumentResponse, "ERROR", RegexOptions.IgnoreCase);
                    if (containsError)
                    {
                        Result = EnumListUser.ResultQuantumDataTest.NT;
                        toExecute.Cells[DEFINE_REMARK].Value += instrumentResponse;
                        break;
                    }

                    // send IR command to tv
                    if (IR_RemoteConnected == null)
                    {
                        Result = EnumListUser.ResultQuantumDataTest.NT;
                        MessageBox.Show("Please check the IR remote connection");
                        toExecute.Cells[DEFINE_REMARK].Value += "IR not connected ";
                        break;
                    }
                    else
                    {
                        #region 2) The image is normally displayed.
                        for (int irsend = 0; irsend < 5; irsend++)
                        {
                            //2) The image is normally displayed.

                            mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "FORMAT");//MENU //FORMAT //SOURCE //INFO
                            Thread.Sleep(1000);// need send twice to change. first click only to appear menu second send to change item

                            mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "FORMAT");//MENU //FORMAT //SOURCE //INFO
                            Thread.Sleep(2000);// enough time to see the picture changing
                                               // save teh bitmap?
                            SavePictureForReference(currentSourceFormat, "Picture_Format");
                            mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "BACK");
                            Thread.Sleep(2000);


                            if (mysaal.NoFreezed && mysaal.NoReboot && mysaal.SMPTEBarDisplay)
                            {
                                Result = EnumListUser.ResultQuantumDataTest.OK;
                            }
                            else
                            {
                                Result = EnumListUser.ResultQuantumDataTest.NG1;
                                if(!mysaal.NoFreezed)
                                    toExecute.Cells[DEFINE_REMARK].Value += "TV freezed ";
                                else if (!mysaal.NoReboot)
                                    toExecute.Cells[DEFINE_REMARK].Value += "TV reboot ";
                                else if (!mysaal.SMPTEBarDisplay)
                                    toExecute.Cells[DEFINE_REMARK].Value += "Pattern Not display ";
                                break;
                            }


                            Regex formatRegex_special = new Regex(@"1080i30|1080P59|1080P60");
                            match = formatRegex_special.Match(currentSourceFormat);
                            // notes:
                            // 1) for VGA 1080i30 are Non-corespondence Resolution so catched already at above logic
                            // 2) for VGA 1080P59 and 1080P60  no need OCR.
                            Boolean interfaceTestisVGA = Regex.IsMatch(currentInterface, "VGA", RegexOptions.IgnoreCase);

                            #region 3) Confirm the all picture format is same as Picture Format Info
                            if (match.Success && !interfaceTestisVGA)// extra special case //1080P60 //1080P59 //1080i30 need OCR
                            {
                                //3) Confirm the all picture format is same as Picture Format Info 

                                mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "INFO");//MENU //FORMAT //SOURCE //INFO
                                Thread.Sleep(2000);// enough time to see the picture changing
                                                   // save teh bitmap?
                                SavePictureForReference(currentSourceFormat, "Picture_Format_info");
                                mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "BACK");
                                Thread.Sleep(2000);

                                Result = EnumListUser.ResultQuantumDataTest.PEND;

                            }
                            #endregion //
                        }
                        #endregion //
                    }



                }
                #endregion //

                Thread.Sleep(2000);
            }

            return Result;
        }

        delegate void SavePictureForReferenceDelegate(String additional1, String additional2);

        private void SavePictureForReference(String additional1, String additional2)
        {
            
            if (pictureBox.InvokeRequired)
            {
                pictureBox.BeginInvoke(new SavePictureForReferenceDelegate(SavePictureForReference), additional1, additional2);
            }
            else
            {
            
                String filename = timeStamp.Year.ToString() + "_" + timeStamp.Day.ToString() + "_" + timeStamp.Month.ToString() + "_" + timeStamp.Hour.ToString() + "_" + timeStamp.Minute.ToString() + "_" + timeStamp.Second.ToString();

                if (!System.IO.Directory.Exists("QuantumData"))
                {
                    System.IO.Directory.CreateDirectory("QuantumData");
                }

                //lock (pictureBox)
                //{
                    //pictureBox.InvokeRequired()
                    //pictureBoxPopup.Image.Save("QuantumData\\" + additional1 + "_" + filename + "_" + additional2 + ".jpg",ImageFormat.Jpeg );
                    //pictureBox.Image.Save("QuantumData\\" + additional1 + "_" + filename + "_" + additional2 + "_.jpg", ImageFormat.Jpeg);
                    bm.Save("QuantumData\\" + additional1 + "_" + filename + "_" + additional2 + "_.png", ImageFormat.Png);
            //}
            }
        }

        public RibbonTab EquipmentSetting
        {
            get
            {
                makeAConnection();
                return ribbonEquipmentSetting;
            }
        }

        public string Name
        {
            get
            {
                return "QuantumAllFormat Test";
            }
        }

        // constructor
        public QuantumAllFormat()
        {
           // EquipmentInit();
            PopulateUI();
            PopulateUI_Camera();
            PopulateUI_Extra();

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker_DoWork);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker_ProgressChanged);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = false;

            connectedUART = null;
            ClickEventFlag = dataGridView_CellClick;
        }

        ~QuantumAllFormat()
        {
            Console.WriteLine("QuantumAllFormat Destructor");

            mysaal.UART_ClosePort();
            mysaal.QD882_ClosePort();
        }


        void makeAConnection()
        {
            if (backgroundWorker.IsBusy != true)
            {
                inProgress = 0;
                progressBar.Visible = true;
                buttonTestEquipment.Enabled = false;
                buttonResetEquipment.Visible = false;
                buttonCheckEquipment.Visible = false;

                connectedCOM = null;
                listBoxSignalType.Enabled = false;
                listBoxSignalFormat.Enabled = false;
                listBoxSignalPattern.Enabled = false;

                //buttonUARTWeatherSpooning.Visible = false;

                QD882Label.BackColor = System.Drawing.SystemColors.Control;
                QD882Label.Text = "QD";
                UARTLabel.BackColor = System.Drawing.SystemColors.Control;
                UARTLabel.Text = "UART";
                RemoteIRLabel.BackColor = System.Drawing.SystemColors.Control;
                RemoteIRLabel.Text = "Remote";

                backgroundWorker.RunWorkerAsync("Test");
            }
        }


        delegate void UpdateCurrentSettingUIDelegate(string format, string instrumentResponse);

        void UpdateCurrentSettingUI(string format, string instrumentResponse)
        {

            if (listBoxSignalFormat.InvokeRequired)
            {
                listBoxSignalFormat.Invoke(new UpdateCurrentSettingUIDelegate(UpdateCurrentSettingUI), format, instrumentResponse);
            } 
            else
            {
                listBoxSignalFormat.SelectedIndex = -1;
                ErrorInfoForFormat.Clear();
                listBoxSignalFormat.SelectedItem = format;

                bool contains = Regex.IsMatch(instrumentResponse, "ERROR", RegexOptions.IgnoreCase);
                if (contains)

                {
                    ErrorInfoForFormat.SetError(listBoxSignalFormat, instrumentResponse);
                }
            }
                
        }

        void checkCurrentSetting()
        {
            Match match;
            String returnStr;

            Regex hdmiD = new Regex("3");
            Regex hdmiH = new Regex("4");
            Regex vga = new Regex("9");
            Regex unknown = new Regex(@"Unknown");

            //check for Interface
            returnStr = mysaal.QD882_SendCmd(connectedCOM, "XVSI?");
            match = unknown.Match(returnStr);
            if (match.Success)
            {
                ErrorInfoForInterface.SetError(listBoxSignalType, returnStr);
            }

            Boolean foundInterfaceInList = false;
            match = hdmiD.Match(returnStr);
            if (match.Success)
            {
                Console.WriteLine("This is HDMI-D interface");
                listBoxSignalType.SelectedItem = "HDMI_D";
                foundInterfaceInList = true;
            }
            match = hdmiH.Match(returnStr);
            if (match.Success)
            {
                Console.WriteLine("This is HDMI-H interface");
                listBoxSignalType.SelectedItem = "HDMI_H";
                foundInterfaceInList = true;
            }
            match = vga.Match(returnStr);
            if (match.Success)
            {
                Console.WriteLine("This is HDMI-H interface");
                listBoxSignalType.SelectedItem = "VGA";
                foundInterfaceInList = true;
            }
            if (!foundInterfaceInList)
            {
                ErrorInfoForInterface.SetError(listBoxSignalType, "Interface not in the test list " + returnStr);
            }

            if (listBoxSignalType.SelectedItem != null)
                currentInterface = listBoxSignalType.SelectedItem.ToString();// potential bugs

            //check for Format
            returnStr = mysaal.QD882_SendCmd(connectedCOM, "FMTL?");
            match = unknown.Match(returnStr);
            if (match.Success)
            {
                ErrorInfoForFormat.SetError(listBoxSignalFormat, returnStr);
            }

            Boolean foundFormatInList = false;
            String[] format = EnumListUser.SourceFormat_QD882_FORMAT;
            foreach (string src in format)
            {
                //Regex formatRegex = new Regex(src);
                bool contains = Regex.IsMatch(returnStr, src, RegexOptions.IgnoreCase);
                //match = formatRegex.Match(returnStr);
                if (contains) //(match.Success)
                {
                    Console.WriteLine("This is " + src + " Format");
                    listBoxSignalFormat.SelectedItem = src;
                    foundFormatInList = true;
                }
            }
            if (!foundFormatInList)
            {
                ErrorInfoForFormat.SetError(listBoxSignalFormat, "Format not in the test list " + returnStr);
            }

            // check for Pattern
            returnStr = mysaal.QD882_SendCmd(connectedCOM, "IMGL?");
            match = unknown.Match(returnStr);
            if (match.Success)
            {
                ErrorInfoForPattern.SetError(listBoxSignalPattern, returnStr);
            }

            Boolean foundPaternInList = false;
            String[] pattern = EnumListUser.SourceContent_QD882_FORMAT;
            foreach (string src in pattern)
            {
                //Regex formatRegex = new Regex(src);
                bool contains = Regex.IsMatch(returnStr, src, RegexOptions.IgnoreCase);
                //match = formatRegex.Match(returnStr);
                if (contains) //(match.Success)
                {
                    Console.WriteLine("This is " + src + " Pattern");
                    listBoxSignalPattern.SelectedItem = src;
                    foundPaternInList = true;
                }
            }
            if (!foundPaternInList)
            {
                ErrorInfoForPattern.SetError(listBoxSignalPattern, "Pattern not in the test list "+returnStr);
            }

        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            inProgress = 100;
            Console.WriteLine("ribbonButtonConnectionTest_click Active COM is:[" + connectedCOM + "]");

            connectedUART = mysaal.UART_Test();
            Console.WriteLine("Check the UART doggle is connected or not UART is:[" + connectedUART + "]");

            IR_RemoteConnected = mysaal.IRTRANS_Test();

            if (connectedUART != null)
            {
                mysaal.UART_ClosePort();
                mysaal.UART_Setup(connectedUART, 115200);
                
            }

            if (connectedCOM != null)
            {
                mysaal.QD882_Setup(connectedCOM, 9600);// setup once.

                QD882Label.BackColor = Color.Lime;
                QD882Label.Text = mysaal.QuantumDataLabel+" @" + connectedCOM;

                //check current status... current interface and current format
                checkCurrentSetting();
                buttonResetEquipment.Visible = true;
                buttonCheckEquipment.Visible = true;

                listBoxSignalType.Enabled = true;
                listBoxSignalFormat.Enabled = true;
                listBoxSignalPattern.Enabled = true;
            }
            else
            {
                 QD882Label.BackColor = Color.Red;
                 QD882Label.Text = mysaal.QuantumDataLabel + " FAILED";
                 mysaal.QD882_ClosePort();// close once.
            }
                
            if (connectedUART != null)
            {
                 UARTLabel.BackColor = Color.Lime;
                 UARTLabel.Text = "UART @" + connectedUART;
            }
            else
            {
                 UARTLabel.BackColor = Color.Red;
                 UARTLabel.Text = "UART FAILED";
            }

            if ((IR_RemoteConnected != null) && (mysaal.IRTRANS_Init()))
            {
                RemoteIRLabel.BackColor = Color.Lime;
                RemoteIRLabel.Text = "Remote OK";
            }
            else
            {
                RemoteIRLabel.BackColor = Color.Red;
                RemoteIRLabel.Text = "Remote Failed";
            }
            progressBar.Value = 100;
            progressBar.Visible = false;
            buttonTestEquipment.Enabled = true;
        }

        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;    
        }

        void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (e.Argument.ToString() == "Test")
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                connectedCOM = mysaal.QD882_Test(worker);
            }

            if (e.Argument.ToString() == "Reset")
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                for (int i = 0; i < 5;i++ )
                {
                    Thread.Sleep(1000);
                    worker.ReportProgress(i *20);
                }
                // waited 5 seconds after reset  
            }
        }

        private void ribbonButtonConnectionTest_click(object sender, EventArgs e)
        {     
            //Console.WriteLine(sender.ToString());
            //Console.WriteLine(e.ToString());
            String select = sender.ToString().Remove(0, 35);
            //Console.WriteLine("-" + sender.ToString().Remove(0, 14));
            /*
             *  System.Windows.Forms.Button, Text: Test (35)
                RibbonButton: Test (14)
             * */

            if (select == "Test")
            {
                makeAConnection();
            }

            if (select == "Reset")
            {
                if(connectedCOM ==null)
                {
                    MessageBox.Show("Please check the equipment connection");
                }
                else
                {
                    String returnStr = mysaal.QD882_SendCmd(connectedCOM, "*RST");

                    if (backgroundWorker.IsBusy != true)
                    {
                        inProgress = 0;
                        progressBar.Visible = true;
                        buttonTestEquipment.Enabled = false;
                        buttonResetEquipment.Visible = false;
                        buttonCheckEquipment.Visible = false;

                        backgroundWorker.RunWorkerAsync("Reset");
                    }
                }
            }

            if (select == "Check")
            {
                if (connectedCOM == null)
                {
                    MessageBox.Show("Please check the equipment connection");
                }
                else
                {
                    listBoxSignalFormat.SelectedIndex = -1;
                    listBoxSignalType.SelectedIndex = -1;
                    listBoxSignalPattern.SelectedIndex = -1; 
                    ErrorInfoForInterface.Clear();
                    ErrorInfoForFormat.Clear();
                    ErrorInfoForPattern.Clear();
                    checkCurrentSetting();
                }
            }
            if (select == "Weatherspoon")
            {
                if (connectedUART == null)
                {
                    MessageBox.Show("Please check the UART connection");
                }
                else
                {
                    mysaal.UART_SendCmd("WEATHER_SPOON 1");
                }
            }

            if (select == "IRSend")
            {
                if (IR_RemoteConnected == null)
                {
                    MessageBox.Show("Please check the IR remote connection");
                }
                else
                {
                    mysaal.IRTRANS_SendCMD("PHILIPS_RC5", "MENU");//MENU //FORMAT //SOURCE //INFO
                }
            }
            
        }

        private void ribbonButton_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
            Console.WriteLine(listBoxSignalType.Text + "-" + sender.ToString().Remove(0, 14));

            if (connectedCOM == null)
            {
                MessageBox.Show("Please check the equipment connection");
            }
            else
                mysaal.QD882_SetSource(connectedCOM,listBoxSignalType.Text, sender.ToString().Remove(0, 14));
        }

        private void listBoxSignalType_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString()); //System.Windows.Forms.MouseEventArgs or System.Windows.Forms.KeyPressEventArgs accepted

            listBoxInterfaceCommand();
        }

        private void listBoxSignalTypeKeys_Enter(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString()); //System.Windows.Forms.MouseEventArgs or System.Windows.Forms.KeyPressEventArgs accepted

            if(e.KeyChar == (char)Keys.Return)
            { 
                listBoxInterfaceCommand();                    
            }
        }

        private void listBoxInterfaceCommand()
        {
            if (connectedCOM == null)
            {
                MessageBox.Show("Please check the equipment connection");
            }
            else
            {
                ErrorInfoForInterface.Clear();
                String instrumentResponse = mysaal.QD882_SetSource(connectedCOM, listBoxSignalType.Text);

                bool contains = Regex.IsMatch(instrumentResponse, @"ERROR", RegexOptions.IgnoreCase);
                if (contains) //(match.Success)
                {
                    ErrorInfoForInterface.SetError(listBoxSignalType, instrumentResponse);
                }
            }
        }

        private void listBoxSignalType_mouseHover(object sender, EventArgs e)
        {
            toolTipInfo.SetToolTip(listBoxSignalType, "Double Click or Press Enter to send command for Signal Type");
        }

        private void listBoxSignalFormat_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());

            listBoxFormatCommand();
        }

        private void listBoxSignalFormat_Enter(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString()); //System.Windows.Forms.MouseEventArgs or System.Windows.Forms.KeyPressEventArgs accepted

            if (e.KeyChar == (char)Keys.Return)
            {
                listBoxFormatCommand();
            }
        }

        private void listBoxFormatCommand()
        {
            if (connectedCOM == null)
            {
                MessageBox.Show("Please check the equipment connection");
            }
            else
            {
                ErrorInfoForFormat.Clear();
                String instrumentResponse = mysaal.QD882_SetSourceFormat(connectedCOM, listBoxSignalFormat.Text);

                bool contains = Regex.IsMatch(instrumentResponse, @"ERROR", RegexOptions.IgnoreCase);
                if (contains) //(match.Success)
                {
                    ErrorInfoForFormat.SetError(listBoxSignalFormat, instrumentResponse);
                }
            }
        }

        private void listBoxSignalFormat_mouseHover(object sender, EventArgs e)
        {
            toolTipInfo.SetToolTip(listBoxSignalFormat, "Double Click or Press Enter to send command for Signal Format");
        }

        private void listBoxSignalPattern_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());

            listBoxPatternCommand();
        }

        private void listBoxSignalPattern_Enter(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString()); //System.Windows.Forms.MouseEventArgs or System.Windows.Forms.KeyPressEventArgs accepted

            if (e.KeyChar == (char)Keys.Return)
            {
                listBoxPatternCommand();
            }
        }

        private void listBoxPatternCommand()
        {
            if (connectedCOM == null)
            {
                MessageBox.Show("Please check the equipment connection");
            }
            else
            {
                ErrorInfoForPattern.Clear();
                String instrumentResponse = mysaal.QD882_SetSourcePattern(connectedCOM,listBoxSignalPattern.Text);

                bool contains = Regex.IsMatch(instrumentResponse, @"ERROR", RegexOptions.IgnoreCase);
                if (contains) //(match.Success)
                {
                    ErrorInfoForPattern.SetError(listBoxSignalPattern, instrumentResponse);
                }
            }
        }
        private void listBoxSignalPattern_mouseHover(object sender, EventArgs e)
        {
            toolTipInfo.SetToolTip(listBoxSignalPattern, "Double Click or Press Enter to send command for Pattern Format");
        }

        private void ribbonCameraAndMicrofonExpand_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
        }
        private void pictureBox_Click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
            pictureBox.Size = new System.Drawing.Size(100, 70);
        }
        private void pictureBox_DoubleClick(object sender, EventArgs e)
        {
            popupBox.BringToFront();
            popupBox.WindowState = FormWindowState.Minimized;
            popupBox.Show();
            popupBox.WindowState = FormWindowState.Normal;
        }


        private void camMicButtonOn_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
            //mysaal.videoDisplay.Start();
            

        }

        private void camMicButtonOff_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
            //mysaal.videoDisplay.Stop();
            //_cameraThread.Abort();

        }

        private void CaptureCamera()
        {
            _cameraThread = new Thread(new ThreadStart(CaptureCameraCallback));
            _cameraThread.Start();
        }

        private void CaptureCameraCallback()
        {
            FontAndOverlaySetting();

            eyeDetectionSetting();
            ColorSegmentationSetting();




            CvCapture cap = CvCapture.FromCamera(cbCamListIndex); //max for HD WEBCAM C525 1280 x 720 screen resolution
            Cv.SetCaptureProperty(cap, CaptureProperty.FrameWidth, 1280);
            Cv.SetCaptureProperty(cap, CaptureProperty.FrameHeight, 720);

            //cap.FrameHeight = 1280;
            //cap.FrameWidth = 720;

            while (true)
            {
                imgOri = cap.QueryFrame();
                //imgPreProcess = new IplImage(imgOri.Size, imgOri.Depth, imgOri.NChannels);
                //imgPreProcess = imgOri;
                imgPreProcess = imgOri.Clone();

                //fourEyeDetection();
                ColorSegmentation();

                timeStamp = DateTime.Now;
                String filename = timeStamp.Year.ToString() + "_" + timeStamp.Day.ToString() + "_" + timeStamp.Month.ToString() + "_" + timeStamp.Hour.ToString() + "_" + timeStamp.Minute.ToString() + "_" + timeStamp.Second.ToString();
                imgOri.PutText("Timestamp: " + currentSourceFormat + "_" + filename + "_" , testpoint, fontTimeStamp, CvColor.Green);

                // showing
                bm = BitmapConverter.ToBitmap(imgOri);
                bm.SetResolution(imgOri.Size.Width, imgOri.Size.Height);
                pictureBox.Image = bm;

                bm2 = BitmapConverter.ToBitmap(imgPreProcess);
                bm2.SetResolution(pictureBoxPopup.Width, pictureBoxPopup.Height);

                pictureBoxPopup.Image = bm2;

                imgOri = null;
                imgPreProcess = null;
                //bm = null;
                bm2 = null;

                
            }
        }

        private void FontAndOverlaySetting()
        {
            Cv.InitFont(out font, FontFace.HersheyComplex, 0.5, 0.5);
            Cv.InitFont(out fontTimeStamp, FontFace.HersheyComplex, 0.8, 0.8);
        }

        private void eyeDetectionSetting()
        {
            cascade = CvHaarClassifierCascade.FromFile("haarcascade_eye.xml");
            MinSize = new CvSize(30, 30);
        }

        private void fourEyeDetection()
        {
            CvSeq<CvAvgComp> eyes = Cv.HaarDetectObjects(imgOri, cascade, Cv.CreateMemStorage(), ScaleFactor, MinNeighbors, HaarDetectionType.DoCannyPruning, MinSize);

            int num_pattern_detect = 0;

            foreach (CvAvgComp eye in eyes.AsParallel())
            {
                if (num_pattern_detect == 0)
                {
                    testpoint1.X = (eye.Rect.Location.X) + (eye.Rect.Size.Width / 2);
                    testpoint1.Y = (eye.Rect.Location.Y) + (eye.Rect.Size.Height / 2);
                }

                if (num_pattern_detect == 1)
                {
                    testpoint2.X = (eye.Rect.Location.X) + (eye.Rect.Size.Width / 2);
                    testpoint2.Y = (eye.Rect.Location.Y) + (eye.Rect.Size.Height / 2);
                }

                if (num_pattern_detect == 2)
                {
                    testpoint3.X = (eye.Rect.Location.X) + (eye.Rect.Size.Width / 2);
                    testpoint3.Y = (eye.Rect.Location.Y) + (eye.Rect.Size.Height / 2);
                }

                if (num_pattern_detect == 3)
                {
                    testpoint4.X = (eye.Rect.Location.X) + (eye.Rect.Size.Width / 2);
                    testpoint4.Y = (eye.Rect.Location.Y) + (eye.Rect.Size.Height / 2);
                }
                num_pattern_detect++;
                imgOri.DrawRect(eye.Rect, CvColor.Red, 3);
                testx = (eye.Rect.Location.X) + (eye.Rect.Size.Width / 2);
                testy = (eye.Rect.Location.Y) + (eye.Rect.Size.Height / 2);

            }

            imgOri.DrawCircle(testx, testy, 10, CvColor.Red, 3);// again
            testpoint.X = testx;
            testpoint.Y = testy;

            imgOri.PutText("Marker x of " + num_pattern_detect.ToString(), testpoint, font, CvColor.Red);

            double ratio;
            testpoint.X = testx;
            testpoint.Y = testy + 12;
            if (num_pattern_detect == 4)
            {
                imgOri.PutText("Valid marker !", testpoint, font, CvColor.Green);

                imgOri.DrawLine(testpoint1, testpoint2, CvColor.Green, 3);
                imgOri.DrawLine(testpoint1, testpoint4, CvColor.Green, 3);

                try
                {
                    ratio = ((double)testpoint2.X - (double)testpoint1.X) / ((double)testpoint4.Y - (double)testpoint1.Y);
                }
                catch
                {
                    // div 0
                    ratio = 0.0;
                }

                double ratio_tester1, ratio_tester2;

                ratio_tester1 = ((Math.Abs(ratio) - 1.33) / 1.33) * 100;

                ratio_tester2 = ((Math.Abs(ratio) - 1.78) / 1.78) * 100;

                testpoint.Y = testy + 24;
                imgOri.PutText("Ratio is (" + testpoint2.X.ToString() + " - " + testpoint1.X.ToString() + ") / (" + testpoint4.Y.ToString() + " - " + testpoint1.Y.ToString() + ")", testpoint, font, CvColor.Blue);
                testpoint.Y = testy + 38;
                imgOri.PutText("= " + Math.Abs(ratio).ToString(), testpoint, font, CvColor.Blue);

                /*testpoint.Y = testy + 60;
                if ((Math.Abs(ratio) > 1.6))
                    imgOri.PutText("16:9 Widescreen", testpoint, font, CvColor.Blue);
                else
                    imgOri.PutText("4:3 Standard", testpoint, font, CvColor.Blue);*/
            }
        }

        private void ColorSegmentationSetting()
        {
            testpoint.X = 0; testpoint.Y = 20;
        }
        private void ColorSegmentation()
        {

            CvSeq comp;
            CvMemStorage storage = new CvMemStorage();
            Cv.PyrSegmentation(imgOri, imgPreProcess, storage, out comp,2, 30, 50);//An unhandled exception of type 'OpenCvSharp.OpenCVException' occurred in OpenCvSharp.dll
            //Additional information: Failed to allocate 11313176 bytes

            //public static void PyrSegmentation(IplImage src, IplImage dst, CvMemStorage storage, out CvSeq comp, int level, double threshold1, double threshold2);

            imgPreProcess.PutText("Number Of Color " + comp.Total.ToString(), testpoint, font, CvColor.Red);
            if(comp.Total >15) // heiristic cut off value..
            {
                mysaal.SMPTEBarDisplay = true;
            }
            else
            {
                mysaal.SMPTEBarDisplay = false;
            }
            storage.Clear();
        }

        private void PopulateUI()
        {
            // UI creation on the fly
            ribbonEquipmentSetting = new RibbonTab();
            ribbonEquipmentSetting.Text = "Equipment Setting";


            ConnectionToolsHost = new RibbonHost();
            ConnectionCheckTools = new Panel();
            ConnectionCheckTools.Size = new System.Drawing.Size(300, 68);//155
            buttonTestEquipment = new Button();
            progressBar = new ProgressBar();
            QD882Label = new Label();
            UARTLabel = new Label();
            RemoteIRLabel = new Label();
            buttonResetEquipment = new Button();
            buttonCheckEquipment = new Button();
            buttonUARTWeatherSpooning = new Button();
            buttonIRSend = new Button();

            buttonTestEquipment.Location = new System.Drawing.Point(6, 6);
            buttonTestEquipment.Name = "Test";
            buttonTestEquipment.Size = new System.Drawing.Size(45, 60);
            buttonTestEquipment.TabIndex = 0;
            buttonTestEquipment.Text = "Test";
            buttonTestEquipment.UseVisualStyleBackColor = true;
            buttonTestEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);

            buttonResetEquipment.Location = new System.Drawing.Point(150, 6);
            buttonResetEquipment.Name = "Reset";
            buttonResetEquipment.Size = new System.Drawing.Size(45, 20);
            buttonResetEquipment.TabIndex = 0;
            buttonResetEquipment.Text = "Reset";
            buttonResetEquipment.UseVisualStyleBackColor = true;
            buttonResetEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);

            buttonCheckEquipment.Location = new System.Drawing.Point(200, 6);
            buttonCheckEquipment.Name = "Check";
            buttonCheckEquipment.Size = new System.Drawing.Size(48, 20);
            buttonCheckEquipment.TabIndex = 0;
            buttonCheckEquipment.Text = "Check";
            buttonCheckEquipment.UseVisualStyleBackColor = true;
            buttonCheckEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);


            buttonUARTWeatherSpooning.Location = new System.Drawing.Point(150, 30);
            buttonUARTWeatherSpooning.Name = "Weatherspoon";
            buttonUARTWeatherSpooning.Size = new System.Drawing.Size(120, 20);
            buttonUARTWeatherSpooning.TabIndex = 0;
            buttonUARTWeatherSpooning.Text = "Weatherspoon";
            buttonUARTWeatherSpooning.UseVisualStyleBackColor = true;
            buttonUARTWeatherSpooning.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);
            buttonUARTWeatherSpooning.Visible = true;

            buttonIRSend.Location = new System.Drawing.Point(150, 50);
            buttonIRSend.Name = "IRSend";
            buttonIRSend.Size = new System.Drawing.Size(60, 20);
            buttonIRSend.TabIndex = 0;
            buttonIRSend.Text = "IRSend";
            buttonIRSend.UseVisualStyleBackColor = true;
            buttonIRSend.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);
            buttonIRSend.Visible = true;


            progressBar.Location = new System.Drawing.Point(50, 6);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(100, 60);
            progressBar.TabIndex = 1;
            progressBar.Step = 1;
            progressBar.Visible = false;

            QD882Label.AutoSize = true;     
            QD882Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            QD882Label.Location = new System.Drawing.Point(50, 6);
            QD882Label.Name = "label1";
            QD882Label.Size = new System.Drawing.Size(146, 31);
            QD882Label.TabIndex = 2;
            QD882Label.BackColor = System.Drawing.SystemColors.Control;
            QD882Label.Text = "QD";


            UARTLabel.AutoSize = true;
            UARTLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            UARTLabel.Location = new System.Drawing.Point(50, 38);
            UARTLabel.Name = "label1";
            UARTLabel.Size = new System.Drawing.Size(146, 31);
            UARTLabel.TabIndex = 2;
            UARTLabel.BackColor = System.Drawing.SystemColors.Control;
            UARTLabel.Text = "UART";

            RemoteIRLabel.AutoSize = true;
            RemoteIRLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            RemoteIRLabel.Location = new System.Drawing.Point(50, 52);
            RemoteIRLabel.Name = "label1";
            RemoteIRLabel.Size = new System.Drawing.Size(146, 31);
            RemoteIRLabel.TabIndex = 2;
            RemoteIRLabel.BackColor = System.Drawing.SystemColors.Control;
            RemoteIRLabel.Text = "";

            ConnectionCheckTools.Controls.Add(buttonTestEquipment);
            ConnectionCheckTools.Controls.Add(progressBar);
            ConnectionCheckTools.Controls.Add(QD882Label);
            ConnectionCheckTools.Controls.Add(UARTLabel);
            ConnectionCheckTools.Controls.Add(RemoteIRLabel);
            ConnectionCheckTools.Controls.Add(buttonResetEquipment);
            ConnectionCheckTools.Controls.Add(buttonCheckEquipment);
            ConnectionCheckTools.Controls.Add(buttonUARTWeatherSpooning);
            ConnectionCheckTools.Controls.Add(buttonIRSend);

            RibbonPanel connectionTest = new RibbonPanel();
            connectionTest.Text = "Connection Test";
            connectionTest.Items.Add(ConnectionToolsHost);
            ConnectionToolsHost.HostedControl = ConnectionCheckTools;



            signalTypeHost = new RibbonHost();
            toolTipInfo = new System.Windows.Forms.ToolTip(); 
            listBoxSignalType = new ListBox();
            listBoxSignalType.FormattingEnabled = true;
            listBoxSignalType.Location = new System.Drawing.Point(0, 0);
            listBoxSignalType.Name = "listBoxSignalType";
            listBoxSignalType.Size = new System.Drawing.Size(80, 43);
            listBoxSignalType.TabIndex = 3;
            listBoxSignalType.KeyPress += new KeyPressEventHandler(this.listBoxSignalTypeKeys_Enter);
            listBoxSignalType.DoubleClick += new EventHandler(this.listBoxSignalType_click);
            listBoxSignalType.MouseHover += new EventHandler(this.listBoxSignalType_mouseHover);

            signalFormatHost = new RibbonHost();
            listBoxSignalFormat = new ListBox();
            listBoxSignalFormat.FormattingEnabled = true;
            listBoxSignalFormat.Location = new System.Drawing.Point(0, 0);
            listBoxSignalFormat.Name = "listBoxSignalFormat";
            listBoxSignalFormat.Size = new System.Drawing.Size(120, 43);
            listBoxSignalFormat.TabIndex = 3;
            listBoxSignalFormat.KeyPress += new KeyPressEventHandler(this.listBoxSignalFormat_Enter);
            listBoxSignalFormat.DoubleClick += new EventHandler(this.listBoxSignalFormat_click);
            listBoxSignalFormat.MouseHover += new EventHandler(this.listBoxSignalFormat_mouseHover);


            signalPatternHost = new RibbonHost();
            listBoxSignalPattern = new ListBox();
            listBoxSignalPattern.FormattingEnabled = true;
            listBoxSignalPattern.Location = new System.Drawing.Point(0, 0);
            listBoxSignalPattern.Name = "listBoxSignalPattern";
            listBoxSignalPattern.Size = new System.Drawing.Size(120, 43);
            listBoxSignalPattern.TabIndex = 3;
            listBoxSignalPattern.KeyPress += new KeyPressEventHandler(this.listBoxSignalPattern_Enter);
            listBoxSignalPattern.DoubleClick += new EventHandler(this.listBoxSignalPattern_click);
            listBoxSignalPattern.MouseHover += new EventHandler(this.listBoxSignalPattern_mouseHover);


            RibbonPanel signalType = new RibbonPanel();
            signalType.Text = "Signal Type";

            signalType.Items.Add(signalTypeHost);
            signalTypeHost.HostedControl = listBoxSignalType;


            String[] type = EnumListUser.SourceType_QD882_FORMAT;
            foreach (string src in type)
            {
                listBoxSignalType.Items.Add(src);
            }

            RibbonPanel tvformat = new RibbonPanel();
            tvformat.Text = "TV Format";

            tvformat.Items.Add(signalFormatHost);
            signalFormatHost.HostedControl = listBoxSignalFormat;

            String[] format = EnumListUser.SourceFormat_QD882_FORMAT;
            foreach (string src in format)
            {
                listBoxSignalFormat.Items.Add(src);
            }


            ErrorInfoForInterface = new ErrorProvider();
            ErrorInfoForInterface.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.BlinkIfDifferentError;
            ErrorInfoForFormat = new ErrorProvider();
            ErrorInfoForFormat.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.BlinkIfDifferentError;
            ErrorInfoForPattern = new ErrorProvider();
            ErrorInfoForPattern.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.BlinkIfDifferentError;


            RibbonPanel videopattern = new RibbonPanel();
            videopattern.Text = "Video Pattern";

            videopattern.Items.Add(signalPatternHost);
            signalPatternHost.HostedControl = listBoxSignalPattern;


            String[] pattern = EnumListUser.SourceContent_QD882_FORMAT;
            foreach (string src in pattern)
            {
                listBoxSignalPattern.Items.Add(src);
            }

            ribbonEquipmentSetting.Panels.Add(connectionTest);
            ribbonEquipmentSetting.Panels.Add(signalType);
            ribbonEquipmentSetting.Panels.Add(tvformat);
            ribbonEquipmentSetting.Panels.Add(videopattern);

            /*
            RibbonPanel cameraAndMic = new RibbonPanel();
            cameraAndMic.Text = "Camera And Microfon";
            ribbonEquipmentSetting.Panels.Add(cameraAndMic);

            RibbonPanel previewImageCam = new RibbonPanel();
            ribbonEquipmentSetting.Panels.Add(previewImageCam);
            */

        }

        private void PopulateUI_Camera()
        {
            // Panel for Camera
            RibbonPanel cameraSetting = new RibbonPanel();
            cameraSetting.Text = "Camera Setting";
            ribbonEquipmentSetting.Panels.Add(cameraSetting);

            // Picture Box Settings
            pictureBox = new PictureBox();
            pictureBox.Size = new System.Drawing.Size(100, 50);// preview..
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.DoubleClick += new EventHandler(pictureBox_DoubleClick);
            RibbonHost pictureBoxHost = new RibbonHost();
            pictureBoxHost.HostedControl = pictureBox;
            pictureBoxHost.ToolTip = "Click to enlarge";

            // Popup Picture Box Settings
            pictureBoxPopup = new PictureBox();
            pictureBoxPopup.Size = new System.Drawing.Size(ResWidth, ResHeight);
            pictureBoxPopup.SizeMode = PictureBoxSizeMode.StretchImage;
            popupBox = new Form();
            popupBox.MaximizeBox = false;
            popupBox.ControlBox = true;
            popupBox.FormBorderStyle = FormBorderStyle.FixedSingle;
            popupBox.Size = new System.Drawing.Size(pictureBoxPopup.Size.Width, pictureBoxPopup.Size.Height);
            popupBox.FormClosing += new FormClosingEventHandler(popupBox_FormClosing);
            popupBox.StartPosition = FormStartPosition.CenterScreen;
            popupBox.Controls.Add(pictureBoxPopup);
            pictureBoxPopup.Location = new Point(popupBox.Location.X, popupBox.Location.Y);

            Panel panelCam = new Panel();
            panelCam.Size = new Size(100, 50);
            //panelCam.BackColor = Color.Transparent;

            // On/Off Button Settings
            btOnOffCam = new Button();
            btOnOffCam.Text = "On";
            btOnOffCam.Size = new Size(100, 20);
            btOnOffCam.Enabled = true;
            btOnOffCam.FlatStyle = FlatStyle.Popup;
            btOnOffCam.BackColor = Color.LightBlue;
            btOnOffCam.Click += new EventHandler(btOnOffCam_Click);
            panelCam.Controls.Add(btOnOffCam);

            // Camera List
            cbCamList = new ComboBox();
            cbCamList.Size = new Size(100, 20);
            cbCamList.SelectedIndexChanged += new EventHandler(cbCamList_SelectedIndexChanged);
            cbCamList.Location = new Point(panelCam.Controls[panelCam.Controls.Count - 1].Location.X,
                    panelCam.Controls[panelCam.Controls.Count - 1].Location.Y + btOnOffCam.Height + 5);
            panelCam.Controls.Add(cbCamList);

            RibbonHost panelCamHost = new RibbonHost();
            panelCamHost.HostedControl = panelCam;

            // Add items to Camera Setting Panel
            cameraSetting.Items.Add(pictureBoxHost);
            cameraSetting.Items.Add(panelCamHost);

            initCamera();


        }

        private void PopulateUI_Extra()
        {
            RibbonPanel recordSetting = new RibbonPanel();
            recordSetting.Text = "Record";
            ribbonEquipmentSetting.Panels.Add(recordSetting);

            Panel panelExtra = new Panel();
            panelExtra.Size = new Size(155, 50);

            Label lblTester = new Label();
            lblTester.Text = "Tester :";
            lblTester.TextAlign = ContentAlignment.MiddleRight;
            lblTester.Size = new Size(50, 20);
            panelExtra.Controls.Add(lblTester);

            tbTester = new TextBox();
            tbTester.Size = new Size(100, 20);
            tbTester.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblTester.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(tbTester);

            Label lblVer = new Label();
            lblVer.Text = "Version :";
            lblVer.TextAlign = ContentAlignment.MiddleRight;
            lblVer.Size = new Size(50, 20);
            lblVer.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 2].Location.X,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y + lblTester.Height + 5);
            panelExtra.Controls.Add(lblVer);

            tbVer = new TextBox();
            tbVer.Size = new Size(100, 20);
            tbVer.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblVer.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(tbVer);

            RibbonHost panelExtraHost = new RibbonHost();
            panelExtraHost.HostedControl = panelExtra;

            recordSetting.Items.Add(panelExtraHost);
        }

        private void initCamera()
        {
            string[] cameras = null;

            // Init for Camera
            mysaal.CAM_GetCAMList(out cameras);
            //Saal.CAM_Init(cameras[0]);
            //Saal.videoDisplay.NewFrame += new NewFrameEventHandler(video_NewFrame);

            foreach (string cam in cameras)
            {
                cbCamList.Items.Add(cam);
            }
            cbCamList.SelectedIndex = 0;
            /*
            // Set Camera resolution
            for (int i = 0; i < Saal.videoDisplay.VideoCapabilities.Length; i++)
            {
                //Console.WriteLine(Saal.videoDisplay.VideoCapabilities[i].FrameSize.ToString());
                if (Saal.videoDisplay.VideoCapabilities[i].FrameSize.Height == ResHeight &&
                    Saal.videoDisplay.VideoCapabilities[i].FrameSize.Width == ResWidth)
                {
                    Saal.videoDisplay.VideoResolution = Saal.videoDisplay.VideoCapabilities[i];
                    Console.WriteLine(Saal.videoDisplay.VideoResolution.FrameSize.ToString());
                    break;
                }
            }*/
        }

        private void popupBox_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            popupBox.Hide();
        }

        private void btOnOffCam_Click(object sender, EventArgs e)
        {
            if (btOnOffCam.Text == "On")
            {
                btOnOffCam.BackColor = Color.PaleVioletRed;
                btOnOffCam.Text = "Off";
                CaptureCamera();

                cbCamListIndex = this.cbCamList.SelectedIndex;
                cbCamList.Enabled = false;
            }
            else
            {/*
                if (Saal.videoDisplay.IsRunning)
                {
                    btOnOffCam.BackColor = Color.LightBlue;
                    btOnOffCam.Text = "On";
                    Saal.videoDisplay.SignalToStop();
                    Saal.videoDisplay.WaitForStop();
                    if (Picture.Image != null)
                    {
                        Picture.Image.Dispose();
                        Picture.Image = null;
                    }
                }*/
                btOnOffCam.Text = "On";
                _cameraThread.Abort();
                cbCamList.Enabled = true;
            }
        }

        private void cbCamList_SelectedIndexChanged(object sender, EventArgs e)
        {/*
            if (Saal.videoDisplay.IsRunning)
            {
                Saal.videoDisplay.Stop();
            }

            foreach (FilterInfo device in Saal.videoDevice)
            {
                if (device.Name == cbCamList.SelectedText)
                {
                    Saal.CAM_Init(device.Name);
                    Saal.videoDisplay.NewFrame += new NewFrameEventHandler(video_NewFrame);

                    if (Saal.videoDisplay.IsRunning == false)
                    {
                        Saal.videoDisplay.Start();
                    }
                }
            }*/
        }

        private void InitCamera_Mic()
        {
            // Init for camera box
            string[] cameras = null, microphone = null;

            mysaal.CAM_GetCAMList(out cameras);
            mysaal.CAM_Init(cameras[0]);
            //mysaal.videoDisplay.NewFrame += new NewFrameEventHandler(video_NewFrame);
            /*
            foreach (string cam in cameras)
            {
                Console.WriteLine(cam);
                RibbonLabel listcam = new RibbonLabel();
                listcam.Text = cam;
                cameraList.DropDownItems.Add(listcam);

                //cameraResolution.DropDownItems.Add(listcam);
            }
            //cameraList.SelectedItem = 0;            //To display the camera's name on the cmbCamera combobox
            */
            mysaal.MIC_GetMICList(out microphone);
            mysaal.MIC_Init(microphone[0]);
            /*foreach (string mic in microphone)
            {
                Console.WriteLine(mic);
                RibbonLabel listmic = new RibbonLabel();
                listmic.Text = mic;
                microfon.DropDownItems.Add(listmic);
            }
            //microfon.SelectedItem.SelectedIndex = 0;        //To display the microphone's name on the cmbMicrophone combobox
             */
            
        }
/*
        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Size size = new Size(640, 480);
            Bitmap video = new Bitmap(size.Width, size.Height);

            using (Graphics g = Graphics.FromImage((System.Drawing.Image)video))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage((Bitmap)eventArgs.Frame.Clone(), 0, 0, size.Width, size.Height);
            }

            if (video != null)
            { 
                pictureBox.Image = (Bitmap)video.Clone();
            }

            video.Dispose();
        }
*/
        private void EquipmentInit()
        {
            ManagementObjectCollection ManObjReturn;
            ManagementObjectSearcher ManObjSearch;
            ManObjSearch = new ManagementObjectSearcher("Select * from WIN32_PnPEntity");
            ManObjReturn = ManObjSearch.Get();

            Console.WriteLine("list of equipment....");
            int i = 0;
            foreach (ManagementObject ManObj in ManObjReturn)
            {
                if (ManObj["Manufacturer"] != null)
                {
                    Console.WriteLine(i + "->"+ManObj["Caption"].ToString());
                    i++;
                    if (ManObj["Manufacturer"].ToString().Contains("Quantum Data"))
                    {
                        Console.WriteLine(ManObj["Manufacturer"].ToString());
                        if (ManObj["Caption"] != null)
                        {
                            string[] portnames = SerialPort.GetPortNames();

                            foreach (string port in portnames)
                            {
                                if (ManObj["Caption"].ToString().Contains(port))
                                {
                                    Console.WriteLine(ManObj["Caption"].ToString());
                                    String resultport = Regex.Match(port, @"\d+").Value;
                                    Console.WriteLine(resultport);
                                    Byte portnum = Byte.Parse(resultport);

                                   // IRToySettings = new IrToySettings { ComPort = portnum, UseHandshake = true, UseNotifyOnComplete = true, UseTransmitCount = true };

                                   // IRToyKeySend.Enabled = true;

                                    //connection

                                    break;
                                }
                            }
                        }
                    }
                }
            }

        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Tag = imageFolderPath;

            if (!(e.RowIndex >= 0) || !(e.ColumnIndex >= 0))
                return;

            var datagridview = sender as DataGridView;

            //e.Argument.ToString() ==
            

            Console.WriteLine(e.RowIndex.ToString() + "   " + e.ColumnIndex.ToString());

            //string[] paths = datagridview.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag as string[];
            string[] paths = datagridview.Rows[e.RowIndex].Cells[DEFINE_PICTURE_REF].Tag as string[];

            //Console.WriteLine("1)" + Environment.CurrentDirectory + " " + paths[0].ToString() + "   " + paths[1].ToString());

            //test
            /*
            string path0,path1;
            path0 = "QuantumData";
            path1 = "1035i29";

            //var imageviewer = new PictureViewer(paths[0], paths[1]);
            var imageviewer = new PictureViewer(path0, path1);
            imageviewer.Text = path0;
            imageviewer.Show();
            */
            if (paths == null || paths.Length == 0)
                return;

            Console.WriteLine("2)" + Environment.CurrentDirectory + " "+ paths[0].ToString() + "   " + paths[1].ToString());

            //if (paths[0].ToString().Contains(Environment.CurrentDirectory))
            //{
                var imageviewer = new PictureViewer(paths[0], paths[1]);
                //var imageviewer = new PictureViewer(path0, path1);
                imageviewer.Text = paths[0];
                imageviewer.Show();
            //}
        }
    }
}