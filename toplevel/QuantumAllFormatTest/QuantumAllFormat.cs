using System;
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
using AForge.Video;
using AForge.Video.DirectShow;
using AForge.Imaging;
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

        enum CameraView
        {
            Original,
            Preprocess
        }

        private CameraView camview;

        #region Member: IR
        private ComboBox cbRCFormatList;
        private ComboBox cbRCKeyList;
        private string IrFormat;
        private string IrKey;
        private Panel panelRCFormat;
        private Button btRCKeyTransmit;
        private EnumListUser.IR_RC_FORMAT m_formatRC;
        #endregion


        private TextBox tbWeatherSpoon;


        #region Member: Extra
        private TextBox tbTester;
        private TextBox tbVer;
        private TextBox tbDirectory;
        #endregion

        IplImage imgOri, imgbeforePreProcess, imgPreProcess;
        Bitmap bm,bm2;
        String m_filename;
        
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


        private PictureBox pictureBox;

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
        private Panel pnlEquipment;
        private Button buttonTestEquipment;
        private ProgressBar progressBar;
        private Int32 inProgress;
        private TextBox QD882Label;
        private TextBox UARTLabel;
        private TextBox RemoteIRLabel;
        private Button buttonResetEquipment;
        private Button buttonCheckEquipment;


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

            if ((connectedCOM == null) || (connectedUART==null) || (IR_RemoteConnected == null))
            {
                MessageBox.Show("Please check the equipment connection");

                if (connectedCOM == null)
                    toExecute.Cells[DEFINE_REMARK].Value += " (#ERROR) Device Connection ";

                if (connectedUART == null)
                    toExecute.Cells[DEFINE_REMARK].Value += "UART not connected ";

                if (IR_RemoteConnected == null)
                    toExecute.Cells[DEFINE_REMARK].Value += "IR not connected ";

                Result = EnumListUser.ResultQuantumDataTest.NT;
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

            #region Parse Result

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

                

            }
            #endregion
            SavePictureForReference(currentSourceFormat, "");

            Thread.Sleep(2000);
            BusyFlag = false; 
        }

        private EnumListUser.ResultQuantumDataTest DecideResult(Boolean InstrumentError)
        {
            if (!InstrumentError)
            {
                if (mysaal.NoFreezed && mysaal.NoReboot && mysaal.SMPTEBarDisplay)
                {
                    return EnumListUser.ResultQuantumDataTest.OK;
                }
                else
                {
                    if (!mysaal.NoFreezed)
                        toExecute.Cells[DEFINE_REMARK].Value += " TV freezed ";
                    else if (!mysaal.NoReboot)
                        toExecute.Cells[DEFINE_REMARK].Value += " TV reboot ";
                    else if (!mysaal.SMPTEBarDisplay)
                        toExecute.Cells[DEFINE_REMARK].Value += " Pattern Not display ";

                    

                    return EnumListUser.ResultQuantumDataTest.NG1;
                }
            }
            else
            {
                SavePictureForReference(currentSourceFormat, "Error Command");

                if (mysaal.NoFreezed && mysaal.NoReboot)
                {
                    return EnumListUser.ResultQuantumDataTest.OK;
                }
                else
                {
                    if (!mysaal.NoFreezed)
                        toExecute.Cells[DEFINE_REMARK].Value += " TV freezed ";
                    else if (!mysaal.NoReboot)
                        toExecute.Cells[DEFINE_REMARK].Value += " TV reboot ";

                    return EnumListUser.ResultQuantumDataTest.NG1;
                }

            }
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

            if (!System.IO.Directory.Exists(tbDirectory.Text))
            {
                System.IO.Directory.CreateDirectory(tbDirectory.Text);
            }

            var paths = new[] { tbDirectory.Text, currentSourceFormat };///
            toExecute.Cells[DEFINE_PICTURE_REF].Tag = paths;////

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
                        // sometime purposely sent the invalid command to see if the tv reboot or freeze
                        //Result = EnumListUser.ResultQuantumDataTest.NT;
                        toExecute.Cells[DEFINE_REMARK].Value += instrumentResponse;
                        Result = DecideResult(true);
                        break;
                    }

                    Thread.Sleep(2000);// enough time to see the picture
                                       // send IR command to tv

                    for (int irsend = 0; irsend < 2; irsend++) // to check the TV not freezed
                    {
                        sendIRTrans("MENU");
                        Thread.Sleep(2000);// enough time to see the picture changing
                                           // save the bitmap?
                        SavePictureForReference(currentSourceFormat,"Menu");
                        sendIRTrans("BACK");
                        Thread.Sleep(2000);
                    }
                    Result = DecideResult(false);
                    if (Result == EnumListUser.ResultQuantumDataTest.NG1)
                        break;
                   
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
                        // sometime purposely sent the invalid command to see if the tv reboot or freeze
                        //Result = EnumListUser.ResultQuantumDataTest.NT;
                        toExecute.Cells[DEFINE_REMARK].Value += instrumentResponse;
                        Result = DecideResult(true);
                        break;
                    }

                    // send IR command to tv

                        #region 2) The image is normally displayed.
                        for (int irsend = 0; irsend < 5; irsend++)
                        {
                            //2) The image is normally displayed.

                            sendIRTrans("FORMAT");
                            Thread.Sleep(1000);// need send twice to change. first click only to appear menu second send to change item

                            sendIRTrans("FORMAT");
                            Thread.Sleep(4000);// enough time to see the picture changing
                                               // save teh bitmap?
                            SavePictureForReference(currentSourceFormat, "Picture_Format");
                            sendIRTrans("BACK");
                            Thread.Sleep(2000);

                            Result = DecideResult(false);
                            if (Result == EnumListUser.ResultQuantumDataTest.NG1)
                                break;


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

                                sendIRTrans("INFO");
                                Thread.Sleep(2000);// enough time to see the picture changing
                                                   // save teh bitmap?
                                SavePictureForReference(currentSourceFormat, "Picture_Format_info");
                                sendIRTrans("BACK");
                                Thread.Sleep(2000);

                                Result = EnumListUser.ResultQuantumDataTest.PEND;

                            }
                            #endregion //
                        }
                        #endregion //


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
                bm.Save(tbDirectory.Text + "\\" + additional1 + "_" + m_filename + "_" + additional2 + "_.png", ImageFormat.Png);
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
            PopulateUI_Device();
            PopulateUI_QuantumData();
            //PopulateUI_DebugMsg();
            PopulateUI_IR();
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

                QD882Label.BackColor = System.Drawing.SystemColors.Control;
                QD882Label.Text = "QD";
                UARTLabel.BackColor = System.Drawing.SystemColors.Control;
                UARTLabel.Text = "UART";
                RemoteIRLabel.BackColor = System.Drawing.SystemColors.Control;
                RemoteIRLabel.Text = "Remote";
                pnlEquipment.BackColor = SystemColors.Control;
                buttonTestEquipment.BackColor = Color.LightBlue;

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
                mysaal.UART_Setup(connectedUART, 115200);//  issue when the device not physically connected
                
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
                 pnlEquipment.BackColor = Color.Red;
                 buttonTestEquipment.BackColor = Color.Red;
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
                 pnlEquipment.BackColor = Color.Red;
                 buttonTestEquipment.BackColor = Color.Red;
            }

            if ((IR_RemoteConnected != null) && (mysaal.IRTRANS_Init()))
            {
                RemoteIRLabel.BackColor = Color.Lime;
                RemoteIRLabel.Text = "Remote OK";

                panelRCFormat.Enabled = true;
                btRCKeyTransmit.Enabled = true;
            }
            else
            {
                RemoteIRLabel.BackColor = Color.Red;
                RemoteIRLabel.Text = "Remote Failed";
                pnlEquipment.BackColor = Color.Red;
                buttonTestEquipment.BackColor = Color.Red;
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

                if (connectedUART == null)
                {
                    MessageBox.Show("Please check the UART connection");
                }
                else
                {
                    mysaal.UART_SendCmd("WEATHER_SPOON 1");
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
        }
        private void pictureBox_DoubleClick(object sender, EventArgs e)
        {
            popupBox.BringToFront();
            popupBox.WindowState = FormWindowState.Minimized;
            popupBox.Show();
            popupBox.WindowState = FormWindowState.Normal;

            //switch view
            if (camview == CameraView.Original)
            {
                camview = CameraView.Preprocess;
            }
            else if (camview == CameraView.Preprocess)
            {
                camview = CameraView.Original;
            }
        }


        private void camMicButtonOn_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
        }

        private void camMicButtonOff_click(object sender, EventArgs e)
        {
            Console.WriteLine(sender.ToString());
            Console.WriteLine(e.ToString());
        }

        private void CaptureCamera()
        {
            _cameraThread = new Thread(new ThreadStart(CaptureCameraCallback));
            _cameraThread.Start();
        }

        delegate void UpdatePictureinPopupDelegate(String additional1, String additional2);

        private void UpdatePictureinPopup(String additional1, String additional2)
        {

            if (pictureBoxPopup.InvokeRequired)
            {
                pictureBoxPopup.BeginInvoke(new UpdatePictureinPopupDelegate(UpdatePictureinPopup), additional1, additional2);
            }
            else
            {
                if (ImageHandler.CurrentFrame != null)
                { 
                    var currentImage = new Bitmap(ImageHandler.CurrentFrame);

                    imgOri = new IplImage(currentImage.Width, currentImage.Height, BitDepth.U8, 3);  //creates the OpenCvSharp IplImage;
                    imgOri.CopyFrom(currentImage); // copies the bitmap data to the IplImage

                    CvSize size = new CvSize(currentImage.Width / 2, currentImage.Height / 2);// resize to half

                    imgbeforePreProcess = Cv.CreateImage(size, BitDepth.U8, 3);
                    Cv.Resize(imgOri, imgbeforePreProcess);// color segmentation need lower resolution to meet realtime processing

                    imgPreProcess = imgbeforePreProcess.Clone();

                    ColorSegmentation();


                    timeStamp = DateTime.Now;
                    m_filename = timeStamp.Year.ToString() + "_" + timeStamp.Day.ToString() + "_" + timeStamp.Month.ToString() + "_" + timeStamp.Hour.ToString() + "_" + timeStamp.Minute.ToString() + "_" + timeStamp.Second.ToString();
                    imgOri.PutText("Timestamp: " + currentSourceFormat + "_[" + currentImage.Width + " x " + currentImage.Height + "] " + m_filename + "_", testpoint, fontTimeStamp, CvColor.Green);

                    bm = BitmapConverter.ToBitmap(imgOri);
                    bm.SetResolution(imgOri.Size.Width, imgOri.Size.Height);

                    bm2 = BitmapConverter.ToBitmap(imgPreProcess);
                    bm2.SetResolution(imgPreProcess.Width, imgPreProcess.Height);

                    if (camview == CameraView.Original)
                        pictureBoxPopup.Image = bm;
                    else if (camview == CameraView.Preprocess)
                        pictureBoxPopup.Image = bm2;

                    imgOri = null;
                    imgPreProcess = null;
                    imgbeforePreProcess = null;

                }
            }
        }

        private void CaptureCameraCallback()
        {
            FontAndOverlaySetting();
            ColorSegmentationSetting();

            while (true)
            {
                Thread.Sleep(500);
                UpdatePictureinPopup("", "");
            }
        }

        private void FontAndOverlaySetting()
        {
            Cv.InitFont(out font, FontFace.HersheyComplex, 0.5, 0.5);
            Cv.InitFont(out fontTimeStamp, FontFace.HersheyComplex, 0.8, 0.8);
        }

        private void ColorSegmentationSetting()
        {
            testpoint.X = 0; testpoint.Y = 20;
        }
        private void ColorSegmentation()
        {
            CvSeq comp;
            CvMemStorage storage = new CvMemStorage();
            Cv.PyrSegmentation(imgbeforePreProcess, imgPreProcess, storage, out comp, 2, 30, 50);//An unhandled exception of type 'OpenCvSharp.OpenCVException' occurred in OpenCvSharp.dll
            //Additional information: Failed to allocate 11313176 bytes

            //public static void PyrSegmentation(IplImage src, IplImage dst, CvMemStorage storage, out CvSeq comp, int level, double threshold1, double threshold2);
            imgPreProcess.PutText("[" + imgPreProcess.Width + " x " + imgPreProcess.Height+"]" +" Number Of Color " + comp.Total.ToString(), testpoint, font, CvColor.Red);
            if (comp.Total > 15) // heiristic cut off value..
            {
                mysaal.SMPTEBarDisplay = true;
            }
            else
            {
                mysaal.SMPTEBarDisplay = false;
            }
            storage.Clear();
        }

        private void PopulateUI_Device()
        {
            // UI creation on the fly
            ribbonEquipmentSetting = new RibbonTab();
            ribbonEquipmentSetting.Text = "Equipment Setting";

            buttonTestEquipment = new Button();
            progressBar = new ProgressBar();

            ConnectionCheckTools = new Panel();

            

            pnlEquipment = new Panel();
            pnlEquipment.AutoScroll = true;

            QD882Label = new TextBox();
            UARTLabel = new TextBox();
            RemoteIRLabel = new TextBox();
            buttonResetEquipment = new Button();
            buttonCheckEquipment = new Button();

            QD882Label.ReadOnly = true;
            UARTLabel.ReadOnly = true;
            RemoteIRLabel.ReadOnly = true;

            ConnectionCheckTools.Size = new Size(180, 55);
            buttonTestEquipment.Size = new System.Drawing.Size(50, 55);
            pnlEquipment.Size = new Size(ConnectionCheckTools.Size.Width-buttonTestEquipment.Size.Width, 55);
            progressBar.Size  = new Size(pnlEquipment.Size.Width, pnlEquipment.Size.Height);

            pnlEquipment.Controls.Add(QD882Label);
            pnlEquipment.Controls.Add(buttonResetEquipment);
            pnlEquipment.Controls.Add(buttonCheckEquipment);
            pnlEquipment.Controls.Add(UARTLabel);
            pnlEquipment.Controls.Add(RemoteIRLabel);

            QD882Label.Size = new Size(pnlEquipment.Size.Width - 20, 10);
            buttonResetEquipment.Size = new Size(QD882Label.Size.Width / 2, QD882Label.Size.Height);
            buttonCheckEquipment.Size = new Size(QD882Label.Size.Width / 2, QD882Label.Size.Height);
            UARTLabel.Size = new Size(QD882Label.Size.Width, QD882Label.Size.Height);
            RemoteIRLabel.Size = new Size(QD882Label.Size.Width, QD882Label.Size.Height);

            QD882Label.Location = new Point(0, 0);
            buttonResetEquipment.Location = new Point(QD882Label.Location.X, QD882Label.Location.Y + QD882Label.Size.Height);
            buttonCheckEquipment.Location = new Point(QD882Label.Location.X + buttonResetEquipment.Size.Width, QD882Label.Location.Y + QD882Label.Size.Height);
            UARTLabel.Location = new Point(buttonResetEquipment.Location.X, buttonResetEquipment.Location.Y + QD882Label.Size.Height);
            RemoteIRLabel.Location = new Point(UARTLabel.Location.X, UARTLabel.Location.Y + UARTLabel.Size.Height);

            


            //QD882Label.AutoSize = true;
            QD882Label.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));   
            QD882Label.Name = "label1";       
            QD882Label.TabIndex = 2;
            QD882Label.BackColor = System.Drawing.SystemColors.Control;
            QD882Label.Text = "QD";

            buttonResetEquipment.FlatStyle = FlatStyle.Popup;
            buttonResetEquipment.BackColor = Color.LightBlue;
            buttonResetEquipment.Name = "Reset"; 
            buttonResetEquipment.TabIndex = 0;
            buttonResetEquipment.Text = "Reset";
            buttonResetEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);

            buttonCheckEquipment.FlatStyle = FlatStyle.Popup;
            buttonCheckEquipment.BackColor = Color.LightBlue;
            buttonCheckEquipment.Name = "Check";
            buttonCheckEquipment.TabIndex = 0;
            buttonCheckEquipment.Text = "Check";
            buttonCheckEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);

  
            //UARTLabel.AutoSize = true;
            UARTLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            UARTLabel.Name = "label1";
            UARTLabel.TabIndex = 2;
            UARTLabel.BackColor = System.Drawing.SystemColors.Control;
            UARTLabel.Text = "UART";

            //RemoteIRLabel.AutoSize = true;
            RemoteIRLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            RemoteIRLabel.Name = "label1";
            RemoteIRLabel.TabIndex = 2;
            RemoteIRLabel.BackColor = System.Drawing.SystemColors.Control;
            RemoteIRLabel.Text = "";

            ConnectionToolsHost = new RibbonHost();
            
            buttonTestEquipment.FlatStyle = FlatStyle.Popup;
            buttonTestEquipment.BackColor = Color.LightBlue;
            buttonTestEquipment.UseVisualStyleBackColor = true;
            buttonTestEquipment.Name = "Test";           
            buttonTestEquipment.TabIndex = 0;
            buttonTestEquipment.Text = "Test";  
            buttonTestEquipment.Click += new System.EventHandler(this.ribbonButtonConnectionTest_click);

            buttonTestEquipment.Location = new Point(0, 0);
            pnlEquipment.Location = new Point(buttonTestEquipment.Location.X + buttonTestEquipment.Size.Width, buttonTestEquipment.Location.Y);
            progressBar.Location  = new Point(buttonTestEquipment.Location.X + buttonTestEquipment.Size.Width, buttonTestEquipment.Location.Y);

            progressBar.Name = "progressBar";   
            progressBar.TabIndex = 1;
            progressBar.Step = 1;
            progressBar.Visible = false;

            ConnectionCheckTools.Controls.Add(buttonTestEquipment);
            ConnectionCheckTools.Controls.Add(progressBar);
            ConnectionCheckTools.Controls.Add(pnlEquipment);

            RibbonPanel connectionTest = new RibbonPanel();
            connectionTest.Text = "Connection Test";
            connectionTest.Items.Add(ConnectionToolsHost);
            ConnectionToolsHost.HostedControl = ConnectionCheckTools;

            ribbonEquipmentSetting.Panels.Add(connectionTest);
        }

        private void PopulateUI_QuantumData()
        {
            signalTypeHost = new RibbonHost();
            toolTipInfo = new System.Windows.Forms.ToolTip();
            listBoxSignalType = new ListBox();
            listBoxSignalType.FormattingEnabled = true;
            listBoxSignalType.Location = new System.Drawing.Point(0, 0);
            listBoxSignalType.Name = "listBoxSignalType";
            listBoxSignalType.Size = new System.Drawing.Size(60, 43);
            listBoxSignalType.TabIndex = 3;
            listBoxSignalType.KeyPress += new KeyPressEventHandler(this.listBoxSignalTypeKeys_Enter);
            listBoxSignalType.DoubleClick += new EventHandler(this.listBoxSignalType_click);
            listBoxSignalType.MouseHover += new EventHandler(this.listBoxSignalType_mouseHover);

            signalFormatHost = new RibbonHost();
            listBoxSignalFormat = new ListBox();
            listBoxSignalFormat.FormattingEnabled = true;
            listBoxSignalFormat.Location = new System.Drawing.Point(0, 0);
            listBoxSignalFormat.Name = "listBoxSignalFormat";
            listBoxSignalFormat.Size = new System.Drawing.Size(80, 43);
            listBoxSignalFormat.TabIndex = 3;
            listBoxSignalFormat.KeyPress += new KeyPressEventHandler(this.listBoxSignalFormat_Enter);
            listBoxSignalFormat.DoubleClick += new EventHandler(this.listBoxSignalFormat_click);
            listBoxSignalFormat.MouseHover += new EventHandler(this.listBoxSignalFormat_mouseHover);


            signalPatternHost = new RibbonHost();
            listBoxSignalPattern = new ListBox();
            listBoxSignalPattern.FormattingEnabled = true;
            listBoxSignalPattern.Location = new System.Drawing.Point(0, 0);
            listBoxSignalPattern.Name = "listBoxSignalPattern";
            listBoxSignalPattern.Size = new System.Drawing.Size(80, 43);
            listBoxSignalPattern.TabIndex = 3;
            listBoxSignalPattern.KeyPress += new KeyPressEventHandler(this.listBoxSignalPattern_Enter);
            listBoxSignalPattern.DoubleClick += new EventHandler(this.listBoxSignalPattern_click);
            listBoxSignalPattern.MouseHover += new EventHandler(this.listBoxSignalPattern_mouseHover);


            RibbonPanel signalType = new RibbonPanel();
            signalType.Text = "Interface";

            signalType.Items.Add(signalTypeHost);
            signalTypeHost.HostedControl = listBoxSignalType;


            String[] type = EnumListUser.SourceType_QD882_FORMAT;
            foreach (string src in type)
            {
                listBoxSignalType.Items.Add(src);
            }

            RibbonPanel tvformat = new RibbonPanel();
            tvformat.Text = "Source";

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
            videopattern.Text = "Content";

            videopattern.Items.Add(signalPatternHost);
            signalPatternHost.HostedControl = listBoxSignalPattern;


            String[] pattern = EnumListUser.SourceContent_QD882_FORMAT;
            foreach (string src in pattern)
            {
                listBoxSignalPattern.Items.Add(src);
            }

            
            ribbonEquipmentSetting.Panels.Add(signalType);
            ribbonEquipmentSetting.Panels.Add(tvformat);
            ribbonEquipmentSetting.Panels.Add(videopattern);

        }

        private void PopulateUI_IR()
        {
            RibbonPanel IRSetting = new RibbonPanel();
            IRSetting.Text = "IR Setting";
            ribbonEquipmentSetting.Panels.Add(IRSetting);

            panelRCFormat = new Panel();
            panelRCFormat.Size = new Size(100, 50);

            // RC Format List
            cbRCFormatList = new ComboBox();
            cbRCFormatList.Size = new Size(100, 20);
            cbRCFormatList.SelectedIndexChanged += new EventHandler(cbRCFormatList_SelectedIndexChanged);
            foreach (string fmt in EnumListUser.RCFORMAT)
            {
                cbRCFormatList.Items.Add(fmt);
            }
            panelRCFormat.Controls.Add(cbRCFormatList);

            // RC Format List
            cbRCKeyList = new ComboBox();
            cbRCKeyList.Size = new Size(100, 20);
            cbRCKeyList.SelectedIndexChanged += new EventHandler(cbRCKeyList_SelectedIndexChanged);
            cbRCKeyList.Location = new Point(panelRCFormat.Controls[panelRCFormat.Controls.Count - 1].Location.X,
                    panelRCFormat.Controls[panelRCFormat.Controls.Count - 1].Location.Y + cbRCFormatList.Height + 5);

            cbRCFormatList.SelectedIndex = (int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC5;// initialized
            foreach (string key in EnumListUser.KEY_PHILIPS_RC5)
            {
                cbRCKeyList.Items.Add(key);
            } 
            panelRCFormat.Controls.Add(cbRCKeyList);

            RibbonHost panelRCFormatHost = new RibbonHost();
            panelRCFormatHost.HostedControl = panelRCFormat;

            // Send Button Settings
            btRCKeyTransmit = new Button();
            btRCKeyTransmit.Text = "Send";
            btRCKeyTransmit.Size = new Size(50, 55);
            btRCKeyTransmit.Enabled = true;
            btRCKeyTransmit.FlatStyle = FlatStyle.Popup;
            btRCKeyTransmit.BackColor = Color.LightBlue;
            btRCKeyTransmit.Click += new EventHandler(btRCKeyTransmit_Click);

            RibbonHost btRCTransmitHost = new RibbonHost();
            btRCTransmitHost.HostedControl = btRCKeyTransmit;

            IRSetting.Items.Add(panelRCFormatHost);
            IRSetting.Items.Add(btRCTransmitHost);


            panelRCFormat.Enabled = false;
            btRCKeyTransmit.Enabled = false;

        }

        private void PopulateUI_DebugMsg()
        {
            RibbonPanel weatherSpoon = new RibbonPanel();
            weatherSpoon.Text = "Weather Spoon";
            ribbonEquipmentSetting.Panels.Add(weatherSpoon);

            Panel panelExtra = new Panel();
            panelExtra.Size = new Size(155, 50);

            tbWeatherSpoon = new TextBox();
            tbWeatherSpoon.Size = new Size(panelExtra.Size.Width, panelExtra.Size.Height);
            tbWeatherSpoon.Multiline = true;
            panelExtra.Controls.Add(tbWeatherSpoon);

            RibbonHost panelExtraHost = new RibbonHost();
            panelExtraHost.HostedControl = panelExtra;

            weatherSpoon.Items.Add(panelExtraHost);
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
            pictureBox.Click += new EventHandler(pictureBox_Click);
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
            panelExtra.Size = new Size(185, 50);
            panelExtra.AutoScroll = true;

            Label lblTester = new Label();
            lblTester.Text = "Tester :";
            lblTester.TextAlign = ContentAlignment.MiddleRight;
            lblTester.Size = new Size(60, 20);
            panelExtra.Controls.Add(lblTester);

            tbTester = new TextBox();
            tbTester.Size = new Size(100, 20);
            tbTester.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblTester.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(tbTester);

            Label lblVer = new Label();
            lblVer.Text = "Version :";
            lblVer.TextAlign = ContentAlignment.MiddleRight;
            lblVer.Size = new Size(60, 20);
            lblVer.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 2].Location.X,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y + lblTester.Height + 5);
            panelExtra.Controls.Add(lblVer);

            tbVer = new TextBox();
            tbVer.Size = new Size(100, 20);
            tbVer.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblVer.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(tbVer);


            Label lblDirectory = new Label();
            lblDirectory.Text = "Directory :";
            lblDirectory.TextAlign = ContentAlignment.MiddleRight;
            lblDirectory.Size = new Size(60, 20);
            lblDirectory.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 2].Location.X,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y + tbVer.Height + 5);
            panelExtra.Controls.Add(lblDirectory);

            tbDirectory = new TextBox();
            tbDirectory.Size = new Size(100, 20);
            tbDirectory.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblDirectory.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(tbDirectory);

            tbDirectory.Text = "QuantumData";

            RibbonHost panelExtraHost = new RibbonHost();
            panelExtraHost.HostedControl = panelExtra;

            recordSetting.Items.Add(panelExtraHost);
        }


        private void initCamera()
        {
            string[] cameras = null;

            // Init for Camera
            mysaal.CAM_GetCAMList(out cameras);
            if (cameras.Length > 0)
            {
                ImageHandler.CamName = cameras[0]; // set camera device
                ImageHandler.ImgFromParent = Picture; // send picture box control
                ImageHandler.ResSize.Width = ResWidth;
                ImageHandler.ResSize.Height = ResHeight;
                //var imageGuide = new Bitmap(ResWidth, ResHeight);
                //imageGuide.Tag = "QuantumAllFormat";

                //ImageHandler.gridBitmap = imageGuide;
                ImageHandler.InitCamera(); // start camera

                foreach (string cam in cameras)
                {
                    cbCamList.Items.Add(cam);
                }
                cbCamList.SelectedIndex = 0;
            }
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
                ImageHandler.StartCamera();
                CaptureCamera();
            }
            else
            {
                btOnOffCam.BackColor = Color.LightBlue;
                btOnOffCam.Text = "On";
                ImageHandler.StopCamera();
                if (Picture.Image != null)
                {
                    Picture.Image.Dispose();
                    Picture.Image = null;
                }
                _cameraThread.Abort();
            }
        }

        private void cbCamList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImageHandler.CamName = cbCamList.Text;
            ImageHandler.ImgFromParent = Picture;
            ImageHandler.InitCamera();

            btOnOffCam.BackColor = Color.LightBlue;
            btOnOffCam.Text = "On";
            if (Picture.Image != null)
            {
                Picture.Image.Dispose();
                Picture.Image = null;
            }
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (!(e.RowIndex >= 0) || !(e.ColumnIndex >= 0))
                return;

            var datagridview = sender as DataGridView;
            
            Console.WriteLine(e.RowIndex.ToString() + "   " + e.ColumnIndex.ToString());
            string[] paths = datagridview.Rows[e.RowIndex].Cells[DEFINE_PICTURE_REF].Tag as string[];

            if (paths == null || paths.Length == 0)
                return;

            Console.WriteLine("2)" + Environment.CurrentDirectory + " "+ paths[0].ToString() + "   " + paths[1].ToString());

            var imageviewer = new PictureViewer(paths[0], paths[1]);
            imageviewer.Text = paths[0];
            imageviewer.Show();
        }

        private void cbRCFormatList_SelectedIndexChanged(object sender, EventArgs e)
        {
            IrFormat = cbRCFormatList.Text;
            cbRCKeyList.Items.Clear();
            
            if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC5)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.PHILIPS_RC5;
                foreach (string key in EnumListUser.KEY_PHILIPS_RC5)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC6)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.PHILIPS_RC6;
                foreach (string key in EnumListUser.KEY_PHILIPS_RC6)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.Matsushita)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.Matsushita;
                foreach (string key in EnumListUser.KEY_MATSUSHITA)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC_THAI)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.NEC_THAI;
                foreach (string key in EnumListUser.KEY_NEC_THAI)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC_INDIA)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.NEC_THAI;
                foreach (string key in EnumListUser.KEY_NEC_INDIA)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.NEC;
                foreach (string key in EnumListUser.KEY_NEC)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.SANYO)
            {
                m_formatRC = EnumListUser.IR_RC_FORMAT.SANYO;
                foreach (string key in EnumListUser.KEY_SANYO)
                {
                    cbRCKeyList.Items.Add(key);
                }
            }
            
            cbRCKeyList.SelectedIndex = 0;
        }

        private void cbRCKeyList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC5)
            {
                IrKey = EnumListUser.KEY_PHILIPS_RC5[cbRCKeyList.SelectedIndex];
            }

            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC6)
            {
                IrKey = EnumListUser.KEY_PHILIPS_RC6[cbRCKeyList.SelectedIndex];
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.Matsushita)
            {
                IrKey = EnumListUser.KEY_MATSUSHITA[cbRCKeyList.SelectedIndex];
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC_THAI)
            {
                IrKey = EnumListUser.KEY_NEC_THAI[cbRCKeyList.SelectedIndex];
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC_INDIA)
            {
                IrKey = EnumListUser.KEY_NEC_INDIA[cbRCKeyList.SelectedIndex];
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.NEC)
            {
                IrKey = EnumListUser.KEY_NEC[cbRCKeyList.SelectedIndex];
            }
            else if (cbRCFormatList.SelectedIndex == (int)EnumListUser.IR_RC_FORMAT.SANYO)
            {
                IrKey = EnumListUser.KEY_SANYO[cbRCKeyList.SelectedIndex];
            }

            
        }

        private void btRCKeyTransmit_Click(object sender, EventArgs e)
        {
            if (IR_RemoteConnected == null)
            {
                MessageBox.Show("Please check the IR remote connection");
            }
            else
            {
                //IrFormat = EnumListUser.RCFORMAT[(int)EnumListUser.IR_RC_FORMAT.PHILIPS_RC5];
                //IrKey = EnumListUser.KEY_PHILIPS_RC5[(int)EnumListUser.IR_PHILIP_RC5.MENU];
                Console.WriteLine(IrFormat + "  " + IrKey);
                mysaal.IRTRANS_SendCMD(IrFormat, IrKey);
            }
        }

        delegate void UpdateWeatherSpooningUIDelegate(string info);

        void UpdateWeatherSpooningUI(string info)
        {

            if (tbWeatherSpoon.InvokeRequired)
            {
                tbWeatherSpoon.Invoke(new UpdateWeatherSpooningUIDelegate(UpdateWeatherSpooningUI), info);
            }
            else
            {
                tbWeatherSpoon.Text = info;
            }

        }

        // each remote have differet key name event it's refering to the same purpose
        private void sendIRTrans(String Key)
        {
            switch (Key)
            {
                case "MENU":
                    switch (m_formatRC)
                    {
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC5:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC5[(int)EnumListUser.IR_PHILIP_RC5.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC6:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC6[(int)EnumListUser.IR_PHILIP_RC6.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.Matsushita:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_MATSUSHITA[(int)EnumListUser.IR_MATSUSHITA.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_THAI:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_THAI[(int)EnumListUser.IR_NEC_THAI.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_INDIA:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_INDIA[(int)EnumListUser.IR_NEC_INDIA.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC[(int)EnumListUser.IR_NEC.MENU]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.SANYO:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_SANYO[(int)EnumListUser.IR_SANYO.MENU]);
                            break;
                    }
                    break;

                case "FORMAT":
                    switch (m_formatRC)
                    {
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC5:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC5[(int)EnumListUser.IR_PHILIP_RC5.FORMAT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC6:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC6[(int)EnumListUser.IR_PHILIP_RC6.FORMAT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.Matsushita:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_MATSUSHITA[(int)EnumListUser.IR_MATSUSHITA.FORMAT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_THAI:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_THAI[(int)EnumListUser.IR_NEC_THAI.ASPECT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_INDIA:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_INDIA[(int)EnumListUser.IR_NEC_INDIA.ASPECT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC[(int)EnumListUser.IR_NEC.ASPECT]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.SANYO:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_SANYO[(int)EnumListUser.IR_SANYO.PIXSHAPE]);
                            break;
                    }
                    break;

                case "BACK":
                    switch (m_formatRC)
                    {
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC5:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC5[(int)EnumListUser.IR_PHILIP_RC5.BACK]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC6:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC6[(int)EnumListUser.IR_PHILIP_RC6.BACK]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.Matsushita:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_MATSUSHITA[(int)EnumListUser.IR_MATSUSHITA.BACK]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_THAI:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_THAI[(int)EnumListUser.IR_NEC_THAI.EXIT]);//assume
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_INDIA:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_INDIA[(int)EnumListUser.IR_NEC_INDIA.EXIT]);//assume
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC[(int)EnumListUser.IR_NEC.EXIT]);//assume
                            break;
                        case EnumListUser.IR_RC_FORMAT.SANYO:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_SANYO[(int)EnumListUser.IR_SANYO.BACK]);
                            break;
                    }
                    break;

                case "INFO":
                    switch (m_formatRC)
                    {
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC5:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC5[(int)EnumListUser.IR_PHILIP_RC5.INFO]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.PHILIPS_RC6:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_PHILIPS_RC6[(int)EnumListUser.IR_PHILIP_RC6.INFO]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.Matsushita:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_MATSUSHITA[(int)EnumListUser.IR_MATSUSHITA.INFO]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_THAI:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_THAI[(int)EnumListUser.IR_NEC_THAI.INFO]);
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC_INDIA:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC_INDIA[(int)EnumListUser.IR_NEC_INDIA.DISPLAY]);//assume
                            break;
                        case EnumListUser.IR_RC_FORMAT.NEC:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_NEC[(int)EnumListUser.IR_NEC.DISPLAY]);//assume
                            break;
                        case EnumListUser.IR_RC_FORMAT.SANYO:
                            mysaal.IRTRANS_SendCMD(IrFormat, EnumListUser.KEY_SANYO[(int)EnumListUser.IR_SANYO.INFO]);
                            break;
                    }
                    break;
            }

        }
    }
}
