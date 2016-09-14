using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginContracts;
using System.Windows.Forms;
using System.Data;
using System.Timers;
using System.Text.RegularExpressions;
using System.IO.Ports;
using mainUI;
using SAAL;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.OCR;
using Emgu.Util;
using Emgu.CV.UI;

#region Camera include files
using System.Drawing;
using System.Drawing.Imaging;
using AForge.Video;
using AForge.Video.DirectShow;
using AForge.Imaging;
using System.ComponentModel;
using System.Threading;
using System.IO;
#endregion

namespace Test
{
    public class ExtInput : IPlugin
    {
        SAAL_Interface Saal = new SAAL_Interface();

        #region Member: Generic
        private RibbonTab ribbonEquipmentSetting;
        private DataGridViewRow toExecute;
        private PictureBox pictureBox;
        private bool BusyFlag;
        private Dictionary<int, string> _validateDict;
        private BackgroundWorker bgWorker;
        private BackgroundWorker bgWorkerFromTopForm;
        private System.Timers.Timer _timer;
        private EN_TEST_MODE test_mode;
        private Mat frame_cv;
        private string IrFormat;
        private EN_INPUT_LIST cur_input = EN_INPUT_LIST.EN_INPUT_MAX;
        private string _executionTime;
        private Bitmap NGImage = null;
        private Bitmap OKImage = null;
        private int noOfRow;
        private int noOfRowTemp;
        private double elapsedTime;

        // Change to true when running specific test
        private Boolean testUARTBypass = false;
        private Boolean testImageBypass = false;

        private enum EN_ID_TIMER
        {
            ID_TIMER_CLEAR_DICT,
            ID_TIMER_CONFIG_MENU,
            ID_TIMER_IR_MENU,
            ID_TIMER_OPERATION,
            ID_TIMER_MAX
        }

        private enum EN_TEST_MODE
        {
            TEST_MODE_1,
            TEST_MODE_2,
            TEST_MODE_MAX
        }

        private enum EN_INPUT_LIST
        {
            EN_INPUT_CAST,
            EN_INPUT_TV,
            EN_INPUT_HDMI1,
            EN_INPUT_HDMI2,
            EN_INPUT_HDMI3,
            EN_INPUT_HDMI4,
            EN_INPUT_VIDEO,
            EN_INPUT_PC,
            EN_INPUT_USB,
            EN_INPUT_MAX
        }
        #endregion

        #region Member: Device Settings
        private DEVICE_SETTINGS[] _deviceSettings;
        private Int32 deviceNum = 3;
        private EN_CON_METHOD con_state;
        private CheckBox[] cbEquipment;
        private Button btTestCon;
        private Form frmConfig;
        private RadioButton rbConfigUART, rbConfigLAN;
        private TextBox tbConfigUART, tbConfigLAN;
        private Label lblConfig;

        private struct DEVICE_SETTINGS
        {
            public String DevName; // eg) TG45, TG59
            public String PortID; // eg) COM1 / 192.168.1.0
            public EN_CON_METHOD Con_Method; // EN_CON_METHOD (Uart/Lan)
            public Boolean Con_Status; // 1 : Connected, 0 : Not Connected
        }

        private enum EN_CON_METHOD
        {
            CON_UART,
            CON_LAN,
            CON_MAX
        }
        #endregion

        #region Member: IR
        private ComboBox cbRCFormatList;
        private ComboBox cbRCKeyList;
        private Form frmIR;
        private RadioButton rbIRmodeAUTO, rbIRmodeMAN;
        private TextBox tbIRport;
        private EN_IR_PORT ir_port = EN_IR_PORT.IR_PORT_MAX;

        private enum EN_IR_PORT
        {
            IR_PORT_AUTO,
            IR_PORT_MANUAL,
            IR_PORT_MAX
        }
        #endregion

        #region Member: Extra
        private TextBox tbTester;
        private TextBox tbVer;
        #endregion

        #region Member: Camera
        // Parameters for Camera
        private ComboBox cbCamList;
        private Button btOnOffCam;
        private Button btCamCalib;
        #endregion      

        #region Member: Public properties
        public string Name { get { return "ExtInput Test"; } }
        public DataGridViewRow ToExecute { get { return toExecute; } set { toExecute = value; DataRowProcess(); } }
        public RibbonTab EquipmentSetting { get { init(); return ribbonEquipmentSetting; } }
        public PictureBox Picture { get { return pictureBox; } set { pictureBox = value; } }
        public bool Busy { get { return BusyFlag; } }

        public bool DoExecute { get; private set; } = true;
        public BeforeExecute BeforeExecute { get; }
        public AfterExecute AfterExecute { get; }
        public CellClickEvent ClickEvent { get; }
        public BeforeExecuteButAfterSelection BeforeExecuteButAfterSelection { get; }
        public ProgressBarChangedEvent ProgressBarChangedEvent { get; }

        #endregion

        // Constructor
        public ExtInput()
        {
            BeforeExecute = InvokeBeforeExecute;
            AfterExecute = InvokeAfterExecute;
            ClickEvent = DataGridView_CellClick;
            BeforeExecuteButAfterSelection = InvokeBeforeExecuteButAfterSelection;
            ProgressBarChangedEvent = InvokeProgressBarChangedEvent;
        }

        #region init
        private void init()
        {
            _deviceSettings = new DEVICE_SETTINGS[deviceNum];
            initDevice();

            // Add input source column number to validate
            _validateDict = new Dictionary<int, string>();
            _validateDict.Add(1, ""); // Input
            _validateDict.Add(3, ""); // Signal Generator
            _validateDict.Add(4, ""); // Format Name
            _validateDict.Add(5, ""); // Input Resolution
            _validateDict.Add(6, ""); // Operation

            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.DoWork += new DoWorkEventHandler(this.bgWorker_DoWork);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(this.bgWorker_ProgressChanged);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.bgWorker_RunWorkerCompleted);
            bgWorker.WorkerSupportsCancellation = true;

            PopulateUI();

            test_mode = EN_TEST_MODE.TEST_MODE_MAX;
            con_state = EN_CON_METHOD.CON_MAX;

            //TestConnection();
            BusyFlag = false;

            // Instantiate timer
            //_timer = new System.Timers.Timer();

            ImageHandler.IsEnOCR = true;
            ImageHandler.IsEnBlurryCheck = true;
            ImageHandler.IsEnSimilarityCheck = true;
            ImageHandler.IsEnPictureFormat = false;
        }

        private void initDevice()
        {
            // Add device lists for this application
            string[] _devName = { "TG59", "TG45", "TG35" };

            for (int i = 0; i < deviceNum; i++)
            {
                _deviceSettings[i].DevName = _devName[i];
                _deviceSettings[i].PortID = "";
                _deviceSettings[i].Con_Method = EN_CON_METHOD.CON_UART; // default setting
                _deviceSettings[i].Con_Status = false;
            }
        }

        private void initCamera()
        {
            string[] cameras = null;
            frame_cv = new Mat();

            // Init for Camera
            Saal.CAM_GetCAMList(out cameras);
            if (cameras.Length > 0)
            {
                ImageHandler.CamName = cameras[0]; // set camera device
                ImageHandler.ImgFromParent = pictureBox; // send picture box control
                ImageHandler.InitCamera(); // start camera

                foreach (string cam in cameras)
                {
                    cbCamList.Items.Add(cam);
                }
                cbCamList.SelectedIndex = 0;
            }
        }
        #endregion

        #region Populate UI
        private void PopulateUI()
        {
            ribbonEquipmentSetting = new RibbonTab();
            ribbonEquipmentSetting.Text = "Equipment Setting";

            PopulateUI_Device();
            PopulateUI_IR();
            PopulateUI_Camera();
            PopulateUI_Extra();
        }

        private void PopulateUI_Device()
        {
            RibbonPanel connectionTest = new RibbonPanel();
            connectionTest.Text = "Connection Test";
            ribbonEquipmentSetting.Panels.Add(connectionTest);

            // Get equipment lists            
            Panel pnlEquipment = new Panel();
            pnlEquipment.Size = new Size(100, 55);
            pnlEquipment.AutoScroll = true;
            //pnlEquipment.BackColor = Color.Transparent;

            cbEquipment = new CheckBox[_deviceSettings.Count()];

            Int32 index = 0;
            for (int i = 0; i < deviceNum; i++)
            {
                cbEquipment[index] = new CheckBox();
                cbEquipment[index].Text = _deviceSettings[i].DevName;
                cbEquipment[index].AutoCheck = false;
                cbEquipment[index].Appearance = Appearance.Button;
                cbEquipment[index].TextAlign = ContentAlignment.MiddleLeft;
                cbEquipment[index].AutoSize = false;
                cbEquipment[index].Size = new Size(80, 20);
                cbEquipment[index].MouseEnter += new EventHandler(cbEquipment_MouseEnter);
                cbEquipment[index].MouseLeave += new EventHandler(cbEquipment_MouseLeave);

                if (pnlEquipment.Controls.Count == 0)
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                else
                {
                    cbEquipment[index].Location = new Point(pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.X,
                        pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.Y + cbEquipment[index].Height);
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                }
                index++;
            }

            RibbonHost pnlEquipmentHost = new RibbonHost();
            pnlEquipmentHost.HostedControl = pnlEquipment;
            connectionTest.Items.Add(pnlEquipmentHost);

            btTestCon = new Button();
            btTestCon.Text = "Test";
            btTestCon.Size = new Size(50, 55);
            btTestCon.FlatStyle = FlatStyle.Popup;
            btTestCon.BackColor = Color.LightBlue;
            btTestCon.Click += new EventHandler(btTestCon_Click);
            RibbonHost btTestConHost = new RibbonHost();
            btTestConHost.HostedControl = btTestCon;
            connectionTest.Items.Add(btTestConHost);

            //Config Form
            frmConfig = new Form();
            frmConfig.Size = new Size(200, 100);
            frmConfig.FormBorderStyle = FormBorderStyle.None;
            frmConfig.MouseLeave += new EventHandler(frmConfig_MouseLeave);
            frmConfig.MouseEnter += new EventHandler(frmConfig_MouseEnter);
            frmConfig.BackColor = System.Drawing.Color.LightYellow;

            lblConfig = new Label();
            lblConfig.Text = "Device Settings For ";
            lblConfig.Location = new Point(2, 2);
            lblConfig.Size = new Size(200, 20);

            rbConfigUART = new RadioButton();
            rbConfigUART.Text = "UART";
            rbConfigUART.Location = new Point(2, lblConfig.Location.Y + lblConfig.Height);
            rbConfigUART.CheckedChanged += new EventHandler(rbConfig_CheckedChanged);
            rbConfigUART.Size = new Size(70, rbConfigUART.Height);
            rbConfigUART.Checked = true;
            tbConfigUART = new TextBox();
            tbConfigUART.Location = new Point(rbConfigUART.Location.X + rbConfigUART.Width + 1, rbConfigUART.Location.Y);
            tbConfigUART.ReadOnly = true;
            tbConfigUART.Enabled = false;

            rbConfigLAN = new RadioButton();
            rbConfigLAN.Text = "LAN";
            rbConfigLAN.Location = new Point(rbConfigUART.Location.X, rbConfigUART.Location.Y + rbConfigUART.Height + 1);
            rbConfigLAN.CheckedChanged += new EventHandler(rbConfig_CheckedChanged);
            rbConfigLAN.Size = new Size(70, rbConfigLAN.Height);
            tbConfigLAN = new TextBox();
            tbConfigLAN.Location = new Point(rbConfigLAN.Location.X + rbConfigLAN.Width + 1, rbConfigLAN.Location.Y);
            tbConfigLAN.Enabled = false;

            Button btnConfigSave = new Button();
            btnConfigSave.Text = "Save";
            btnConfigSave.Location = new Point(rbConfigUART.Location.X, rbConfigLAN.Location.Y + rbConfigLAN.Height + 1);
            btnConfigSave.Click += new EventHandler(btnConfigSave_Click);

            frmConfig.Controls.Add(lblConfig);
            frmConfig.Controls.Add(rbConfigUART);
            frmConfig.Controls.Add(tbConfigUART);
            frmConfig.Controls.Add(rbConfigLAN);
            frmConfig.Controls.Add(tbConfigLAN);
            frmConfig.Controls.Add(btnConfigSave);
            frmConfig.Hide();
            //frmConfig
        }

        private void PopulateUI_IR()
        {
            RibbonPanel IRSetting = new RibbonPanel();
            IRSetting.Text = "IR Setting";
            ribbonEquipmentSetting.Panels.Add(IRSetting);

            Panel panelRCFormat = new Panel();
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
            /*
            foreach (string fmt in EnumListUser.RCKEY)
            {
                cbRCKeyList.Items.Add(fmt);
            }*/
            panelRCFormat.Controls.Add(cbRCKeyList);

            RibbonHost panelRCFormatHost = new RibbonHost();
            panelRCFormatHost.HostedControl = panelRCFormat;

            // Send Button Settings
            Button btRCKeyTransmit = new Button();
            btRCKeyTransmit.Text = "Send";
            btRCKeyTransmit.Size = new Size(50, 55);
            btRCKeyTransmit.Enabled = true;
            btRCKeyTransmit.FlatStyle = FlatStyle.Popup;
            btRCKeyTransmit.BackColor = Color.LightBlue;
            btRCKeyTransmit.Click += new EventHandler(btRCKeyTransmit_Click);
            btRCKeyTransmit.MouseEnter += new EventHandler(btRCKeyTransmit_MouseEnter);
            btRCKeyTransmit.MouseLeave += new EventHandler(btRCKeyTransmit_MouseLeave);

            RibbonHost btRCTransmitHost = new RibbonHost();
            btRCTransmitHost.HostedControl = btRCKeyTransmit;

            IRSetting.Items.Add(panelRCFormatHost);
            IRSetting.Items.Add(btRCTransmitHost);

            //IR config Form
            frmIR = new Form();
            frmIR.Size = new Size(200, 100);
            frmIR.FormBorderStyle = FormBorderStyle.None;
            frmIR.MouseLeave += new EventHandler(frmIR_MouseLeave);
            frmIR.MouseEnter += new EventHandler(frmIR_MouseEnter);
            frmIR.BackColor = System.Drawing.Color.LightYellow;

            Label lblIR = new Label();
            lblIR.Text = "IR Mode Setting ";
            lblIR.Location = new Point(2, 2);
            lblIR.Size = new Size(200, 20);

            rbIRmodeAUTO = new RadioButton();
            rbIRmodeAUTO.Text = "Auto";
            rbIRmodeAUTO.Location = new Point(lblIR.Location.X, lblIR.Location.Y + lblIR.Height + 5);
            rbIRmodeAUTO.CheckedChanged += new EventHandler(rbIRmodeAUTO_CheckedChanged);
            rbIRmodeAUTO.Size = new Size(70, rbIRmodeAUTO.Height);
            rbIRmodeAUTO.Checked = true;

            rbIRmodeMAN = new RadioButton();
            rbIRmodeMAN.Text = "Manual";
            rbIRmodeMAN.Location = new Point(rbIRmodeAUTO.Location.X + rbIRmodeAUTO.Width + 2, rbIRmodeAUTO.Location.Y);
            rbIRmodeMAN.CheckedChanged += new EventHandler(rbIRmodeMAN_CheckedChanged);
            rbIRmodeMAN.Size = new Size(70, rbIRmodeMAN.Height);
            rbIRmodeMAN.Checked = false;

            tbIRport = new TextBox();
            tbIRport.Location = new Point(rbIRmodeMAN.Location.X, rbIRmodeMAN.Location.Y + rbIRmodeMAN.Height + 2);
            tbIRport.Enabled = false;

            frmIR.Controls.Add(lblIR);
            frmIR.Controls.Add(rbIRmodeAUTO);
            frmIR.Controls.Add(rbIRmodeMAN);
            frmIR.Controls.Add(tbIRport);
            frmIR.Hide();

            initIR();
        }

        private void PopulateUI_Camera()
        {
            // Panel for Camera
            RibbonPanel cameraSetting = new RibbonPanel();
            cameraSetting.Text = "Camera Setting";
            ribbonEquipmentSetting.Panels.Add(cameraSetting);

            // Picture Box Settings
            pictureBox = new PictureBox();
            pictureBox.Size = new System.Drawing.Size(100, 50);
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;

            RibbonHost pictureBoxHost = new RibbonHost();
            pictureBoxHost.HostedControl = pictureBox;
            pictureBoxHost.ToolTip = "Click to enlarge";

            // Camera list and OnOff button
            Panel panelCam = new Panel();
            panelCam.Size = new Size(100, 50);

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

            // OCR calibration
            Panel panelOCR = new Panel();
            panelOCR.Size = new Size(80, 50);

            btCamCalib = new Button();
            btCamCalib.Text = "Camera Calibration";
            btCamCalib.Size = new Size(80, 50);
            btCamCalib.Enabled = true;
            btCamCalib.FlatStyle = FlatStyle.Popup;
            btCamCalib.BackColor = Color.LightBlue;
            btCamCalib.Click += new EventHandler(btCamCalib_Click);
            panelOCR.Controls.Add(btCamCalib);

            RibbonHost panelOCRHost = new RibbonHost();
            panelOCRHost.HostedControl = panelOCR;


            // Add items to Camera Setting Panel
            cameraSetting.Items.Add(pictureBoxHost);
            cameraSetting.Items.Add(panelCamHost);
            cameraSetting.Items.Add(panelOCRHost);

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
        #endregion

        #region Camera - Event Handler

            // On/Off button event for camera
        private void btOnOffCam_Click(object sender, EventArgs e)
        {
            if (btOnOffCam.Text == "On")
            {
                btOnOffCam.BackColor = Color.PaleVioletRed;
                btOnOffCam.Text = "Off";
                ImageHandler.StartCamera();
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
            }
        }

        // Camera list item click event
        private void cbCamList_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImageHandler.CamName = cbCamList.Text;
            ImageHandler.ImgFromParent = pictureBox;
            ImageHandler.InitCamera();

            btOnOffCam.BackColor = Color.LightBlue;
            btOnOffCam.Text = "On";
            if (Picture.Image != null)
            {
                Picture.Image.Dispose();
                Picture.Image = null;
            }
        }

        private void btCamCalib_Click(object sender, EventArgs eventArgs)
        {
            if (btOnOffCam.Text == "Off")
            {
                ImageHandler.OpenCameraCalibForm();
            }
            else
                MessageBox.Show("Please on the camera");
        }

        #endregion

        #region Test Connection - Event Handler
        private void TestConnection()
        {
            if (bgWorker.IsBusy != true)
            {
                btTestCon.Enabled = false;
                bgWorker.RunWorkerAsync("Test Connection");
            }
        }

        private void btTestCon_Click(object sender, EventArgs e)
        {
            TestConnection();
        }

        private void cbEquipment_MouseEnter(object sender, EventArgs e)
        {
            Console.WriteLine("cbEquipment_MouseEnter");
            var ControlName = (CheckBox)sender;

            if (frmConfig.Visible != true)
            {
                if (frmIR.Visible == true)
                {
                    KillTimer();
                    frmIR.Hide();
                }

                Console.WriteLine(ControlName.Text);
                lblConfig.Text = "Device Settings For " + ControlName.Text;
                rbConfig_Update(ControlName.Text);
                frmConfig.Location = new Point(Cursor.Position.X + 1, Cursor.Position.Y - 10);
                frmConfig.Show();
            }
        }

        private void cbEquipment_MouseLeave(object sender, EventArgs e)
        {
            Console.WriteLine("cbEquipment_MouseLeave");
            if (frmConfig.Visible == true)
                frmConfig_close();
        }

        private void frmConfig_MouseEnter(object sender, EventArgs e)
        {
            Console.WriteLine("frmConfig_MouseEnter");
            KillTimer();
        }

        private void frmConfig_MouseLeave(object sender, EventArgs e)
        {
            Console.WriteLine("frmConfig_MouseLeave");
            Rectangle screenBounds = new Rectangle(frmConfig.Location, frmConfig.Size);
            if (!screenBounds.Contains(Control.MousePosition))
            {
                frmConfig_close();
            }
        }

        private void frmConfig_close()
        {
            KillTimer();
            SetTimer(EN_ID_TIMER.ID_TIMER_CONFIG_MENU);
        }

        private void rbConfig_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton btnConfigDef = sender as RadioButton;
            if (rbConfigUART.Checked)
            {
                con_state = EN_CON_METHOD.CON_UART;
                if (tbConfigLAN != null)
                    tbConfigLAN.Enabled = false;
            }
            else if (rbConfigLAN.Checked)
            {
                con_state = EN_CON_METHOD.CON_LAN;
                if (tbConfigLAN != null)
                    tbConfigLAN.Enabled = true;
            }
        }

        private void btnConfigSave_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < deviceNum; i++)
            {
                Regex regex = new Regex(@_deviceSettings[i].DevName);
                Match match = regex.Match(lblConfig.Text);
                if (match.Success)
                {
                    switch (con_state)
                    {
                        case EN_CON_METHOD.CON_UART:
                            _deviceSettings[i].PortID = tbConfigUART.Text.ToString();
                            _deviceSettings[i].Con_Method = EN_CON_METHOD.CON_UART;
                            break;

                        case EN_CON_METHOD.CON_LAN:
                            _deviceSettings[i].PortID = tbConfigLAN.Text.ToString();
                            _deviceSettings[i].Con_Method = EN_CON_METHOD.CON_LAN;
                            break;
                    }
                    break;
                }
            }
            frmConfig.Hide();
        }

        private void rbConfig_Update(string devID)
        {
            for (int i = 0; i < deviceNum; i++)
            {
                Regex regex = new Regex(@_deviceSettings[i].DevName);
                Match match = regex.Match(devID);
                if (match.Success)
                {
                    switch (_deviceSettings[i].Con_Method)
                    {
                        case EN_CON_METHOD.CON_UART:
                            con_state = EN_CON_METHOD.CON_UART;
                            rbConfigUART.Checked = true;
                            tbConfigLAN.Text = "";
                            tbConfigUART.Text = _deviceSettings[i].PortID;
                            break;

                        case EN_CON_METHOD.CON_LAN:
                            con_state = EN_CON_METHOD.CON_LAN;
                            rbConfigLAN.Checked = true;
                            tbConfigLAN.Text = _deviceSettings[i].PortID;
                            tbConfigUART.Text = "";
                            break;
                    }
                    break;
                }
            }
        }
        #endregion

        #region IR - Event Handler
        private void btRCKeyTransmit_MouseEnter(object sender, EventArgs e)
        {
            Console.WriteLine("btRCKeyTransmit_MouseEnter");
            if (frmIR.Visible != true)
            {
                if (frmConfig.Visible == true)
                {
                    KillTimer();
                    frmConfig.Hide();
                }

                frmIR.Location = new Point(Cursor.Position.X + 1, Cursor.Position.Y - 10);
                frmIR.Show();
            }
        }

        private void btRCKeyTransmit_MouseLeave(object sender, EventArgs e)
        {
            Console.WriteLine("btRCKeyTransmit_MouseLeave");
            if (frmIR.Visible == true)
                frmIR_close();
        }

        private void frmIR_MouseEnter(object sender, EventArgs e)
        {
            Console.WriteLine("frmIR_MouseEnter");
            KillTimer();
        }

        private void frmIR_MouseLeave(object sender, EventArgs e)
        {
            Console.WriteLine("frmIR_MouseLeave");
            Rectangle screenBounds = new Rectangle(frmIR.Location, frmIR.Size);
            if (!screenBounds.Contains(Control.MousePosition))
            {
                frmIR_close();
            }
        }

        private void frmIR_close()
        {
            KillTimer();
            SetTimer(EN_ID_TIMER.ID_TIMER_IR_MENU);
        }

        private void rbIRmodeAUTO_CheckedChanged(object sender, EventArgs e)
        {
            ir_port = EN_IR_PORT.IR_PORT_AUTO;
            if (tbIRport != null)
                tbIRport.Enabled = false;
        }

        private void rbIRmodeMAN_CheckedChanged(object sender, EventArgs e)
        {
            ir_port = EN_IR_PORT.IR_PORT_MANUAL;
            if (tbIRport != null)
                tbIRport.Enabled = true;
        }
        #endregion

        #region Cell Click - Event Handler
        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!(e.RowIndex >= 0) || !(e.ColumnIndex >= 0))
                return;

            var datagridview = sender as DataGridView;
            string[] paths = datagridview.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag as string[];

            if (paths == null || paths.Length == 0)
                return;

            if (paths[0].ToString().Contains(Environment.CurrentDirectory))
            {
                var imageviewer = new PictureViewer(paths[0], paths[1]);
                imageviewer.Text = paths[0];
                imageviewer.Show();
            }
        }

        // Sub API to check connection method
        private EN_CON_METHOD check_con_method(string devID)
        {
            EN_CON_METHOD ret = EN_CON_METHOD.CON_MAX;
            for (int i = 0; i < deviceNum; i++)
            {
                if (_deviceSettings[i].DevName == devID)
                {
                    switch (_deviceSettings[i].Con_Method)
                    {
                        case EN_CON_METHOD.CON_LAN:
                            ret = EN_CON_METHOD.CON_LAN;
                            break;
                        case EN_CON_METHOD.CON_UART:
                            ret = EN_CON_METHOD.CON_UART;
                            break;
                    }
                    break;
                }
            }
            return ret;
        }

        #endregion

        #region Timer Handler
        private void SetTimer(EN_ID_TIMER timerID)
        {
            if (_timer != null)
            {
                if (_timer.Enabled == true)
                    return;
                else
                    _timer.Dispose();
            }            

            Console.WriteLine("Timer_Handler::SetTimer -> " + timerID.ToString());            

            _timer = new System.Timers.Timer();

            switch (timerID) // Set timer for respective timer ID
            {
                case EN_ID_TIMER.ID_TIMER_CLEAR_DICT:
                    _timer.Interval = 1000; // 10 sec                    
                    _timer.Elapsed += (sender, e) => OnTimedEvent(sender, timerID);
                    _timer.AutoReset = true;
                    break;
                case EN_ID_TIMER.ID_TIMER_CONFIG_MENU:
                    _timer.Interval = 500; // 5 sec
                    _timer.Elapsed += (sender, e) => OnTimedEvent(sender, timerID);
                    _timer.AutoReset = true;
                    break;
                case EN_ID_TIMER.ID_TIMER_IR_MENU:
                    _timer.Interval = 500; // 5 sec
                    _timer.Elapsed += (sender, e) => OnTimedEvent(sender, timerID);
                    _timer.AutoReset = true;
                    break;
                case EN_ID_TIMER.ID_TIMER_OPERATION:
                    _timer.Interval = 100; // 1 sec
                    _timer.Elapsed += (sender, e) => OnTimedEvent(sender, timerID);
                    _timer.AutoReset = true;
                    break;
                default:
                    _timer.Interval = 100; // 1 sec
                    break;
            }            

            _timer.Enabled = true;
            _timer.Start();
        }

        private void OnTimedEvent(Object sender, EN_ID_TIMER timerID)
        {            
            if (_timer == null || _timer.Enabled == false)
                return;
            
            Console.WriteLine("Timer_Handler::StopTimer -> " + timerID.ToString());

            switch (timerID)
            {
                case EN_ID_TIMER.ID_TIMER_CLEAR_DICT:
                    Console.WriteLine("OnTimedEvent::ID_TIMER_CLEAR_DICT");
                    ICollection<int> keys = new List<int>(_validateDict.Keys);
                    foreach (int k in keys)
                    {
                        _validateDict[k] = "";
                    }
                    _timer.Enabled = false;
                    _timer.Stop();                    
                    break;
                case EN_ID_TIMER.ID_TIMER_CONFIG_MENU:
                    Console.WriteLine("OnTimedEvent::ID_TIMER_CONFIG_MENU");
                    if (frmConfig.InvokeRequired)
                    {
                        frmConfig.Invoke(new Action(() =>
                        {
                            if(frmConfig.Visible)
                                frmConfig.Hide();
                        }));
                    }
                    _timer.Enabled = false;
                    _timer.Stop();
                    break;
                case EN_ID_TIMER.ID_TIMER_IR_MENU:
                    Console.WriteLine("OnTimedEvent::ID_TIMER_IR_MENU");
                    if (frmIR.InvokeRequired)
                    {
                        frmIR.Invoke(new Action(() =>
                        {
                            if (frmIR.Visible)
                                frmIR.Hide();
                        }));
                    }
                    _timer.Enabled = false;
                    _timer.Stop();
                    break;
                case EN_ID_TIMER.ID_TIMER_OPERATION:
                    elapsedTime++;
                    break;
            }                    
        }

        private void KillTimer()
        {
            if (_timer == null || _timer.Enabled == false)
                return;

            Console.WriteLine("Timer_Handler::KillTimer");

            _timer.Enabled = false;
            _timer.Stop();
        }

        #endregion

        #region Image Handler
        private Boolean CompareImgResult(string input)
        {
            bool ret = false;
            bool infoflag = true;
            int trueCnt = 0;
            int falseCnt = 0;
            int tryAgain = 0;
            string tempRes = "";

            
            if (NGImage != null)
                NGImage.Dispose();
            if (OKImage != null)
                OKImage.Dispose();

            // ----- Checklist Start -----
            Regex imgResult = new Regex(@"Resolution as (\d*\w*)");
            Match match_comp1 = imgResult.Match(input);

            imgResult = new Regex(@"OSD of \[(\d*\w*)\]");
            Match match_comp2 = imgResult.Match(input);

            imgResult = new Regex(@"Resolution.*blank");
            Match match_comp3 = imgResult.Match(input);

            imgResult = new Regex(@"image.* is displayed");
            Match match_comp4 = imgResult.Match(input);

            imgResult = new Regex(@"image.* is not displayed");
            Match match_comp5 = imgResult.Match(input);

            imgResult = new Regex(@"is erased");
            Match match_comp6 = imgResult.Match(input);

            imgResult = new Regex(@"is not erased");
            Match match_comp7 = imgResult.Match(input);

            imgResult = new Regex(@"INFO.*invalid");
            Match match_comp8 = imgResult.Match(input);
            // ----- Checklist End -----

            if (IR_checkConnection())
            {
                //Saal.IRTOYS_OpenCon();
                ImageHandler.OperationDialog_updLog("", Color.White);
                ImageHandler.OperationDialog_updLog("##### Start image comparing #####", Color.White);
                ImageHandler.OperationDialog_updLog("", Color.White);

                while (!ret)
                {
                    if (match_comp1.Success || match_comp2.Success)
                    {
                        // TEST CASE #1,#2) Check resolution is displayed
                        if (match_comp1.Success)
                            tempRes = match_comp1.Groups[1].ToString();
                        else
                            tempRes = match_comp2.Groups[1].ToString();

                        Console.WriteLine("CASE 1,2) String : " + input + " Resolution : " + tempRes);

                        // I - Attempting to open Info menu
                        if (infoflag)
                        {                            
                            ImageHandler.OperationDialog_updLog("<KEY> Press INFO", Color.White);
                            ImageHandler.OperationDialog_updLog("Waiting TV to open info menu...", Color.White);
                            IR_sendCommand(IrFormat, "INFO");
                            infoflag = false;
                            Thread.Sleep(5000);

                            // II - Checking if Info menu already opened
                            ImageHandler.OperationDialog_updLog("Checking if info menu already opened...", Color.White);
                            if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_OPT1) != "Close")
                            {
                                ImageHandler.OperationDialog_updLog("<KEY> Press BACK", Color.White);
                                ImageHandler.OperationDialog_updLog("Detecting info menu is not opened. Attempting to open again...", Color.White);
                                IR_sendCommand(IrFormat, "BACK");
                                Thread.Sleep(500);
                                IR_sendCommand(IrFormat, "BACK");
                                Thread.Sleep(500);
                                infoflag = true;
                                continue;
                            }
                            ImageHandler.OperationDialog_updLog("Detecting info menu is opened...", Color.White);
                        }

                        // III - Checking if resolution match with OCR
                        if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_INFO).Contains(tempRes))
                        {                            
                            OKImage = new Bitmap(ImageHandler.CurrentFrame);
                            ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                            ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                            IR_sendCommand(IrFormat, "BACK");
                            ret = true;
                            break;
                        }
                        // IV - Attempting to execute OCR again if resolution doesn't match
                        else
                        {
                            // V - Checking if number of attempts is 15. Stop execution.
                            if (tryAgain != 0 && tryAgain % 15 == 0)
                            {
                                NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                IR_sendCommand(IrFormat, "BACK");
                                break;
                            }
                            // VI - Checking if number of attempts is 5,10. Reopen Info menu.
                            else if (tryAgain != 0 && tryAgain % 5 == 0)
                            {                             
                                ImageHandler.OperationDialog_updLog("OCR failed. Attempting to reopen info menu and execute OCR (#" + tryAgain + ")...", Color.White);
                                IR_sendCommand(IrFormat, "BACK");
                                infoflag = true;
                                Thread.Sleep(2000);
                            }
                            // VII - Checking if number of attempts is not 5,10,15. Try OCR again.
                            else
                            {
                                ImageHandler.OperationDialog_updLog("Attempting to execute OCR (#" + tryAgain + ")...", Color.White);
                            }                           
                            
                            tryAgain++;
                        }
                    }
                    else if(match_comp3.Success)
                    {
                        // TEST CASE #3) Check blank resolution is displayed
                        Console.WriteLine("CASE 3) String " + input);

                        // I - Attempting to open Info menu
                        if (infoflag)
                        {
                            ImageHandler.OperationDialog_updLog("<KEY> Press INFO", Color.White);
                            ImageHandler.OperationDialog_updLog("Waiting TV to open info menu...", Color.White);
                            IR_sendCommand(IrFormat, "INFO");
                            infoflag = false;
                            Thread.Sleep(5000);

                            // II - Checking if Info menu already opened
                            ImageHandler.OperationDialog_updLog("Checking if info menu already opened...", Color.White);
                            if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_OPT1) != "Close")
                            {
                                ImageHandler.OperationDialog_updLog("<KEY> Press BACK", Color.White);
                                ImageHandler.OperationDialog_updLog("Detecting info menu is not opened. Attempting to open again...", Color.White);
                                IR_sendCommand(IrFormat, "BACK");
                                Thread.Sleep(500);
                                IR_sendCommand(IrFormat, "BACK");
                                Thread.Sleep(500);
                                infoflag = true;
                                continue;
                            }
                            ImageHandler.OperationDialog_updLog("Detecting info menu is opened...", Color.White);
                        }

                        // III - Checking if blank resolution match with OCR
                        if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_INFO) == "")
                        {
                            OKImage = new Bitmap(ImageHandler.CurrentFrame);
                            ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                            ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                            IR_sendCommand(IrFormat, "BACK");
                            ret = true;
                            break;
                        }
                        // IV - Attempting to execute OCR again if resolution doesn't match
                        else
                        {
                            // V - Checking if number of attempts is 15. Stop execution.
                            if (tryAgain != 0 && tryAgain % 15 == 0)
                            {
                                NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                IR_sendCommand(IrFormat, "BACK");
                                break;
                            }
                            // VI - Checking if number of attempts is 5,10. Reopen Info menu.
                            else if (tryAgain != 0 && tryAgain % 5 == 0)
                            {
                                ImageHandler.OperationDialog_updLog("OCR failed. Attempting to reopen info menu and execute OCR (#" + tryAgain + ")...", Color.White);
                                IR_sendCommand(IrFormat, "BACK");
                                infoflag = true;
                                Thread.Sleep(2000);
                            }
                            // VII - Checking if number of attempts is not 5,10,15. Try OCR again.
                            else
                            {
                                ImageHandler.OperationDialog_updLog("Attempting to execute OCR (#" + tryAgain + ")...", Color.White);
                            }

                            tryAgain++;
                        }

                        break;
                    }
                    else if (match_comp4.Success || match_comp5.Success)
                    {
                        // TEST CASE #4,#5) Check image is displayed or not displayed
                        Console.WriteLine("CASE 4,5) String " + input);

                        // I - Executing binary calculation to get black percentage
                        var binaryImage = ImageHandler.GetBinarialImage_GlobalCrop(ImageHandler.CurrentFrame, 70);                        
                        double perc = (int)ImageHandler.GetBlackPixelPercentage(binaryImage);
                        ImageHandler.OperationDialog_updLog("<Binarization Process> Black pixel percentage : " + perc + "%", Color.White);

                        // II - Checking if black percentage is higher or lower than threshold
                        if (perc < 40)
                        {
                            Console.WriteLine("White");
                            // Case OK for image is displayed (#4)
                            if (match_comp4.Success)
                            {
                                OKImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                                ret = true;
                            }
                            // Case NG for image is not displayed (#5)
                            else
                            {
                                NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                ret = false;
                            }                            
                        }
                        else
                        {
                            Console.WriteLine("Black");
                            // Case NG for image is displayed (#4)
                            if (match_comp4.Success)
                            {
                                NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                ret = false;
                            }
                            // Case OK for image is not displayed (#5)
                            else
                            {
                                OKImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                                ret = true;
                            }
                        }                        
                        break;
                    }
                    else if (match_comp6.Success || match_comp7.Success)
                    {
                        // TEST CASE #6,#7) Check not support resolution is erased / not erased
                        Console.WriteLine("CASE 6,7) String " + input);

                        // I - Return to LIVE mode from any menu
                        ImageHandler.OperationDialog_updLog("<KEY> Press BACK", Color.White);
                        IR_sendCommand(IrFormat, "BACK");
                        Thread.Sleep(1000);
                        ImageHandler.OperationDialog_updLog("<KEY> Press BACK", Color.White);
                        IR_sendCommand(IrFormat, "BACK");
                        ImageHandler.OperationDialog_updLog("Waiting OSD label to appear...", Color.White);
                        Thread.Sleep(10000);

                        if (match_comp6.Success)
                        {
                            // CASE 6 - OSD label erased
                            // II - Checking if resolution match with OCR
                            if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_LABEL) == "null")
                            {
                                OKImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                                ret = true;
                                break;
                            }
                            // III - Attempting to execute OCR again if resolution doesn't match
                            else
                            {
                                // IV - Checking if number of attempts is 15. Stop execution.
                                if (tryAgain != 0 && tryAgain % 15 == 0)
                                {
                                    NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                    ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                    ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                    break;
                                }
                                ImageHandler.OperationDialog_updLog("Attempting to execute OCR (#" + tryAgain + ")...", Color.White);
                                tryAgain++;
                            }
                        }
                        else
                        {
                            // CASE7 - OSD label not erased
                            // II - Checking if resolution match with OCR
                            if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_LABEL) != "Please change source resolution")
                            {
                                OKImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                                ret = true;
                                break;
                            }
                            // III - Attempting to execute OCR again if resolution doesn't match
                            else
                            {
                                // IV - Checking if number of attempts is 15. Stop execution.
                                if (tryAgain != 0 && tryAgain % 15 == 0)
                                {
                                    NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                    ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                    ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                    break;
                                }
                                ImageHandler.OperationDialog_updLog("Attempting to execute OCR (#" + tryAgain + ")...", Color.White);
                                tryAgain++;
                            }
                        }

                    }
                    else if (match_comp8.Success)
                    {
                        // TEST CASE 8) Check INFO key is invalid
                        Console.WriteLine("CASE 8) String " + input);

                        // I - Attempting to open Info menu
                        ImageHandler.OperationDialog_updLog("<KEY> Press INFO", Color.White);
                        ImageHandler.OperationDialog_updLog("Waiting TV to open info menu...", Color.White);
                        IR_sendCommand(IrFormat, "INFO");
                        Thread.Sleep(5000);

                        // II - Checking if number of attempts is 15. Stop execution.
                        if (tryAgain != 0 && tryAgain % 15 == 0)
                        {
                            if(trueCnt > falseCnt)
                            {
                                NGImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)NGImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result NG> Saving NG Image...", Color.Red);
                                IR_sendCommand(IrFormat, "BACK");
                            }
                            else
                            {
                                OKImage = new Bitmap(ImageHandler.CurrentFrame);
                                ImageHandler.OperationDialog_updRefCom((Bitmap)OKImage.Clone());
                                ImageHandler.OperationDialog_updLog("<Result OK> Saving OK Image...", Color.Green);
                                ret = true;
                            }
                            break;
                        }
                        tryAgain++;

                        // Checking INFO menu availability
                        // II - Checking if Info menu already opened
                        if (ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_OPT1) != "Close")
                        {
                            ImageHandler.OperationDialog_updLog("Detecting info menu is not opened. Verifying again...", Color.White);
                            falseCnt++;
                        }
                        else
                        {
                            ImageHandler.OperationDialog_updLog("Detecting info menu is opened. Verifying again...", Color.White);
                            trueCnt++;
                        }
                    }
                }
            }
            return ret;
        }
        #endregion

        #region IR Handler

        private void cbRCFormatList_SelectedIndexChanged(object sender, EventArgs e)
        {
            IrFormat = cbRCFormatList.Text;
            cbRCKeyList.Items.Clear();
            if (cbRCFormatList.Text == "PHILIPS_RC5")
            {
                for (Int32 index = 0; index < Saal.PHILIPSRC5.Length - 1; index++)
                {
                    if (Saal.PHILIPSRC5[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.PHILIPSRC5[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "PHILIPS_RC6")
            {
                for (Int32 index = 0; index < Saal.PHILIPSRC6.Length - 1; index++)
                {
                    if (Saal.PHILIPSRC6[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.PHILIPSRC6[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "Matsushita")
            {
                for (Int32 index = 0; index < Saal.MATSUSHITA.Length - 1; index++)
                {
                    if (Saal.MATSUSHITA[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.MATSUSHITA[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "NEC_THAI")
            {
                for (Int32 index = 0; index < Saal.NECTHAI.Length - 1; index++)
                {
                    if (Saal.NECTHAI[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.NECTHAI[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "NEC_INDIA")
            {
                for (Int32 index = 0; index < Saal.NECINDIA.Length - 1; index++)
                {
                    if (Saal.NECINDIA[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.NECINDIA[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "NEC")
            {
                for (Int32 index = 0; index < Saal.NECNEW.Length - 1; index++)
                {
                    if (Saal.NECNEW[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.NECNEW[index].KEYLABEL);
                }
            }
            else if (cbRCFormatList.Text == "SANYO")
            {
                for (Int32 index = 0; index < Saal.SANYO.Length - 1; index++)
                {
                    if (Saal.SANYO[index].KEYLABEL != null)
                        cbRCKeyList.Items.Add(Saal.SANYO[index].KEYLABEL);
                }
            }

            cbRCKeyList.SelectedIndex = 0;
        }

        private void cbRCKeyList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btRCKeyTransmit_Click(object sender, EventArgs e)
        {
            if (IR_checkConnection())
            {
                //Saal.IRTOYS_OpenCon();
                Console.WriteLine("Format : " + cbRCFormatList.Text + ", Key : " + cbRCKeyList.Text);
                IR_sendCommand(cbRCFormatList.Text, cbRCKeyList.Text);
                //Saal.IRTOYS_CloseCon();
            }
            else
            {
                MessageBox.Show("COM port cannot find, please manually reassign correct COM port.");
            }
        }

        private void initIR()
        {
            Saal.IRTOYS_init();
            cbRCFormatList.SelectedIndex = 0;
        }

        private Boolean IR_checkConnection()
        {
            bool mode = true;

            if (ir_port == EN_IR_PORT.IR_PORT_MANUAL)
            {
                if (tbIRport.Text != "")
                    mode = false;
                else
                    MessageBox.Show("Please insert correct COM port");
            }

            if (Saal.IRTOYS_Setup(mode, tbIRport.Text))
            {
                Console.WriteLine("Connected");
                return true;
            }
            else
            {
                Console.WriteLine("Not Connected");
                return false;
            }
        }

        private Boolean IR_sendCommand(string format, string key)
        {
            return Saal.IRTOYS_SendCMD(format, key);
        }
        #endregion

        #region Send Command Handler
        private Boolean SendCommand(string device, string comm, EN_CON_METHOD con_method)
        {
            bool result1 = false;
            bool result2 = false;
            bool isUART = true;

            if (testUARTBypass)
                return true;

            if (con_method == EN_CON_METHOD.CON_LAN)
                isUART = false;

            if (test_mode == EN_TEST_MODE.TEST_MODE_1)
            {
                Console.WriteLine("Test Mode 1");
                result1 = Saal.TG45_59_SetFormat(toExecute.Cells[4].Value.ToString(), comm, isUART); // Change Format (col 4)
                result2 = true;
            }
            else if (test_mode == EN_TEST_MODE.TEST_MODE_2)
            {
                Console.WriteLine("Test Mode 2");
                result1 = Saal.TG45_59_SetFormat(toExecute.Cells[4].Value.ToString(), comm, isUART); // Change Format (col 4)
                Thread.Sleep(1000);
                result2 = Saal.TG45_59_SetFormat(toExecute.Cells[6].Value.ToString(), comm, isUART); // Change Format (col 6)
            }

            Console.WriteLine("return1 : " + result1 + " return2 : " + result2);
            if (result1 && result2)
                return true;
            else
                return false;
        }
        #endregion

        #region Validation Process
        private bool validateProcess()
        {
            bool flagCh = false;
            bool flagRet = true;

            foreach (KeyValuePair<int, string> pair in _validateDict)
            {
                if ((toExecute.Cells[pair.Key].Value.ToString() != "")
                    && (toExecute.Cells[pair.Key].Value.ToString() != "#ERROR"))
                    flagCh = true;
                else
                {
                    // Store previous value into new cell, input Error if previous value empty
                    if (pair.Value.ToString() != "")
                        toExecute.Cells[pair.Key].Value = pair.Value;
                    else
                        toExecute.Cells[pair.Key].Value = "#ERROR";
                }

                if (toExecute.Cells[pair.Key].Value.ToString() == "#ERROR")
                {
                    ImageHandler.OperationDialog_updLog("<Validation Error> Input at column #" + pair.Key + " is empty. Skipping this row...", Color.White);
                    flagRet = false;
                }
            }

            if (flagCh)
            {
                ICollection<int> keys = new List<int>(_validateDict.Keys);
                foreach (int k in keys)
                {
                    _validateDict[k] = toExecute.Cells[k].Value.ToString();
                }
            }

            // Test mode selection
            Regex regex = new Regex(@"INFO");
            Match match = regex.Match(_validateDict[6].ToString());
            if (match.Success)
            {
                test_mode = EN_TEST_MODE.TEST_MODE_1;
            }
            else
            {
                test_mode = EN_TEST_MODE.TEST_MODE_2;                
            }

            if (!IR_checkConnection())
            {
                ImageHandler.OperationDialog_updLog("<Validation Error> IR is not connected", Color.Red);
                flagRet = false;
            }

            return flagRet;
        }
        #endregion  

        #region Background Worker
        void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (e.Result.ToString() == "Test OCR")
                {
                    MessageBox.Show("OCR completed");
                }

                else if (e.Result.ToString() == "Test Connection")
                {
                    Int32 index = 0;
                    for (int i = 0; i < deviceNum; i++)
                    {
                        Console.WriteLine(_deviceSettings[i].DevName + ", " + _deviceSettings[i].PortID);
                        if (_deviceSettings[i].Con_Status)
                            cbEquipment[index].BackColor = Color.Green;
                        else
                            cbEquipment[index].BackColor = Color.Red;
                        index++;
                    }
                    btTestCon.Enabled = true;
                    MessageBox.Show("Connection test completed");
                }
            }
            catch
            {

            }

        }

        void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (e.Argument.ToString() == "Test Connection")
            {
                for (int i = 0; i < deviceNum; i++)
                {
                    switch (_deviceSettings[i].Con_Method)
                    {
                        case EN_CON_METHOD.CON_UART:
                            if (testUARTBypass)
                            {
                                _deviceSettings[i].PortID = "COM1";
                                _deviceSettings[i].Con_Status = true;
                                break;
                            }
                            // Check if connection OK   
                            string resPort = Saal.TG45_59_TestPerCon(_deviceSettings[i].DevName);
                            if (resPort != "")
                            {
                                _deviceSettings[i].PortID = resPort;
                                _deviceSettings[i].Con_Status = true;
                            }
                            else
                            {
                                _deviceSettings[i].Con_Status = false;
                            }
                            break;
                        case EN_CON_METHOD.CON_LAN:
                            if (testUARTBypass)
                            {
                                _deviceSettings[i].PortID = "127.0.0.1";
                                _deviceSettings[i].Con_Status = true;
                                break;
                            }

                            if (Saal.TG45_59_Lan_TestPerCon(_deviceSettings[i].PortID))
                                _deviceSettings[i].Con_Status = true;
                            else
                                _deviceSettings[i].Con_Status = false;
                            break;
                    }
                }
                e.Result = "Test Connection";
            }

        }
        #endregion

        #region ProgressBar
        public void _ProgressBarChangedEvent(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage != 0)
            {
                string remain = "Current Status : " + (noOfRow - noOfRowTemp).ToString() + "/" + noOfRow.ToString() + "";
                ImageHandler.OperationDialog_updPbar(e.ProgressPercentage, remain);
            }
        }

        private void InvokeProgressBarChangedEvent(object sender, ProgressChangedEventArgs e)
        {
            _ProgressBarChangedEvent(sender, e);
        }
        #endregion

        #region Change source
        private Boolean ChangeSource(EN_INPUT_LIST src)
        {
            EN_INPUT_LIST source = EN_INPUT_LIST.EN_INPUT_MAX;
            bool ret = false;
            bool sourceflag = true;
            int tryAgain = 0;
            int tryAgainFull = 0;
            if (IR_checkConnection())
            {
                //Saal.IRTOYS_OpenCon();
                ImageHandler.OperationDialog_updLog("", Color.White);
                ImageHandler.OperationDialog_updLog("##### Start changing source #####", Color.White);
                ImageHandler.OperationDialog_updLog("", Color.White);
                while (!ret)
                {
                    if (sourceflag)
                    {
                        ImageHandler.OperationDialog_updLog("<KEY> Press SOURCE", Color.White);
                        ImageHandler.OperationDialog_updLog("Waiting the TV to open source menu...", Color.White);
                        IR_sendCommand(IrFormat, "SOURCE");
                        sourceflag = false;
                        Thread.Sleep(5000);
                    }

                    source = CurrentInput(ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_SOURCE));

                    if ((int)source == (int)EN_INPUT_LIST.EN_INPUT_MAX)
                    {                        
                        Thread.Sleep(500);

                        if (tryAgain == 5)
                        {
                            ImageHandler.OperationDialog_updLog("<KEY> Press BACK", Color.White);
                            ImageHandler.OperationDialog_updLog("Unable to find anything. Selecting current source and attempting to find again...", Color.White);
                            sourceflag = true;
                            tryAgain = 0;
                            IR_sendCommand(IrFormat, "BACK");
                            Thread.Sleep(1000);
                            IR_sendCommand(IrFormat, "BACK");
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            ImageHandler.OperationDialog_updLog("Attempting to decode OCR (#" + tryAgain + ")...", Color.White);
                        }
                        tryAgain++;
                    }
                    else if ((int)source == (int)src)
                    {
                        ImageHandler.OperationDialog_updLog("Found input source !!!", Color.White);
                        ImageHandler.OperationDialog_updLog("<KEY> Press OK", Color.White);
                        IR_sendCommand(IrFormat, "OK");

                        // Recheck if already press OK button
                        Thread.Sleep(3000);
                        //EN_INPUT_LIST sourceReCheck = EN_INPUT_LIST.EN_INPUT_MAX;
                        ImageHandler.OperationDialog_updLog("Rechecking if still in source menu...", Color.White);
                        while (source == CurrentInput(ImageHandler.ExtractText(ImageHandler.EN_OCR_ITEM.OCR_ITEM_SOURCE)))
                        {
                            ImageHandler.OperationDialog_updLog("Still in source menu...", Color.White);
                            ImageHandler.OperationDialog_updLog("<KEY> Press OK", Color.White);
                            IR_sendCommand(IrFormat, "OK");
                            Thread.Sleep(3000);
                        }
                        ImageHandler.OperationDialog_updLog("Already change menu...", Color.White);
                        ImageHandler.OperationDialog_updLog("Waiting TV to receive signal...", Color.White);                        
                        Thread.Sleep(7000); // Wait 10 sec until image appear
                        ret = true;
                        break;
                    }
                    else if ((int)source < (int)src)
                    {
                        ImageHandler.OperationDialog_updLog("<KEY> Press LEFT", Color.White);
                        ImageHandler.OperationDialog_updLog("Moving to the left icon...", Color.White);
                        IR_sendCommand(IrFormat, "LEFT");
                        tryAgainFull++;
                        Thread.Sleep(500);
                    }
                    else if ((int)source > (int)src)
                    {
                        ImageHandler.OperationDialog_updLog("<KEY> Press RIGHT", Color.White);
                        ImageHandler.OperationDialog_updLog("Moving to the right icon...", Color.White);
                        IR_sendCommand(IrFormat, "RIGHT");
                        tryAgainFull++;
                        Thread.Sleep(500);
                    }

                    if(tryAgainFull == 15)
                    {
                        ImageHandler.OperationDialog_updLog("Unable to find input source. Operation terminated...", Color.White);
                        break;
                    }
                }
                //Saal.IRTOYS_CloseCon();
            }
            else
            {
                MessageBox.Show("COM port cannot find, please manually reassign correct COM port.");
            }
            return ret;
        }

        private EN_INPUT_LIST CurrentInput(string input)
        {
            EN_INPUT_LIST output = EN_INPUT_LIST.EN_INPUT_MAX;

            if (Regex.Match(input, @"CAST").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_CAST;
            }
            else if (Regex.Match(input, @"TV").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_TV;
            }
            else if (Regex.Match(input, @"HDMI1").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_HDMI1;
            }
            else if (Regex.Match(input, @"HDMI2").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_HDMI2;
            }
            else if (Regex.Match(input, @"HDMI3").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_HDMI3;
            }
            else if (Regex.Match(input, @"HDMI4").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_HDMI4;
            }
            else if (Regex.Match(input, @"Video").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_VIDEO;
            }
            else if (Regex.Match(input, @"PC").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_PC;
            }
            else if (Regex.Match(input, @"USB").Success)
            {
                output = EN_INPUT_LIST.EN_INPUT_USB;
            }
            else
            {
                output = EN_INPUT_LIST.EN_INPUT_MAX; // Case cannot find anything
            }
            return output;
        }
        #endregion

        #region Save to folder
        private string[] _SaveImagesToOutputFolder(Bitmap images)
        {
            // create folder in exe level
            var current_app_path = Environment.CurrentDirectory;

            // create output folder
            var outputFolder = current_app_path + "\\output";
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // create output/testparent folder
            var parentImagesFolder = outputFolder + "\\CapturedImagesByRow_ExtInput";
            if (!Directory.Exists(parentImagesFolder))
                Directory.CreateDirectory(parentImagesFolder);

            // create sub folder
            var folderName = toExecute.Cells[0].EditedFormattedValue.ToString();
            var imagesFolder = parentImagesFolder + "\\[" + folderName + "]";
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            var imagePath = imagesFolder + "\\" + _executionTime + ".png";
            if (File.Exists(imagePath))
                File.Delete(imagePath);

            images.Save(imagePath, ImageFormat.Png);
            /*
            for (int i = 0; i < images.Length; ++i)
            {
                var imagePath = imagesFolder + "\\" + _executionTime + (i + 1) + ".png";
                if (File.Exists(imagePath))
                    File.Delete(imagePath);

                images[i].Save(imagePath, ImageFormat.Png);
            }
            */
            return new[] { imagesFolder, _executionTime };
        }

        #endregion

        #region Before and after execute
        private void InvokeBeforeExecute()
        {
            // check camera status
            if (btOnOffCam.Text == "On")
            {
                MessageBox.Show("Do switch on the camera before testing");
                DoExecute = false;
                return;
            }

            // set timestamp
            var curdate = DateTime.Now.ToString("yyyyMMdd_");
            var curtime = DateTime.Now.ToString("hhmm.fff_");
            _executionTime = curdate + curtime;

            ImageHandler.OperationDialog_Open();

            // decide to execute or note
            DoExecute = true;
            ImageHandler.OperationDialog_updLog("Test Start !!!", Color.White);

            elapsedTime = 0;
            SetTimer(EN_ID_TIMER.ID_TIMER_OPERATION);
        }

        private void InvokeAfterExecute()
        {
            string currentSec = (elapsedTime % 60).ToString();
            string currentMin = (elapsedTime / 60).ToString();            
            string currentHour = (double.Parse(currentMin) / 60).ToString();

            ImageHandler.OperationDialog_updLog("Test Finish !!! (Time elapsed : " + currentHour + "hour " + currentMin + "min " + currentSec + "sec)", Color.White);
            ImageHandler.OperationDialog_updLog("------------------------------------", Color.White);
            ImageHandler.UpdOpFlag = true;

            string remain = "Current Status : " + (noOfRow - noOfRowTemp).ToString() + "/" + noOfRow.ToString() + "";
            ImageHandler.OperationDialog_updPbar(100, remain);

            
        }

        private void InvokeBeforeExecuteButAfterSelection(object selection, object datagridview, object owner, BackgroundWorker backgroundWorker)
        {
            bgWorkerFromTopForm = backgroundWorker;
            var numOfRowtemp = selection as List<int>;
            noOfRow = 0;
            noOfRowTemp = 0;
            noOfRow = numOfRowtemp.Count;
            noOfRowTemp = noOfRow;

            string remain = "Current Status : " + (noOfRow - noOfRowTemp).ToString() + "/" + noOfRow.ToString() + "";
            ImageHandler.OperationDialog_updPbar(0, remain);
        }

        
        #endregion

        #region Main process
        private void DataRowProcess()
        {
            ImageHandler.startScan();    
            bgWorkerFromTopForm.ReportProgress((noOfRow - noOfRowTemp) * 100 / noOfRow);            

            KillTimer();

            BusyFlag = true;

            // Validating input
            Console.WriteLine("Start validating input");
            ImageHandler.OperationDialog_updLog("Start validating input...", Color.White);
            if (!validateProcess())
            {
                Console.WriteLine("Validation failed");
                ImageHandler.OperationDialog_updLog("Input validation failed. Skipping this row...", Color.Red);
                BusyFlag = false;
                return;
            }

            Console.WriteLine("Start process");
            ImageHandler.OperationDialog_updLog("Start processing row...", Color.White);

            // Test operation
            bool errFlagDev = true;
            bool errFlagInpSrc = false;
            bool res = false;
            string devtemp = toExecute.Cells[3].Value.ToString();
            for (int i = 0; i < deviceNum; i++)
            {
                // Check test device name
                Regex regex = new Regex(@_deviceSettings[i].DevName);
                Match match = regex.Match(devtemp);
                if (match.Success)
                {
                    // Check test device connection       
                    if (_deviceSettings[i].Con_Status != false)
                    {
                        errFlagDev = false;

                        // Operation to change current input. Break if input source does not exist.
                        EN_INPUT_LIST cur_input_temp = CurrentInput(toExecute.Cells[1].Value.ToString());
                        if (cur_input != cur_input_temp)
                        {
                            if (ChangeSource(cur_input_temp))
                            {                                
                                cur_input = cur_input_temp;
                            }
                            else
                            {
                                errFlagInpSrc = true;
                                break;
                            }
                        }

                        // Operation to change format name and compare image                        
                        res = SendCommand(_deviceSettings[i].DevName, _deviceSettings[i].PortID, _deviceSettings[i].Con_Method)
                                && CompareImgResult(toExecute.Cells[7].Value.ToString());

                        Action<Bitmap> okAction = (bitmapCell) =>
                        {
                            toExecute.Cells[11].Value = "OK";
                            var imageFolderPath = _SaveImagesToOutputFolder(bitmapCell);

                            var resizedImage = new Bitmap(bitmapCell,
                                toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Size);
                            toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Value = resizedImage;
                            toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Tag = imageFolderPath;
                        };

                        Action<Bitmap> ngAction = (bitmapCell) =>
                        {
                            toExecute.Cells[11].Value = "NG1";
                            var imageFolderPath = _SaveImagesToOutputFolder(bitmapCell);

                            var resizedImage = new Bitmap(bitmapCell,
                                toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Size);
                            toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Value = resizedImage;
                            toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Tag = imageFolderPath;
                        };

                        if (res)
                        {
                            ImageHandler.drawWatermark(ref OKImage, @"OK Image");
                            okAction(OKImage);
                        }
                        else
                        {
                            ImageHandler.drawWatermark(ref NGImage, @"NG image");
                            ngAction(NGImage);
                        }

                        toExecute.Cells[12].Value = tbVer.Text;
                        toExecute.Cells[13].Value = tbTester.Text;
                        toExecute.Cells[14].Value = DateTime.Now.ToString("d/M/yyyy");
                    }          
                    break;                
                }
            }

            // ERROR) Test device disconnected
            if (errFlagDev)
            {
                ImageHandler.OperationDialog_updLog("Test device not connected. Skipping this row...", Color.White);
                toExecute.Cells[3].Value += " (#ERROR)";
            }

            // ERROR) Input source does not exist
            if (errFlagInpSrc)
            {
                ImageHandler.OperationDialog_updLog("Input source does not exist. Skipping this row...", Color.White);
                toExecute.Cells[1].Value += " (#ERROR)";
            }

            BusyFlag = false;
            Console.WriteLine("Stop process");
            ImageHandler.OperationDialog_updLog("Finished processing row...", Color.White);
            SetTimer(EN_ID_TIMER.ID_TIMER_CLEAR_DICT);
            ImageHandler.stopScan();
            noOfRowTemp--;
        }

        #endregion
    }
}
