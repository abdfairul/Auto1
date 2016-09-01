using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginContracts;
using System.Windows.Forms;
using System.Data;
using SAAL;

#region Camera/Mic include files
using System.Drawing;
using System.Drawing.Imaging;
using AForge.Video;
using AForge.Video.DirectShow;
using AForge.Imaging;
using System.ComponentModel;
using System.Threading;
#endregion

namespace Test
{
    public class PlugUnplug : IPlugin
    {
        SAAL_Interface Saal = new SAAL_Interface();
        private RibbonTab ribbonEquipmentSetting;
        private DataGridViewRow toExecute;
        private PictureBox pictureBox;
        private bool BusyFlag;
        private bool DoFlag = true;
        private BeforeExecute BeforeFlag;
        private AfterExecute AfterFlag;
        private CellClickEvent ClickEventFlag;

        #region Test Connection - Define parameters
        private List<string> EquipmentList;
        private CheckBox[] cbEquipment;
        #endregion

        #region Camera/Mic - Define parameters
        // Parameters for Camera
        private PictureBox pictureBoxPopup;
        private ComboBox cbCamList;
        private Button btOnOffCam;
        private Form popupBox;
        private int ResWidth = 640;
        private int ResHeight = 480;

        // Parameters for Mic
        private ProgressBar pbVolIndicator;
        private TrackBar tbVolThres;
        private Label lblVolThres;
        private BackgroundWorker bwVolIndicator;
        private ComboBox cbMicList;
        #endregion

        // Constructor
        public PlugUnplug()
        {
            // Add equipment lists for this application
            EquipmentList = new List<string>();
            EquipmentList.Add("TG59");
            EquipmentList.Add("TG45");
            EquipmentList.Add("TG35");

            PopulateUI();

            #region Camera/Mic - Populate UI and Initialize
            PopulateUI_CameraMic();
            InitCamera_Mic();
            #endregion

            ClickEventFlag = dataGridView_CellClick;
        }

        #region Public function to be exposed
        // Passing Name for plugin contract
        public string Name
        {
            get
            {
                return "PlugUnplug Test";
            }
        }

        // Passing DataGridViewRow to and from main
        public DataGridViewRow ToExecute
        {
            get
            {
                return toExecute;
            }
            set
            {
                toExecute = value;
                DataRowProcess();
            }
        }

        // Passing RibbonTab setting to main
        public RibbonTab EquipmentSetting
        {
            get
            {
                return ribbonEquipmentSetting;
            }
        }

        // Passing Picture to main
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
        #endregion

        private void PopulateUI()
        {
            ribbonEquipmentSetting = new RibbonTab();
            ribbonEquipmentSetting.Text = "Equipment Setting";

            #region Test Connection - Populate UI
            RibbonPanel connectionTest = new RibbonPanel();
            connectionTest.Text = "Connection Test";
            ribbonEquipmentSetting.Panels.Add(connectionTest);

            // Get equipment lists            
            Panel pnlEquipment = new Panel();
            pnlEquipment.Size = new Size(100, 55);
            pnlEquipment.AutoScroll = true;
            //pnlEquipment.BackColor = Color.Transparent;

            cbEquipment = new CheckBox[EquipmentList.Count];

            for (int index = 0; index < EquipmentList.Count; index++)
            {
                cbEquipment[index] = new CheckBox();
                cbEquipment[index].Text = EquipmentList[index];
                cbEquipment[index].AutoCheck = false;
                cbEquipment[index].Appearance = Appearance.Button;
                cbEquipment[index].TextAlign = ContentAlignment.MiddleLeft;
                cbEquipment[index].AutoSize = false;
                cbEquipment[index].Size = new Size(80, 20);

                if (TestConnection(EquipmentList[index]))
                    cbEquipment[index].BackColor = Color.Green;
                else
                    cbEquipment[index].BackColor = Color.Red;

                if (pnlEquipment.Controls.Count == 0)
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                else
                {
                    cbEquipment[index].Location = new Point(pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.X,
                        pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.Y + cbEquipment[index].Height);
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                }
            }

            RibbonHost pnlEquipmentHost = new RibbonHost();
            pnlEquipmentHost.HostedControl = pnlEquipment;
            connectionTest.Items.Add(pnlEquipmentHost);

            Button btTestCon = new Button();
            btTestCon.Text = "Test";
            btTestCon.Size = new Size(50, 55);
            btTestCon.FlatStyle = FlatStyle.Popup;
            btTestCon.BackColor = Color.LightBlue;
            btTestCon.Click += new EventHandler(btTestCon_Click);
            RibbonHost btTestConHost = new RibbonHost();
            btTestConHost.HostedControl = btTestCon;
            connectionTest.Items.Add(btTestConHost);
            #endregion

        }

        #region Camera/Mic - Populate UI
        private void PopulateUI_CameraMic()
        {
            #region Camera - Populate UI
            // Panel for Camera
            RibbonPanel cameraSetting = new RibbonPanel();
            cameraSetting.Text = "Camera Setting";
            ribbonEquipmentSetting.Panels.Add(cameraSetting);

            // Picture Box Settings
            pictureBox = new PictureBox();
            pictureBox.Size = new System.Drawing.Size(100, 50);
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
            #endregion

            #region Mic - Populate UI
            // Panel for mic
            RibbonPanel micSetting = new RibbonPanel();
            micSetting.Text = "MIC Setting";
            ribbonEquipmentSetting.Panels.Add(micSetting);

            Panel panelMic1 = new Panel();
            panelMic1.Size = new Size(100, 50);
            //panelMic1.BackColor = Color.Transparent;

            Panel panelMic2 = new Panel();
            panelMic2.Size = new Size(100, 50);
            //panelMic2.BackColor = Color.Transparent;

            // Volume Indicator Progress Bar Settings
            pbVolIndicator = new ProgressBar();
            pbVolIndicator.Size = new Size(100, 20);
            pbVolIndicator.Maximum = 100;
            panelMic1.Controls.Add(pbVolIndicator);

            // Volume Indicator Background Worker Settings
            bwVolIndicator = new BackgroundWorker();
            bwVolIndicator.DoWork += new DoWorkEventHandler(bwVolIndicator_DoWork);
            bwVolIndicator.ProgressChanged += new ProgressChangedEventHandler(bwVolIndicator_ProgressChanged);
            bwVolIndicator.WorkerSupportsCancellation = true;
            bwVolIndicator.WorkerReportsProgress = true;
            bwVolIndicator.RunWorkerAsync();

            // Volume Threshold Track Bar Settings
            tbVolThres = new TrackBar();
            tbVolThres.AutoSize = false;
            tbVolThres.Size = new Size(100, 30);
            tbVolThres.Maximum = 100;
            tbVolThres.TickStyle = TickStyle.None;
            tbVolThres.Scroll += new EventHandler(tbVolThres_Scroll);
            tbVolThres.Location = new Point(panelMic1.Controls[panelMic1.Controls.Count - 1].Location.X,
                    panelMic1.Controls[panelMic1.Controls.Count - 1].Location.Y + pbVolIndicator.Height + 5);
            panelMic1.Controls.Add(tbVolThres);

            // Volume Threshold Label Settings
            lblVolThres = new Label();
            lblVolThres.TextAlign = ContentAlignment.MiddleLeft;
            lblVolThres.BackColor = Color.LightYellow;
            lblVolThres.Size = new Size(100, 20);
            lblVolThres_val(0);
            panelMic2.Controls.Add(lblVolThres);

            // MIC List
            cbMicList = new ComboBox();
            cbMicList.Size = new Size(100, 20);
            cbMicList.SelectedIndexChanged += new EventHandler(cbMicList_SelectedIndexChanged);
            cbMicList.Location = new Point(panelMic2.Controls[panelMic2.Controls.Count - 1].Location.X,
                    panelMic2.Controls[panelMic2.Controls.Count - 1].Location.Y + lblVolThres.Height + 5);
            panelMic2.Controls.Add(cbMicList);

            RibbonHost panelMic1Host = new RibbonHost();
            panelMic1Host.HostedControl = panelMic1;

            RibbonHost panelMic2Host = new RibbonHost();
            panelMic2Host.HostedControl = panelMic2;

            micSetting.Items.Add(panelMic1Host);
            micSetting.Items.Add(panelMic2Host);
            #endregion
        }
        #endregion

        #region Camera/Mic - Initialization
        private void InitCamera_Mic()
        {

            string[] cameras = null, microphone = null;

            // Set Camera resolution
            for (int i = 0; i < Saal.videoDisplay.VideoCapabilities.Length; i++)
            {
                Console.WriteLine(Saal.videoDisplay.VideoCapabilities[i].FrameSize.ToString());
                if (Saal.videoDisplay.VideoCapabilities[i].FrameSize.Height == ResHeight &&
                    Saal.videoDisplay.VideoCapabilities[i].FrameSize.Width == ResWidth)
                {
                    Saal.videoDisplay.VideoResolution = Saal.videoDisplay.VideoCapabilities[i];
                    Console.WriteLine(Saal.videoDisplay.VideoResolution.FrameSize.ToString());
                    break;
                }
            }

            // Init for Camera
            Saal.CAM_GetCAMList(out cameras);
            Saal.CAM_Init(cameras[0]);
            Saal.videoDisplay.NewFrame += new NewFrameEventHandler(video_NewFrame);

            foreach (string cam in cameras)
            {
                cbCamList.Items.Add(cam);
            }
            cbCamList.SelectedIndex = 0;

            // Init for Mic
            Saal.MIC_GetMICList(out microphone);
            Saal.MIC_Init(microphone[0]);

            foreach (string mic in microphone)
            {
                cbMicList.Items.Add(mic);
            }
            cbMicList.SelectedIndex = 0;
        }

        private void video_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Size size = new Size(160, 120);
            Bitmap video = new Bitmap(size.Width, size.Height);

            using (Graphics g = Graphics.FromImage((System.Drawing.Image)video))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage((Bitmap)eventArgs.Frame.Clone(), 0, 0, size.Width, size.Height);
            }

            if (video != null)
            {
                pictureBox.Image = (Bitmap)video.Clone();
                pictureBoxPopup.Image = (Bitmap)video.Clone();
            }

            video.Dispose();
        }
        #endregion

        #region Camera/Mic - Event Handler
        // Open picture box in new form
        private void pictureBox_DoubleClick(object sender, EventArgs e)
        {
            popupBox.BringToFront();
            popupBox.WindowState = FormWindowState.Minimized;
            popupBox.Show();
            popupBox.WindowState = FormWindowState.Normal;
        }

        // Close picture box
        private void popupBox_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            popupBox.Hide();
        }

        // On/Off button event for camera
        private void btOnOffCam_Click(object sender, EventArgs e)
        {
            if (btOnOffCam.Text == "On")
            {
                btOnOffCam.BackColor = Color.PaleVioletRed;
                btOnOffCam.Text = "Off";
                Saal.videoDisplay.Start();
            }
            else
            {
                btOnOffCam.BackColor = Color.LightBlue;
                btOnOffCam.Text = "On";
                Saal.videoDisplay.Stop();
            }
        }

        // Camera list item click event
        private void cbCamList_SelectedIndexChanged(object sender, EventArgs e)
        {
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
            }
        }

        // Mic list item click event
        private void cbMicList_SelectedIndexChanged(object sender, EventArgs e)
        {
            Saal.MIC_Init(cbMicList.SelectedText);
        }

        // Setting threshold value
        private void tbVolThres_Scroll(object sender, EventArgs e)
        {
            lblVolThres_val(tbVolThres.Value);
        }

        // Change threshold label value
        private void lblVolThres_val(int val)
        {
            lblVolThres.Text = "Threshold : " + val;
        }

        // Volume indicator background work
        private void bwVolIndicator_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                if (Saal.MIC_IsAlive() == true)
                {
                    // Get the backgroundworker that raised this event
                    BackgroundWorker worker = sender as BackgroundWorker;

                    worker.ReportProgress(Saal.MIC_Amplitude());
                }
                else
                    pbVolIndicator.Value = 0;

                Thread.Sleep(100);
            }
        }

        // Volume indicator progress bar
        private void bwVolIndicator_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbVolIndicator.Value = e.ProgressPercentage;
        }
        #endregion

        #region Test Connection - Event Handler
        private void btTestCon_Click(object sender, EventArgs e)
        {
            for (int index = 0; index < EquipmentList.Count; index++)
            {
                if (TestConnection(EquipmentList[index]))
                    cbEquipment[index].BackColor = Color.Green;
                else
                    cbEquipment[index].BackColor = Color.Red;
            }
        }

        private bool TestConnection(string Equip)
        {
            Int32 ConStat;

            //Test
            ConStat = 1;

            if (ConStat == 1)
                return true;
            else
                return false;
        }
        #endregion

        #region Click - Event Handler
        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        #endregion

        private void DataRowProcess()
        {
            //Example on how to set toExecute DataGridViewRow
            //toExecute.Cells[9].Value = "OK";
        }
    }
}
