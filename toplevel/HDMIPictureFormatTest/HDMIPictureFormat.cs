using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Security.Cryptography.X509Certificates;
using System.Timers;
using System.Windows.Forms;
using Emgu.CV.Structure;
using PluginContracts;
using PopupControl;
using SAAL;
using ComboBox = System.Windows.Forms.ComboBox;
using Image = System.Drawing.Image;
//bomb

#region Camera/Mic include files

using System.Drawing;
using System.Drawing.Imaging;
using AForge.Video;
using AForge.Video.DirectShow;
using AForge.Imaging;
using System.ComponentModel;
using System.Threading;
using System.Text.RegularExpressions;
using mainUI;
using Emgu.CV;

#endregion

namespace Test
{
    public class HDMIPictureFormat : IPlugin
    {
        #region DEBUG properties
        private Boolean testUARTBypass = false;
        #endregion

        #region Helper Classes and Struct
        /// <summary>
        /// row data collection. input only
        /// </summary>

        private class RowIndexData
        {
            /// <summary>
            /// row, cmd, raw
            /// </summary>
            public Triplet<int, string, string> HDMIFormat;
            /// <summary>
            /// row, cmd, raw
            /// </summary>
            public Triplet<int, string, string> frequency;
            /// <summary>
            /// row, cmd, raw
            /// </summary>
            public Triplet<int, string, string> colorType;
            /// <summary>
            /// row, cmd, raw
            /// </summary>
            public KeyValuePair<int, string> operation;
            public KeyValuePair<int, string> expected;


            //var samplingIndex = "";

            public RowIndexData()
            {
                HDMIFormat = new Triplet<int, string, string>((int)RowIndex.HdmiFormat, "", "");
                frequency = new Triplet<int, string, string>((int)RowIndex.Frequency, "", "");
                colorType = new Triplet<int, string, string>((int)RowIndex.ColorSpace, "", "");

                operation = new KeyValuePair<int, string>((int)RowIndex.Operation, "");
                expected = new KeyValuePair<int, string>((int)RowIndex.Expected, "");
            }

            public RowIndexData(RowIndexData other)
            {
                HDMIFormat = new Triplet<int, string, string>(other.HDMIFormat);
                frequency = new Triplet<int, string, string>(other.frequency);
                colorType = new Triplet<int, string, string>(other.colorType);

                operation = other.operation;
                expected = other.expected;
            }
        }

        private class Triplet<A, B, C>
        {
            public Triplet(A a, B b, C c)
            {
                vA = a;
                vB = b;
                vC = c;
            }
            public A vA { get; set; }
            public B vB { get; set; }
            public C vC { get; set; }

            public Triplet(Triplet<A, B, C> other)
            {
                vA = other.vA;
                vB = other.vB;
                vC = other.vC;
            }
        }

        private struct DeviceSettings
        {
            /// <summary>
            /// eg) COM1 / 192.168.1.0
            /// </summary>
            public String conValue;
            /// <summary>
            /// EN_CON_METHOD (Uart/Lan)
            /// </summary>
            public EN_CON_METHOD conMethod;
            /// <summary>
            /// connected?
            /// </summary>
            public Boolean conStatus;
        }

        private class ProgressBarStuff
        {
            private const int Period1stRow = 15; // in secs
            private Stopwatch rowStopwatch = new Stopwatch();
            private Stopwatch totalStopwatch = new Stopwatch();
            private BackgroundWorker bgWorker;
            private ProgressBar progressBar = new ProgressBar();
            private System.Timers.Timer timer = new System.Timers.Timer();
            private System.Timers.Timer timerForceRefresh = new System.Timers.Timer();

            private List<Triplet<int, int, string>> rowindexTimeName = new List<Triplet<int, int, string>>();
            private int totalEstTime = 0;
            private bool bgCompleted = false;

            public string logger = string.Empty;
            public bool CancellationPending = false;

            public Bitmap camLiveImage
            {
                set
                {
                    progressBar.camPictureBox.Image?.Dispose();
                    progressBar.camPictureBox.Image = value;
                }
            }

            public Bitmap refImage
            {
                set
                {
                    progressBar.refPictureBox.Image?.Dispose();
                    progressBar.refPictureBox.Image = value;
                }
                get
                {
                    return (Bitmap)progressBar.refPictureBox.Image;
                }
            }

            public ProgressBarStuff(ref BackgroundWorker bgW)
            {
                bgWorker = bgW;

                progressBar.Closing += (sender, args) =>
                {
                    args.Cancel = true;
                    progressBar.Hide();
                };

                progressBar.button1.Click += (sender, args) =>
                {
                    bgWorker.CancelAsync();
                    CancellationPending = true;
                    Reset();
                };

                bgWorker.RunWorkerCompleted += (sender, args) =>
                {
                    bgCompleted = true;
                    timer.Enabled = false;
                };

                timer.Interval = 997;
                timerForceRefresh.Interval = 100;

                timer.Elapsed += (sender, args) =>
                {
                    var elapsed = (int)((double)totalStopwatch.ElapsedMilliseconds / 1000);
                    Console.WriteLine("Total estimated (s): " + totalEstTime);
                    var remaining = totalEstTime - elapsed;
                    Console.WriteLine("Total remaining (s): " + remaining);
                    Console.WriteLine("Time elapsed (s): " + elapsed);
                    var percent = (double)elapsed / totalEstTime * 100;
                    percent = percent > 99 ? 99 : percent;
                    remaining = remaining < 1 ? 1 : remaining;

                    var someString = remaining.ToString();
                    if (bgWorker != null && !bgCompleted)
                    {
                        bgWorker.ReportProgress((int)percent, someString);
                    }
                };

                timerForceRefresh.Elapsed += (sender, args) =>
                {
                    progressBar.camPictureBox.Invoke(new Action((() => progressBar.camPictureBox.Refresh())));

                    progressBar.label2.Invoke(new Action(() =>
                    {
                        progressBar.label2.Text = logger;
                        progressBar.label2.Refresh();
                    }));
                    

                };

            }

            public void Reset()
            {
                timer.Enabled = false;
                timer.Stop();
                timerForceRefresh.Enabled = false;
                timerForceRefresh.Stop();

                progressBar.total_time.Text = ".....";
                progressBar.progressBar1.Value = 0;
                progressBar.Close();
                rowindexTimeName.Clear();
                totalStopwatch.Stop();
                totalStopwatch.Reset();
                totalEstTime = 0;
            }

            public string BeginOneIteration(int iterationIndex)
            {
                rowStopwatch.Start();

                foreach (var i in rowindexTimeName)
                {
                    if (i.vA == iterationIndex)
                    {
                        return i.vC;
                    }
                }

                return "Default";
            }

            public void StopOneIteration(int iterationIndex)
            {
                rowStopwatch.Stop();
                var rowTime = rowStopwatch.ElapsedMilliseconds;
                rowStopwatch.Reset();

                var expectedTime = 0;
                foreach (var i in rowindexTimeName)
                {
                    if (i.vA == iterationIndex)
                    {
                        expectedTime = i.vB * 1000;
                        break;
                    }
                }

                var differenceInTime = rowTime - expectedTime;
                totalEstTime += (int)((double)differenceInTime / 1000);
            }

            public void ProgressBarChangedEvent(object sender, ProgressChangedEventArgs e)
            {
                progressBar.progressBar1.Value = e.ProgressPercentage;

                var str0 = e.UserState as string;
                var remaining = Convert.ToDouble(str0);
                var remainingTimeSpan = TimeSpan.FromSeconds(remaining);

                var str = remaining > 60
                    ? (remainingTimeSpan.Minutes + " min, " + remainingTimeSpan.Seconds + " sec")
                    : remaining + " sec";
                  progressBar.total_time.Text = str;
                progressBar.Refresh();
            }

            public void Show(object selection, object datagrid, object owner)
            {
                var list = selection as List<int>;
                var data = datagrid as DataGridViewRowCollection;
                foreach (var i in list)
                {
                    var row = data[i];
                    var str = row.Cells[(int)RowIndex.Operation].Value.ToString();
                    var name = row.Cells[0].EditedFormattedValue.ToString();

                    // 1.1 to add little leverage
                    //8sec here for sendallcommand totalt
                    if (str.Contains("The image is seen for 3 minutes."))
                        rowindexTimeName.Add(new Triplet<int, int, string>(i,
                            (int)(((ImageHandler.SimilarityTestPeriod * 1.1) / 1000) + 8), name));
                    else
                        rowindexTimeName.Add(new Triplet<int, int, string>(i,
                            (int)((Period1stRow * 1.1) + 8), name));
                }

                foreach (var i in rowindexTimeName)
                    totalEstTime += i.vB;

                var ownerForm = owner as Form;
                ownerForm.Invoke(new Action(() =>
                {
                    progressBar.StartPosition = FormStartPosition.CenterScreen;
                    progressBar.Show(ownerForm);

                    //timer.SynchronizingObject = ownerForm;
                    totalStopwatch.Start();
                    timer.Enabled = true;
                    timer.Start();

                    timerForceRefresh.Enabled = true;
                    timerForceRefresh.Start();
                }));
            }
        }
        #endregion

        #region Enum
        private enum EN_CON_METHOD
        {
            CON_UART,
            CON_LAN,
            CON_MAX
        }

        private enum RowIndex
        {
            HdmiFormat = 2,
            Frequency = 4,
            ColorSpace = 5,
            Operation = 7,
            Expected = 8,
            Result = 12,
            Tester = 14,
            Remarks = 16
        }

        #endregion

        #region Member: Generic
        private RibbonTab equipmentSetting;
        private Bitmap _blankImg;
        private List<string> _colorList;
        private RowIndexData _rowIndexData = new RowIndexData();
        private RowIndexData _previousrowIndexData = new RowIndexData();
        private string _executionTime;
        private readonly SAAL_Interface Saal = new SAAL_Interface();
        private DataGridViewRow toExecute;
        private static int blackPixelThreshold = 30;
        private static int binaryThreshold = 70;
        private ProgressBarStuff progressBarStuff;
        private TextBox[] testerInfos = new TextBox[2];
        private string previousColor = "";
        #endregion

        #region Member: Device 
        //private Dictionary<string, string> _tg4559List;
        private CheckBox[] cbEquipment;
        private Button btTestCon;
        private Popup devicePopup;
        //private string _selectedEquipment = "";
        //private EquipmentStatus _equipmentStatus = EquipmentStatus.NotTested;
        private Dictionary<string, DeviceSettings> _deviceSettings;
        #endregion

        #region Member: Camera

        // Parameters for Camera
        private Bitmap currentBitmap;
        private ComboBox cbCamList;
        private Button btOnOffCam;
        private Button btCamCalib;
        private readonly int ResWidth = 800;
        private readonly int ResHeight = 600;
        #endregion

        #region Member: Public properties
        public string Name { get { return "HDMIPictureFormat Test"; } }
        public DataGridViewRow ToExecute { get { return toExecute; } set { toExecute = value; DataRowProcess(); } }
        public RibbonTab EquipmentSetting { get { init(); return equipmentSetting; } }
        public PictureBox Picture { get; set; }
        public bool Busy { get; private set; }
        public bool DoExecute { get; private set; } = true;
        public BeforeExecute BeforeExecute { get; }
        public AfterExecute AfterExecute { get; }
        public CellClickEvent ClickEvent { get; }
        public BeforeExecuteButAfterSelection BeforeExecuteButAfterSelection { get; }
        public ProgressBarChangedEvent ProgressBarChangedEvent { get; }
        #endregion

        // Constructor
        public HDMIPictureFormat()
        {
            BeforeExecute = InvokeBeforeExecute;
            AfterExecute = InvokeAfterExecute;
            ClickEvent = DataGridView_CellClick;
            BeforeExecuteButAfterSelection = InvokeBeforeExecuteButAfterSelection;
            ProgressBarChangedEvent = InvokeProgressBarChangedEvent;
        }

        private void init()
        {
            // Add color spaces
            _colorList = new List<string>();
            _colorList.Add("RGB");
            _colorList.Add("Y444");
            _colorList.Add("Y422");

            InitDevice();
            PopulateUIDevice();
            PopulateUICamera();
            PopulateDebugUI();
            PopulateTestThresholdUI();
            PopulateUI_Extra();

            InitCamera();

            Busy = false;

            ImageHandler.IsEnOCR = false;
            ImageHandler.IsEnBlurryCheck = true;
            ImageHandler.IsEnSimilarityCheck = true;
            ImageHandler.IsEnPictureFormat = false;
        }

        private void PopulateUIDevice()
        {
            devicePopup = new Popup(new DeviceSettingPopup());
            devicePopup.FocusOnOpen = false;
            devicePopup.AutoClose = true;
            devicePopup.Closing += (s, o) => {
                var dVP = devicePopup.Content as DeviceSettingPopup;
                dVP.UpdateCurrentLabelValue();

            };

            equipmentSetting = new RibbonTab();
            equipmentSetting.Text = "HDMI Picture Test";
            var connectionTest = new RibbonPanel();
            connectionTest.Text = "Connection Test";
            equipmentSetting.Panels.Add(connectionTest);

            // Get equipment lists            
            var pnlEquipment = new Panel();
            pnlEquipment.Size = new Size(100, 55);
            pnlEquipment.AutoScroll = true;
            //pnlEquipment.BackColor = Color.Transparent;

            cbEquipment = new CheckBox[_deviceSettings.Count];
            var index = 0;
            foreach (var i in _deviceSettings)
            {
                cbEquipment[index] = new CheckBox();
                cbEquipment[index].Text = i.Key;
                cbEquipment[index].AutoCheck = false;
                cbEquipment[index].Appearance = Appearance.Button;
                cbEquipment[index].TextAlign = ContentAlignment.MiddleLeft;
                cbEquipment[index].AutoSize = false;
                cbEquipment[index].Size = new Size(80, 20);
                cbEquipment[index].UpdateStatusBar_MouseEvents();
                cbEquipment[index].Click += (sender, args) =>
                {
                    var cBox = sender as CheckBox;

                    if (cBox.BackColor == Color.Green)
                    {
                        foreach (var j in cbEquipment)
                        {
                            bool k = cBox.Text.Equals(j.Text);
                            var dS = _deviceSettings[j.Text];
                            j.BackColor = k && dS.conStatus ? Color.Green : Color.Red;
                            dS.conStatus = k && dS.conStatus;
                            _deviceSettings[j.Text] = dS;
                        }
                    }

                    if (devicePopup.Visible)
                        devicePopup.Hide();

                    var dVP = devicePopup.Content as DeviceSettingPopup;
                    dVP.label1.Text = cBox.Text;
                    devicePopup.Show(cBox, new Rectangle(cBox.Width, 0, 0, 0));
                };

                if (pnlEquipment.Controls.Count == 0)
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                else
                {
                    cbEquipment[index].Location =
                        new Point(pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.X,
                            pnlEquipment.Controls[pnlEquipment.Controls.Count - 1].Location.Y +
                            cbEquipment[index].Height);
                    pnlEquipment.Controls.Add(cbEquipment[index]);
                }
                index++;
            }

            var pnlEquipmentHost = new RibbonHost();
            pnlEquipmentHost.HostedControl = pnlEquipment;
            connectionTest.Items.Add(pnlEquipmentHost);

            btTestCon = new Button();
            btTestCon.Text = "Test";
            btTestCon.Size = new Size(50, 55);
            btTestCon.FlatStyle = FlatStyle.Popup;
            btTestCon.BackColor = Color.LightBlue;
            btTestCon.Click += btTestCon_Click;
            btTestCon.UpdateStatusBar_MouseEvents();
            var btTestConHost = new RibbonHost();
            btTestConHost.HostedControl = btTestCon;
            connectionTest.Items.Add(btTestConHost);
            devicePopup.Closing += (sender, e) =>
            {
                if (e.CloseReason == ToolStripDropDownCloseReason.CloseCalled)
                {
                    ((Control)sender).Hide();
                    e.Cancel = true;
                }
            };
        }

        private void PopulateUI_Extra()
        {
            RibbonPanel recordSetting = new RibbonPanel();
            recordSetting.Text = "Record";
            equipmentSetting.Panels.Add(recordSetting);

            Panel panelExtra = new Panel();
            panelExtra.Size = new Size(155, 50);

            Label lblTester = new Label();
            lblTester.Text = "Tester :";
            lblTester.TextAlign = ContentAlignment.MiddleRight;
            lblTester.Size = new Size(50, 20);
            panelExtra.Controls.Add(lblTester);

            testerInfos[0] = new TextBox();
            testerInfos[0].Size = new Size(100, 20);
            testerInfos[0].Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblTester.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(testerInfos[0]);

            Label lblVer = new Label();
            lblVer.Text = "Version :";
            lblVer.TextAlign = ContentAlignment.MiddleRight;
            lblVer.Size = new Size(50, 20);
            lblVer.Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 2].Location.X,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y + lblTester.Height + 5);
            panelExtra.Controls.Add(lblVer);

            testerInfos[1] = new TextBox();
            testerInfos[1].Size = new Size(100, 20);
            testerInfos[1].Location = new Point(panelExtra.Controls[panelExtra.Controls.Count - 1].Location.X + lblVer.Width + 2,
                    panelExtra.Controls[panelExtra.Controls.Count - 1].Location.Y);
            panelExtra.Controls.Add(testerInfos[1]);

            RibbonHost panelExtraHost = new RibbonHost();
            panelExtraHost.HostedControl = panelExtra;

            recordSetting.Items.Add(panelExtraHost);
        }

        #region Camera - Populate UI
        private void PopulateUICamera()
        {
            // Panel for Camera
            var cameraSetting = new RibbonPanel();
            cameraSetting.Text = "Camera Setting";
            equipmentSetting.Panels.Add(cameraSetting);

            // Picture Box Settings
            Picture = new PictureBox();
            Picture.Image = null;
            Picture.Size = new Size(100, 50);
            Picture.SizeMode = PictureBoxSizeMode.StretchImage;
            Picture.UpdateStatusBar_MouseEvents();
            var pictureBoxHost = new RibbonHost();
            pictureBoxHost.HostedControl = Picture;
            pictureBoxHost.ToolTip = "Click to enlarge";

            var panelCam = new Panel();
            panelCam.Size = new Size(100, 50);

            // On/Off Button Settings
            btOnOffCam = new Button();
            btOnOffCam.Text = "On";
            btOnOffCam.Size = new Size(100, 20);
            btOnOffCam.Enabled = true;
            btOnOffCam.FlatStyle = FlatStyle.Popup;
            btOnOffCam.BackColor = Color.LightBlue;
            btOnOffCam.Click += btOnOffCam_Click;
            btOnOffCam.UpdateStatusBar_MouseEvents();
            panelCam.Controls.Add(btOnOffCam);
            //btOnOffCam.PerformClick();

            // Camera List
            cbCamList = new ComboBox();
            cbCamList.Size = new Size(100, 20);
            cbCamList.SelectedIndexChanged += cbCamList_SelectedIndexChanged;
            cbCamList.Location = new Point(panelCam.Controls[panelCam.Controls.Count - 1].Location.X,
                panelCam.Controls[panelCam.Controls.Count - 1].Location.Y + btOnOffCam.Height + 5);
            panelCam.Controls.Add(cbCamList);

            var panelCamHost = new RibbonHost();
            panelCamHost.HostedControl = panelCam;

            btCamCalib = new Button();
            btCamCalib.Text = "Camera Calibration";
            btCamCalib.Size = new Size(80, 50);
            btCamCalib.Enabled = true;
            btCamCalib.FlatStyle = FlatStyle.Popup;
            btCamCalib.BackColor = Color.LightBlue;
            btCamCalib.Click += (sender, args) =>
            {
                if (btOnOffCam.Text == "Off")
                {
                    ImageHandler.OpenCameraCalibForm();
                }
                else
                    MessageBox.Show("Please on the camera");
            };

            var cameraCalibrateRibbonHost = new RibbonHost();
            cameraCalibrateRibbonHost.HostedControl = btCamCalib;

            // Add items to Camera Setting Panel
            cameraSetting.Items.Add(pictureBoxHost);
            cameraSetting.Items.Add(panelCamHost);
            cameraSetting.Items.Add(cameraCalibrateRibbonHost);
        }
        #endregion

        #region For Debug - Populate UI
        private void PopulateDebugUI()
        {
            //Populate test image picturebox
            var debugPanel = new RibbonPanel();
            debugPanel.Text = "DEBUG STUFF";

#if DEBUG   // Visible only in DEBUGMODE
            equipmentSetting.Panels.Add(debugPanel);
#endif

            var emulateDevice = new CheckBox();
            emulateDevice.Text = "Device True";
            emulateDevice.CheckedChanged += (sender, args) =>
            {
                var i = sender as CheckBox;
                testUARTBypass = i.Checked;
            };

            var buttonShowForm = new Button();
            buttonShowForm.Text = "ALGO";
            buttonShowForm.Size = new Size(80, 20);
            buttonShowForm.Location = new Point(emulateDevice.Location.X, emulateDevice.Location.Y + emulateDevice.Height + 1);
            buttonShowForm.Click += (sender, args) =>
            {
                ImageHandler.StopCamera();
                var hoho = new ImageAlgoTester();
                hoho.Show();
            };

            var debugChildPanel = new Panel();
            debugChildPanel.Controls.Add(emulateDevice);
            debugChildPanel.Controls.Add(buttonShowForm);
            debugChildPanel.Size = new Size(100, 50);

            var debugRibbonHost = new RibbonHost();
            debugRibbonHost.HostedControl = debugChildPanel;
            debugPanel.Items.Add(debugRibbonHost);
        }
        #endregion

        #region Test parameter - PopulateUI
        private void PopulateTestThresholdUI()
        {
            var testPanel = new RibbonPanel();
            testPanel.Text = "Test Parameter";
            equipmentSetting.Panels.Add(testPanel);

            var testPanelLabel = new Panel();
            testPanelLabel.Size = new Size(100, 50);

            var binaryButton = new Button();
            binaryButton.Text = "Binary Test";
            binaryButton.Size = new Size(80, 20);
            binaryButton.Location = new Point(10, 2);
            binaryButton.Click += (sender, args) =>
            {
                var Popup = new Popup(new BinaryTestPropPopup());
                var bTPP = Popup.Content as BinaryTestPropPopup;
                Popup.Opened += (o, eventArgs) =>
                {
                    bTPP.textBox1.Text = blackPixelThreshold.ToString();
                    bTPP.textBox2.Text = binaryThreshold.ToString();
                };
                Popup.Closing += (o, eventArgs) =>
                {
                    try
                    {
                        blackPixelThreshold = Convert.ToInt32(bTPP.textBox1.Text);
                        binaryThreshold = Convert.ToInt32(bTPP.textBox2.Text);
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(exp.Message);
                        eventArgs.Cancel = true;
                    }
                };
                bTPP.button1.Click += (o, eventArgs) =>
                {
                    blackPixelThreshold = 30;
                    binaryThreshold = 70;

                    bTPP.textBox1.Text = blackPixelThreshold.ToString();
                    bTPP.textBox2.Text = binaryThreshold.ToString();
                };


                Popup.Show((Control)sender, new Rectangle(((Control)sender).Width, 0, 0, 0));
            };

            Action resetSimilarityParam = () =>
            {
                ImageHandler.BlurryTestThreshold = 100;
                ImageHandler.SimilarityTestInterval = 500;
                ImageHandler.SimilarityTestPeriod = 180000; // 3min
                ImageHandler.SimilarityTestSSIMThreshold = 100;
                ImageHandler.SimilarityTestPSNRThreshold = 34;
            };

            resetSimilarityParam();

            var simButton = new Button();
            simButton.Text = "Same+Blur Test";
            simButton.Size = new Size(80, 20);
            simButton.Location = new Point(10, 28);
            simButton.Click += (sender, args) =>
            {
                var Popup = new Popup(new SimilarityTestPropPopup());
                var sTPP = Popup.Content as SimilarityTestPropPopup;
                Popup.Opened += (o, eventArgs) =>
                {
                    sTPP.blurThreshold.Text = ImageHandler.BlurryTestThreshold.ToString();
                    sTPP.simPeriodInt.Text = ((double)ImageHandler.SimilarityTestInterval / 1000).ToString();
                    sTPP.simPeriodTotal.Text = ((double)ImageHandler.SimilarityTestPeriod / 1000).ToString();
                    sTPP.simThresMSSIM.Text = ImageHandler.SimilarityTestSSIMThreshold.ToString();
                    sTPP.simThresPSNR.Text = ImageHandler.SimilarityTestPSNRThreshold.ToString();
                };
                Popup.Closing += (o, eventArgs) =>
                {
                    Func<string, int, int> convertToInt = (str, multi) =>
                    {
                        var ret = 0;
                        try
                        {
                            var dbl = Convert.ToDouble(str);
                            dbl *= multi;
                            ret = (int)Math.Ceiling(dbl);
                        }
                        catch (Exception exp)
                        {
                            MessageBox.Show(exp.Message);
                            eventArgs.Cancel = true;
                            return ret;
                        }
                        return ret;
                    };

                    ImageHandler.BlurryTestThreshold = convertToInt(sTPP.blurThreshold.Text, 1);
                    ImageHandler.SimilarityTestInterval = convertToInt(sTPP.simPeriodInt.Text, 1000);
                    ImageHandler.SimilarityTestPeriod = convertToInt(sTPP.simPeriodTotal.Text, 1000);
                    ImageHandler.SimilarityTestSSIMThreshold = convertToInt(sTPP.simThresMSSIM.Text, 1);
                    ImageHandler.SimilarityTestPSNRThreshold = convertToInt(sTPP.simThresPSNR.Text, 1);
                };
                sTPP.button1.Click += (o, eventArgs) =>
                {
                    resetSimilarityParam();

                    sTPP.blurThreshold.Text = ImageHandler.BlurryTestThreshold.ToString();
                    sTPP.simPeriodInt.Text = ((double)ImageHandler.SimilarityTestInterval / 1000).ToString();
                    sTPP.simPeriodTotal.Text = ((double)ImageHandler.SimilarityTestPeriod / 1000).ToString();
                    sTPP.simThresMSSIM.Text = ImageHandler.SimilarityTestSSIMThreshold.ToString();
                    sTPP.simThresPSNR.Text = ImageHandler.SimilarityTestPSNRThreshold.ToString();
                };

                Popup.Show((Control)sender, new Rectangle(((Control)sender).Width, 0, 0, 0));
            };

            testPanelLabel.Controls.Add(binaryButton);
            testPanelLabel.Controls.Add(simButton);

            var testPanelRibbonHost = new RibbonHost();
            testPanelRibbonHost.HostedControl = testPanelLabel;
            testPanel.Items.Add(testPanelRibbonHost);
        }

        #endregion

        #region Execution Handler

        private bool SendCommandTG4559(string i_strCmd)
        {
            bool retVal = false;
            string retStr = string.Empty;
            string conValue = string.Empty;
            bool isUART = false;

            foreach (var i in _deviceSettings)
            {
                if (i.Value.conStatus)
                {
                    conValue = i.Value.conValue;
                    isUART = i.Value.conMethod == EN_CON_METHOD.CON_UART;
                    break;
                }
            }

            if (conValue.Equals(string.Empty))
                return false;

            if (isUART)
            {
                Saal.TG45_59_Setup(conValue, 115200);
                retStr = Saal.TG45_59_SendCmd(i_strCmd + "\n");
                retVal = retStr.Contains("OK");
                Saal.TG45_59_ClosePort();
            }
            else
            {
                Saal.TG45_59_Lan_Setup(conValue);
                retStr = Saal.TG45_59_Lan_SendCmd(i_strCmd + "\n");
                retVal = retStr.Contains("OK");
                Saal.TG45_59_Lan_Close();
            }

            return retVal;
        }

        private string SendAllCommand()
        {
            // set 8 sec
            int sleepTime = 8000 / 4;

            var hdmiFormat = _rowIndexData.HDMIFormat.vB;
            var freqIndex = _rowIndexData.frequency.vB;

            var colorIndex = Regex.IsMatch(_rowIndexData.colorType.vB, @"RGB") ? "0" : "1";
            var samplingIndex = !Regex.IsMatch(_rowIndexData.colorType.vB, @"4:2:2") ? "0" : "1";

            //1. set resolution
            // check if the previous is the same, then no need to change
            if (!_previousrowIndexData.HDMIFormat.vB.Equals(_rowIndexData.HDMIFormat.vB))
            {
                foreach (var i in _deviceSettings)
                {
                    if (i.Value.conStatus)
                    {
                        if (!Saal.TG45_59_SetFormat(hdmiFormat,
                            i.Value.conValue,
                            i.Value.conMethod == EN_CON_METHOD.CON_UART))
                            return "Set resolution failed";

                        break;
                    }
                }

            }



            //2. set frequency (60/1 or 59.94/1.0001)
            // check if the previous is the same, then no need to change
            Thread.Sleep(sleepTime);

            if (!_previousrowIndexData.frequency.vB.Equals(_rowIndexData.frequency.vB))
            {
                string strCmd = "mfrq " + freqIndex;
                if (!SendCommandTG4559(strCmd))
                    return "Set frequency modifier failed";
            }



            //3. set color (YCbCr or RGB)
            Thread.Sleep(sleepTime);
            if (!_previousrowIndexData.colorType.vB.Equals(_rowIndexData.colorType.vB))
            {
                string strCmd = "chgcol " + colorIndex;
                if (!SendCommandTG4559(strCmd))
                    return "Set color failed";
            }

            //4. set sampling (444 pr 442)
            Thread.Sleep(sleepTime);

            if (!_previousrowIndexData.colorType.vB.Equals(_rowIndexData.colorType.vB))
            {
                //note: assume bit is 8bit (0, 2nd param)
                string strCmd = "003_vid " + samplingIndex + "," + "0" + "," + "1";
                if (!SendCommandTG4559(strCmd))
                    return "Set color sampling failed";
            }

            //5. set color bar
            Thread.Sleep(sleepTime);


                string strCmd1 = "sg 1,0,4";
                if (!SendCommandTG4559(strCmd1))
                    return "Set colobar failed";
   

            return "OK";
        }

        #endregion

        private void TestConnection()
        {
            foreach (var i in _deviceSettings.ToList())
            {
                var value = _deviceSettings[i.Key];

                var dVP = devicePopup.Content as DeviceSettingPopup;
                var deviceProp = dVP.switchDictionary[i.Key];

                value.conMethod = deviceProp.IsUarTorLan == DeviceSettingPopup.UARTorLAN.UART ?
                    EN_CON_METHOD.CON_UART : EN_CON_METHOD.CON_LAN;

                switch (value.conMethod)
                {
                    case EN_CON_METHOD.CON_UART:
                        {
                            if (testUARTBypass)
                            {
                                value.conValue = deviceProp.comPort;
                                value.conStatus = true;
                            }
                            else
                            {
                                string resPort = Saal.TG45_59_TestPerCon(i.Key);
                                value.conValue = resPort;
                                value.conStatus = resPort != "";
                            }

                            _deviceSettings[i.Key] = value;
                            break;
                        }
                    case EN_CON_METHOD.CON_LAN:
                        {
                            if (testUARTBypass) // for test
                            {
                                value.conValue = deviceProp.ipAddress;
                                value.conStatus = true;
                            }
                            else
                            {
                                value.conValue = deviceProp.ipAddress;
                                value.conStatus = Saal.TG45_59_Lan_TestPerCon(value.conValue);
                            }

                            _deviceSettings[i.Key] = value;
                            break;
                        }
                }

                _deviceSettings[i.Key] = value;
            }

            // highlight buttons
            var valid_single = new List<bool>();
            foreach (var j in cbEquipment)
            {
                var ok = _deviceSettings[j.Text].conStatus;
                j.BackColor = ok ? Color.Green : Color.Red;

                if (ok) valid_single.Add(ok);
            }

            if (valid_single.Count > 1)
                MessageBox.Show("Multiple valid device detected! \nPlease select only one valid device by clicking at the button");
            else if (valid_single.Count == 0)
                MessageBox.Show("No valid device. \nPlease make sure devices connected properly and run the connection test again");
        }


        /// <summary>
        /// return folderpath with executiontime
        /// </summary>
        /// <param name="images"></param>
        /// <returns></returns>
        private string[] _SaveImagesToOutputFolder(Bitmap[] images, string rowName)
        {
            // create folder in exe level
            var tempPath = Path.GetTempPath();

            // create output folder
            var outputFolder = tempPath + "FunaiAutomation\\output";
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // create output/testparent folder
            var parentImagesFolder = outputFolder + "\\CapturedImagesByRow_HDMIPictureTest";
            if (!Directory.Exists(parentImagesFolder))
                Directory.CreateDirectory(parentImagesFolder);

            // create sub folder


            var imagesFolder = parentImagesFolder + "\\[" + rowName + "]";
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            for (int i = 0; i < images.Length; ++i)
            {
                var imagePath = imagesFolder + "\\" + _executionTime + (i + 1) + ".png";
                if (File.Exists(imagePath))
                    File.Delete(imagePath);

                images[i].Save(imagePath, ImageFormat.Png);
            }

            return new[] { imagesFolder, _executionTime };
        }

        private void DataRowProcess()
        {
            if (progressBarStuff.CancellationPending)
                return;

            var rowName = progressBarStuff.BeginOneIteration(toExecute.Index);

            _previousrowIndexData = new RowIndexData(_rowIndexData);

            if (!CollectRowIndexData())
            {
                MessageBox.Show("Invalid Input.\nPlease select from row with valid input and extend from there", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //send commands
            if (!testUARTBypass)
            {
                var result = SendAllCommand();
                if (!result.Equals("OK"))
                {
                    MessageBox.Show(result + "\n" + "Test #" + rowName,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var currentImage = new Bitmap(ImageHandler.CurrentFrame);

            Action<bool, Bitmap[], Bitmap> commonAction = (result, bitmapArray, bitmapCell) =>
            {
                toExecute.Cells[(int)RowIndex.Result].Value = result? "OK" : "NG1";
                var imageFolderPath = _SaveImagesToOutputFolder(bitmapArray, rowName);

                var resizedImage = new Bitmap(bitmapCell,
                    toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Size);
                toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Value = resizedImage;
                toExecute.Cells[toExecute.DataGridView.ColumnCount - 1].Tag = imageFolderPath;
            };

            if (_rowIndexData.operation.Value.Contains("The displayed image is confirmed") &&
                _rowIndexData.expected.Value.Contains("The color space is normal"))
            {
                var str = "Test #[" + rowName + "] : Checking if screen is black of not...";
                progressBarStuff.logger = str;

                // wait for TV to start
                Thread.Sleep(10000);

                //- check if image is displayed. - compare with
                var binaryImage = ImageHandler.GetBinarialImage_Global(currentImage, binaryThreshold);
                var blackPercent = (int)ImageHandler.GetBlackPixelPercentage(binaryImage);

                Thread.Sleep(1000);

                var refImage = new Bitmap(progressBarStuff.refImage);
                refImage.Tag = progressBarStuff.refImage.Tag;

                ImageHandler.drawWatermark(ref refImage, @"Reference image");
                ImageHandler.drawWatermark(ref refImage, refImage.Tag as string, 18,
                    new Size(0, refImage.Height - 20));

                var result = blackPercent > blackPixelThreshold;
                var conditionalText = result? @"OK " : @"NOT OK ";

                ImageHandler.drawWatermark(ref binaryImage,
                     conditionalText + "image from Binary Test : ( " + blackPercent + "% Black pixels )");

                commonAction(result, new[] { refImage, binaryImage }, binaryImage);

                refImage.Dispose();
                currentImage.Dispose();
                binaryImage.Dispose();
            }
            else if (_rowIndexData.operation.Value.Contains("The image is seen for 3 minutes.") &&
                _rowIndexData.expected.Value.Contains("There is neither disorder nor a blur of the image (3 minutes)"))
            {
                //- check if image is seen 3 minutes 
                // this check for 3 minutes, 100ms threshold.
                progressBarStuff.logger = "Test #[" + rowName + "] : Checking image for blurriness and disorder...";

                Thread.Sleep(1000);

                var problematicBitmap = new Bitmap(ResWidth, ResHeight);
                var errorInfo = new object();
                var processInfo = new object();

                var testResult = ImageHandler.RunSimilarityAndBlurryCheck(currentImage, ref problematicBitmap,
                    ref errorInfo, ref processInfo, ref progressBarStuff.CancellationPending);
                //- check no blurry, no disorder. use blurry algo, and compare image

                if (progressBarStuff.CancellationPending)
                {
                    currentImage.Dispose();
                    problematicBitmap.Dispose();
                    return;
                }

                var bools = errorInfo as bool[];

                var str = (!bools[0]) ? @"Similarity Test" : @"Blurry Test";
                var info = processInfo as int[];

                var refImage = new Bitmap(progressBarStuff.refImage);
                refImage.Tag = progressBarStuff.refImage.Tag;

                ImageHandler.drawWatermark(ref refImage, @"Reference image");
                ImageHandler.drawWatermark(ref refImage, refImage.Tag as string, 18, new Size(0, refImage.Height - 20));
                ImageHandler.drawWatermark(ref problematicBitmap, problematicBitmap.Tag as string, 18, new Size(0, problematicBitmap.Height - 20));

                var conditionalText = testResult ? @"OK " : @"NOT OK ";

                if (!bools[0])
                    ImageHandler.drawWatermark(ref problematicBitmap,
                        conditionalText + @"image from Similarity Test: " + "( P:" + info[0] + " ; M:" + info[1] + "," + info[2] + "," + info[3] + " )");
                else
                    ImageHandler.drawWatermark(ref problematicBitmap,
                        conditionalText + @"image from Blurry Test: " + "( " + info[4] + " )");

                commonAction(testResult, new[] { refImage, problematicBitmap }, problematicBitmap);
                refImage.Dispose();

                currentImage.Dispose();
                problematicBitmap.Dispose();
            }

            testerInfos[0].Invoke(new Action(() =>
            {
                toExecute.Cells[(int)RowIndex.Tester].Value = testerInfos[0].Text;
            }));
            testerInfos[1].Invoke(new Action(() =>
            {
                toExecute.Cells[(int)RowIndex.Remarks].Value = testerInfos[1].Text;
            }));

            progressBarStuff.StopOneIteration(toExecute.Index);
        }

        #region Camera/Mic - Initialization

        private void InitDevice()
        {
            _deviceSettings = new Dictionary<string, DeviceSettings>
            {
                {"TG45", new DeviceSettings()},
                {"TG59", new DeviceSettings()}
            };

            foreach (var i in _deviceSettings.ToList())
            {
                var dS = new DeviceSettings();
                dS.conValue = "";
                dS.conMethod = EN_CON_METHOD.CON_UART;
                dS.conStatus = false;

                _deviceSettings[i.Key] = dS;
            }
        }

        private void InitCamera()
        {
            string[] cameras = null;

            // Init for Camera
            Saal.CAM_GetCAMList(out cameras);
            if (cameras.Length > 0)
            {
                ImageHandler.CamName = cameras[0]; // set camera device
                ImageHandler.ImgFromParent = Picture; // send picture box control
                ImageHandler.ResSize.Width = ResWidth;
                ImageHandler.ResSize.Height = ResHeight;
                var imageGuide = new Bitmap(ResWidth, ResHeight);
                imageGuide.Tag = "HDMIPictureFormat";
                DrawGuides(imageGuide, new Rectangle(0, 0, ResWidth, ResHeight));
                ImageHandler.gridBitmap = imageGuide;
                ImageHandler.InitCamera(); // start camera

                foreach (string cam in cameras)
                {
                    cbCamList.Items.Add(cam);
                }
                cbCamList.SelectedIndex = 0;
            }
        }

        private static void DrawGuides(Bitmap imageCanvas, Rectangle drawBox)
        {
            // draw guides on top of imageCanvas
            using (var g = Graphics.FromImage(imageCanvas))
            {

                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                drawBox.Inflate(-Convert.ToInt32(drawBox.Width * 0.1), -Convert.ToInt32(drawBox.Height * 0.1));
                var image = Resource.vcolorbar;
                var image_with_alpha = image.SetOpacity((float)0.2);
                g.DrawImage(image_with_alpha, drawBox);
                g.DrawRectangle(new Pen(Color.FromArgb(128, Color.White), 4), drawBox);

                g.DrawLine(new Pen(Color.FromArgb(128, Color.White), 4), drawBox.Width / 2 + drawBox.Left - 1, drawBox.Y,
                drawBox.Width / 2 + drawBox.Left - 1, drawBox.Bottom);

                for (int i = 1; i < 7; ++i)
                {
                    g.DrawLine(new Pen(Color.FromArgb(128, Color.White), 4), drawBox.Left, (drawBox.Height / 7 * i) +
                        drawBox.Top + 2, drawBox.Right, (drawBox.Height / 7 * i) + drawBox.Top + 2);
                }

                string text1 = "[HDMI Picture Test]\nPlease make sure the colorbar position\n" +
                    "matched with test TV as close as possible";
                using (Font font1 = new Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point))
                {
                    StringFormat stringFormat = new StringFormat();
                    stringFormat.Alignment = StringAlignment.Center;
                    stringFormat.LineAlignment = StringAlignment.Center;

                    g.DrawString(text1, font1, new SolidBrush(Color.FromArgb(200, Color.Red)), drawBox, stringFormat);
                }
                image.Dispose();
                image_with_alpha.Dispose();
            }
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
        #endregion

        #region Test Connection - Event Handler
        private void btTestCon_Click(object sender, EventArgs e)
        {
            TestConnection();
        }

        #endregion

        #region Image Handler
        // Capture and return image from camera
        private Bitmap CaptureImg()
        {
            try
            {
                var captureImg = _blankImg;
                if (Picture.Image != null)
                {
                    if (Picture.InvokeRequired)
                    {
                        Picture.Invoke(
                            new Action(() => { captureImg = new Bitmap((Image)Picture.Image.Clone()); }));
                    }
                }

                if (currentBitmap != null)
                    currentBitmap.Dispose();

                currentBitmap = captureImg;

                return captureImg;
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
                throw;
            }
        }

        // Make and return comparison result
        private double CompareImg(Bitmap image1, Bitmap image2)
        {
            double dSimilarityResult = 0;

            var compA = AForge.Imaging.Image.Clone(image1, PixelFormat.Format24bppRgb);
            var compB = AForge.Imaging.Image.Clone(image2, PixelFormat.Format24bppRgb);
            var tm = new ExhaustiveTemplateMatching(0);
            var matchings = tm.ProcessImage(compA, compB);
            if (matchings.Length > 0)
                dSimilarityResult = matchings[0].Similarity;
            else if (matchings.Length <= 0)
                dSimilarityResult = 0;

            //return Math.Round(dSimilarityResult, 2)*100;
            return dSimilarityResult;
        }

        // Compare process and return true/false
        private bool CompareImgResult(string folderName, ref string imagesFolder)
        {
            var simirate = new double[3];
            var simirate_ = new double[2];

            Bitmap compTemp1;
            Bitmap compTemp2;
            Bitmap compTemp3;

            compTemp1 = new Bitmap(CaptureImg());
            Thread.Sleep(50);
            compTemp2 = new Bitmap(CaptureImg());
            Thread.Sleep(50);
            compTemp3 = new Bitmap(CaptureImg());

            // create folder in exe level
            var current_app_path = Environment.CurrentDirectory;

            // create output folder
            var outputFolder = current_app_path + "\\output";
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // create output/testparent folder
            var parentImagesFolder = outputFolder + "\\CapturedImagesByRow_HDMIPictureTest";
            if (!Directory.Exists(parentImagesFolder))
                Directory.CreateDirectory(parentImagesFolder);

            // create sub folder
            imagesFolder = parentImagesFolder + "\\[" + folderName + "]";
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            compTemp1.Tag = imagesFolder + "\\" + _executionTime + "1.png";
            compTemp2.Tag = imagesFolder + "\\" + _executionTime + "2.png";
            compTemp3.Tag = imagesFolder + "\\" + _executionTime + "3.png";

            var files = new List<Bitmap>();
            files.Add(compTemp1);
            files.Add(compTemp2);
            files.Add(compTemp3);

            foreach (var i in files)
            {
                if (File.Exists(i.Tag.ToString()))
                    File.Delete(i.Tag.ToString());
                i.Save(i.Tag.ToString(), ImageFormat.Png);
            }

            //compare 1 to 2,3     
            var similar1v1 = CompareImg(compTemp1, compTemp1);
            var similar1v2 = CompareImg(compTemp1, compTemp2);
            var similar1v3 = CompareImg(compTemp1, compTemp3);
            var similar2v3 = CompareImg(compTemp2, compTemp3);

            Console.WriteLine("Comparison:");
            Console.WriteLine(" 1v1 " + folderName + " " + (similar1v1) + " " + Path.GetFileNameWithoutExtension(compTemp1.Tag.ToString()) + " " + Path.GetFileNameWithoutExtension(compTemp1.Tag.ToString()));
            Console.WriteLine(" 1v2 " + folderName + " " + (similar1v2) + " " + Path.GetFileNameWithoutExtension(compTemp1.Tag.ToString()) + " " + Path.GetFileNameWithoutExtension(compTemp2.Tag.ToString()));
            Console.WriteLine(" 1v3 " + folderName + " " + (similar1v3) + " " + Path.GetFileNameWithoutExtension(compTemp1.Tag.ToString()) + " " + Path.GetFileNameWithoutExtension(compTemp3.Tag.ToString()));
            Console.WriteLine(" 2v3 " + folderName + " " + (similar2v3) + " " + Path.GetFileNameWithoutExtension(compTemp2.Tag.ToString()) + " " + Path.GetFileNameWithoutExtension(compTemp3.Tag.ToString()));

            return true;
        }

        #endregion 

        #region Click - Event Handler

        private void InvokeBeforeExecute()
        {
            // check camera status
            if (btOnOffCam.Text == "On")
            {
                MessageBox.Show("Do switch on the camera before testing");
                DoExecute = false;
                return;
            }

            bool valid_single = false;
            foreach (var i in _deviceSettings)
            {
                if (i.Value.conStatus)
                {
                    if (!valid_single)
                        valid_single = true;
                    else
                    {
                        MessageBox.Show("Please select only one valid device by clicking at the button");
                        DoExecute = false;
                        return;
                    }

                }
            }

            if (!valid_single)
            {
                MessageBox.Show("Please make sure device is connected properly before testing");
                DoExecute = false;
                return;
            }

            // set timestamp
            var curdate = DateTime.Now.ToString("yyyyMMdd_");
            var curtime = DateTime.Now.ToString("hhmm.fff_");
            _executionTime = curdate + curtime;

            // decide to execute or note
            DoExecute = true;
        }

        private void InvokeBeforeExecuteButAfterSelection(object selection, object datagridview, object owner, BackgroundWorker backgroundWorker)
        {
            progressBarStuff = new ProgressBarStuff(ref backgroundWorker);
            progressBarStuff.Show(selection, datagridview, owner);

            var currentImage = new Bitmap(ImageHandler.CurrentFrame);
            currentImage.Tag = "Time taken: " + DateTime.Now;
            progressBarStuff.refImage = currentImage;
        }

        private void InvokeProgressBarChangedEvent(object sender, ProgressChangedEventArgs e)
        {
            progressBarStuff.ProgressBarChangedEvent(sender, e);
            progressBarStuff.camLiveImage = ImageHandler.CurrentFrame;
        }

        private bool CollectRowIndexData()
        {

            /* 
             * 1. Set resolution             * 
             * search for key (1 ~ 34) in the row, then assign the value (HDMI001)
             */
            {
                var hdmiFormatDictionary = new Dictionary<string, string>();
                hdmiFormatDictionary.Add("1", "HDMI001");
                hdmiFormatDictionary.Add("7", "HDMI007");
                hdmiFormatDictionary.Add("3", "HDMI003");
                hdmiFormatDictionary.Add("4", "HDMI004");
                hdmiFormatDictionary.Add("5", "HDMI005");
                hdmiFormatDictionary.Add("16", "HDMI016");
                hdmiFormatDictionary.Add("32", "HDMI032");
                hdmiFormatDictionary.Add("34", "HDMI034");

                var hdmiKey = _rowIndexData.HDMIFormat.vA;
                var hdmiValue = _rowIndexData.HDMIFormat.vB;
                var hdmiRaw = _rowIndexData.HDMIFormat.vC;

                var current_hdmi_raw = toExecute.Cells[hdmiKey].Value.ToString();

                if (current_hdmi_raw.Equals(""))
                    toExecute.Cells[hdmiKey].Value = hdmiRaw; // assign old value
                else
                {
                    hdmiRaw = current_hdmi_raw;
                    var m = Regex.Match(hdmiRaw, @"\d+");
                    if (m.Success)
                        hdmiValue = hdmiFormatDictionary[m.Value];
                }

                if (hdmiValue.Equals(""))
                    return false;

                _rowIndexData.HDMIFormat.vA = hdmiKey;
                _rowIndexData.HDMIFormat.vB = hdmiValue;
                _rowIndexData.HDMIFormat.vC = hdmiRaw;
            }

            /* 
             * 2. Set frequency             
             * search for ".94", which means 1.001 modifier, otherwise 1.0
             * in devicecmd, 1.001 = 1, 1 = 0
             */
            {
                var freqKey = _rowIndexData.frequency.vA;
                var freqValue = _rowIndexData.frequency.vB;
                var freqRaw = _rowIndexData.frequency.vC;

                var current_freq_raw = toExecute.Cells[freqKey].Value.ToString();

                if (current_freq_raw.Equals(""))
                    toExecute.Cells[freqKey].Value = freqRaw; // assign old value
                else
                {
                    freqRaw = current_freq_raw;
                    freqValue = Regex.IsMatch(freqRaw, @".94") ? "1" : "0";
                }

                if (freqValue.Equals(""))
                    return false;

                _rowIndexData.frequency.vA = freqKey;
                _rowIndexData.frequency.vB = freqValue;
                _rowIndexData.frequency.vC = freqRaw;
            }

            /* 
             * 3. Set Color space            
             */
            {
                // add color space
                var colorSpaces = new List<string>();
                colorSpaces.Add("4:4:4");
                colorSpaces.Add("4:2:2");
                colorSpaces.Add("RGB");

                var color_key = _rowIndexData.colorType.vA;
                var color_value = _rowIndexData.colorType.vB;
                var color_raw = _rowIndexData.colorType.vC;

                var current_color_raw = toExecute.Cells[color_key].Value.ToString();

                if (current_color_raw.Equals(""))
                {
                    toExecute.Cells[color_key].Value = color_raw; // assign old value
                }
                else
                {
                    color_raw = current_color_raw;
                    foreach (var i in colorSpaces)
                    {
                        var m = Regex.Match(color_raw, i);
                        if (m.Success)
                        {
                            color_value = i;
                            break;
                        }
                    }
                }

                if (color_value.Equals(""))
                    return false;

                _rowIndexData.colorType.vA = color_key;
                _rowIndexData.colorType.vB = color_value;
                _rowIndexData.colorType.vC = color_raw;
            }

            _rowIndexData.operation = new KeyValuePair<int, string>(_rowIndexData.operation.Key,
                toExecute.Cells[_rowIndexData.operation.Key].Value.ToString());
            _rowIndexData.expected = new KeyValuePair<int, string>(_rowIndexData.expected.Key,
                toExecute.Cells[_rowIndexData.expected.Key].Value.ToString());

            return true;
        }

        private void InvokeAfterExecute()
        {
            progressBarStuff.Reset();
            previousColor = "";
            _rowIndexData = new RowIndexData();
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!(e.RowIndex >= 0) || !(e.ColumnIndex >= 0))
                return;

            var datagridview = sender as DataGridView;
            string[] paths = datagridview.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag as string[];

            if (paths == null || paths.Length == 0)
                return;

            if (paths[0].Contains(Path.GetTempPath()))
            {
                var imageviewer = new PictureViewer(paths[0], paths[1]);
                imageviewer.Text = paths[0];
                imageviewer.Show();
            }
        }
        #endregion
    }

    public static class BitmapExtensions
    {
        public static Image SetOpacity(this Image image, float opacity)
        {
            var colorMatrix = new ColorMatrix();
            colorMatrix.Matrix33 = opacity;
            var imageAttributes = new ImageAttributes();
            imageAttributes.SetColorMatrix(
                colorMatrix,
                ColorMatrixFlag.Default,
                ColorAdjustType.Bitmap);
            var output = new Bitmap(image.Width, image.Height);
            using (var gfx = Graphics.FromImage(output))
            {
                gfx.SmoothingMode = SmoothingMode.AntiAlias;
                gfx.DrawImage(
                    image,
                    new Rectangle(0, 0, image.Width, image.Height),
                    0,
                    0,
                    image.Width,
                    image.Height,
                    GraphicsUnit.Pixel,
                    imageAttributes);
            }
            return output;
        }
    }
}