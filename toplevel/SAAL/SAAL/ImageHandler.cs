using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.OCR;
using Emgu.Util;
using Emgu.CV.UI;
using System.Threading;
using System.IO;
using System.Security;
using System.Xml;
using Emgu.CV.Util;
using System.Timers;
using System.Text.RegularExpressions;

namespace SAAL
{
    public static class ImageHandler
    {
        private static SAAL_Interface Saal;

        #region Grid Calib Parameters
        private static Form fmGridCalib;
        private static PictureBox pbGridCalib;
        private static Rectangle rcGridCalib;
        private static ComboBox cbPreTest;
        private static ComboBox cbGridList;
        private static ComboBox cbOCRProfileList;
        private static Button btOCRProfileSave;
        private static ComboBox cbOCRItemList;
        private static TextBox tbOCR_H;
        private static TextBox tbOCR_W;
        private static TextBox tbOCR_X;
        private static TextBox tbOCR_Y;
        private static TrackBar tbarCamBrightness;
        private static TrackBar tbarCamContrast;
        private static TrackBar tbarCamSharpness;
        private static EN_POS_SIZABLE_RECT nodeSelected = EN_POS_SIZABLE_RECT.POS_NONE;
        private static EN_PRE_TEST preTestItem = EN_PRE_TEST.PRE_TEST_MAX;
        private static bool allowDeformingDuringMovement = false;
        private static int sizeNodeRect = 5;
        private static bool mIsClick = false;
        private static bool mMove = false;
        private static int oldX;
        private static int oldY;
        private static Item_Info[] _item;
        private static int ItemSize = 50;
        private static string XMLfile = "profile.xml";
        private static string test_profile;
        private static string test_item;
        private static Boolean isEnOCR;
        private static Boolean isEnBlurryCheck;
        private static Boolean isEnSimilarityCheck;
        private static Boolean isEnPictureFormat;
        private static string ErrorMsg = "This function is not applicable in your case. Please enable it first.";
        private enum EN_POS_SIZABLE_RECT
        {
            POS_UP_MIDDLE,
            POS_LEFT_MIDDLE,
            POS_LEFT_BOTTOM,
            POS_LEFT_UP,
            POS_RIGHT_UP,
            POS_RIGHT_MIDDLE,
            POS_RIGHT_BOTTOM,
            POS_BOTTOM_MIDDLE,
            POS_NONE
        };

        private enum EN_PRE_TEST
        {
            PRE_TEST_OCR_CALIB,
            PRE_TEST_SIMILARITY,
            PRE_TEST_BLURRY,
            PRE_TEST_PICT_FORMAT,
            PRE_TEST_MAX
        };

        private struct Item_Info
        {
            public string Name;
            public string Source;
            public string Info;
            public string Label;
            public string Opt1;
            public string Opt2;
            public string Opt3;
            public string Full;
        }
        #endregion

        #region OCR Calib Parameters
        private static Tesseract tess;
        private static Form fmOCRCalib;
        private static Button btOCRRefresh;
        private static ImageBox pbOCRCalib1;
        private static ImageBox pbOCRCalib2;
        private static ImageBox pbOCRCalib3;
        private static ImageBox pbOCRCalib4;
        private static ImageBox pbOCRCalib5;
        private static HistogramBox hbOCRCalib1;
        private static HistogramBox hbOCRCalib2;
        private static HistogramBox hbOCRCalib3;
        private static HistogramBox hbOCRCalib4;
        private static HistogramBox hbOCRCalib5;
        private static Panel pnlOCRCalib1;
        private static Panel pnlOCRCalib2;
        private static Panel pnlOCRCalib3;
        private static Panel pnlOCRCalib4;
        private static Panel pnlOCRCalib5;
        private static TextBox tblOCRCalib1;
        private static TextBox tblOCRCalib2;
        private static TextBox tblOCRCalib3;
        private static TextBox tblOCRCalib4;
        private static TextBox tblOCRCalib5;
        private static CheckBox cblOCRCalib3;
        private static CheckBox cblOCRCalib4;
        private static TrackBar tbarCLAHEClip;
        private static TextBox tbCLAHEClip;
        private static TextBox tbCLAHEWSize;
        private static TextBox tbCLAHEHSize;
        private static CheckBox cbInvBinarize;
        private static Int32 CLAHEClip;
        private static Int32 CLAHEHSize;
        private static Int32 CLAHEWSize;
        private static Boolean isCLAHE;
        private static Boolean isSharpenize;
        private static Boolean isInvBinarize;
        private static Int32 thresBinarize;
        private static Double weightAlpha;
        private static Double weightBeta;
        private static Double weightGamma;
        private static Double gaussianSigmaX;
        private static TextBox tbThresBinarize;
        private static TextBox tbWeightAlpha;
        private static TextBox tbWeightBeta;
        private static TextBox tbWeightGamma;
        private static TextBox tbGaussianSigmaX;
        private static double zoomScale;
        private static TextBox tbZoomScale;
        #endregion

        #region Blurry Check Parameters
        private static Form fmBlurryCheck;
        private static ImageBox pbBlurryImage;
        private static CheckBox cbBlurryGaussian;
        private static Boolean isGaussianBlur;
        private static TextBox tbBlurryResult;
        private static TextBox tbBlurryThreshold;
        private static Int32 thresBlurry;

        #region Exposed Blurry Check Parameters
        public static int BlurryTestThreshold { set { thresBlurry = value; } get { return thresBlurry; } }
        #endregion

        #endregion

        #region Simlarity Check Parameters
        private static Form fmSimilarityCheck;
        private static ImageBox pbSimilarityRef;
        private static ImageBox pbSimilarityCur;
        private static TextBox tbSimilarityThresPSNR;
        private static TextBox tbSimilarityThresSSIM;
        private static TextBox tbSimilarityPeriod;
        private static TextBox tbSimilarityInterval;
        private static TextBox tbSimilarityResult;
        private static Int32 period = 10000;
        private static Int32 interval = 500;
        private static Int32 thresPSNR = 35;
        private static Int32 thresSSIM = 90;
        #endregion

        #region Picture Format Parameters
        private static Form fmPictFormat;
        #endregion

        #region Operation Dialog Parameters
        private static Form fmOperationDialog;
        private static PictureBox pbImgRef;
        private static PictureBox pbImgCom;
        private static RichTextBox tbOperationLog;
        private static Boolean updOpFlag;
        private static ProgressBar pbarOpDialog;
        private static Label lblOpDialog;
        #endregion

        #region Video Parameters
        private static string camName;
        private static PictureBox imgFromParent;
        private static Mat frameCapture;
        private static Boolean capture_flag;
        private static Int32 setBrightness;
        private static Int32 setContrast;
        private static Int32 setSharpness;
        private static Bitmap scan_img;
        private static Rectangle rec_scan_img;
        private static System.Timers.Timer scan_timer;
        #endregion

        #region Public properties and methods
        public enum EN_OCR_ITEM
        {
            OCR_ITEM_SOURCE,
            OCR_ITEM_LABEL,
            OCR_ITEM_INFO,
            OCR_ITEM_OPT1,
            OCR_ITEM_OPT2,
            OCR_ITEM_OPT3,
            OCR_ITEM_FULL,
            OCR_ITEM_MAX
        }

        public struct ResSize
        {
            public static int Width = 1280;
            public static int Height = 720;
        }

        public static void OpenCameraCalibForm() { fmGridCalib_Open(); }
        public static void InitCamera() { videoInit(); }
        public static void StartCamera() { videoStart(); }
        public static void StopCamera() { videoStop(); }
        public static String ExtractText(EN_OCR_ITEM item) { if (!isEnOCR) MessageBox.Show(ErrorMsg); return ((isEnOCR) ? TextExtractionFromParent(item) : ""); }
        public static String ExtractTextBitmap(Bitmap item) { if (!isEnOCR) MessageBox.Show(ErrorMsg); return ((isEnOCR) ? TextExtractionBitmap(item) : ""); }
        public static Boolean BlurryCheck(Int32 ThresBlurry) { thresBlurry = ThresBlurry; if (!isEnBlurryCheck) MessageBox.Show(ErrorMsg); return ((isEnBlurryCheck) ? blurryCheck() : false); }
        public static Boolean SimilarityCheck(Int32 Period, Int32 Interval, Int32 ThresPSNR, Int32 ThresSSIM) { interval = Interval; period = Period; thresPSNR = ThresPSNR; thresSSIM = ThresSSIM; if (!isEnSimilarityCheck) MessageBox.Show(ErrorMsg); return ((isEnSimilarityCheck) ? similarityCheckFromParent() : false); }
        public static Boolean RunSimilarityAndBlurryCheck(Bitmap referenceBitmap, ref Bitmap probBitmap, 
            ref object errorInfo, ref object processInfo, ref bool stop)
        {
            if (!isEnSimilarityCheck || !isEnBlurryCheck)
                MessageBox.Show(ErrorMsg);

            return ((isEnSimilarityCheck && isEnBlurryCheck) ? similarity_wBlurryCheck(referenceBitmap, ref probBitmap, 
                ref errorInfo, ref processInfo, ref stop) : false);
        }
        public static String CamName { set { camName = value; } get { return camName; } }
        public static PictureBox ImgFromParent { set { imgFromParent = value; } get { return imgFromParent; } }
        public static Boolean IsEnOCR { set { isEnOCR = value; } get { return isEnOCR; } }
        public static Boolean IsEnBlurryCheck { set { isEnBlurryCheck = value; } get { return isEnBlurryCheck; } }
        public static Boolean IsEnSimilarityCheck { set { isEnSimilarityCheck = value; } get { return isEnSimilarityCheck; } }
        public static Boolean IsEnPictureFormat { set { isEnPictureFormat = value; } get { return isEnPictureFormat; } }
        public static Bitmap CurrentFrame { get { return frameCapture.Bitmap; } }
        public static Int32 ThresBinarize { set { thresBinarize = value; } get { return thresBinarize; } }
        public static Double WeightAlpha { set { weightAlpha = value; } get { return weightAlpha; } }
        public static Double WeightBeta { set { weightBeta = value; } get { return weightBeta; } }
        public static Double WeightGamma { set { weightGamma = value; } get { return weightGamma; } }
        public static Double GaussianSigmaX { set { gaussianSigmaX = value; } get { return gaussianSigmaX; } }
        public static int SimilarityTestPeriod { set { period = value; } get { return period; } }
        public static int SimilarityTestInterval { set { interval = value; } get { return interval; } }
        public static int SimilarityTestPSNRThreshold { set { thresPSNR = value; } get { return thresPSNR; } }
        public static int SimilarityTestSSIMThreshold { set { thresSSIM = value; } get { return thresSSIM; } }
        public static bool UpdOpFlag { set { updOpFlag = value; } get { return updOpFlag; } }
        #endregion

        static ImageHandler()
        {
            Saal = new SAAL_Interface();
            tess = new Tesseract("", "eng", OcrEngineMode.TesseractCubeCombined);
            //tess = new Tesseract("", "eng", OcrEngineMode.TesseractOnly);
            /*
            tess = new Tesseract();
            tess.SetVariable("load_system_dawg", "F");
            tess.SetVariable("load_freq_dawg", "F");
            tess.SetVariable("user_words_suffix", "user-words");
            tess.SetVariable("user_patterns_suffix", "user-patterns");
            tess.SetVariable("tessedit_char_whitelist", "1234567890p1234567890");
            tess.Init("", "eng", OcrEngineMode.TesseractOnly);
            */
            
            XML_fetch_data();

            isEnOCR = false;
            isEnBlurryCheck = false;
            isEnSimilarityCheck = false;
            isEnPictureFormat = false;

            setBrightness = 0;
            setContrast = 0;
            setSharpness = 0;

            // OCR parameters
            isCLAHE = false;
            isSharpenize = true;
            CLAHEClip = 2;
            CLAHEHSize = 10;
            CLAHEWSize = 10;
            WeightAlpha = 1.5;
            WeightBeta = -1.5;
            WeightGamma = 0;
            GaussianSigmaX = 1.5;
            ThresBinarize = 210;
            zoomScale = 4;

            // Start Operation Dialog
            fmOperationDialog_Open();
        }

        #region Grid properties 
        private static Dictionary<string, Bitmap> gridBitmaps = new Dictionary<string, Bitmap>();
		public static Bitmap gridBitmap{set { gridBitmaps.Add(value.Tag.ToString(), value); } }
        #endregion

        #region Grid Calibration
        private static Boolean fmGridCalib_Open()
        {
            if (fmGridCalib != null && !fmGridCalib.IsDisposed)
                return false;

            fmGridCalib = new Form();
            fmGridCalib.ClientSize = new Size(800 + 200, 600);
            fmGridCalib.MaximizeBox = false;
            fmGridCalib.ControlBox = true;
            fmGridCalib.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmGridCalib.FormClosing += fmGridCalib_Close;
            fmGridCalib.StartPosition = FormStartPosition.CenterScreen;
            rcGridCalib = new Rectangle();
            rcGridCalib.Size = new Size(160, 120);

            pbGridCalib = new PictureBox();
            pbGridCalib.ClientSize = new System.Drawing.Size(800, 600);
            pbGridCalib.BackColor = Color.LightBlue;
            pbGridCalib.SizeMode = PictureBoxSizeMode.StretchImage;
            pbGridCalib.MouseDown += new MouseEventHandler(pbGridCalib_MouseDown);
            pbGridCalib.MouseUp += new MouseEventHandler(pbGridCalib_MouseUp);
            pbGridCalib.MouseMove += pbGridCalib_MouseMove;
            pbGridCalib.Paint += new PaintEventHandler(pbGridCalib_Paint);
            fmGridCalib.Controls.Add(pbGridCalib);

            Panel pnlGridAssistant = new Panel();
            pnlGridAssistant.Size = new Size(180, 80);
            pnlGridAssistant.BackColor = Color.LightYellow;
            pnlGridAssistant.Location = new Point(pbGridCalib.Location.X + pbGridCalib.Width + 2, pbGridCalib.Location.Y);
            fmGridCalib.Controls.Add(pnlGridAssistant);

            Label lblGridAssistant = new Label();
            lblGridAssistant.Text = "Grid Assistant";
            lblGridAssistant.Font = new Font(lblGridAssistant.Font, FontStyle.Bold);
            lblGridAssistant.Size = new Size(100, 30);
            lblGridAssistant.Location = new Point(2, 2);
            pnlGridAssistant.Controls.Add(lblGridAssistant);

            cbGridList = new ComboBox();
            cbGridList.Size = new Size(100, 30);
            cbGridList.Location = new Point(lblGridAssistant.Location.X + 5, lblGridAssistant.Location.Y + lblGridAssistant.Height + 10);
            cbGridList.Items.Add("None");
            foreach (var i in gridBitmaps)
                cbGridList.Items.Add(i.Key);
            cbGridList.SelectedIndex = 0;
            pnlGridAssistant.Controls.Add(cbGridList);

            Panel pnlCamSetting = new Panel();
            pnlCamSetting.Size = new Size(180, 170);
            pnlCamSetting.BackColor = Color.LightYellow;
            pnlCamSetting.Location = new Point(pnlGridAssistant.Location.X, pnlGridAssistant.Location.Y + pnlGridAssistant.Height + 2);
            fmGridCalib.Controls.Add(pnlCamSetting);

            Label lblCamSettings = new Label();
            lblCamSettings.Text = "Camera Settings";
            lblCamSettings.Font = new Font(lblCamSettings.Font, FontStyle.Bold);
            lblCamSettings.Size = new Size(100, 30);
            lblCamSettings.Location = new Point(2, 2);
            pnlCamSetting.Controls.Add(lblCamSettings);

            Label lblCamBrightness = new Label();
            lblCamBrightness.Text = "Brightness :";
            lblCamBrightness.Size = new Size(70, 30);
            lblCamBrightness.Location = new Point(lblCamSettings.Location.X, lblCamSettings.Location.Y + lblCamSettings.Height + 10);
            pnlCamSetting.Controls.Add(lblCamBrightness);

            tbarCamBrightness = new TrackBar();
            tbarCamBrightness.Size = new Size(100, 20);
            tbarCamBrightness.Maximum = 500;
            tbarCamBrightness.TickStyle = TickStyle.None;
            tbarCamBrightness.Value = 50;
            tbarCamBrightness.Scroll += new EventHandler(tbarCamBrightness_Scroll);
            tbarCamBrightness.Location = new Point(lblCamBrightness.Location.X + lblCamBrightness.Width + 2, lblCamBrightness.Location.Y);
            pnlCamSetting.Controls.Add(tbarCamBrightness);

            Label lblCamContrast = new Label();
            lblCamContrast.Text = "Contrast :";
            lblCamContrast.Size = new Size(70, 30);
            lblCamContrast.Location = new Point(lblCamBrightness.Location.X, lblCamBrightness.Location.Y + lblCamBrightness.Height + 15);
            pnlCamSetting.Controls.Add(lblCamContrast);

            tbarCamContrast = new TrackBar();
            tbarCamContrast.Size = new Size(100, 20);
            tbarCamContrast.Maximum = 500;
            tbarCamContrast.TickStyle = TickStyle.None;
            tbarCamContrast.Value = 50;
            tbarCamContrast.Scroll += new EventHandler(tbarCamContrast_Scroll);
            tbarCamContrast.Location = new Point(lblCamContrast.Location.X + lblCamContrast.Width + 2, lblCamContrast.Location.Y);
            pnlCamSetting.Controls.Add(tbarCamContrast);

            Label lblCamSharpness = new Label();
            lblCamSharpness.Text = "Sharpness :";
            lblCamSharpness.Size = new Size(70, 30);
            lblCamSharpness.Location = new Point(lblCamContrast.Location.X, lblCamContrast.Location.Y + lblCamContrast.Height + 15);
            pnlCamSetting.Controls.Add(lblCamSharpness);

            tbarCamSharpness = new TrackBar();
            tbarCamSharpness.Size = new Size(100, 20);
            tbarCamSharpness.Maximum = 500;
            tbarCamSharpness.TickStyle = TickStyle.None;
            tbarCamSharpness.Value = 50;
            tbarCamSharpness.Scroll += new EventHandler(tbarCamSharpness_Scroll);
            tbarCamSharpness.Location = new Point(lblCamSharpness.Location.X + lblCamSharpness.Width + 2, lblCamSharpness.Location.Y);
            pnlCamSetting.Controls.Add(tbarCamSharpness);

            Panel pnlOCRAssistant = new Panel();
            pnlOCRAssistant.Size = new Size(180, 200);
            pnlOCRAssistant.BackColor = Color.LightYellow;
            pnlOCRAssistant.Location = new Point(pnlCamSetting.Location.X, pnlCamSetting.Location.Y + pnlCamSetting.Height + 2);
            fmGridCalib.Controls.Add(pnlOCRAssistant);

            Label lblOCRAssistant = new Label();
            lblOCRAssistant.Text = "OCR Assistant";
            lblOCRAssistant.Font = new Font(lblGridAssistant.Font, FontStyle.Bold);
            lblOCRAssistant.Size = new Size(100, 30);
            lblOCRAssistant.Location = new Point(2, 2);
            pnlOCRAssistant.Controls.Add(lblOCRAssistant);

            Label lblOCRProfileList = new Label();
            lblOCRProfileList.Text = "Profile";
            lblOCRProfileList.Size = new Size(60, 20);
            lblOCRProfileList.Location = new Point(lblOCRAssistant.Location.X + 5, lblOCRAssistant.Location.Y + lblOCRAssistant.Height + 10);
            pnlOCRAssistant.Controls.Add(lblOCRProfileList);

            btOCRProfileSave = new Button();
            btOCRProfileSave.Text = "Save";
            btOCRProfileSave.Size = new Size(50, 20);
            btOCRProfileSave.Location = new Point(lblOCRProfileList.Location.X + lblOCRProfileList.Width + 20, lblOCRProfileList.Location.Y);
            btOCRProfileSave.Click += new EventHandler(btOCRProfileSave_Click);
            if (!isEnOCR)
                btOCRProfileSave.Enabled = false;
            pnlOCRAssistant.Controls.Add(btOCRProfileSave);

            cbOCRProfileList = new ComboBox();
            cbOCRProfileList.Size = new Size(130, 30);
            cbOCRProfileList.Location = new Point(lblOCRProfileList.Location.X, lblOCRProfileList.Location.Y + lblOCRProfileList.Height + 1);
            cbOCRProfileList_update();
            cbOCRProfileList.SelectedIndexChanged += new EventHandler(cbOCRProfileList_SelectedIndexChanged);
            if (!isEnOCR)
                cbOCRProfileList.Enabled = false;
            pnlOCRAssistant.Controls.Add(cbOCRProfileList);

            Label lblOCRItemList = new Label();
            lblOCRItemList.Text = "Items";
            lblOCRItemList.Size = new Size(100, 20);
            lblOCRItemList.Location = new Point(cbOCRProfileList.Location.X, cbOCRProfileList.Location.Y + cbOCRProfileList.Height + 5);
            pnlOCRAssistant.Controls.Add(lblOCRItemList);

            cbOCRItemList = new ComboBox();
            cbOCRItemList.Size = new Size(130, 30);
            cbOCRItemList.Location = new Point(lblOCRItemList.Location.X, lblOCRItemList.Location.Y + lblOCRItemList.Height + 1);
            cbOCRItemList.Items.Add("Source");
            cbOCRItemList.Items.Add("Label");
            cbOCRItemList.Items.Add("Info");
            cbOCRItemList.Items.Add("Option1");
            cbOCRItemList.Items.Add("Option2");
            cbOCRItemList.Items.Add("Option3");
            cbOCRItemList.Items.Add("Full");
            cbOCRItemList.SelectedIndexChanged += new EventHandler(cbOCRItemList_SelectedIndexChanged);
            if (!isEnOCR)
                cbOCRItemList.Enabled = false;
            pnlOCRAssistant.Controls.Add(cbOCRItemList);

            Label lblOCRItemW = new Label();
            lblOCRItemW.Text = "W :";
            lblOCRItemW.Size = new Size(30, 20);
            lblOCRItemW.Location = new Point(cbOCRItemList.Location.X, cbOCRItemList.Location.Y + cbOCRItemList.Height + 10);
            pnlOCRAssistant.Controls.Add(lblOCRItemW);

            tbOCR_W = new TextBox();
            tbOCR_W.Size = new Size(30, 30);
            tbOCR_W.ReadOnly = true;
            tbOCR_W.Location = new Point(lblOCRItemW.Location.X + lblOCRItemW.Width + 1, lblOCRItemW.Location.Y);
            if (!isEnOCR)
                tbOCR_W.Enabled = false;
            pnlOCRAssistant.Controls.Add(tbOCR_W);

            Label lblOCRItemH = new Label();
            lblOCRItemH.Text = "H :";
            lblOCRItemH.Size = new Size(30, 20);
            lblOCRItemH.Location = new Point(tbOCR_W.Location.X + tbOCR_W.Width + 5, tbOCR_W.Location.Y);
            pnlOCRAssistant.Controls.Add(lblOCRItemH);

            tbOCR_H = new TextBox();
            tbOCR_H.Size = new Size(30, 30);
            tbOCR_H.ReadOnly = true;
            tbOCR_H.Location = new Point(lblOCRItemH.Location.X + lblOCRItemH.Width + 1, lblOCRItemH.Location.Y);
            if (!isEnOCR)
                tbOCR_H.Enabled = false;
            pnlOCRAssistant.Controls.Add(tbOCR_H);

            Label lblOCRItemX = new Label();
            lblOCRItemX.Text = "X :";
            lblOCRItemX.Size = new Size(30, 20);
            lblOCRItemX.Location = new Point(lblOCRItemW.Location.X, lblOCRItemW.Location.Y + lblOCRItemW.Width + 5);
            pnlOCRAssistant.Controls.Add(lblOCRItemX);

            tbOCR_X = new TextBox();
            tbOCR_X.Size = new Size(30, 30);
            tbOCR_X.ReadOnly = true;
            tbOCR_X.Location = new Point(lblOCRItemX.Location.X + lblOCRItemX.Width + 1, lblOCRItemX.Location.Y);
            if (!isEnOCR)
                tbOCR_X.Enabled = false;
            pnlOCRAssistant.Controls.Add(tbOCR_X);

            Label lblOCRItemY = new Label();
            lblOCRItemY.Text = "Y :";
            lblOCRItemY.Size = new Size(30, 20);
            lblOCRItemY.Location = new Point(tbOCR_X.Location.X + tbOCR_X.Width + 5, tbOCR_X.Location.Y);
            pnlOCRAssistant.Controls.Add(lblOCRItemY);

            tbOCR_Y = new TextBox();
            tbOCR_Y.Size = new Size(30, 30);
            tbOCR_Y.ReadOnly = true;
            tbOCR_Y.Location = new Point(lblOCRItemY.Location.X + lblOCRItemY.Width + 1, lblOCRItemY.Location.Y);
            if (!isEnOCR)
                tbOCR_Y.Enabled = false;
            pnlOCRAssistant.Controls.Add(tbOCR_Y);

            Panel pnlPreTest = new Panel();
            pnlPreTest.Size = new Size(180, 200);
            pnlPreTest.BackColor = Color.LightYellow;
            pnlPreTest.Location = new Point(pnlOCRAssistant.Location.X, pnlOCRAssistant.Location.Y + pnlOCRAssistant.Height + 5);
            fmGridCalib.Controls.Add(pnlPreTest);

            Label lblPreTest = new Label();
            lblPreTest.Text = "Pre-Test";
            lblPreTest.Font = new Font(lblPreTest.Font, FontStyle.Bold);
            lblPreTest.Size = new Size(100, 30);
            lblPreTest.Location = new Point(2, 2);
            pnlPreTest.Controls.Add(lblPreTest);

            cbPreTest = new ComboBox();
            cbPreTest.Size = new Size(130, 30);
            cbPreTest.Location = new Point(lblPreTest.Location.X, lblPreTest.Location.Y + lblPreTest.Height + 10);
            cbPreTest.Items.Add("OCR Calibration");
            cbPreTest.Items.Add("Similarity Check");
            cbPreTest.Items.Add("Blurry Check");
            cbPreTest.Items.Add("Picture Format");
            cbPreTest.SelectedIndexChanged += cbPreTest_SelectedIndexChanged;
            pnlPreTest.Controls.Add(cbPreTest);

            Button btPreTest = new Button();
            btPreTest.Size = new Size(50, 20);
            btPreTest.Text = "Open";
            btPreTest.Location = new Point(cbPreTest.Location.X + cbPreTest.Width - btPreTest.Width, cbPreTest.Location.Y + cbPreTest.Height + 2);
            btPreTest.Click += btPreTest_Click;
            pnlPreTest.Controls.Add(btPreTest);

            videoStart();

            fmGridCalib.Show();

            capture_flag = true;

            return true;
        }

        private static void cbOCRProfileList_update()
        {
            cbOCRProfileList.Items.Clear();
            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == null || _item[i].Name == "")
                    break;

                cbOCRProfileList.Items.Add(_item[i].Name);
            }

            cbOCRProfileList.SelectedIndex = 0;

            if (cbOCRItemList != null)
                cbOCRItemList.SelectedIndex = 0;

            if (cbOCRProfileList.Text != "")
                test_profile = cbOCRProfileList.Text;

            if (cbOCRItemList != null && cbOCRItemList.Text != "")
                test_item = cbOCRItemList.Text;

        }

        private static void cbOCRProfileList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] words = new string[4];

            if (cbOCRProfileList.Text == "" || cbOCRItemList.Text == "")
                return;

            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Source")
                {
                    words = _item[i].Source.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Label")
                {
                    words = _item[i].Label.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Info")
                {
                    words = _item[i].Info.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option1")
                {
                    words = _item[i].Opt1.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option2")
                {
                    words = _item[i].Opt2.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option3")
                {
                    words = _item[i].Opt3.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Full")
                {
                    words = _item[i].Full.Split(';');
                    break;
                }
            }

            if (words != null && words.Length != 0)
            {
                tbOCR_W.Text = words[0];
                tbOCR_H.Text = words[1];
                tbOCR_X.Text = words[2];
                tbOCR_Y.Text = words[3];
                reDraw(int.Parse(tbOCR_W.Text), int.Parse(tbOCR_H.Text), int.Parse(tbOCR_X.Text), int.Parse(tbOCR_Y.Text));
            }

            test_profile = cbOCRProfileList.Text;
            test_item = cbOCRItemList.Text;
        }

        private static void cbOCRItemList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] words = new string[4];

            if (cbOCRProfileList.Text == "" || cbOCRItemList.Text == "")
                return;

            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Source")
                {
                    words = _item[i].Source.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Label")
                {
                    words = _item[i].Label.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Info")
                {
                    words = _item[i].Info.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option1")
                {
                    words = _item[i].Opt1.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option2")
                {
                    words = _item[i].Opt2.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Option3")
                {
                    words = _item[i].Opt3.Split(';');
                    break;
                }
                else if (_item[i].Name == cbOCRProfileList.Text && cbOCRItemList.Text == "Full")
                {
                    words = _item[i].Full.Split(';');
                    break;
                }
            }

            if (words != null && words.Length != 0)
            {
                tbOCR_W.Text = words[0];
                tbOCR_H.Text = words[1];
                tbOCR_X.Text = words[2];
                tbOCR_Y.Text = words[3];
                reDraw(int.Parse(tbOCR_W.Text), int.Parse(tbOCR_H.Text), int.Parse(tbOCR_X.Text), int.Parse(tbOCR_Y.Text));
            }

            test_profile = cbOCRProfileList.Text;
            test_item = cbOCRItemList.Text;
        }

        private static void cbPreTest_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbPreTest.Text == "OCR Calibration")
                preTestItem = EN_PRE_TEST.PRE_TEST_OCR_CALIB;
            else if (cbPreTest.Text == "Similarity Check")
                preTestItem = EN_PRE_TEST.PRE_TEST_SIMILARITY;
            else if (cbPreTest.Text == "Blurry Check")
                preTestItem = EN_PRE_TEST.PRE_TEST_BLURRY;
            else if (cbPreTest.Text == "Picture Format")
                preTestItem = EN_PRE_TEST.PRE_TEST_PICT_FORMAT;
            else
                preTestItem = EN_PRE_TEST.PRE_TEST_MAX;
        }

        private static void btPreTest_Click(object sender, EventArgs e)
        {
            switch (preTestItem)
            {
                case EN_PRE_TEST.PRE_TEST_OCR_CALIB:
                    fmOCRCalib_Open();
                    break;
                case EN_PRE_TEST.PRE_TEST_SIMILARITY:
                    fmSimilarityCheck_Open();
                    break;
                case EN_PRE_TEST.PRE_TEST_BLURRY:
                    fmBlurryCheck_Open();
                    break;
                case EN_PRE_TEST.PRE_TEST_PICT_FORMAT:
                    fmPictFormat_Open();
                    break;
                default:
                    break;
            }
        }

        private static void btOCRProfileSave_Click(object sender, EventArgs e)
        {
            add_data();
            cbOCRProfileList_update();
        }

        private static void tbarCamBrightness_Scroll(object sender, EventArgs e)
        {
            SetBrightness(tbarCamBrightness.Value);
        }

        private static void tbarCamContrast_Scroll(object sender, EventArgs e)
        {
            SetContrast(tbarCamContrast.Value);
        }

        private static void tbarCamSharpness_Scroll(object sender, EventArgs e)
        {
            SetSharpness(tbarCamSharpness.Value);
        }

        private static void fmGridCalib_Close(object sender, FormClosingEventArgs e)
        {
            if (fmGridCalib != null && fmGridCalib.IsDisposed)
                return;

            fmGridCalib.Hide();

            capture_flag = false;
            Thread.Sleep(1000);
            pbGridCalib.Dispose();
            fmGridCalib.Dispose();
        }

        private static void pbGridCalib_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                Draw(e.Graphics);
            }
            catch (Exception exp)
            {
                System.Console.WriteLine(exp.Message);
            }
        }

        private static void pbGridCalib_MouseDown(object sender, MouseEventArgs e)
        {
            mIsClick = true;

            nodeSelected = EN_POS_SIZABLE_RECT.POS_NONE;
            nodeSelected = GetNodeSelectable(e.Location);

            if (rcGridCalib.Contains(new Point(e.X, e.Y)))
            {
                mMove = true;
            }
            oldX = e.X;
            oldY = e.Y;
        }

        private static void pbGridCalib_MouseUp(object sender, MouseEventArgs e)
        {
            mIsClick = false;
            mMove = false;
        }

        private static void pbGridCalib_MouseMove(object sender, MouseEventArgs e)
        {
            ChangeCursor(e.Location);
            if (mIsClick == false)
            {
                return;
            }

            Rectangle backupRect = rcGridCalib;

            switch (nodeSelected)
            {
                case EN_POS_SIZABLE_RECT.POS_LEFT_UP:
                    rcGridCalib.X += e.X - oldX;
                    rcGridCalib.Width -= e.X - oldX;
                    rcGridCalib.Y += e.Y - oldY;
                    rcGridCalib.Height -= e.Y - oldY;
                    break;
                case EN_POS_SIZABLE_RECT.POS_LEFT_MIDDLE:
                    rcGridCalib.X += e.X - oldX;
                    rcGridCalib.Width -= e.X - oldX;
                    break;
                case EN_POS_SIZABLE_RECT.POS_LEFT_BOTTOM:
                    rcGridCalib.Width -= e.X - oldX;
                    rcGridCalib.X += e.X - oldX;
                    rcGridCalib.Height += e.Y - oldY;
                    break;
                case EN_POS_SIZABLE_RECT.POS_BOTTOM_MIDDLE:
                    rcGridCalib.Height += e.Y - oldY;
                    break;
                case EN_POS_SIZABLE_RECT.POS_RIGHT_UP:
                    rcGridCalib.Width += e.X - oldX;
                    rcGridCalib.Y += e.Y - oldY;
                    rcGridCalib.Height -= e.Y - oldY;
                    break;
                case EN_POS_SIZABLE_RECT.POS_RIGHT_BOTTOM:
                    rcGridCalib.Width += e.X - oldX;
                    rcGridCalib.Height += e.Y - oldY;
                    break;
                case EN_POS_SIZABLE_RECT.POS_RIGHT_MIDDLE:
                    rcGridCalib.Width += e.X - oldX;
                    break;

                case EN_POS_SIZABLE_RECT.POS_UP_MIDDLE:
                    rcGridCalib.Y += e.Y - oldY;
                    rcGridCalib.Height -= e.Y - oldY;
                    break;

                default:
                    if (mMove)
                    {
                        rcGridCalib.X = rcGridCalib.X + e.X - oldX;
                        rcGridCalib.Y = rcGridCalib.Y + e.Y - oldY;
                    }
                    break;
            }
            oldX = e.X;
            oldY = e.Y;

            if (rcGridCalib.Width < 5 || rcGridCalib.Height < 5)
            {
                rcGridCalib = backupRect;
            }

            TestIfRectInsideArea();

            pbGridCalib.Invalidate();
        }

        private static void TestIfRectInsideArea()
        {
            // Test if rectangle still inside the area.
            if (rcGridCalib.X < 0) rcGridCalib.X = 0;
            if (rcGridCalib.Y < 0) rcGridCalib.Y = 0;
            if (rcGridCalib.Width <= 0) rcGridCalib.Width = 1;
            if (rcGridCalib.Height <= 0) rcGridCalib.Height = 1;

            if (rcGridCalib.X + rcGridCalib.Width > pbGridCalib.Width)
            {
                rcGridCalib.Width = pbGridCalib.Width - rcGridCalib.X - 1; // -1 to be still show 
                if (allowDeformingDuringMovement == false)
                {
                    mIsClick = false;
                }
            }
            if (rcGridCalib.Y + rcGridCalib.Height > pbGridCalib.Height)
            {
                rcGridCalib.Height = pbGridCalib.Height - rcGridCalib.Y - 1;// -1 to be still show 
                if (allowDeformingDuringMovement == false)
                {
                    mIsClick = false;
                }
            }
        }

        private static void reDraw(int W, int H, int X, int Y)
        {
            rcGridCalib.Size = new Size(W, H);
            rcGridCalib.Location = new Point(X, Y);
            pbGridCalib.Invalidate();
        }

        private static void Draw(Graphics g)
        {
            if (isEnOCR)
            {
                g.DrawRectangle(new Pen(Color.Red), rcGridCalib);

                foreach (EN_POS_SIZABLE_RECT pos in Enum.GetValues(typeof(EN_POS_SIZABLE_RECT)))
                {
                    g.DrawRectangle(new Pen(Color.Red), GetRect(pos));
                }

                tbOCR_H.Text = rcGridCalib.Height.ToString();
                tbOCR_W.Text = rcGridCalib.Width.ToString();
                tbOCR_X.Text = rcGridCalib.X.ToString();
                tbOCR_Y.Text = rcGridCalib.Y.ToString();
            }

            if (cbGridList != null)
            {
                var items = cbGridList.Items;

                foreach (var i in items)
                {
                    var j = i as string;
                    var k = cbGridList.SelectedText;

                    if (gridBitmaps.ContainsKey(k))
                        g.DrawImage(gridBitmaps[k], 0, 0);
                }
            }
        }

        private static Rectangle GetRect(EN_POS_SIZABLE_RECT p)
        {
            switch (p)
            {
                case EN_POS_SIZABLE_RECT.POS_LEFT_UP:
                    return CreateRectSizableNode(rcGridCalib.X, rcGridCalib.Y);

                case EN_POS_SIZABLE_RECT.POS_LEFT_MIDDLE:
                    return CreateRectSizableNode(rcGridCalib.X, rcGridCalib.Y + rcGridCalib.Height / 2);

                case EN_POS_SIZABLE_RECT.POS_LEFT_BOTTOM:
                    return CreateRectSizableNode(rcGridCalib.X, rcGridCalib.Y + rcGridCalib.Height);

                case EN_POS_SIZABLE_RECT.POS_BOTTOM_MIDDLE:
                    return CreateRectSizableNode(rcGridCalib.X + rcGridCalib.Width / 2, rcGridCalib.Y + rcGridCalib.Height);

                case EN_POS_SIZABLE_RECT.POS_RIGHT_UP:
                    return CreateRectSizableNode(rcGridCalib.X + rcGridCalib.Width, rcGridCalib.Y);

                case EN_POS_SIZABLE_RECT.POS_RIGHT_BOTTOM:
                    return CreateRectSizableNode(rcGridCalib.X + rcGridCalib.Width, rcGridCalib.Y + rcGridCalib.Height);

                case EN_POS_SIZABLE_RECT.POS_RIGHT_MIDDLE:
                    return CreateRectSizableNode(rcGridCalib.X + rcGridCalib.Width, rcGridCalib.Y + rcGridCalib.Height / 2);

                case EN_POS_SIZABLE_RECT.POS_UP_MIDDLE:
                    return CreateRectSizableNode(rcGridCalib.X + rcGridCalib.Width / 2, rcGridCalib.Y);
                default:
                    return new Rectangle();
            }
        }

        private static Rectangle CreateRectSizableNode(int x, int y)
        {
            return new Rectangle(x - sizeNodeRect / 2, y - sizeNodeRect / 2, sizeNodeRect, sizeNodeRect);
        }

        private static EN_POS_SIZABLE_RECT GetNodeSelectable(Point p)
        {
            foreach (EN_POS_SIZABLE_RECT r in Enum.GetValues(typeof(EN_POS_SIZABLE_RECT)))
            {
                if (GetRect(r).Contains(p))
                {
                    return r;
                }
            }
            return EN_POS_SIZABLE_RECT.POS_NONE;
        }

        private static void ChangeCursor(Point p)
        {
            pbGridCalib.Cursor = GetCursor(GetNodeSelectable(p));
        }

        private static Cursor GetCursor(EN_POS_SIZABLE_RECT p)
        {
            switch (p)
            {
                case EN_POS_SIZABLE_RECT.POS_LEFT_UP:
                    return Cursors.SizeNWSE;

                case EN_POS_SIZABLE_RECT.POS_LEFT_MIDDLE:
                    return Cursors.SizeWE;

                case EN_POS_SIZABLE_RECT.POS_LEFT_BOTTOM:
                    return Cursors.SizeNESW;

                case EN_POS_SIZABLE_RECT.POS_BOTTOM_MIDDLE:
                    return Cursors.SizeNS;

                case EN_POS_SIZABLE_RECT.POS_RIGHT_UP:
                    return Cursors.SizeNESW;

                case EN_POS_SIZABLE_RECT.POS_RIGHT_BOTTOM:
                    return Cursors.SizeNWSE;

                case EN_POS_SIZABLE_RECT.POS_RIGHT_MIDDLE:
                    return Cursors.SizeWE;

                case EN_POS_SIZABLE_RECT.POS_UP_MIDDLE:
                    return Cursors.SizeNS;
                default:
                    return Cursors.Default;
            }
        }
        #endregion

        #region OCR Calibration
        private static Boolean fmOCRCalib_Open()
        {
            if (fmOCRCalib != null && !fmOCRCalib.IsDisposed)
                return false;

            fmOCRCalib = new Form();
            fmOCRCalib.MaximizeBox = true;
            fmOCRCalib.AutoScroll = true;
            fmOCRCalib.ControlBox = true;
            fmOCRCalib.Size = new Size(1040, 600);
            fmOCRCalib.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmOCRCalib.FormClosing += new FormClosingEventHandler(fmOCRCalib_Close);
            fmOCRCalib.StartPosition = FormStartPosition.CenterScreen;

            Panel pnlTopOCRCalib = new Panel();
            pnlTopOCRCalib.Size = new Size(1000, 40);
            pnlTopOCRCalib.Location = new Point(fmOCRCalib.Location.X, fmOCRCalib.Location.Y);
            fmOCRCalib.Controls.Add(pnlTopOCRCalib);

            Label lblOCRSettings = new Label();
            lblOCRSettings.Text = "OCR Settings";
            lblOCRSettings.Font = new Font(lblOCRSettings.Font, FontStyle.Bold);
            lblOCRSettings.Location = new Point(10, 10);
            pnlTopOCRCalib.Controls.Add(lblOCRSettings);

            // (1)         
            pnlOCRCalib1 = new Panel();
            pnlOCRCalib1.Size = new Size(160, 180);
            pnlOCRCalib1.BackColor = Color.LightGray;

            Label lblOCRCalib1 = new Label();
            lblOCRCalib1.Text = "Original Image";
            lblOCRCalib1.Font = new Font(lblOCRCalib1.Font, FontStyle.Bold);
            lblOCRCalib1.Location = new Point(10, 10);
            pnlOCRCalib1.Controls.Add(lblOCRCalib1);

            Label lblZoomScale = new Label();
            lblZoomScale.Text = "Scale :";
            lblZoomScale.Size = new Size(45, 20);
            lblZoomScale.Location = new Point(lblOCRCalib1.Location.X, lblOCRCalib1.Location.Y + lblOCRCalib1.Height + 10);
            pnlOCRCalib1.Controls.Add(lblZoomScale);

            tbZoomScale = new TextBox();
            tbZoomScale.Text = "8";
            tbZoomScale.Size = new Size(30, 20);
            tbZoomScale.TextChanged += tbZoomScale_TextChanged;
            tbZoomScale.Location = new Point(lblZoomScale.Location.X + lblZoomScale.Width + 2, lblZoomScale.Location.Y - 1);
            tbZoomScale.TextAlign = HorizontalAlignment.Center;
            pnlOCRCalib1.Controls.Add(tbZoomScale);

            pbOCRCalib1 = new ImageBox();
            pbOCRCalib1.Size = new System.Drawing.Size(320, 180);
            pbOCRCalib1.SizeMode = PictureBoxSizeMode.StretchImage;
            pbOCRCalib1.BackColor = Color.LightBlue;
            pbOCRCalib1.Location = new Point(pnlOCRCalib1.Location.X + pnlOCRCalib1.Width + 1,
                pnlOCRCalib1.Location.Y);

            hbOCRCalib1 = new HistogramBox();
            hbOCRCalib1.Size = new System.Drawing.Size(320, 180);
            hbOCRCalib1.BackColor = Color.LightBlue;
            hbOCRCalib1.Location = new Point(pbOCRCalib1.Location.X + pbOCRCalib1.Width + 1,
                pbOCRCalib1.Location.Y);

            tblOCRCalib1 = new TextBox();
            tblOCRCalib1.Size = new System.Drawing.Size(200, 180);
            tblOCRCalib1.Multiline = true;
            tblOCRCalib1.WordWrap = true;
            tblOCRCalib1.ScrollBars = ScrollBars.Vertical;
            tblOCRCalib1.Location = new Point(hbOCRCalib1.Location.X + hbOCRCalib1.Width + 1,
                hbOCRCalib1.Location.Y);
            /*
            // (2)
            pnlOCRCalib2 = new Panel();
            pnlOCRCalib2.Size = new Size(160, 180);
            pnlOCRCalib2.BackColor = Color.LightGray;
            pnlOCRCalib2.Location = new Point(pnlOCRCalib1.Location.X,
               pnlOCRCalib1.Location.Y + pnlOCRCalib1.Height + 1);

            Label lblOCRCalib2 = new Label();
            lblOCRCalib2.Text = "Gray Scale";
            lblOCRCalib2.Font = new Font(lblOCRCalib2.Font, FontStyle.Bold);
            lblOCRCalib2.Size = new Size(100, 20);
            lblOCRCalib2.TextAlign = ContentAlignment.MiddleLeft;
            lblOCRCalib2.Location = new Point(10, 10);
            pnlOCRCalib2.Controls.Add(lblOCRCalib2);

            pbOCRCalib2 = new ImageBox();
            pbOCRCalib2.Size = new System.Drawing.Size(320, 180);
            pbOCRCalib2.SizeMode = PictureBoxSizeMode.StretchImage;
            pbOCRCalib2.BackColor = Color.LightBlue;
            pbOCRCalib2.Location = new Point(pbOCRCalib1.Location.X,
                pbOCRCalib1.Location.Y + pbOCRCalib1.Height + 1);

            hbOCRCalib2 = new HistogramBox();
            hbOCRCalib2.Size = new System.Drawing.Size(320, 180);
            hbOCRCalib2.BackColor = Color.LightBlue;
            hbOCRCalib2.Location = new Point(pbOCRCalib2.Location.X + pbOCRCalib2.Width + 1,
                pbOCRCalib2.Location.Y);

            tblOCRCalib2 = new TextBox();
            tblOCRCalib2.Size = new System.Drawing.Size(200, 180);
            tblOCRCalib2.Multiline = true;
            tblOCRCalib2.WordWrap = true;
            tblOCRCalib2.ScrollBars = ScrollBars.Vertical;
            tblOCRCalib2.Location = new Point(hbOCRCalib2.Location.X + hbOCRCalib2.Width + 1,
                hbOCRCalib2.Location.Y);
                */
            // (3)
            pnlOCRCalib3 = new Panel();
            pnlOCRCalib3.Size = new Size(160, ResSize.Height / 4);
            pnlOCRCalib3.BackColor = Color.LightGray;
            pnlOCRCalib3.Location = new Point(pnlOCRCalib1.Location.X,
               pnlOCRCalib1.Location.Y + pnlOCRCalib1.Height + 1);

            cblOCRCalib3 = new CheckBox();
            cblOCRCalib3.Text = "CLAHE";
            cblOCRCalib3.Font = new Font(cblOCRCalib3.Font, FontStyle.Bold);
            cblOCRCalib3.TextAlign = ContentAlignment.MiddleLeft;
            cblOCRCalib3.CheckStateChanged += new EventHandler(cbPopup3_CheckStateChanged);
            cblOCRCalib3.Location = new Point(10, 10);
            pnlOCRCalib3.Controls.Add(cblOCRCalib3);

            Label lblCLAHESize = new Label();
            lblCLAHESize.Text = "Size : ";
            lblCLAHESize.Size = new Size(70, 20);
            lblCLAHESize.TextAlign = ContentAlignment.MiddleRight;
            lblCLAHESize.Location = new Point(cblOCRCalib3.Location.X, cblOCRCalib3.Location.Y + cblOCRCalib3.Height + 10);
            pnlOCRCalib3.Controls.Add(lblCLAHESize);

            tbCLAHEWSize = new TextBox();
            tbCLAHEWSize.Text = "8";
            tbCLAHEWSize.Size = new Size(30, 20);
            tbCLAHEWSize.TextChanged += tbCLAHEWSize_TextChanged;
            tbCLAHEWSize.Location = new Point(lblCLAHESize.Location.X + lblCLAHESize.Width + 2, lblCLAHESize.Location.Y);
            tbCLAHEWSize.TextAlign = HorizontalAlignment.Center;
            pnlOCRCalib3.Controls.Add(tbCLAHEWSize);

            Label lblCLAHESizeX = new Label();
            lblCLAHESizeX.Text = "x";
            lblCLAHESizeX.Size = new Size(8, 20);
            lblCLAHESizeX.Location = new Point(tbCLAHEWSize.Location.X + tbCLAHEWSize.Width + 2, tbCLAHEWSize.Location.Y);
            pnlOCRCalib3.Controls.Add(lblCLAHESizeX);

            tbCLAHEHSize = new TextBox();
            tbCLAHEHSize.Text = "8";
            tbCLAHEHSize.Size = new Size(30, 20);
            tbCLAHEHSize.TextChanged += tbCLAHEHSize_TextChanged;
            tbCLAHEHSize.Location = new Point(lblCLAHESizeX.Location.X + lblCLAHESizeX.Width + 2, lblCLAHESizeX.Location.Y);
            tbCLAHEHSize.TextAlign = HorizontalAlignment.Center;
            pnlOCRCalib3.Controls.Add(tbCLAHEHSize);

            Label lblCLAHEClip = new Label();
            lblCLAHEClip.Text = "Clip Limit : ";
            lblCLAHEClip.Size = new Size(70, 20);
            lblCLAHEClip.TextAlign = ContentAlignment.MiddleRight;
            lblCLAHEClip.Location = new Point(lblCLAHESize.Location.X, lblCLAHESize.Location.Y + lblCLAHESize.Height + 5);
            pnlOCRCalib3.Controls.Add(lblCLAHEClip);

            tbarCLAHEClip = new TrackBar();
            tbarCLAHEClip.Size = new Size(40, 20);
            tbarCLAHEClip.Maximum = 100;
            tbarCLAHEClip.TickStyle = TickStyle.None;
            tbarCLAHEClip.Value = 40;
            tbarCLAHEClip.Scroll += new EventHandler(tbarCLAHEClip_Scroll);
            tbarCLAHEClip.Location = new Point(lblCLAHEClip.Location.X + lblCLAHEClip.Width + 2, lblCLAHEClip.Location.Y);
            pnlOCRCalib3.Controls.Add(tbarCLAHEClip);

            tbCLAHEClip = new TextBox();
            tbCLAHEClip.Text = "40";
            tbCLAHEClip.Size = new Size(30, 20);
            tbCLAHEClip.TextChanged += new EventHandler(tbCLAHEClip_TextChanged);
            tbCLAHEClip.Location = new Point(tbarCLAHEClip.Location.X + tbarCLAHEClip.Width + 2, tbarCLAHEClip.Location.Y);
            tbCLAHEClip.TextAlign = HorizontalAlignment.Center;
            pnlOCRCalib3.Controls.Add(tbCLAHEClip);

            pbOCRCalib3 = new ImageBox();
            pbOCRCalib3.Size = new System.Drawing.Size(320, 180);
            pbOCRCalib3.SizeMode = PictureBoxSizeMode.StretchImage;
            pbOCRCalib3.BackColor = Color.LightBlue;
            pbOCRCalib3.Location = new Point(pbOCRCalib1.Location.X,
                pbOCRCalib1.Location.Y + pbOCRCalib1.Height + 1);

            hbOCRCalib3 = new HistogramBox();
            hbOCRCalib3.Size = new System.Drawing.Size(320, 180);
            hbOCRCalib3.BackColor = Color.LightBlue;
            hbOCRCalib3.Location = new Point(pbOCRCalib3.Location.X + pbOCRCalib3.Width + 1,
                pbOCRCalib3.Location.Y);

            tblOCRCalib3 = new TextBox();
            tblOCRCalib3.Size = new System.Drawing.Size(200, 180);
            tblOCRCalib3.Multiline = true;
            tblOCRCalib3.WordWrap = true;
            tblOCRCalib3.ScrollBars = ScrollBars.Vertical;
            tblOCRCalib3.Location = new Point(hbOCRCalib3.Location.X + hbOCRCalib3.Width + 1,
                hbOCRCalib3.Location.Y);

            // (4)
            pnlOCRCalib4 = new Panel();
            pnlOCRCalib4.Size = new Size(160, 180);
            pnlOCRCalib4.BackColor = Color.LightGray;
            pnlOCRCalib4.Location = new Point(pnlOCRCalib3.Location.X,
               pnlOCRCalib3.Location.Y + pnlOCRCalib3.Height + 1);

            cblOCRCalib4 = new CheckBox();
            cblOCRCalib4.Text = "Sharpen";
            cblOCRCalib4.Font = new Font(cblOCRCalib4.Font, FontStyle.Bold);
            cblOCRCalib4.TextAlign = ContentAlignment.MiddleLeft;
            cblOCRCalib4.CheckStateChanged += new EventHandler(cbPopup4_CheckStateChanged);
            cblOCRCalib4.Location = new Point(10, 10);
            pnlOCRCalib4.Controls.Add(cblOCRCalib4);

            Label lblGaussianSigmaX = new Label();
            lblGaussianSigmaX.Text = "Gaussian Sigma X :";
            lblGaussianSigmaX.Location = new Point(cblOCRCalib4.Location.X, cblOCRCalib4.Location.Y + cblOCRCalib4.Height + 10);
            pnlOCRCalib4.Controls.Add(lblGaussianSigmaX);

            tbGaussianSigmaX = new TextBox();
            tbGaussianSigmaX.Size = new System.Drawing.Size(30, 30);
            tbGaussianSigmaX.TextChanged += new EventHandler(tbGaussianSigmaX_TextChanged);
            tbGaussianSigmaX.Location = new Point(lblGaussianSigmaX.Location.X + lblGaussianSigmaX.Width + 1,
                lblGaussianSigmaX.Location.Y);
            pnlOCRCalib4.Controls.Add(tbGaussianSigmaX);

            Label lblWeightAlpha = new Label();
            lblWeightAlpha.Text = "Alpha Weight :";
            lblWeightAlpha.Location = new Point(lblGaussianSigmaX.Location.X, lblGaussianSigmaX.Location.Y + lblGaussianSigmaX.Height + 2);
            pnlOCRCalib4.Controls.Add(lblWeightAlpha);

            tbWeightAlpha = new TextBox();
            tbWeightAlpha.Size = new System.Drawing.Size(30, 30);
            tbWeightAlpha.TextChanged += new EventHandler(tbWeightAlpha_TextChanged);
            tbWeightAlpha.Location = new Point(lblWeightAlpha.Location.X + lblWeightAlpha.Width + 1,
                lblWeightAlpha.Location.Y);
            pnlOCRCalib4.Controls.Add(tbWeightAlpha);

            Label lblWeightBeta = new Label();
            lblWeightBeta.Text = "Beta Weight :";
            lblWeightBeta.Location = new Point(lblWeightAlpha.Location.X, lblWeightAlpha.Location.Y + lblWeightAlpha.Height + 2);
            pnlOCRCalib4.Controls.Add(lblWeightBeta);

            tbWeightBeta = new TextBox();
            tbWeightBeta.Size = new System.Drawing.Size(30, 30);
            tbWeightBeta.TextChanged += new EventHandler(tbWeightBeta_TextChanged);
            tbWeightBeta.Location = new Point(lblWeightBeta.Location.X + lblWeightBeta.Width + 1,
                lblWeightBeta.Location.Y);
            pnlOCRCalib4.Controls.Add(tbWeightBeta);

            Label lblWeightGamma = new Label();
            lblWeightGamma.Text = "Gamma Weight :";
            lblWeightGamma.Location = new Point(lblWeightBeta.Location.X, lblWeightBeta.Location.Y + lblWeightBeta.Height + 2);
            pnlOCRCalib4.Controls.Add(lblWeightGamma);

            tbWeightGamma = new TextBox();
            tbWeightGamma.Size = new System.Drawing.Size(30, 30);
            tbWeightGamma.TextChanged += new EventHandler(tbWeightGamma_TextChanged);
            tbWeightGamma.Location = new Point(lblWeightGamma.Location.X + lblWeightGamma.Width + 1,
                lblWeightGamma.Location.Y);
            pnlOCRCalib4.Controls.Add(tbWeightGamma);

            pbOCRCalib4 = new ImageBox();
            pbOCRCalib4.Size = new System.Drawing.Size(320, 180);
            pbOCRCalib4.SizeMode = PictureBoxSizeMode.StretchImage;
            pbOCRCalib4.BackColor = Color.LightBlue;
            pbOCRCalib4.Location = new Point(pbOCRCalib3.Location.X,
                pbOCRCalib3.Location.Y + pbOCRCalib3.Height + 1);

            hbOCRCalib4 = new HistogramBox();
            hbOCRCalib4.Size = new System.Drawing.Size(320, 180);
            hbOCRCalib4.BackColor = Color.LightBlue;
            hbOCRCalib4.Location = new Point(pbOCRCalib4.Location.X + pbOCRCalib4.Width + 1,
                pbOCRCalib4.Location.Y);

            tblOCRCalib4 = new TextBox();
            tblOCRCalib4.Size = new System.Drawing.Size(200, 180);
            tblOCRCalib4.Multiline = true;
            tblOCRCalib4.WordWrap = true;
            tblOCRCalib4.ScrollBars = ScrollBars.Vertical;
            tblOCRCalib4.Location = new Point(hbOCRCalib4.Location.X + hbOCRCalib4.Width + 1,
                hbOCRCalib4.Location.Y);

            // (5)
            pnlOCRCalib5 = new Panel();
            pnlOCRCalib5.Size = new Size(160, 180);
            pnlOCRCalib5.BackColor = Color.LightGray;
            pnlOCRCalib5.Location = new Point(pnlOCRCalib4.Location.X,
               pnlOCRCalib4.Location.Y + pnlOCRCalib4.Height + 1);

            Label lblOCRCalib5 = new Label();
            lblOCRCalib5.Text = "Binarization";
            lblOCRCalib5.Font = new Font(lblOCRCalib5.Font, FontStyle.Bold);
            lblOCRCalib5.Location = new Point(10, 10);
            pnlOCRCalib5.Controls.Add(lblOCRCalib5);

            cbInvBinarize = new CheckBox();
            cbInvBinarize.Text = "Invert Binarization";
            cbInvBinarize.Size = new Size(200, 20);
            cbInvBinarize.Location = new Point(lblOCRCalib5.Location.X, lblOCRCalib5.Location.Y + lblOCRCalib5.Height + 10);
            cbInvBinarize.CheckStateChanged += new EventHandler(cbInvBinarize_CheckStateChanged);
            pnlOCRCalib5.Controls.Add(cbInvBinarize);

            Label lblThresBinarize = new Label();
            lblThresBinarize.Text = "Threshold :";
            lblThresBinarize.Location = new Point(cbInvBinarize.Location.X, cbInvBinarize.Location.Y + cbInvBinarize.Height + 2);
            pnlOCRCalib5.Controls.Add(lblThresBinarize);

            tbThresBinarize = new TextBox();
            tbThresBinarize.Size = new System.Drawing.Size(30, 30);
            tbThresBinarize.TextChanged += new EventHandler(tbThresBinarize_TextChanged);
            tbThresBinarize.Location = new Point(lblThresBinarize.Location.X + lblThresBinarize.Width + 1,
                lblThresBinarize.Location.Y);
            pnlOCRCalib5.Controls.Add(tbThresBinarize);

            pbOCRCalib5 = new ImageBox();
            pbOCRCalib5.Size = new System.Drawing.Size(320, 180);
            pbOCRCalib5.SizeMode = PictureBoxSizeMode.StretchImage;
            pbOCRCalib5.BackColor = Color.LightBlue;
            pbOCRCalib5.Location = new Point(pbOCRCalib4.Location.X,
                pbOCRCalib4.Location.Y + pbOCRCalib4.Height + 1);

            hbOCRCalib5 = new HistogramBox();
            hbOCRCalib5.Size = new System.Drawing.Size(320, 180);
            hbOCRCalib5.BackColor = Color.LightBlue;
            hbOCRCalib5.Location = new Point(pbOCRCalib5.Location.X + pbOCRCalib5.Width + 1,
                pbOCRCalib5.Location.Y);

            tblOCRCalib5 = new TextBox();
            tblOCRCalib5.Size = new System.Drawing.Size(200, 180);
            tblOCRCalib5.Multiline = true;
            tblOCRCalib5.WordWrap = true;
            tblOCRCalib5.ScrollBars = ScrollBars.Vertical;
            tblOCRCalib5.Location = new Point(hbOCRCalib5.Location.X + hbOCRCalib5.Width + 1,
                hbOCRCalib5.Location.Y);

            Panel pnlBottomOCRCalib = new Panel();
            pnlBottomOCRCalib.Size = new Size(1000, 1000);
            pnlBottomOCRCalib.Controls.Add(pnlOCRCalib1);
            pnlBottomOCRCalib.Controls.Add(pnlOCRCalib2);
            pnlBottomOCRCalib.Controls.Add(pnlOCRCalib3);
            pnlBottomOCRCalib.Controls.Add(pnlOCRCalib4);
            pnlBottomOCRCalib.Controls.Add(pnlOCRCalib5);
            pnlBottomOCRCalib.Controls.Add(pbOCRCalib1);
            pnlBottomOCRCalib.Controls.Add(pbOCRCalib2);
            pnlBottomOCRCalib.Controls.Add(pbOCRCalib3);
            pnlBottomOCRCalib.Controls.Add(pbOCRCalib4);
            pnlBottomOCRCalib.Controls.Add(pbOCRCalib5);
            pnlBottomOCRCalib.Controls.Add(hbOCRCalib1);
            pnlBottomOCRCalib.Controls.Add(hbOCRCalib2);
            pnlBottomOCRCalib.Controls.Add(hbOCRCalib3);
            pnlBottomOCRCalib.Controls.Add(hbOCRCalib4);
            pnlBottomOCRCalib.Controls.Add(hbOCRCalib5);
            pnlBottomOCRCalib.Controls.Add(tblOCRCalib1);
            pnlBottomOCRCalib.Controls.Add(tblOCRCalib2);
            pnlBottomOCRCalib.Controls.Add(tblOCRCalib3);
            pnlBottomOCRCalib.Controls.Add(tblOCRCalib4);
            pnlBottomOCRCalib.Controls.Add(tblOCRCalib5);
            fmOCRCalib.Controls.Add(pnlBottomOCRCalib);
            pnlBottomOCRCalib.Location = new Point(pnlTopOCRCalib.Location.X, pnlTopOCRCalib.Location.Y + pnlTopOCRCalib.Height + 1);

            btOCRRefresh = new Button();
            btOCRRefresh.Text = "Refresh";
            btOCRRefresh.Click += new EventHandler(btOCRRefresh_Click);
            btOCRRefresh.Location = new Point(pnlTopOCRCalib.Location.X + pnlTopOCRCalib.Width - btOCRRefresh.Width - 5, lblOCRSettings.Location.Y);
            pnlTopOCRCalib.Controls.Add(btOCRRefresh);

            OCRSetting_init();

            fmOCRCalib.Show();
            return true;
        }

        private static void fmOCRCalib_Close(object sender, FormClosingEventArgs e)
        {
            if (fmOCRCalib != null && fmOCRCalib.IsDisposed)
                return;

            fmOCRCalib.Dispose();
        }

        private static void OCRSetting_init()
        {
            cblOCRCalib3.Checked = false;
            cblOCRCalib4.Checked = true;

            tbCLAHEClip.Text = "2";
            tbCLAHEHSize.Text = "10";
            tbCLAHEWSize.Text = "10";
            tbWeightAlpha.Text = "1.5";
            tbWeightBeta.Text = "-1.5";
            tbWeightGamma.Text = "0";
            tbGaussianSigmaX.Text = "1.5";
            tbThresBinarize.Text = "210";
            tbZoomScale.Text = "4";
        }

        private static void btOCRRefresh_Click(object sender, EventArgs eventArgs)
        {
            TextExtraction(cbOCRItemList.Text);
        }

        private static void cbInvBinarize_CheckStateChanged(object sender, EventArgs e)
        {
            if (cbInvBinarize.CheckState == CheckState.Checked)
                isInvBinarize = true;
            else
                isInvBinarize = false;
        }

        private static void tbZoomScale_TextChanged(object sender, EventArgs e)
        {
            if (tbZoomScale.Text != "")
                zoomScale = Double.Parse(tbZoomScale.Text);
        }

        private static void tbGaussianSigmaX_TextChanged(object sender, EventArgs e)
        {
            if (tbGaussianSigmaX.Text != "")
                gaussianSigmaX = Double.Parse(tbGaussianSigmaX.Text);
        }

        private static void tbWeightAlpha_TextChanged(object sender, EventArgs e)
        {
            if (tbWeightAlpha.Text != "")
                weightAlpha = Double.Parse(tbWeightAlpha.Text);
        }

        private static void tbWeightBeta_TextChanged(object sender, EventArgs e)
        {
            if (tbWeightBeta.Text != "" && tbWeightBeta.Text != "-")
                weightBeta = Double.Parse(tbWeightBeta.Text);
        }

        private static void tbWeightGamma_TextChanged(object sender, EventArgs e)
        {
            if (tbWeightGamma.Text != "")
                weightGamma = Double.Parse(tbWeightGamma.Text);
        }

        private static void tbThresBinarize_TextChanged(object sender, EventArgs e)
        {
            if (tbThresBinarize.Text != "" && Int32.Parse(tbThresBinarize.Text) > -1)
                thresBinarize = Int32.Parse(tbThresBinarize.Text);
        }

        private static void tbCLAHEClip_TextChanged(object sender, EventArgs e)
        {
            if (tbCLAHEClip.Text != "" && int.Parse(tbCLAHEClip.Text) > -1 && int.Parse(tbCLAHEClip.Text) < 101)
            {
                tbarCLAHEClip.Value = int.Parse(tbCLAHEClip.Text);
                CLAHEClip = int.Parse(tbCLAHEClip.Text);
            }
        }

        private static void tbCLAHEHSize_TextChanged(object sender, EventArgs e)
        {
            if (tbCLAHEHSize.Text != "" && int.Parse(tbCLAHEHSize.Text) > -1)
            {
                CLAHEHSize = int.Parse(tbCLAHEHSize.Text);
            }
        }

        private static void tbCLAHEWSize_TextChanged(object sender, EventArgs e)
        {
            if (tbCLAHEWSize.Text != "" && int.Parse(tbCLAHEWSize.Text) > -1)
            {
                CLAHEWSize = int.Parse(tbCLAHEWSize.Text);
            }
        }

        private static void tbarCLAHEClip_Scroll(object sender, EventArgs e)
        {
            tbCLAHEClip.Text = tbarCLAHEClip.Value.ToString();
        }

        private static void cbPopup3_CheckStateChanged(object sender, EventArgs e)
        {
            if (cblOCRCalib3.CheckState == CheckState.Checked)
                isCLAHE = true;
            else
                isCLAHE = false;
        }

        private static void cbPopup4_CheckStateChanged(object sender, EventArgs e)
        {
            if (cblOCRCalib4.CheckState == CheckState.Checked)
                isSharpenize = true;
            else
                isSharpenize = false;
        }

        private static Mat CropImage(EN_OCR_ITEM item, Image<Bgr, Byte> processImg)
        {
            string[] words = new string[4];
            
            processImg.ROI = Rectangle.Empty;

            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == test_profile)
                {
                    switch (item)
                    {
                        case EN_OCR_ITEM.OCR_ITEM_SOURCE:
                            words = _item[i].Source.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_LABEL:
                            words = _item[i].Label.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_INFO:
                            words = _item[i].Info.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_OPT1:
                            words = _item[i].Opt1.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_OPT2:
                            words = _item[i].Opt2.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_OPT3:
                            words = _item[i].Opt3.Split(';');
                            break;
                        case EN_OCR_ITEM.OCR_ITEM_FULL:
                            words = _item[i].Full.Split(';');
                            break;
                    }
                    Rectangle roi = new Rectangle(convRealSize(int.Parse(words[2]), false), convRealSize(int.Parse(words[3]), true), convRealSize(int.Parse(words[0]), false), convRealSize(int.Parse(words[1]), true));
                    processImg.ROI = roi;
                }
            }
            processImg.Resize(zoomScale, Emgu.CV.CvEnum.Inter.Linear);

            return processImg.Mat;
        }

        private static Int32 convRealSize(Int32 orgSize, Boolean mode)
        {
            // mode 1 : Height, 0 : Width
            Int32 ConvertSize = 0;

            if (orgSize != 0)
            {
                if (mode)
                    ConvertSize = orgSize * ResSize.Height / 600;
                else
                    ConvertSize = orgSize * ResSize.Width / 800;
            }

            return ConvertSize;
        }

        [Obsolete]
        private static String ConvertText(EN_OCR_ITEM item)
        {
            string ret = "";
            switch (item)
            {
                case EN_OCR_ITEM.OCR_ITEM_INFO:
                    ret = "Info";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_LABEL:
                    ret = "Label";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_SOURCE:
                    ret = "Source";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_OPT1:
                    ret = "Option1";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_OPT2:
                    ret = "Option2";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_OPT3:
                    ret = "Option3";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_FULL:
                    ret = "Full";
                    break;
            }

            return ret;
        }

        private static EN_OCR_ITEM ConvertText(String item)
        {
            EN_OCR_ITEM ret = EN_OCR_ITEM.OCR_ITEM_MAX;
            if (item == "Source")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_SOURCE;
            }
            else if (item == "Label")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_LABEL;
            }
            else if (item == "Info")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_INFO;
            }
            else if (item == "Option1")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_OPT1;
            }
            else if (item == "Option2")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_OPT2;
            }
            else if (item == "Option3")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_OPT3;
            }
            else if (item == "Full")
            {
                ret = EN_OCR_ITEM.OCR_ITEM_FULL;
            }
            return ret;
        }

        private static String TextExtractionFromParent(EN_OCR_ITEM item)
        {
            Mat processImg = new Mat();
            List<string> tempList = new List<String>();
            string retStr = "null";

            // Do 10 attemps for OCR processing
            for (int i = 0; i < 10; i++)
            {
                Saal.videoDisplay_cv.Retrieve(processImg, 0);

                processImg = CropImage(item, processImg.ToImage<Bgr, Byte>());

                if (isCLAHE)
                {
                    processImg = CLAHEHandler(processImg);
                    OperationDialog_updLog("<OCR Pre-processing> Running Clahe...", Color.White);
                    OperationDialog_updRefImg(processImg.Bitmap);
                }

                processImg = GrayScaleHandler(processImg);

                if (isSharpenize)
                {
                    processImg = SharpenHandler(processImg);
                    OperationDialog_updLog("<OCR Pre-processing> Running Sharpenize...", Color.White);
                    OperationDialog_updRefImg(processImg.Bitmap);
                }
                processImg = BinarizationHandler(processImg);
                OperationDialog_updLog("<OCR Pre-processing> Running Binarization...", Color.White);
                OperationDialog_updRefImg(processImg.Bitmap);

                OperationDialog_updLog("<OCR Processing> Running OCR...", Color.White);

                string temp = OCRHandler(item, processImg.ToImage<Gray, Byte>());
                tempList.Add(temp);
            }

            int tempCnt = 0;
            foreach (var grp in tempList.GroupBy(k => k))
            {
                Console.WriteLine("OCR result --> {0} : {1}", grp.Key, grp.Count());
                // Case for checking resolution
                if (item == EN_OCR_ITEM.OCR_ITEM_INFO)
                {
                    if (grp.Count() > tempCnt)
                    {
                        tempCnt = grp.Count();
                        retStr = grp.Key;
                        ImageHandler.OperationDialog_updLog("<OCR Post-processing> OCR Highest frequency : " + retStr + " (" + tempCnt + "/10)", Color.White);
                    }
                    continue;
                }
                // Other case
                else
                {
                    if (grp.Count() > tempCnt && grp.Key != "null")
                    {
                        tempCnt = grp.Count();
                        retStr = grp.Key;
                        ImageHandler.OperationDialog_updLog("<OCR Post-processing> OCR Highest frequency : " + retStr + " (" + tempCnt + "/10)", Color.White);
                    }
                }                
            }
            updOpFlag = true;
            return retStr;
        }

        private static String TextExtractionBitmap(Bitmap img)
        {
            Image<Gray, Byte> processImg = new Image<Gray, Byte>(img);
            tess.Recognize(processImg);
            return tess.GetText();
        }

        private static void TextExtraction(string itemStr)
        {
            Mat processImg = new Mat();

            Saal.videoDisplay_cv.Retrieve(processImg, 0);
            EN_OCR_ITEM item = EN_OCR_ITEM.OCR_ITEM_MAX;

            item = ConvertText(itemStr);

            processImg = CropImage(item, processImg.ToImage<Bgr, Byte>());

            // (1)
            // Initialize
            pbOCRCalib1.Image = null;
            hbOCRCalib1.ClearHistogram();
            tblOCRCalib1.Clear();
            // Do operation
            pnlOCRCalib1.BackColor = Color.LightGreen;
            pnlOCRCalib1.Refresh();
            pbOCRCalib1.Image = processImg;
            hbOCRCalib1.GenerateHistograms(processImg.ToImage<Bgr, Byte>(), 256);
            hbOCRCalib1.Enabled = true;
            tblOCRCalib1.Text = OCRHandler(item, processImg.ToImage<Gray, Byte>());
            tblOCRCalib1.Refresh();
            pnlOCRCalib1.BackColor = Color.LightGray;
            // Refresh
            pbOCRCalib1.Refresh();
            hbOCRCalib1.Refresh();
            pnlOCRCalib1.Refresh();

            /*
            // (2) Convert to grayscale 
            // Initialize
            pbOCRCalib2.Image = null;
            hbOCRCalib2.ClearHistogram();
            tblOCRCalib2.Clear();
            // Do operation                        
            pnlOCRCalib2.BackColor = Color.LightGreen;
            pnlOCRCalib2.Refresh();
            processImg = GrayScaleHandler(processImg);
            pbOCRCalib2.Image = processImg;
            hbOCRCalib2.GenerateHistograms(processImg.ToImage<Gray, Byte>(), 256);
            hbOCRCalib2.Enabled = true;
            tblOCRCalib2.Text = OCRHandler(processImg.ToImage<Gray, Byte>());
            pnlOCRCalib2.BackColor = Color.LightGray;
            pnlOCRCalib2.Refresh();                        
            // Refresh
            pbOCRCalib2.Refresh();
            hbOCRCalib2.Refresh();
            tblOCRCalib2.Refresh();
            */

            // (3) Add contrast using Contrast Limited Adaptive Histogram Equalization
            // Initialize
            pbOCRCalib3.Image = null;
            hbOCRCalib3.ClearHistogram();
            tblOCRCalib3.Clear();
            // Do operation
            if (isCLAHE)
            {
                pnlOCRCalib3.BackColor = Color.LightGreen;
                pnlOCRCalib3.Refresh();
                processImg = CLAHEHandler(processImg);
                pbOCRCalib3.Image = processImg;
                hbOCRCalib3.GenerateHistograms(processImg.ToImage<Bgr, Byte>(), 256);
                hbOCRCalib3.Enabled = true;
                tblOCRCalib3.Text = OCRHandler(item, processImg.ToImage<Gray, Byte>());
                pnlOCRCalib3.BackColor = Color.LightGray;
                pnlOCRCalib3.Refresh();
            }
            // Refresh
            pbOCRCalib3.Refresh();
            hbOCRCalib3.Refresh();
            tblOCRCalib3.Refresh();

            // Change to grayscale
            processImg = GrayScaleHandler(processImg);

            // (4) Sharpen text details and edge
            // Initialize
            pbOCRCalib4.Image = null;
            hbOCRCalib4.ClearHistogram();
            tblOCRCalib4.Clear();
            // Do operation
            if (isSharpenize)
            {
                pnlOCRCalib4.BackColor = Color.LightGreen;
                pnlOCRCalib4.Refresh();
                processImg = SharpenHandler(processImg);
                pbOCRCalib4.Image = processImg;
                hbOCRCalib4.GenerateHistograms(processImg.ToImage<Gray, Byte>(), 256);
                hbOCRCalib4.Enabled = true;
                tblOCRCalib4.Text = OCRHandler(item, processImg.ToImage<Gray, Byte>());
                pnlOCRCalib4.BackColor = Color.LightGray;
                pnlOCRCalib4.Refresh();
            }
            // Refresh
            pbOCRCalib4.Refresh();
            hbOCRCalib4.Refresh();
            tblOCRCalib4.Refresh();

            // (5) Binarization
            // Initialize            
            pbOCRCalib5.Image = null;
            hbOCRCalib5.ClearHistogram();
            tblOCRCalib5.Clear();
            // Do operation
            pnlOCRCalib5.BackColor = Color.LightGreen;
            pnlOCRCalib5.Refresh();
            processImg = BinarizationHandler(processImg);
            pbOCRCalib5.Image = processImg;
            hbOCRCalib5.GenerateHistograms(processImg.ToImage<Gray, Byte>(), 256);
            hbOCRCalib5.Enabled = true;
            tblOCRCalib5.Text = OCRHandler(item, processImg.ToImage<Gray, Byte>());
            pnlOCRCalib5.BackColor = Color.LightGray;
            pnlOCRCalib5.Refresh();
            // Refresh
            pbOCRCalib5.Refresh();
            hbOCRCalib5.Refresh();
            tblOCRCalib5.Refresh();

            processImg.Dispose();
            MessageBox.Show("Done !!!");
        }

        private static Mat GrayScaleHandler(Mat processImgBefore)
        {
            Console.WriteLine("Convert to gray scale...");
            Mat processImgAfter = new Mat();
            CvInvoke.CvtColor(processImgBefore, processImgAfter, ColorConversion.Bgr2Gray);
            return processImgAfter;
        }

        private static Mat CLAHEHandler(Mat processImgBefore)
        {
            Console.WriteLine("Running CLAHE...");
            Mat processImgHSV = new Mat();
            Mat processImgHSVAfter = new Mat();
            Mat processImgAfter = new Mat();
            Double clipLimit = CLAHEClip; // default setting 40
            Size tileGridSize = new Size(CLAHEWSize, CLAHEHSize); // default setting (8,8)
            Console.WriteLine("0");

            // Convert to HSV
            CvInvoke.CvtColor(processImgBefore, processImgHSV, ColorConversion.Bgr2Hsv);
            Image<Hsv, byte> ImgHSV = processImgHSV.ToImage<Hsv, byte>();

            // Split channels
            Mat[] HSVchannelsBef = ImgHSV.Mat.Split();
            Mat HSVchannelsAft = new Mat();

            // Do CLAHE only to V (H,S remain the same)
            CvInvoke.CLAHE(HSVchannelsBef[2], clipLimit, tileGridSize, HSVchannelsAft);

            // Merge channels
            using (VectorOfMat vm = new VectorOfMat(HSVchannelsBef[0], HSVchannelsBef[1], HSVchannelsAft))
            {
                CvInvoke.Merge(vm, processImgHSVAfter);
            }

            // Convert to RGB
            CvInvoke.CvtColor(processImgHSVAfter, processImgAfter, ColorConversion.Hsv2Bgr);

            HSVchannelsBef[0].Dispose();
            HSVchannelsBef[1].Dispose();
            HSVchannelsBef[2].Dispose();
            HSVchannelsAft.Dispose();
            processImgHSVAfter.Dispose();
            processImgHSV.Dispose();

            return processImgAfter;
        }

        private static Mat SharpenHandler(Mat processImgBefore)
        {
            Console.WriteLine("Sharpening the image...");
            Mat processImgAfter = new Mat();
            Mat processImgGaussian = new Mat();
            Mat processImgMask = new Mat();
            Size gaussianSize = new Size(3, 3);
            //Double gaussianSigmaX = 1.5;

            // Blur the image
            CvInvoke.GaussianBlur(processImgBefore, processImgGaussian, gaussianSize, gaussianSigmaX);

            // Subtract blurred image from original image
            CvInvoke.Subtract(processImgBefore, processImgGaussian, processImgMask);

            // Add the mask to original
            CvInvoke.AddWeighted(processImgBefore, weightAlpha, processImgMask, weightBeta, weightGamma, processImgAfter);
            //CvInvoke.Add(processImgBefore, processImgMask, processImgAfter);

            return processImgAfter;
        }

        private static Mat BinarizationHandler(Mat processImgBefore)
        {
            Console.WriteLine("Binarization...");
            Mat processImgAfter = new Mat();
            Mat processImgThreshold = new Mat();

            if (isInvBinarize)
            {
                //CvInvoke.Threshold(processImgBefore, processImgAfter, 0, 255, ThresholdType.Otsu | ThresholdType.BinaryInv);
                CvInvoke.Threshold(processImgBefore, processImgAfter, thresBinarize, 255, ThresholdType.BinaryInv);
            }

            else
            {
                //CvInvoke.Threshold(processImgBefore, processImgAfter, 0, 255, ThresholdType.Otsu | ThresholdType.Binary);
                CvInvoke.Threshold(processImgBefore, processImgAfter, thresBinarize, 255, ThresholdType.Binary);
            }

            return processImgAfter;
            //return MorphologicalThinning(processImgAfter);
        }

        private static Mat MorphologicalThinning(Mat processImgBefore)
        {
            Mat processImgAfter = new Mat(processImgBefore.Size, DepthType.Cv8U, 4);

            Mat element = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Cross, new Size(3, 3), new Point(1, 1));

            // Morphology

            //CvInvoke.MorphologyEx(processImgBefore, processImgAfter, MorphOp.Gradient, element, new Point(1, 1), 1, BorderType.Default, new MCvScalar(1));
            //CvInvoke.Erode(processImgBefore, processImgAfter, element, new Point(1, 1), 3, BorderType.Default, new MCvScalar(1));

            // Skeleton algorithm
            if (true)
            {
                Mat processImgEroded = new Mat(processImgBefore.Size, DepthType.Cv8U, 4);
                Mat processImgTemp = new Mat(processImgBefore.Size, DepthType.Cv8U, 4);
                Mat processImgSkel = new Mat(processImgBefore.Size, DepthType.Cv8U, 4);
                processImgSkel.SetTo(new MCvScalar(0, 0, 0, 0));
                bool done;
                do
                {

                    CvInvoke.Erode(processImgBefore, processImgEroded, element, new Point(1, 1), 1, BorderType.Default, new MCvScalar(1));
                    CvInvoke.Dilate(processImgEroded, processImgTemp, element, new Point(1, 1), 1, BorderType.Default, new MCvScalar(1)); // temp = open(img)
                    CvInvoke.Subtract(processImgBefore, processImgTemp, processImgTemp);
                    //CvInvoke.BitwiseNot(processImgTemp, processImgTemp);
                    //CvInvoke.BitwiseAnd(processImgBefore, processImgTemp, processImgTemp);
                    CvInvoke.BitwiseOr(processImgSkel, processImgTemp, processImgAfter);
                    //processImgSkel = processImgSkel | processImgTemp;

                    processImgEroded.CopyTo(processImgBefore);
                    done = (CvInvoke.CountNonZero(processImgBefore) == 0);
                } while (!done);
                processImgEroded.Dispose();
                processImgTemp.Dispose();
                processImgSkel.Dispose();
            }
            return processImgAfter;
        }

        private static Mat CannyHandler(Mat processImgBefore)
        {
            Console.WriteLine("Canny...");
            Mat processImgAfter = new Mat();
            double th1 = 0.1;
            double th2 = 0.2;
            CvInvoke.Canny(processImgBefore, processImgAfter, th1, th2);

            return processImgAfter;
        }

        private static String OCRHandler(EN_OCR_ITEM item, Image<Gray, Byte> processImg)
        {
            Console.WriteLine("Running OCR...");
            string ret_str = "null";

            tess.Recognize(processImg);
            ret_str = tess.GetText();
            Console.WriteLine("--------------------------------START--------------------------------------");
            Console.WriteLine("DEBUG OCR RESULT : " + ret_str);
            Console.WriteLine("--------------------------------FINISH--------------------------------------");
            return OCRpostprocessing(item, ret_str);
        }

        private static String OCRpostprocessing(EN_OCR_ITEM item, string strInp)
        {
            Console.WriteLine("Running OCR Post processing...");
            string strOut = "null";     
            switch (item)
            {
                case EN_OCR_ITEM.OCR_ITEM_INFO:
                    //if (Regex.Match(strInp, @"SI\)|S0|SD|5I|50|5D").Success)
                    //    strOut = "SD";
                    char[] whitespace = new char[] { ' ', '\t' };
                    string[] strInp_split = strInp.Split(whitespace);

                    // Check splitted array number
                    int cnt = 0;
                    foreach(string str in strInp_split)
                    {
                        cnt++;
                    }
                    if (cnt < 2)
                    {
                        strOut = strInp_split[0];
                        break;
                    }

                    // Start comparison
                    Int32 comp = 0;
                    Int32 compTemp = 0;
                    foreach (string pat in EnumListUser.OCR_RESOLUTION)
                    {
                        comp = OCRpostprocessing_Info(strInp_split[1], pat);
                        if (comp > compTemp)
                        {
                            Console.WriteLine(strInp_split[1] + " -> " + pat + "(" + comp + "%)");
                            compTemp = comp;
                            strOut = pat;
                        }
                    }                    
                    strOut = strInp_split[0] + " " + strOut;
                    break;
                case EN_OCR_ITEM.OCR_ITEM_OPT1:
                    if (Regex.Match(strInp, @"Close|Cl|ose|C|se|lo").Success)
                        strOut = "Close";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_SOURCE:
                    if (Regex.Match(strInp, @"CAST|ST|CAS").Success)
                        strOut = "CAST";
                    else if (Regex.Match(strInp, @"TV").Success)
                        strOut = "TV";
                    else if (Regex.Match(strInp, @"HDMI1|I1|IL|II|11").Success)
                        strOut = "HDMI1";
                    else if (Regex.Match(strInp, @"HDMI2|I2|IZ|12").Success)
                        strOut = "HDMI2";
                    else if (Regex.Match(strInp, @"HDMI3|I3|13").Success)
                        strOut = "HDMI3";
                    else if (Regex.Match(strInp, @"HDMI4|I4|14").Success)
                        strOut = "HDMI4";
                    else if (Regex.Match(strInp, @"Video|Vdeo|deo|Vito|Who|tho").Success)
                        strOut = "Video";
                    else if (Regex.Match(strInp, @"PC").Success)
                        strOut = "PC";
                    else if (Regex.Match(strInp, @"USB|sb|us").Success)
                        strOut = "USB";
                    break;
                case EN_OCR_ITEM.OCR_ITEM_LABEL:
                    if (Regex.Match(strInp, @"Please|ase|change|cha|source|sou|resolution|res").Success)
                        strOut = "Please change source resolution";
                    else if (Regex.Match(strInp, @"No|Signal|sig|slg").Success)
                        strOut = "No signal";
                    else
                        strOut = "null";
                    break;
            }

            return strOut;
        }

        private static Int32 OCRpostprocessing_Info(string input, string pattern)
        {
            Int32 perc = 0;
            char[] inp_arrs = input.ToCharArray();
            char[] pat_arrs = pattern.ToCharArray();
            int j_temp = 0;
            int match_cnt = 0;
            //Console.WriteLine("Compare : " + input + " -- " + pattern);
            for (int i = 0; i < pat_arrs.Length; i++)
            {
                for (int j = j_temp; j < inp_arrs.Length; j++)
                {
                    //Console.WriteLine("Compare " + pat_arrs[i] + " -- " + inp_arrs[j]);
                    if (pat_arrs[i] == inp_arrs[j])
                    {
                        j_temp = j;
                        match_cnt++;
                        //Console.WriteLine("Found ");
                        break;
                    }
                }
            }

            if(pat_arrs.Length > 0)
                perc = match_cnt * 100 / pat_arrs.Length;

            return perc;
        }

        #endregion

        #region Blurry Check
        private static Boolean fmBlurryCheck_Open()
        {
            if (fmBlurryCheck != null && !fmBlurryCheck.IsDisposed)
                return false;

            fmBlurryCheck = new Form();
            fmBlurryCheck.Size = new System.Drawing.Size(470, 330);
            fmBlurryCheck.MaximizeBox = false;
            fmBlurryCheck.ControlBox = true;
            fmBlurryCheck.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmBlurryCheck.FormClosing += new FormClosingEventHandler(fmBlurryCheck_Close);
            fmBlurryCheck.StartPosition = FormStartPosition.CenterScreen;

            Label lblBlurrySettings = new Label();
            lblBlurrySettings.Text = "Blurry Settings";
            lblBlurrySettings.Font = new Font(lblBlurrySettings.Font, FontStyle.Bold);
            lblBlurrySettings.Location = new Point(10, 10);
            fmBlurryCheck.Controls.Add(lblBlurrySettings);

            Button btBlurryRefresh = new Button();
            btBlurryRefresh.Text = "Refresh";
            btBlurryRefresh.Click += new EventHandler(btBlurryRefresh_Click);
            fmBlurryCheck.Controls.Add(btBlurryRefresh);

            Label lblBlurryThreshold = new Label();
            lblBlurryThreshold.Text = "Threshold :";
            lblBlurryThreshold.Size = new Size(60, 30);
            lblBlurryThreshold.Location = new Point(lblBlurrySettings.Location.X, lblBlurrySettings.Location.Y + lblBlurrySettings.Height + 10);
            fmBlurryCheck.Controls.Add(lblBlurryThreshold);

            tbBlurryThreshold = new TextBox();
            tbBlurryThreshold.Text = "";
            tbBlurryThreshold.Size = new Size(50, 30);
            tbBlurryThreshold.TextChanged += new EventHandler(tbBlurryThreshold_TextChanged);
            tbBlurryThreshold.Location = new Point(lblBlurryThreshold.Location.X + lblBlurryThreshold.Width + 5, lblBlurryThreshold.Location.Y);
            fmBlurryCheck.Controls.Add(tbBlurryThreshold);

            tbBlurryResult = new TextBox();
            tbBlurryResult.Text = "";
            tbBlurryResult.Size = new Size(115, 190);
            tbBlurryResult.Multiline = true;
            tbBlurryResult.WordWrap = true;
            tbBlurryResult.Location = new Point(lblBlurryThreshold.Location.X, lblBlurryThreshold.Location.Y + lblBlurryThreshold.Height + 5);
            fmBlurryCheck.Controls.Add(tbBlurryResult);

            cbBlurryGaussian = new CheckBox();
            cbBlurryGaussian.Text = "Gaussian Blur";
            cbBlurryGaussian.Location = new Point(tbBlurryResult.Location.X, tbBlurryResult.Location.Y + tbBlurryResult.Height + 1);
            cbBlurryGaussian.CheckStateChanged += new EventHandler(cbBlurryGaussian_CheckStateChanged);
            fmBlurryCheck.Controls.Add(cbBlurryGaussian);

            pbBlurryImage = new ImageBox();
            pbBlurryImage.Size = new System.Drawing.Size(320, 240);
            pbBlurryImage.SizeMode = PictureBoxSizeMode.StretchImage;
            pbBlurryImage.BackColor = Color.LightBlue;
            pbBlurryImage.Location = new Point(tbBlurryThreshold.Location.X + tbBlurryThreshold.Width + 5,
                tbBlurryThreshold.Location.Y);
            fmBlurryCheck.Controls.Add(pbBlurryImage);

            btBlurryRefresh.Location = new Point(pbBlurryImage.Location.X + pbBlurryImage.Width - btBlurryRefresh.Width, lblBlurrySettings.Location.Y);

            fmBlurryCheck.Show();
            init_BlurryCheck();

            return true;
        }

        private static void fmBlurryCheck_Close(object sender, FormClosingEventArgs e)
        {
            if (fmBlurryCheck != null && fmBlurryCheck.IsDisposed)
                return;

            fmBlurryCheck.Dispose();
        }

        private static void btBlurryRefresh_Click(object sender, EventArgs eventArgs)
        {
            blurryCheck();
        }

        private static void tbBlurryThreshold_TextChanged(object sender, EventArgs e)
        {
            if (tbBlurryThreshold.Text != "" && int.Parse(tbBlurryThreshold.Text) > -1)
                thresBlurry = int.Parse(tbBlurryThreshold.Text);
        }

        private static void cbBlurryGaussian_CheckStateChanged(object sender, EventArgs e)
        {
            if (cbBlurryGaussian.CheckState == CheckState.Checked)
                isGaussianBlur = true;
            else
                isGaussianBlur = false;
        }

        private static void init_BlurryCheck()
        {
            cbBlurryGaussian.CheckState = CheckState.Checked;
            tbBlurryThreshold.Text = "100";
        }

        private static Boolean blurryCheck()
        {
            bool ret = false;
            Mat processImg = new Mat();

            tbBlurryResult.Text = "Running...";

            Saal.videoDisplay_cv.Retrieve(processImg, 0);
            pbBlurryImage.Image = processImg;
            pbBlurryImage.Refresh();

            // Remove noise by blurring with Gaussian Filter
            if (isGaussianBlur)
            {
                processImg = GaussianBlurHandler(processImg);
                pbBlurryImage.Image = processImg;
                pbBlurryImage.Refresh();
            }

            // Change color image to gray
            processImg = GrayScaleHandler(processImg);
            pbBlurryImage.Image = processImg;
            pbBlurryImage.Refresh();

            // Implement Laplace algorithm
            processImg = LaplacianHandler(processImg);

            List<double> data = new List<double>();

            Image<Gray, byte> img = processImg.ToImage<Gray, Byte>();
            for (int i = 0; i < img.Rows; i++)
            {
                for (int j = 0; j < img.Cols; j++)
                {
                    data.Add(img.Data[i, j, 0]);
                }
            }

            if (data.Variance() > thresBlurry)
                ret = true;
            else
                ret = false;

            tbBlurryResult.Text = "";
            tbBlurryResult.Text = "Variance value : " + data.Variance().ToString() + "\r\n";
            tbBlurryResult.Text = tbBlurryResult.Text + "\r\n";
            tbBlurryResult.Text = tbBlurryResult.Text + "Result : " + ret.ToString();

            processImg.Dispose();

            return ret;
        }

        #region Extension
        public static double Mean(this List<double> values, int start, int end)
        {
            double s = 0;

            for (int i = start; i < end; i++)
            {
                s += values[i];
            }

            return s / (end - start);
        }

        public static double Mean(this List<double> values)
        {
            return values.Count == 0 ? 0 : values.Mean(0, values.Count);
        }

        public static double Variance(this List<double> values)
        {
            return values.Variance(values.Mean(), 0, values.Count);
        }

        public static double Variance(this List<double> values, double mean)
        {
            return values.Variance(mean, 0, values.Count);
        }

        public static double Variance(this List<double> values, double mean, int start, int end)
        {
            double variance = 0;

            for (int i = start; i < end; i++)
            {
                variance += Math.Pow((values[i] - mean), 2);
            }

            int n = end - start;
            if (start > 0) n -= 1;

            return variance / (n);
        }
        #endregion

        private static Mat GaussianBlurHandler(Mat processImgBefore)
        {
            Console.WriteLine("Gaussian Blur...");
            Mat processImgAfter = new Mat();
            Size sz = new Size(3, 3);
            CvInvoke.GaussianBlur(processImgBefore, processImgAfter, sz, 0, 0, BorderType.Default);
            return processImgAfter;
        }

        private static Mat LaplacianHandler(Mat processImgBefore)
        {
            Console.WriteLine("Laplacian...");
            int kernel_size = 3;
            int scale = 1;
            int delta = 0;
            Mat processImgAfter = new Mat();
            CvInvoke.Laplacian(processImgBefore, processImgAfter, DepthType.Cv64F, kernel_size, scale, delta, BorderType.Default);

            return processImgAfter;
        }
        #endregion

        #region Similarity Check
        private static Boolean fmSimilarityCheck_Open()
        {
            if (fmSimilarityCheck != null && !fmSimilarityCheck.IsDisposed)
                return false;

            fmSimilarityCheck = new Form();
            fmSimilarityCheck.Size = new System.Drawing.Size(690, 480);
            fmSimilarityCheck.MaximizeBox = false;
            fmSimilarityCheck.ControlBox = true;
            fmSimilarityCheck.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmSimilarityCheck.FormClosing += new FormClosingEventHandler(fmSimilarityCheck_Close);
            fmSimilarityCheck.StartPosition = FormStartPosition.CenterScreen;

            Label lblSimilarityCheck = new Label();
            lblSimilarityCheck.Text = "Similarity Check Settings";
            lblSimilarityCheck.Font = new Font(lblSimilarityCheck.Font, FontStyle.Bold);
            lblSimilarityCheck.Location = new Point(10, 10);
            fmSimilarityCheck.Controls.Add(lblSimilarityCheck);

            // Reference
            Panel pnlSimilarityRef = new Panel();
            pnlSimilarityRef.Size = new Size(330, 280);
            pnlSimilarityRef.Location = new Point(lblSimilarityCheck.Location.X, lblSimilarityCheck.Location.Y + lblSimilarityCheck.Height + 10);
            fmSimilarityCheck.Controls.Add(pnlSimilarityRef);

            Label lblSimilarityRef = new Label();
            lblSimilarityRef.Text = "Reference Image";
            lblSimilarityRef.Location = new Point(pnlSimilarityRef.Width / 2 - lblSimilarityRef.Width / 2, 1);
            pnlSimilarityRef.Controls.Add(lblSimilarityRef);

            pbSimilarityRef = new ImageBox();
            pbSimilarityRef.Size = new System.Drawing.Size(320, 240);
            pbSimilarityRef.SizeMode = PictureBoxSizeMode.StretchImage;
            pbSimilarityRef.BackColor = Color.LightBlue;
            pbSimilarityRef.Location = new Point(0,
                lblSimilarityRef.Location.Y + lblSimilarityRef.Height + 1);
            pnlSimilarityRef.Controls.Add(pbSimilarityRef);

            // Current
            Panel pnlSimilarityCur = new Panel();
            pnlSimilarityCur.Size = new Size(330, 280);
            pnlSimilarityCur.Location = new Point(pnlSimilarityRef.Location.X + pnlSimilarityRef.Width + 5, pnlSimilarityRef.Location.Y);
            fmSimilarityCheck.Controls.Add(pnlSimilarityCur);

            Label lblSimilarityCur = new Label();
            lblSimilarityCur.Text = "NG Image";
            lblSimilarityCur.Location = new Point(pnlSimilarityCur.Width / 2 - lblSimilarityCur.Width / 2, 1);
            pnlSimilarityCur.Controls.Add(lblSimilarityCur);

            pbSimilarityCur = new ImageBox();
            pbSimilarityCur.Size = new System.Drawing.Size(320, 240);
            pbSimilarityCur.SizeMode = PictureBoxSizeMode.StretchImage;
            pbSimilarityCur.BackColor = Color.LightBlue;
            pbSimilarityCur.Location = new Point(0,
                lblSimilarityCur.Location.Y + lblSimilarityCur.Height + 1);
            pnlSimilarityCur.Controls.Add(pbSimilarityCur);

            // Bottom
            Panel pnlSimilarityBottom = new Panel();
            pnlSimilarityBottom.Size = new Size(665, 150);
            pnlSimilarityBottom.Location = new Point(pnlSimilarityRef.Location.X, pnlSimilarityRef.Location.Y + pnlSimilarityRef.Height + 5);
            fmSimilarityCheck.Controls.Add(pnlSimilarityBottom);

            Label lblSimilarityThresPSNR = new Label();
            lblSimilarityThresPSNR.Text = "PSNR Threshold : ";
            lblSimilarityThresPSNR.Location = new Point(10, 10);
            pnlSimilarityBottom.Controls.Add(lblSimilarityThresPSNR);

            tbSimilarityThresPSNR = new TextBox();
            tbSimilarityThresPSNR.Text = "";
            tbSimilarityThresPSNR.Size = new Size(50, 30);
            tbSimilarityThresPSNR.TextChanged += new EventHandler(tbSimilarityThresPSNR_TextChanged);
            tbSimilarityThresPSNR.Location = new Point(lblSimilarityThresPSNR.Location.X + lblSimilarityThresPSNR.Width + 2, lblSimilarityThresPSNR.Location.Y);
            pnlSimilarityBottom.Controls.Add(tbSimilarityThresPSNR);

            Label lblSimilarityThresSSIM = new Label();
            lblSimilarityThresSSIM.Text = "SSIM Threshold : ";
            lblSimilarityThresSSIM.Location = new Point(tbSimilarityThresPSNR.Location.X + tbSimilarityThresPSNR.Width + 10, tbSimilarityThresPSNR.Location.Y);
            pnlSimilarityBottom.Controls.Add(lblSimilarityThresSSIM);

            tbSimilarityThresSSIM = new TextBox();
            tbSimilarityThresSSIM.Text = "";
            tbSimilarityThresSSIM.Size = new Size(50, 30);
            tbSimilarityThresSSIM.TextChanged += new EventHandler(tbSimilarityThresSSIM_TextChanged);
            tbSimilarityThresSSIM.Location = new Point(lblSimilarityThresSSIM.Location.X + lblSimilarityThresSSIM.Width + 2, lblSimilarityThresSSIM.Location.Y);
            pnlSimilarityBottom.Controls.Add(tbSimilarityThresSSIM);

            Label lblSimilarityPeriod = new Label();
            lblSimilarityPeriod.Text = "Period (sec) : ";
            lblSimilarityPeriod.Location = new Point(lblSimilarityThresPSNR.Location.X, lblSimilarityThresPSNR.Location.Y + lblSimilarityThresPSNR.Height + 5);
            pnlSimilarityBottom.Controls.Add(lblSimilarityPeriod);

            tbSimilarityPeriod = new TextBox();
            tbSimilarityPeriod.Text = "";
            tbSimilarityPeriod.Size = new Size(50, 30);
            tbSimilarityPeriod.TextChanged += new EventHandler(tbSimilarityPeriod_TextChanged);
            tbSimilarityPeriod.Location = new Point(lblSimilarityPeriod.Location.X + lblSimilarityPeriod.Width + 2, lblSimilarityPeriod.Location.Y);
            pnlSimilarityBottom.Controls.Add(tbSimilarityPeriod);

            Label lblSimilarityInterval = new Label();
            lblSimilarityInterval.Text = "Interval (sec) : ";
            lblSimilarityInterval.Location = new Point(tbSimilarityPeriod.Location.X + tbSimilarityPeriod.Width + 10, tbSimilarityPeriod.Location.Y);
            pnlSimilarityBottom.Controls.Add(lblSimilarityInterval);

            tbSimilarityInterval = new TextBox();
            tbSimilarityInterval.Text = "";
            tbSimilarityInterval.Size = new Size(50, 30);
            tbSimilarityInterval.TextChanged += new EventHandler(tbSimilarityInterval_TextChanged);
            tbSimilarityInterval.Location = new Point(lblSimilarityInterval.Location.X + lblSimilarityInterval.Width + 2, lblSimilarityInterval.Location.Y);
            pnlSimilarityBottom.Controls.Add(tbSimilarityInterval);

            tbSimilarityResult = new TextBox();
            tbSimilarityResult.Text = "";
            tbSimilarityResult.Size = new Size(320, 80);
            tbSimilarityResult.ScrollBars = ScrollBars.Vertical;
            tbSimilarityResult.WordWrap = true;
            tbSimilarityResult.Multiline = true;
            tbSimilarityResult.Location = new Point(tbSimilarityThresSSIM.Location.X + tbSimilarityThresSSIM.Width + 10, tbSimilarityThresSSIM.Location.Y);
            pnlSimilarityBottom.Controls.Add(tbSimilarityResult);

            // Refresh
            Button btSimilarityRefresh = new Button();
            btSimilarityRefresh.Text = "Refresh";
            btSimilarityRefresh.Click += new EventHandler(btSimilarityRefresh_Click);
            btSimilarityRefresh.Location = new Point(pnlSimilarityCur.Location.X + pnlSimilarityCur.Width - btSimilarityRefresh.Width, lblSimilarityCheck.Location.Y);
            fmSimilarityCheck.Controls.Add(btSimilarityRefresh);

            init_similarity();

            fmSimilarityCheck.Show();

            return true;
        }

        private static void fmSimilarityCheck_Close(object sender, FormClosingEventArgs e)
        {
            if (fmSimilarityCheck != null && fmSimilarityCheck.IsDisposed)
                return;

            fmSimilarityCheck.Dispose();
        }

        private static void btSimilarityRefresh_Click(object sender, EventArgs eventArgs)
        {
            if (similarityCheck())
                tbSimilarityResult.AppendText("\r\nResult OK !!!");
            else
                tbSimilarityResult.AppendText("\r\nResult NG !!!");

            MessageBox.Show("Done !!!");
        }

        private static void tbSimilarityThresPSNR_TextChanged(object sender, EventArgs e)
        {
            if (tbSimilarityThresPSNR.Text != "" && int.Parse(tbSimilarityThresPSNR.Text) > -1)
                thresPSNR = int.Parse(tbSimilarityThresPSNR.Text);
        }

        private static void tbSimilarityThresSSIM_TextChanged(object sender, EventArgs e)
        {
            if (tbSimilarityThresSSIM.Text != "" && int.Parse(tbSimilarityThresSSIM.Text) > -1)
                thresSSIM = int.Parse(tbSimilarityThresSSIM.Text);
        }

        private static void tbSimilarityPeriod_TextChanged(object sender, EventArgs e)
        {
            if (tbSimilarityPeriod.Text != "" && int.Parse(tbSimilarityPeriod.Text) > 0)
                period = int.Parse(tbSimilarityPeriod.Text) * 1000;
        }

        private static void tbSimilarityInterval_TextChanged(object sender, EventArgs e)
        {
            if (tbSimilarityInterval.Text != "" && float.Parse(tbSimilarityInterval.Text) > 0)
                interval = (int)(float.Parse(tbSimilarityInterval.Text) * 1000);
        }

        private static void init_similarity()
        {
            tbSimilarityThresPSNR.Text = thresPSNR.ToString();
            tbSimilarityThresSSIM.Text = thresSSIM.ToString();
            tbSimilarityPeriod.Text = (period/1000).ToString();
            tbSimilarityInterval.Text = (interval/1000).ToString();
        }

        public static void drawWatermark(ref Bitmap bmp, string str, float fontSize = 18, Size position = default(Size))
        {
            //this is to handle error (graphic cannot be created from indexed image)
            var newBitmap = new Bitmap(bmp);
            newBitmap.Tag = bmp.Tag;
            bmp.Dispose();
            bmp = newBitmap;

            using (Graphics g = Graphics.FromImage(bmp))
            {
                string text1 = str;
                using (Font font1 = new Font("Arial", fontSize, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    StringFormat stringFormat = new StringFormat();
                    stringFormat.Alignment = StringAlignment.Near;
                    stringFormat.LineAlignment = StringAlignment.Near;

                    var drawBox = new Rectangle(position.Width, position.Height, bmp.Width, bmp.Height);

                    g.DrawString(text1, font1, new SolidBrush(Color.FromArgb(200, Color.Red)), drawBox, stringFormat);
                }
            }
        }

        private static Boolean similarityCheckFromParent()
        {
            Mat processImgRef = new Mat();
            Mat processImg = new Mat();
            double psnrV;
            MCvScalar mssimV;
            Int32 periodCnt = 0;
            bool ret = true;

            if (period > 0 && interval > 0)
                periodCnt = period / interval;

            Saal.videoDisplay_cv.Retrieve(processImgRef, 0); // Reference            

            for (int i = 0; i < periodCnt; i++)
            {
                Saal.videoDisplay_cv.Retrieve(processImg, 0);

                psnrV = getPSNR(processImgRef, processImg);

                if (psnrV < thresPSNR)
                {
                    mssimV = getMSSIM(processImgRef, processImg);
                    if (((double)(mssimV.V2 * 100) < (double)thresSSIM) && ((double)(mssimV.V1 * 100) < (double)thresSSIM) && ((double)(mssimV.V0 * 100) < (double)thresSSIM))
                    {
                        ret = false;
                        break;
                    }
                }
                Thread.Sleep(interval);
            }

            processImgRef.Dispose();
            processImg.Dispose();

            return ret;
        }

        private static Boolean similarity_wBlurryCheck(Bitmap referenceBitmap, ref Bitmap probBitmap, 
            ref object errorInfo, ref object processInfo, ref bool stop)
        {
            double psnrV = 0.0;
            MCvScalar mssimV = new MCvScalar();
            Int32 periodCnt = 0;
            bool retSimilarTest = true;
            bool retBlurTest = true;

            if (period > 0 && interval > 0)
                periodCnt = period / interval;
            
            var processImgRef = new Image<Bgr, Byte>(referenceBitmap).Mat;
            
            for (int i = 0; i < periodCnt; i++)
            {
                if (stop)
                    break;

                Mat processImg = new Mat();
                Saal.videoDisplay_cv.Retrieve(processImg, 0);
                var captureTime = "Time taken: " + DateTime.Now;

                // begin similarityTest
                var imageSimilarityTest = processImg.Clone();
                psnrV = getPSNR(processImgRef, imageSimilarityTest);
                if (psnrV < thresPSNR)
                {
                    mssimV = getMSSIM(processImgRef, imageSimilarityTest);


                    bool similarityTest = ((double)(mssimV.V2 * 100) < (double)thresSSIM) &&
                                          ((double)(mssimV.V1 * 100) < (double)thresSSIM) &&
                                          ((double)(mssimV.V0 * 100) < (double)thresSSIM);

                    if (similarityTest)
                        retSimilarTest = false;
                }

                imageSimilarityTest.Dispose();
                // end similarityTest

                if (!retSimilarTest)
                {
                    probBitmap = processImg.Clone().Bitmap;
                    probBitmap.Tag = captureTime;
                    processImg.Dispose();
                    processInfo = new[] { (int)psnrV, (int)(mssimV.V0 * 100), (int)(mssimV.V1 * 100), (int)(mssimV.V2 * 100), 0 };
                    break;
                }

                // begin blurTest
                // action to delete old resource before assign new one
                // ref cannot be used here :(
                Func<Mat, Mat, Mat> replaceResources = (oldmat, newmat) =>
                {
                    oldmat.Dispose();
                    return newmat;
                };

                var imageBlurTest = processImg.Clone();

                // Remove noise by blurring with Gaussian Filter
                if (isGaussianBlur)
                {
                    imageBlurTest = replaceResources(imageBlurTest, GaussianBlurHandler(imageBlurTest));
                }

                // Change color image to gray
                imageBlurTest = replaceResources(imageBlurTest, GrayScaleHandler(imageBlurTest));

                // Implement Laplace algorithm
                imageBlurTest = replaceResources(imageBlurTest, LaplacianHandler(imageBlurTest));

                List<double> data = new List<double>();

                Image<Gray, byte> img = imageBlurTest.ToImage<Gray, Byte>();
                for (int j = 0; j < img.Rows; j++)
                {
                    for (int k = 0; k < img.Cols; k++)
                    {
                        data.Add(img.Data[j, k, 0]);
                    }
                }

                retBlurTest = data.Variance() > thresBlurry;

                imageBlurTest.Dispose();
                // end blurTest

                if (!retBlurTest)
                {
                    probBitmap = processImg.Clone().Bitmap;
                    probBitmap.Tag = captureTime;
                    processImg.Dispose();
                    processInfo = new[] { (int)psnrV, (int)(mssimV.V0 * 100), (int)(mssimV.V1 * 100), (int)(mssimV.V2 * 100), (int)data.Variance() };
                    break;
                }

                processImg.Dispose();

                Thread.Sleep(interval);
            }

            processImgRef.Dispose();

            errorInfo = new [] { retSimilarTest, retBlurTest };
            return (retSimilarTest && retBlurTest);
        }

        private static Boolean similarityCheck()
        {
            Mat processImgRef = new Mat();
            Mat processImg = new Mat();
            double psnrV;
            MCvScalar mssimV;
            Int32 periodCnt = 0;
            bool ret = true;

            if (period > 0 && interval > 0)
                periodCnt = period / interval;

            Saal.videoDisplay_cv.Retrieve(processImgRef, 0); // Reference
            tbSimilarityResult.Text = "";
            tbSimilarityResult.Refresh();
            pbSimilarityRef.Image = processImgRef;
            pbSimilarityRef.Refresh();

            for (int i = 0; i < periodCnt; i++)
            {
                Saal.videoDisplay_cv.Retrieve(processImg, 0);

                psnrV = getPSNR(processImgRef, processImg);

                if (psnrV < thresPSNR)
                {
                    mssimV = getMSSIM(processImgRef, processImg);
                    tbSimilarityResult.AppendText("Frame " + (i + 1).ToString() + "# - PSNR: " + Math.Round(psnrV, 2) + "dB - MSSIM: R " + Math.Round((double)(mssimV.V2 * 100), 2) + "% G " + Math.Round((double)(mssimV.V1 * 100), 2) + "% B " + Math.Round((double)(mssimV.V0 * 100), 2) + "%\r\n");
                    tbSimilarityResult.Refresh();
                    if (((double)(mssimV.V2 * 100) < (double)thresSSIM) && ((double)(mssimV.V1 * 100) < (double)thresSSIM) && ((double)(mssimV.V0 * 100) < (double)thresSSIM))
                    {
                        pbSimilarityCur.Image = processImg;
                        pbSimilarityCur.Refresh();
                        ret = false;
                        break;
                    }
                }
                else
                {
                    tbSimilarityResult.AppendText("Frame " + (i + 1).ToString() + "# - PSNR: " + Math.Round(psnrV, 2) + "dB\r\n");
                    tbSimilarityResult.Refresh();
                }
                Thread.Sleep(interval);
            }

            processImgRef.Dispose();
            processImg.Dispose();

            return ret;
        }

        private static double getPSNR(Mat inpA, Mat inpB)
        {
            double ret = 0;
            Mat res = new Mat();
            CvInvoke.AbsDiff(inpA, inpB, res); // |inpA - inpB|
            res.ConvertTo(res, DepthType.Cv32F); // Convert from 8bit to 32bit
            CvInvoke.Multiply(res, res, res); // |inpA - inpB|^2

            MCvScalar s = CvInvoke.Sum(res); // sum elements per channel

            double sse = s.V0 + s.V1 + s.V2; // sum channels

            if (sse > 1e-10)
            {
                double mse = sse / (double)(inpA.NumberOfChannels * (int)inpA.Total);
                double psnr = 10.0 * Math.Log10(255 * 255 / mse);
                ret = psnr;
            }

            res.Dispose();
            return ret;
        }

        private static MCvScalar getMSSIM(Mat inpA, Mat inpB)
        {
            const double C1 = 6.5025, C2 = 58.5225;

            /***************************** INITS **********************************/
            Mat _inpA = new Mat();
            Mat _inpB = new Mat();

            inpA.ConvertTo(_inpA, DepthType.Cv32F);         // cannot calculate on one byte large values
            inpB.ConvertTo(_inpB, DepthType.Cv32F);

            Mat _inpA_2 = new Mat();
            Mat _inpB_2 = new Mat();
            Mat _inpA_inpB = new Mat();

            CvInvoke.Multiply(_inpA, _inpA, _inpA_2); // inpA ^ 2
            CvInvoke.Multiply(_inpB, _inpB, _inpB_2); // inpB ^ 2
            CvInvoke.Multiply(_inpA, _inpB, _inpA_inpB); // inpA * inpB

            /***********************PRELIMINARY COMPUTING ******************************/
            Mat mu1 = new Mat();
            Mat mu2 = new Mat();

            CvInvoke.GaussianBlur(_inpA, mu1, new Size(11, 11), 1.5);
            CvInvoke.GaussianBlur(_inpB, mu2, new Size(11, 11), 1.5);

            Mat mu1_2 = new Mat();
            Mat mu2_2 = new Mat();
            Mat mu1_mu2 = new Mat();

            CvInvoke.Multiply(mu1, mu1, mu1_2); // mu1 ^ 2
            CvInvoke.Multiply(mu2, mu2, mu2_2); // mu2 ^ 2
            CvInvoke.Multiply(mu1, mu2, mu1_mu2); // mu1 * mu2

            Mat sigma1_2 = new Mat();
            Mat sigma2_2 = new Mat();
            Mat sigma12 = new Mat();

            CvInvoke.GaussianBlur(_inpA_2, sigma1_2, new Size(11, 11), 1.5);
            CvInvoke.GaussianBlur(_inpB_2, sigma2_2, new Size(11, 11), 1.5);
            CvInvoke.GaussianBlur(_inpA_inpB, sigma12, new Size(11, 11), 1.5);

            CvInvoke.Subtract(sigma1_2, mu1_2, sigma1_2);
            CvInvoke.Subtract(sigma2_2, mu2_2, sigma2_2);
            CvInvoke.Subtract(sigma12, mu1_mu2, sigma12);

            ///////////////////////////////// FORMULA ////////////////////////////////
            Mat t1 = new Mat();
            Mat t2 = new Mat();
            Mat t3 = new Mat();

            CvInvoke.AddWeighted(mu1_mu2, 1, mu1_mu2, 1, C1, t1);
            CvInvoke.AddWeighted(sigma12, 1, sigma12, 1, C2, t2);
            CvInvoke.Multiply(t1, t2, t3);   // t3 = ((2*mu1_mu2 + C1).*(2*sigma12 + C2))

            CvInvoke.AddWeighted(mu1_2, 1, mu2_2, 1, C1, t1);
            CvInvoke.AddWeighted(sigma1_2, 1, sigma2_2, 1, C2, t2);
            CvInvoke.Multiply(t1, t2, t1);  // t1 =((mu1_2 + mu2_2 + C1).*(sigma1_2 + sigma2_2 + C2))

            Mat ssim_map = new Mat();
            CvInvoke.Divide(t3, t1, ssim_map);      // ssim_map =  t3./t1;

            MCvScalar mssim = CvInvoke.Mean(ssim_map); // mssim = average of ssim map

            _inpA.Dispose();
            _inpB.Dispose();
            _inpA_2.Dispose();
            _inpB_2.Dispose();
            _inpA_inpB.Dispose();
            mu1.Dispose();
            mu2.Dispose();
            mu1_2.Dispose();
            mu2_2.Dispose();
            mu1_mu2.Dispose();
            sigma1_2.Dispose();
            sigma2_2.Dispose();
            sigma12.Dispose();
            t1.Dispose();
            t2.Dispose();
            t3.Dispose();
            ssim_map.Dispose();

            return mssim;
        }

        #endregion

        #region Picture Format
        private static Boolean fmPictFormat_Open()
        {
            if (fmPictFormat != null && !fmPictFormat.IsDisposed)
                return false;

            fmPictFormat = new Form();
            fmPictFormat.Size = new System.Drawing.Size(ResSize.Width + 200, ResSize.Height);
            fmPictFormat.MaximizeBox = false;
            fmPictFormat.ControlBox = true;
            fmPictFormat.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmPictFormat.FormClosing += new FormClosingEventHandler(fmPictFormat_Close);
            fmPictFormat.StartPosition = FormStartPosition.CenterScreen;

            fmPictFormat.Show();

            return true;
        }

        private static void fmPictFormat_Close(object sender, FormClosingEventArgs e)
        {
            if (fmPictFormat != null && fmPictFormat.IsDisposed)
                return;

            e.Cancel = true;
            fmPictFormat.Dispose();
        }
        #endregion

        #region Operation Dialog
        private static Boolean fmOperationDialog_Open()
        {
            if (fmOperationDialog != null && !fmOperationDialog.IsDisposed)
                return false;

            fmOperationDialog = new Form();
            fmOperationDialog.Size = new System.Drawing.Size(690, 560);
            fmOperationDialog.MaximizeBox = false;
            fmOperationDialog.ControlBox = true;
            fmOperationDialog.FormBorderStyle = FormBorderStyle.FixedSingle;
            fmOperationDialog.FormClosing += new FormClosingEventHandler(fmOperationDialog_Close);
            fmOperationDialog.BackColor = ColorTranslator.FromHtml("#404040");
            fmOperationDialog.StartPosition = FormStartPosition.CenterScreen;

            Label lblOperationDialog = new Label();
            lblOperationDialog.Text = "Running Operation ...";
            lblOperationDialog.ForeColor = Color.White;
            lblOperationDialog.Font = new Font(FontFamily.GenericSansSerif, 20, FontStyle.Bold);
            lblOperationDialog.Size = new Size(600, 40);
            fmOperationDialog.Controls.Add(lblOperationDialog);

            Label lblImgRef = new Label();
            lblImgRef.Text = "Processing Image";
            //lblImgRef.ForeColor = Color.White;
            lblImgRef.Font = new Font(FontFamily.GenericSansSerif, 16, FontStyle.Bold);
            lblImgRef.BackColor = ColorTranslator.FromHtml("#808080");
            lblImgRef.Size = new Size(320, 30);
            lblImgRef.TextAlign = ContentAlignment.MiddleCenter;
            fmOperationDialog.Controls.Add(lblImgRef);

            pbImgRef = new PictureBox();
            pbImgRef.Size = new System.Drawing.Size(320, 240);
            pbImgRef.SizeMode = PictureBoxSizeMode.StretchImage;
            pbImgRef.BackColor = Color.LightBlue;
            pbImgRef.Paint += new PaintEventHandler(pbImgRef_Paint);
            fmOperationDialog.Controls.Add(pbImgRef);

            Label lblImgCom = new Label();
            lblImgCom.Text = "NG Image";
            //lblImgCom.ForeColor = Color.White;
            lblImgCom.Font = new Font(lblImgCom.Font, FontStyle.Bold);
            lblImgCom.Font = new Font(FontFamily.GenericSansSerif, 16, FontStyle.Bold);
            lblImgCom.BackColor = ColorTranslator.FromHtml("#808080");
            lblImgCom.TextAlign = ContentAlignment.MiddleCenter;
            lblImgCom.Size = new Size(320, 30);
            fmOperationDialog.Controls.Add(lblImgCom);

            pbImgCom = new PictureBox();
            pbImgCom.Size = new System.Drawing.Size(320, 240);
            pbImgCom.SizeMode = PictureBoxSizeMode.StretchImage;
            pbImgCom.BackColor = Color.LightBlue;            
            fmOperationDialog.Controls.Add(pbImgCom);
            
            tbOperationLog = new RichTextBox();
            tbOperationLog.Text = "";
            tbOperationLog.Size = new Size(645, 150);
            tbOperationLog.Multiline = true;
            tbOperationLog.ScrollBars = RichTextBoxScrollBars.Vertical;
            tbOperationLog.ReadOnly = true;
            tbOperationLog.BackColor = Color.Black;
            tbOperationLog.ForeColor = Color.WhiteSmoke;
            fmOperationDialog.Controls.Add(tbOperationLog);

            pbarOpDialog = new ProgressBar();
            pbarOpDialog.Size = new Size(100, 20);
            pbarOpDialog.Style = ProgressBarStyle.Marquee;
            fmOperationDialog.Controls.Add(pbarOpDialog);
            
            lblOpDialog = new Label();
            lblOpDialog.Text = "Current Status : 0/0";
            lblOpDialog.ForeColor = Color.White;
            lblOpDialog.TextAlign = ContentAlignment.MiddleRight;
            lblOpDialog.Size = new Size(200, 20);
            fmOperationDialog.Controls.Add(lblOpDialog);
            
            lblOperationDialog.Location = new Point(10, 10);            
            pbImgRef.Location = new Point(lblOperationDialog.Location.X, lblOperationDialog.Location.Y + lblOperationDialog.Height + 40);
            lblImgRef.Location = new Point(pbImgRef.Location.X, pbImgRef.Location.Y - lblImgRef.Height);            
            pbImgCom.Location = new Point(pbImgRef.Location.X + pbImgRef.Width + 5, pbImgRef.Location.Y);
            lblImgCom.Location = new Point(pbImgCom.Location.X, pbImgCom.Location.Y - lblImgCom.Height);
            tbOperationLog.Location = new Point(pbImgRef.Location.X, pbImgRef.Location.Y + pbImgRef.Height + 10);
            pbarOpDialog.Location = new Point(tbOperationLog.Location.X + tbOperationLog.Width - pbarOpDialog.Width, tbOperationLog.Location.Y + tbOperationLog.Height + 1);
            lblOpDialog.Location = new Point(pbarOpDialog.Location.X - lblOpDialog.Width - 1, pbarOpDialog.Location.Y + 3);

            addScanBackground();

            scan_timer = new System.Timers.Timer();
            scan_timer.Elapsed += (sender, e) => TimeTick(sender, e);
            scan_timer.AutoReset = true;
            scan_timer.Interval = 100;                        
            return true;
        }

        private static void fmOperationDialog_Close(object sender, FormClosingEventArgs e)
        {
            if (!fmOperationDialog.Visible)
                return;

            e.Cancel = true;

            scan_timer.Close();
            scan_timer.Enabled = false;
            fmOperationDialog.Hide();
        }

        private static void TimeTick(Object sender, ElapsedEventArgs e)
        {
            if (fmOperationDialog.Visible)
            {
                if (scan_timer.Enabled == false)
                    return;

                rec_scan_img.Y += 10;
                if (rec_scan_img.Y > pbImgCom.Height - 1)
                    rec_scan_img.Y = -90;

                pbImgRef.Invalidate();
            }
        }

        public static void stopScan()
        {
            if(scan_timer.Enabled)
            {
                scan_timer.Close();
                scan_timer.Enabled = false;
                rec_scan_img.Y = pbImgCom.Height;
            }            
        }

        public static void startScan()
        {
            scan_timer.Start();
            scan_timer.Enabled = true;            
        }

        public static void OperationDialog_Open()
        {
            if (fmOperationDialog.Visible)
                return;
            
            updOpFlag = true;
            fmOperationDialog.Show();            
        }

        public static void OperationDialog_updRefImg(Bitmap img)
        {
            if (fmOperationDialog.Visible)
            {
                updOpFlag = false;
                Action actionOp = () =>
                {
                    pbImgRef.Image = img;
                    pbImgRef.Refresh();
                };
                if (pbImgRef.InvokeRequired)
                    pbImgRef.BeginInvoke(actionOp);
                else
                    actionOp();                
            }
        }

        public static void OperationDialog_updRefCom(Bitmap img)
        {
            if (fmOperationDialog.Visible)
            {
                updOpFlag = false;
                Action actionOp = () =>
                {
                    pbImgCom.Image = img;
                    pbImgCom.Refresh();
                };

                if (pbImgCom.InvokeRequired)
                    pbImgCom.BeginInvoke(actionOp);
                else
                    actionOp();                
            }
        }

        public static void OperationDialog_updLog(string str, Color clr)
        {
            if (fmOperationDialog.Visible)
            {
                Action actionOp = () =>
                {
                    //tbOperationLog.ForeColor = clr;
                    tbOperationLog.SelectionColor = clr;
                    tbOperationLog.AppendText("\r\n" + str);
                    tbOperationLog.SelectionColor = Color.White;
                };

                if (tbOperationLog.InvokeRequired)
                    tbOperationLog.BeginInvoke(actionOp);
                else
                    actionOp();
            }
        }

        public static void OperationDialog_updPbar(int val, string str)
        {
            if (fmOperationDialog.Visible)
            {
                Action actionOp = () =>
                {
                    if (val == 100)
                    {
                        pbarOpDialog.Hide();
                    }
                    else
                    {
                        pbarOpDialog.Show();
                    }                       
                };
                
                if (pbarOpDialog.InvokeRequired)
                    pbarOpDialog.BeginInvoke(actionOp);
                else
                    actionOp();

                Action actionOplbl = () =>
                {
                    if (val == 100)
                    {
                        lblOpDialog.Hide();
                    }
                    else
                    {
                        lblOpDialog.Text = str;
                        lblOpDialog.Refresh();
                        lblOpDialog.Show();
                    }
                    
                };

                if (lblOpDialog.InvokeRequired)
                    lblOpDialog.BeginInvoke(actionOplbl);
                else
                    actionOplbl();
                    
            }            
        }
        #endregion

        #region Video Handler
        private static Boolean videoInit()
        {
            frameCapture = new Mat();

            videoStop();
            Thread.Sleep(500);
            Saal.CAM_Init_cv(camName);
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.FrameWidth, ResSize.Width);
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.FrameHeight, ResSize.Height);
            Saal.videoDisplay_cv.ImageGrabbed += videoCapture;

            return true;
        }

        private static Boolean videoReInit()
        {
            if (frameCapture != null)
                frameCapture.Dispose();
            frameCapture = new Mat();

            videoStop();
            Thread.Sleep(500);
            Saal.CAM_Init_cv(camName);
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.FrameWidth, ResSize.Width);
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.FrameHeight, ResSize.Height);
            Saal.videoDisplay_cv.ImageGrabbed += videoCapture;

            videoStart();
            return true;
        }

        private static Boolean videoStart()
        {
            if (Saal.videoDisplay_cv != null)
            {
                Saal.videoDisplay_cv.Start();
            }
            return true;
        }

        private static Boolean videoStop()
        {
            if (Saal.videoDisplay_cv != null)
            {
                Saal.videoDisplay_cv.Stop();
            }
            return true;
        }

        private static void SetBrightness(int bright)
        {
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.Brightness, bright);
        }

        private static void SetContrast(int contrast)
        {
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.Contrast, contrast);
        }

        private static void SetSharpness(int sharpness)
        {
            Saal.videoDisplay_cv.SetCaptureProperty(CapProp.Sharpness, sharpness);
        }

        private static void videoCapture(object sender, EventArgs eventArgs)
        {
            Saal.videoDisplay_cv.Retrieve(frameCapture, 0);

            if (frameCapture.IsEmpty)
                return;

            try
            {
                if (capture_flag)
                {
                    if (ImgFromParent != null)
                        imgFromParent.Image = null;

                    Action action = () =>
                        pbGridCalib.Image = (Bitmap)frameCapture.Bitmap.Clone();

                    if (pbGridCalib.InvokeRequired)
                        pbGridCalib.BeginInvoke(action);
                    else
                        action();
                }
                else
                {

                    if (ImgFromParent == null)
                        return;

                    Action action = () =>
                        imgFromParent.Image = (Bitmap)frameCapture.Bitmap.Clone();

                    if (imgFromParent.InvokeRequired)
                        imgFromParent.BeginInvoke(action);
                    else
                        action();

                    if (fmOperationDialog.Visible && updOpFlag)
                    {
                        Action actionOp = () =>
                        pbImgRef.Image = (Bitmap)frameCapture.Bitmap.Clone();

                        if (pbImgRef.InvokeRequired)
                            pbImgRef.BeginInvoke(actionOp);
                        else
                            actionOp();
                    }
                }
            }
            catch
            {

            }

        }

        private static void addScanBackground()
        {
            var image = Properties.Resources.scan_image;
            scan_img = new Bitmap((Bitmap)image);
            scan_img.MakeTransparent();

            rec_scan_img = new Rectangle(new Point(0, -90), new Size(320, 100));
        }
        
        private static void pbImgRef_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                DrawImgRef(e.Graphics, 0, 0);
            }
            catch (Exception exp)
            {
                System.Console.WriteLine(exp.Message);
            }
        }
        
        private static void DrawImgRef(Graphics g, int X, int Y)
        {
            if (fmOperationDialog.Visible)
            {                                
                g.DrawImage(scan_img, rec_scan_img);
            }
        }

        #endregion

        #region Check black image
        private static Mat GlobalBinarizationHandler(Mat processImgBefore, double threshold = 0, double maxValue = 255)
        {
            Mat processImgAfter = new Mat();
            CvInvoke.Threshold(processImgBefore, processImgAfter, threshold, maxValue, ThresholdType.Binary);
            return processImgAfter;
        }

        private static Mat AdaptiveBinarizationHandler(Mat processImgBefore, double maxValue = 255, int blockSize = 11, double param1 = 2)
        {
            Mat processImgAfter = new Mat();
            CvInvoke.AdaptiveThreshold(processImgBefore, processImgAfter, maxValue, AdaptiveThresholdType.GaussianC,
                ThresholdType.Binary, blockSize, param1);
            return processImgAfter;
        }

        public static Bitmap GetBinarialImage_Adaptive(Bitmap bmp, double maxValue = 255, int blockSize = 11, double param1 = 2)
        {
            using (var processImg = new Image<Bgr, Byte>(bmp))
            {
                Saal.videoDisplay_cv.Retrieve(processImg, 0);
                var grey = GrayScaleHandler(processImg.Mat);
                var binary = AdaptiveBinarizationHandler(grey, maxValue, blockSize, param1);

                var bitmap = processImg.Bitmap;
                grey.Dispose();
                binary.Dispose();
                return bitmap;
            }

        }

        public static Bitmap GetBinarialImage_Global(Bitmap bmp, double threshold = 0, double maxValue = 255)
        {
            using (var processImg = new Image<Bgr, Byte>(bmp))
            {
                //Saal.videoDisplay_cv.Retrieve(processImg, 0);

                var grey = GrayScaleHandler(processImg.Mat);
                var binary = GlobalBinarizationHandler(grey, threshold, maxValue);

                var bitmap = binary.Bitmap;
                grey.Dispose();
                binary.Dispose();
                return bitmap;
            }
        }

        public static Bitmap GetBinarialImage_GlobalCrop(Bitmap bmp, double threshold = 0, double maxValue = 255)
        {
            using (var processImg = new Image<Bgr, Byte>(bmp))
            {
                //Saal.videoDisplay_cv.Retrieve(processImg, 0);

                var grey = GrayScaleHandler(CropImage(EN_OCR_ITEM.OCR_ITEM_FULL, processImg));
                var binary = GlobalBinarizationHandler(grey, threshold, maxValue);

                var bitmap = binary.Bitmap;
                grey.Dispose();
                binary.Dispose();
                return bitmap;
            }
        }

        public static double GetBlackPixelPercentage(Bitmap bmp)
        {
            using (var ImageBgr = new Image<Bgr, Byte>(bmp))
            {
                Mat ff = new Mat();

                byte[,,] data = ImageBgr.Data;
                int count = 0;
                for (int i = ImageBgr.Rows - 1; i >= 0; i--)
                {
                    for (int j = ImageBgr.Cols - 1; j >= 0; j--)
                    {
                        if (data[i, j, 0] == 0 && data[i, j, 1] == 0 && data[i, j, 2] == 0)
                            count++;
                    }
                }

                return (double)count / (ImageBgr.Rows * ImageBgr.Cols) * 100;
            }
        }

        public static Bitmap GetBinarialImage_InRange(Bitmap src, int v1, int v2, int v3, int v4)
        {
            using (var processImg = new Image<Bgr, Byte>(src))
            {
                Saal.videoDisplay_cv.Retrieve(processImg, 0);

                using (var ImageHsV = processImg.Convert<Hsv, Byte>())
                {
                    //extract the hue and value channels
                    Image<Gray, Byte>[] channels = ImageHsV.Split();  //split into components
                    Image<Gray, Byte> imghue = channels[0];            //hsv, so channels[0] is hue.
                    Image<Gray, Byte> imgval = channels[2];            //hsv, so channels[2] is value.

                    //filter out all but "the color you want"...seems to be 0 to 128 ?
                    Image<Gray, byte> huefilter = imghue.InRange(new Gray(v1),
                        new Gray(v2));

                    //use the value channel to filter out all but brighter colors
                    Image<Gray, byte> valfilter = imgval.InRange(new Gray(v3), new Gray(v4));

                    //now and the two to get the parts of the imaged that are colored and above some brightness.
                    Image<Gray, byte> colordetimg = huefilter.And(valfilter);

                    return colordetimg.Bitmap;
                }
            }
        }
        #endregion

        #region Color Detection
        /*
        private static void ColorDetection()
        {
            Mat processImg = new Mat();
            Saal.videoDisplay_cv.Retrieve(processImg, 0);

            processImg = CropImage(EN_OCR_ITEM.OCR_ITEM_FULL, processImg.ToImage<Bgr, Byte>());

            Image<Gray, Byte> ImageFrameDetection = cvAndHsvImage(processImg.ToImage<Bgr, byte>);

            if (iB2C == 0) imageBox2.Image = ImageFrameDetection;

            if (iB2C == 1)
            {
                Image<Bgr, Byte> imgF = new Image<Bgr, Byte>(ImageFrame.Width, ImageFrame.Height);
                Image<Bgr, Byte> imgD = ImageFrameDetection.Convert<Bgr, Byte>();
                CvInvoke.cvAnd(ImageFrame, imgD, imgF, IntPtr.Zero);
                imageBox2.Image = imgF;
            }

            if (iB2C == 2)
            {
                Image<Bgr, Byte> imgF = new Image<Bgr, Byte>(ImageFrame.Width, ImageFrame.Height);
                Image<Bgr, Byte> imgD = ImageFrameDetection.Convert<Bgr, Byte>();
                CvInvoke.cvAnd(ImageFrame, imgD, imgF, IntPtr.Zero);
                for (int x = 0; x < imgF.Width; x++)
                    for (int y = 0; y < imgF.Height; y++)
                    {
                        {
                            Bgr c = imgF[y, x];
                            if (c.Red == 0 && c.Blue == 0 && c.Green == 0)
                            {
                                imgF[y, x] = new Bgr(255, 255, 255);
                            }
                        }
                    }

                imageBox2.Image = imgF;
            }


            if (checkBox_VAr.Checked) RecDetection(ImageFrameDetection, ImageFrame, trackBar_VAr.Value);

            Image<Gray, Byte> ImageHSVwheelDetection = cvAndHsvImage(
               ImageHSVwheel,
               Convert.ToInt32(numeric_HL.Value), Convert.ToInt32(numeric_HH.Value),
               Convert.ToInt32(numeric_SL.Value), Convert.ToInt32(numeric_SH.Value),
               Convert.ToInt32(numeric_VL.Value), Convert.ToInt32(numeric_VH.Value),
               checkBox_EH.Checked, checkBox_ES.Checked, checkBox_EV.Checked, checkBox_IV.Checked);
            imageBox4.Image = ImageHSVwheelDetection;
        }

        private static void RecDetection(Image<Gray, Byte> img, Image<Bgr, Byte> showRecOnImg, int areaV)
        {
            Image<Gray, Byte> imgForContour = new Image<Gray, byte>(img.Width, img.Height);
            CvInvoke.cvCopy(img, imgForContour, System.IntPtr.Zero);


            IntPtr storage = CvInvoke.cvCreateMemStorage(0);
            IntPtr contour = new IntPtr();

            CvInvoke.cvFindContours(
                imgForContour,
                storage,
                ref contour,
                System.Runtime.InteropServices.Marshal.SizeOf(typeof(MCvContour)),
                Emgu.CV.CvEnum.RETR_TYPE.CV_RETR_EXTERNAL,
                Emgu.CV.CvEnum.CHAIN_APPROX_METHOD.CV_CHAIN_APPROX_NONE,
                new Point(0, 0));


            Seq<Point> seq = new Seq<Point>(contour, null);

            for (; seq != null && seq.Ptr.ToInt64() != 0; seq = seq.HNext)
            {
                Rectangle bndRec = CvInvoke.cvBoundingRect(seq, 2);
                double areaC = CvInvoke.cvContourArea(seq, MCvSlice.WholeSeq, 1) * -1;
                if (areaC > areaV)
                {
                    CvInvoke.cvRectangle(showRecOnImg, new Point(bndRec.X, bndRec.Y),
                        new Point(bndRec.X + bndRec.Width, bndRec.Y + bndRec.Height),
                        new MCvScalar(0, 0, 255), 2, LINE_TYPE.CV_AA, 0);
                }

            }

        }
        private static Image<Gray, Byte> cvAndHsvImage(Image<Bgr, Byte> imgFame)
        {
            Image<Hsv, Byte> hsvImage = imgFame.Convert<Hsv, Byte>();
            Image<Gray, Byte> ResultImage = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height);
            Image<Gray, Byte> ResultImageH = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height);
            Image<Gray, Byte> ResultImageS = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height);
            Image<Gray, Byte> ResultImageV = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height);

            int L1 = 0;
            int H1 = 255;
            int L2 = 10;
            int H2 = 245;
            int L3 = 20;
            int H3 = 235;

            Image<Gray, Byte> img1 = inRangeImage(hsvImage, L1, H1, 0);
            Image<Gray, Byte> img2 = inRangeImage(hsvImage, L2, H2, 1);
            Image<Gray, Byte> img3 = inRangeImage(hsvImage, L3, H3, 2);
            Image<Gray, Byte> img4 = inRangeImage(hsvImage, 0, L1, 0);
            Image<Gray, Byte> img5 = inRangeImage(hsvImage, H1, 180, 0);

            bool H = true;
            bool I = true;
            bool S = true;
            bool V = true;

            if (H)
            {
                if (I)
                {
                    CvInvoke.BitwiseOr(img4, img5, img4);
                    ResultImageH = img4;
                }
                else { ResultImageH = img1; }
            }

            if (S) ResultImageS = img2;
            if (V) ResultImageV = img3;

            if (H && !S && !V) ResultImage = ResultImageH;
            if (!H && S && !V) ResultImage = ResultImageS;
            if (!H && !S && V) ResultImage = ResultImageV;

            if (H && S && !V)
            {
                CvInvoke.BitwiseAnd(ResultImageH, ResultImageS, ResultImageH);
                ResultImage = ResultImageH;
            }

            if (H && !S && V)
            {
                CvInvoke.BitwiseAnd(ResultImageH, ResultImageV, ResultImageH);
                ResultImage = ResultImageH;
            }

            if (!H && S && V)
            {
                CvInvoke.BitwiseAnd(ResultImageS, ResultImageV, ResultImageS);
                ResultImage = ResultImageS;
            }

            if (H && S && V)
            {
                CvInvoke.BitwiseAnd(ResultImageH, ResultImageS, ResultImageH);
                CvInvoke.BitwiseAnd(ResultImageH, ResultImageV, ResultImageH);
                ResultImage = ResultImageH;
            }

            CvInvoke.Erode(ResultImage, ResultImage, (IntPtr)null, 1);

            return ResultImage;
        }
        private static Image<Gray, Byte> inRangeImage(Image<Hsv, Byte> hsvImage, int Lo, int Hi, int con)
        {
            Image<Gray, Byte> ResultImage = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height);
            Image<Gray, Byte> IlowCh = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height, new Gray(Lo));
            Image<Gray, Byte> IHiCh = new Image<Gray, Byte>(hsvImage.Width, hsvImage.Height, new Gray(Hi));
            CvInvoke.InRange(hsvImage[con], IlowCh, IHiCh, ResultImage);
            return ResultImage;
        }
        private static int[] scaleImage(int wP, int hP)
        {
            int[] dR = new int[2];
            int ra;
            if (wP != 0)
            {
                ra = (100 * 320) / wP;
                wP = 320;
                hP = (hP * ra) / 100;
                if (hP != 0 && hP > 240)
                {
                    ra = (100 * 240) / hP;
                    hP = 240;
                    wP = (wP * ra) / 100;
                }
                dR[0] = wP;
                dR[1] = hP;
            }
            return dR;
        }
        */
        #endregion

        #region XML
        private static void XML_fetch_data()
        {
            string curFile = @XMLfile;
            if (!File.Exists(curFile))
            {
                XML_new_file();
            }

            _item = new Item_Info[ItemSize];

            XDocument xdoc = XDocument.Load(XMLfile);
            Console.WriteLine("Reading xml file...");

            var tempProfiles = from itm in xdoc.Descendants("OCR").Descendants("TV")
                               select itm;

            int i = 0;
            foreach (var profile in tempProfiles)
            {
                _item[i].Name = profile.Element("NAME").Value.ToString();
                _item[i].Source = profile.Element("SOURCE").Value.ToString();
                _item[i].Info = profile.Element("INFO").Value.ToString();
                _item[i].Label = profile.Element("LABEL").Value.ToString();
                _item[i].Opt1 = profile.Element("OPTION1").Value.ToString();
                _item[i].Opt2 = profile.Element("OPTION2").Value.ToString();
                _item[i].Opt3 = profile.Element("OPTION3").Value.ToString();
                _item[i].Full = profile.Element("FULL").Value.ToString();
                i++;
            }

            if (_item[0].Name != null && _item[0].Source != null)
            {
                test_profile = _item[0].Name;
                test_item = _item[0].Source;
            }
        }

        private static void add_data()
        {
            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == cbOCRProfileList.Text)
                {
                    // Modify profile
                    if (cbOCRItemList.Text == "Source")
                    {
                        _item[i].Source = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Label")
                    {
                        _item[i].Label = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Info")
                    {
                        _item[i].Info = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Option1")
                    {
                        _item[i].Opt1 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Option2")
                    {
                        _item[i].Opt2 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Option3")
                    {
                        _item[i].Opt3 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    else if (cbOCRItemList.Text == "Full")
                    {
                        _item[i].Full = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    break;
                }
                else if (_item[i].Name == null || _item[i].Name == "")
                {
                    // Add new profile
                    _item[i].Name = cbOCRProfileList.Text;
                    if (cbOCRItemList.Text == "Source")
                    {
                        _item[i].Source = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Label")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Info")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Option1")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Option2")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Full")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                        _item[i].Full = "10;10;10;10";
                    }
                    else if (cbOCRItemList.Text == "Option3")
                    {
                        _item[i].Source = "10;10;10;10";
                        _item[i].Label = "10;10;10;10";
                        _item[i].Info = "10;10;10;10";
                        _item[i].Opt1 = "10;10;10;10";
                        _item[i].Opt2 = "10;10;10;10";
                        _item[i].Opt3 = "10;10;10;10";
                        _item[i].Full = tbOCR_W.Text + ";" + tbOCR_H.Text + ";" + tbOCR_X.Text + ";" + tbOCR_Y.Text;
                    }
                    break;
                }
            }
            XML_store_data();
        }

        private static void XML_store_data()
        {
            XDocument xmlFile = XDocument.Load(XMLfile);

            xmlFile.Root.Descendants("OCR").Descendants().Remove();
            xmlFile.Save(XMLfile);


            for (int i = 0; i < ItemSize; i++)
            {
                if (_item[i].Name == null || _item[i].Name == "")
                    break;

                xmlFile.Root.Element("OCR").Add(new XElement("TV",
                    new XElement("NAME", _item[i].Name),
                    new XElement("SOURCE", _item[i].Source),
                    new XElement("LABEL", _item[i].Label),
                    new XElement("INFO", _item[i].Info),
                    new XElement("OPTION1", _item[i].Opt1),
                    new XElement("OPTION2", _item[i].Opt2),
                    new XElement("OPTION3", _item[i].Opt3),
                    new XElement("FULL", _item[i].Full)
                    ));
            }
            xmlFile.Save(XMLfile);
            MessageBox.Show("Done !!!");
        }

        private static void XML_new_file()
        {
            XmlTextWriter writer = new XmlTextWriter(XMLfile, System.Text.Encoding.UTF8);
            writer.WriteStartDocument(true);
            writer.Formatting = Formatting.Indented;
            writer.Indentation = 2;
            writer.WriteStartElement("PROFILE");
            writer.WriteStartElement("GRID");
            //createNodeGRID("SANYO_55W", "10;10;10;10", "10;10;10;10", "10;10;10;10", writer);
            writer.WriteEndElement();
            writer.WriteStartElement("OCR");
            createNodeOCR("SANYO_55W", "10;10;10;10", "10;10;10;10", "10;10;10;10", "10;10;10;10", "10;10;10;10", "10;10;10;10", "10;10;10;10", writer);
            createNodeOCR("PHILIPS_55W", "20;20;20;20", "20;20;20;20", "20;20;20;20", "20;20;20;20", "20;20;20;20", "20;20;20;20", "20;20;20;20", writer);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Close();

        }

        private static void createNodeOCR(string Item1, string Item2, string Item3, string Item4, string Item5, string Item6, string Item7, string Item8, XmlTextWriter writer)
        {
            // Item1 : Name
            // Item2 : Source
            // Item3 : Info
            // Item4 : Label
            writer.WriteStartElement("TV");
            writer.WriteStartElement("NAME");
            writer.WriteString(Item1);
            writer.WriteEndElement();
            writer.WriteStartElement("SOURCE");
            writer.WriteString(Item2);
            writer.WriteEndElement();
            writer.WriteStartElement("INFO");
            writer.WriteString(Item3);
            writer.WriteEndElement();
            writer.WriteStartElement("LABEL");
            writer.WriteString(Item4);
            writer.WriteEndElement();
            writer.WriteStartElement("OPTION1");
            writer.WriteString(Item5);
            writer.WriteEndElement();
            writer.WriteStartElement("OPTION2");
            writer.WriteString(Item6);
            writer.WriteEndElement();
            writer.WriteStartElement("OPTION3");
            writer.WriteString(Item7);
            writer.WriteEndElement();
            writer.WriteStartElement("FULL");
            writer.WriteString(Item8);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void createNodeGRID(string Item1, string Item2, string Item3, string Item4, string Item5, string Item6, string Item7, string Item8, XmlTextWriter writer)
        {

        }

        #endregion

    } 
}
