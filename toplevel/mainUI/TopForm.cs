using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PluginContracts;
using System.Text.RegularExpressions;
using System.IO;
using DataGridViewAutoFilter;

namespace mainUI
{


    public partial class TopForm : Form
    {
        private string AppsName = "SQA Automation Tools";

        private ExcelBinding ExcelBind = new ExcelBinding();
        private Dictionary<string, IPlugin> _Plugins;
        private IPlugin m_plugin;
        private DataGridViewImageColumn imageCol;
        private int rowIndex = 0;
        private string load_file_name = "";
        private Form fmSheetName;
        private ComboBox cbSheetName;
        private MruLoader _MruLoader;
        private BackgroundWorker bgWorker;
        private ProgressBar pbBusy;
        private List<int> rowNumToExec;

        private Button BtnQuantum;
        private Button BtnExtInput;
        private Button BtnHDMIPic;
        private Button BtnEDID;

        private RibbonHost BtnQuantumHost = new RibbonHost();
        private RibbonHost BtnExtInputHost = new RibbonHost();
        private RibbonHost BtnHDMIPicHost = new RibbonHost();
        private RibbonHost BtnEDIDHost = new RibbonHost();

        private RibbonPanel panelTestCases = new RibbonPanel();
        private RibbonTab testCasesWizard = new RibbonTab();

        #region Load TopForm
        public TopForm()
        {
            InitializeComponent();

            _Plugins = new Dictionary<string, IPlugin>();

            ICollection<IPlugin> plugins = PluginLoader.LoadPlugins(".");
            foreach (var item in plugins)
            {
                _Plugins.Add(item.Name, item);
            }

            this.Text = AppsName;

            _MruLoader = new MruLoader(AppsName, ribbonMain.OrbDropDown, 10);
            _MruLoader.FileSelected += _MruLoader_FileSelected;

            StatusBar.toolLabel = toolStripStatusLabel;

            BtnQuantum = new Button();
            BtnQuantum.Size = new System.Drawing.Size(100, 60);
            BtnQuantum.Image = global::mainUI.Properties.Resources.quantum_logo64x64;
            BtnQuantum.Text = "Quantum All Format";
            BtnQuantum.TextAlign = ContentAlignment.BottomCenter;
            BtnQuantum.Tag = "Test using QuantumData device. Relevent Excel document is *QuantumData_All_format*.xls";
            BtnQuantum.Click += new EventHandler(ribbonOrbMenuItemQuantumDataAllFormat_wizard_Click);
            BtnQuantum.UpdateStatusBar_MouseEvents();

            BtnExtInput = new Button();
            BtnExtInput.Size = new System.Drawing.Size(100, 60);
            BtnExtInput.Image = global::mainUI.Properties.Resources.funailogo;
            BtnExtInput.Text = "Ext Input";
            BtnExtInput.TextAlign = ContentAlignment.BottomCenter;
            BtnExtInput.Tag = "Test each external input. Relevent Excel document is *ExtInput（Except PCInput)*.xls";
            BtnExtInput.Click += new EventHandler(ribbonOrbMenuItemExtInput_wizard_Click);
            BtnExtInput.UpdateStatusBar_MouseEvents();

            BtnHDMIPic = new Button();
            BtnHDMIPic.Size = new System.Drawing.Size(100, 60);
            BtnHDMIPic.Image = global::mainUI.Properties.Resources.funailogo;
            BtnHDMIPic.Text = "HDMI Pic Format";
            BtnHDMIPic.TextAlign = ContentAlignment.BottomCenter;
            BtnHDMIPic.Tag = "Test HDMI picture. Relevant Excel document is *HDMI_PICTFORMAT*.xls";
            BtnHDMIPic.Click += new EventHandler(ribbonOrbMenuItemHDMIPictureFormat_wizard_Click);
            BtnHDMIPic.UpdateStatusBar_MouseEvents();

            BtnEDID = new Button();
            BtnEDID.Size = new System.Drawing.Size(100, 60);
            BtnEDID.Image = global::mainUI.Properties.Resources.funailogo;
            BtnEDID.Text = "EDID";
            BtnEDID.TextAlign = ContentAlignment.BottomCenter;
            BtnEDID.Click += new EventHandler(ribbonOrbMenuItemEDIDtest_Click);
            BtnEDID.UpdateStatusBar_MouseEvents();

            this.toolTipMain.SetToolTip(BtnQuantum, "Quantum All Data Test");
            this.toolTipMain.SetToolTip(BtnExtInput, "External Input Test");
            this.toolTipMain.SetToolTip(BtnHDMIPic, "HDMI Picture Format Test");
            this.toolTipMain.SetToolTip(BtnEDID, "EDID Test");

            BtnQuantumHost.HostedControl = BtnQuantum;
            BtnExtInputHost.HostedControl = BtnExtInput;
            BtnHDMIPicHost.HostedControl = BtnHDMIPic;
            BtnEDIDHost.HostedControl = BtnEDID;

            panelTestCases.Text = "Please choose your testcases";
            testCasesWizard.Text = "Testcases";

            panelTestCases.Items.Add(BtnQuantumHost);
            panelTestCases.Items.Add(BtnExtInputHost);
            panelTestCases.Items.Add(BtnHDMIPicHost);
            panelTestCases.Items.Add(BtnEDIDHost);

            testCasesWizard.Panels.Add(panelTestCases);
            this.ribbonMain.Tabs.Add(testCasesWizard);

            this.bgWorker = new BackgroundWorker();
            this.bgWorker.WorkerReportsProgress = true;
            this.bgWorker.DoWork += new DoWorkEventHandler(this.bgWorker_DoWork);
            this.bgWorker.ProgressChanged += new ProgressChangedEventHandler(this.bgWorker_ProgressChanged);
            this.bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.bgWorker_RunWorkerCompleted);
            this.bgWorker.WorkerSupportsCancellation = true;

            this.pbBusy = new ProgressBar();
            this.pbBusy.Anchor = ((AnchorStyles)((AnchorStyles.Top | AnchorStyles.Right)));
            this.pbBusy.Location = new Point(850, 10);
            this.pbBusy.Size = new Size(100, 17);
            this.pbBusy.Style = ProgressBarStyle.Marquee;
            this.Controls.Add(this.pbBusy);
            this.pbBusy.BringToFront();

            this.pbBusy.Visible = false;
        }
        #endregion

        #region Test Items

        // Quantum Data All Format Test
        private void ribbonOrbMenuItemQuantumDataAllFormat_wizard_Click(object sender, EventArgs e)
        {
            m_plugin = _Plugins["QuantumAllFormat Test"];
            this.ribbonMain.Tabs.Clear();
            this.ribbonMain.Tabs.Add(testCasesWizard);
            this.ribbonMain.Tabs.Add(m_plugin.EquipmentSetting);

            BtnManagement(false, true, true, true);

            ribbonOrbMenuItemOpen_Click(null, null);
            this.ribbonMain.ActivateNextTab();
        }

        // ExtInput Test
        private void ribbonOrbMenuItemExtInput_wizard_Click(object sender, EventArgs e)
        {
            m_plugin = _Plugins["ExtInput Test"];
            this.ribbonMain.Tabs.Clear();
            this.ribbonMain.Tabs.Add(testCasesWizard);
            this.ribbonMain.Tabs.Add(m_plugin.EquipmentSetting);

            BtnManagement(true, false, true, true);

            ribbonOrbMenuItemOpen_Click(null, null);
            this.ribbonMain.ActivateNextTab();
        }

        // HDMI Picture Format Test
        private void ribbonOrbMenuItemHDMIPictureFormat_wizard_Click(object sender, EventArgs e)
        {
            m_plugin = _Plugins["HDMIPictureFormat Test"];
            this.ribbonMain.Tabs.Clear();
            this.ribbonMain.Tabs.Add(testCasesWizard);
            this.ribbonMain.Tabs.Add(m_plugin.EquipmentSetting);

            BtnManagement(true, true, false, true);

            ribbonOrbMenuItemOpen_Click(null, null);
            this.ribbonMain.ActivateNextTab();
        }

        // EDID Test 
        private void ribbonOrbMenuItemEDIDtest_Click(object sender, EventArgs e)
        {
            m_plugin = _Plugins["EDID Test"];
            this.ribbonMain.Tabs.Clear();
            this.ribbonMain.Tabs.Add(testCasesWizard);
            this.ribbonMain.Tabs.Add(m_plugin.EquipmentSetting);

            BtnManagement(true, true, true, false);

            ribbonOrbMenuItemOpen_Click(null, null);
            this.ribbonMain.ActivateNextTab();
        }

        // PlugUnplug Test
        private void ribbonOrbMenuItemPlugUnplugTest_Click(object sender, EventArgs e)
        {
            //m_plugin = _Plugins["PlugUnplug Test"];
            //this.ribbon1.Tabs.Clear();
            //this.ribbon1.Tabs.Add(m_plugin.EquipmentSetting);
        }

        // Tuning Test
        private void ribbonOrbMenuItemTuningTest_Click(object sender, EventArgs e)
        {
            //m_plugin = _Plugins["ItemTuning Test"];
            //this.ribbon1.Tabs.Clear();
            //this.ribbon1.Tabs.Add(m_plugin.EquipmentSetting);           
        }

        private void BtnManagement(bool BtnQuantumEn, bool BtnExtInputEn,
            bool BtnHDMIPicEn, bool BtnEDIDEn)
        {
            this.BtnQuantum.Enabled = BtnQuantumEn;
            this.BtnExtInput.Enabled = BtnExtInputEn;
            this.BtnHDMIPic.Enabled = BtnHDMIPicEn;
            this.BtnEDID.Enabled = BtnEDIDEn;
        }
        #endregion

        #region Misc
        // Load a file (Open file dialog)
        private void ribbonOrbMenuItemOpen_Click(object sender, EventArgs e)
        {
            //openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            openFileDialog.Filter = "All Excel Files|*.xls;*.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FileName = "";
            openFileDialog.ShowDialog();
        }

        // Load a file (Press OK event)
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            load_file_name = openFileDialog.FileName;
            OpenFile();
        }

        // Load a file event
        private void OpenFile()
        {
            // Change application's title
            this.Text = AppsName + " - " + Path.GetFileName(load_file_name);

            // Open sheet selection windows
            ExcelBind.SendFileName = load_file_name;

            // Create new popup form
            fmSheetName = new Form();
            fmSheetName.Size = new System.Drawing.Size(390, 50);
            fmSheetName.ControlBox = false;
            fmSheetName.StartPosition = FormStartPosition.CenterParent;
            fmSheetName.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            fmSheetName.TopMost = true;
            fmSheetName.BackColor = System.Drawing.Color.LightGray;

            // Create combobox inside new form for sheet selection
            cbSheetName = new ComboBox();
            cbSheetName.Text = "Please select test sheet";
            cbSheetName.Location = new System.Drawing.Point(10, 10);
            cbSheetName.Size = new System.Drawing.Size(200, 10);
            char[] MyChar = { '$', '\'' };
            Regex regex = new Regex(@"\'.*\$\'$");
            Int16 i = 0;
            foreach (string SN in ExcelBind.GetSheetName)
            {
                Match match = regex.Match(SN);
                if (match.Success)
                {
                    string SN_trim = SN.Trim(MyChar);
                    cbSheetName.Items.Insert(i, SN_trim);
                    i++;
                }
            }
            fmSheetName.Controls.Add(cbSheetName);

            // Create ok button to choose sheet
            Button btSheetSelect = new Button();
            btSheetSelect.Text = "OK";
            btSheetSelect.Click += new System.EventHandler(this.btSheetSelect_Click);
            btSheetSelect.Location = new System.Drawing.Point(cbSheetName.Width + 20, cbSheetName.Location.Y - 1);
            fmSheetName.Controls.Add(btSheetSelect);

            // Create cancel button to close
            Button btSheetCancel = new Button();
            btSheetCancel.Text = "Cancel";
            btSheetCancel.Click += new System.EventHandler(this.btSheetCancel_Click);
            btSheetCancel.Location = new System.Drawing.Point(cbSheetName.Width + btSheetSelect.Width + 20, cbSheetName.Location.Y - 1);
            fmSheetName.Controls.Add(btSheetCancel);

            // Show popup window    
            fmSheetName.ShowDialog(this);
        }

        private void btSheetSelect_Click(object sender, EventArgs e)
        {
            if (cbSheetName.SelectedItem != null)
            {
                ExcelBind.SendSheetName = cbSheetName.SelectedItem.ToString() + "$";
                fmSheetName.Close();
                if (bgWorker.IsBusy != true)
                {
                    bgWorker.RunWorkerAsync("load");
                }
            }
            else
            {
                MessageBox.Show("Please select any sheet !!!");
            }
        }

        private void btSheetCancel_Click(object sender, EventArgs e)
        {
            fmSheetName.Close();
        }

        // Load a file selected from the MRU list.
        private void _MruLoader_FileSelected(string file_name)
        {
            load_file_name = file_name;
            OpenFile();
        }

        // Save a file (Open file dialog)
        private void ribbonOrbMenuItemSave_Click(object sender, EventArgs e)
        {
            if (bgWorker.IsBusy != true)
            {
                bgWorker.RunWorkerAsync("save");
            }
        }

        // Save a file (Open file dialog)
        private void ribbonOrbMenuItemSaveAs_Click(object sender, EventArgs e)
        {
            saveFileDialog.InitialDirectory = Environment.CurrentDirectory;
            saveFileDialog.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.ShowDialog();
        }

        // Save a file (Press OK event)
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            // Change application's title
            load_file_name = saveFileDialog.FileName;
            this.Text = AppsName + " - " + Path.GetFileName(load_file_name);

            if (bgWorker.IsBusy != true)
            {
                bgWorker.RunWorkerAsync("saveAs");
            }
        }

        // About
        private void ribbonOrbOptionButtonAbout_Click(object sender, EventArgs e)
        {

        }

        // Close file
        private void ribbonOrbMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Text = AppsName;
            dataGridView.Columns.Clear();
            this.ribbonMain.Tabs.Clear();
            this.ribbonMain.Tabs.Add(testCasesWizard);
            this.BtnQuantum.Enabled = true;
            this.BtnExtInput.Enabled = true;
            this.BtnHDMIPic.Enabled = true;
        }

        // Exit
        private void ribbonOrbMenuItemExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Get datagrid index
        private void dataGridView_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.rowIndex = e.RowIndex;

                if (m_plugin == null)
                {
                    executeToolStripMenuItem.Enabled = false;
                    pauseToolStripMenuItem.Enabled = false;
                    stopToolStripMenuItem.Enabled = false;
                }
                else
                {
                    executeToolStripMenuItem.Enabled = true;
                    pauseToolStripMenuItem.Enabled = true;
                    stopToolStripMenuItem.Enabled = true;
                }
            }
        }

        // Press Execute event
        private void executeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rowNumToExec = new List<int>();
            if (bgWorker.IsBusy != true)
            {
                m_plugin.BeforeExecute?.Invoke();

                if (m_plugin.DoExecute)
                    bgWorker.RunWorkerAsync("execute");
            }
        }

        // Press Pause event
        private void pauseToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        // Press Stop event
        private void stopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bgWorker.CancelAsync();
        }

        // Select all event
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView.Rows[dataGridView.CurrentCell.RowIndex].Selected = false;
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
            {
                Regex regex = new Regex(@"\d+");

                Match match = regex.Match(dataGridView.Rows[i].Cells[0].Value.ToString());
                if (match.Success)
                {
                    this.dataGridView.Rows[i].Selected = true;
                }
            }
        }

        // Select from here event
        private void selectFromHereToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = dataGridView.CurrentCell.RowIndex;
                i < dataGridView.Rows.Count - 1; i++)
            {
                Regex regex = new Regex(@"\d+");

                Match match = regex.Match(dataGridView.Rows[i].Cells[0].Value.ToString());
                if (match.Success)
                {
                    this.dataGridView.Rows[i].Selected = true;
                }
            }
        }

        // Press Delete event
        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!this.dataGridView.Rows[this.rowIndex].IsNewRow)
            {
                this.dataGridView.Rows.RemoveAt(this.rowIndex);
            }
        }

        private void dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (m_plugin != null && m_plugin.ClickEvent != null)
                m_plugin.ClickEvent.Invoke(sender, e);
        }
        #endregion

        # region BG Worker
        void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pbBusy.Visible = false;
            if (e.Cancelled)
            {
                MessageBox.Show("Execution Stop !");
                return;
            }

            try
            {
                if (e.Result.ToString() == "execution")
                {
                    m_plugin.AfterExecute?.Invoke();
                    MessageBox.Show("Execution Done !");
                }
                else if (e.Result.ToString() == "save")
                {
                    MessageBox.Show("File Saved");
                }
                else if (e.Result.ToString() == "saveAs")
                {
                    MessageBox.Show("File Saved");
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbBusy.Visible = true;
            m_plugin.ProgressBarChangedEvent(sender, e);
        }

        delegate void GetFocusUIDelegate(int i);
        void GetFocusUI(int i)
        {
            if (dataGridView.InvokeRequired)
            {
                dataGridView.Invoke(new GetFocusUIDelegate(GetFocusUI), i);
            }
            else
            {
                if (m_plugin.Name == "QuantumAllFormat Test")
                {
                    dataGridView.CurrentCell = dataGridView[8, i];
                }
                else // define later for each testaces which column to focus
                    dataGridView.CurrentCell = dataGridView[0, i];
            }
        }

        void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
           BackgroundWorker worker = (BackgroundWorker)sender;

            // Save function
            if (e.Argument.ToString() == "save")
            {
                if (dataGridView.Columns.Count != 0) //if datagrid view not null
                {
                    //DataTable dtToSave = new DataTable();
                    List<string> noToSave = new List<string>();
                    List<string> resultToSave = new List<string>();
                    List<string> testerToSave = new List<string>();
                    List<string> verToSave = new List<string>();
                    List<string> dateToSave = new List<string>();
                    List<string> remarkToSave = new List<string>();

                    //dtToSave.Columns.Add();

                    // Data cells
                    this.Invoke(new MethodInvoker(delegate
                    {
                        int pointToSaveRes = 0;
                        int pointToSaveVer = 0;
                        int pointToSaveTester = 0;
                        int pointToSaveDate = 0;
                        int pointToSaveRemark = 0;
                        for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                        {
                            DataGridViewRow row = dataGridView.Rows[i];
                            //DataRow dr = dtToSave.NewRow();
                            if (i <= 1)
                            {
                                for (int j = 0; j < dataGridView.Columns.Count - 1; j++)
                                {
                                    if (row.Cells[j].Value == null)
                                        continue;

                                    if (row.Cells[j].Value.ToString() == "Result")
                                    {
                                        pointToSaveRes = j;
                                    }
                                    else if (row.Cells[j].Value.ToString() == "Ver." || row.Cells[j].Value.ToString() == "Ver")
                                    {
                                        pointToSaveVer = j;
                                    }
                                    else if (row.Cells[j].Value.ToString() == "Tester")
                                    {
                                        pointToSaveTester = j;
                                    }
                                    else if (row.Cells[j].Value.ToString() == "Date" || row.Cells[j].Value.ToString() == "Test Date")
                                    {
                                        pointToSaveDate = j;
                                    }
                                    else if (row.Cells[j].Value.ToString() == "Remark" || row.Cells[j].Value.ToString() == "Remarks")
                                    {
                                        pointToSaveRemark = j;
                                    }
                                }
                            }

                            //Console.WriteLine(row.Cells[j].Value.ToString());
                            if (row.Cells[pointToSaveRes].Value.ToString() == "OK" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "NG1" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "NG2" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "NG3" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "NG4" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "PEND" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "NT" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "-" ||
                                row.Cells[pointToSaveRes].Value.ToString() == "OK/NG")
                            {
                                noToSave.Add(row.Cells[0].Value.ToString());
                                resultToSave.Add(row.Cells[pointToSaveRes].Value.ToString());
                                verToSave.Add(row.Cells[pointToSaveVer].Value.ToString());
                                testerToSave.Add(row.Cells[pointToSaveTester].Value.ToString());
                                dateToSave.Add(row.Cells[pointToSaveDate].Value.ToString());
                                remarkToSave.Add(row.Cells[pointToSaveRemark].Value.ToString());
                                //dtToSave.Rows.Add(row.Cells[j].Value);
                            }
                        }
                    }));
                    ExcelBind.save_data(noToSave, resultToSave, verToSave, testerToSave, dateToSave, remarkToSave);
                }
                e.Result = "save";
            }

            // Save As function
            else if (e.Argument.ToString() == "saveAs")
            {
                if (dataGridView.Columns.Count != 0) //if datagrid view not null
                {
                    DataTable dtToSave = new DataTable();

                    // Header columns
                    foreach (DataGridViewColumn column in dataGridView.Columns)
                    {
                        DataColumn dc = new DataColumn(column.Name.ToString());
                        dtToSave.Columns.Add(dc);
                    }

                    // Data cells
                    for (int i = 0; i < dataGridView.Rows.Count; i++)
                    {
                        DataGridViewRow row = dataGridView.Rows[i];
                        DataRow dr = dtToSave.NewRow();
                        for (int j = 0; j < dataGridView.Columns.Count; j++)
                        {
                            dr[j] = (row.Cells[j].Value == null) ? "" : row.Cells[j].Value.ToString();
                        }
                        dtToSave.Rows.Add(dr);
                    }
                    ExcelBind.SendFileName = saveFileDialog.FileName;
                    ExcelBind.SaveAsResult = dtToSave;
                }
                else { MessageBox.Show("Nothing to save"); }
                e.Result = "saveAs";
            }

            // Execute function
            else if (e.Argument.ToString() == "execute")
            {
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    if (dataGridView.Rows[i].Selected == true)
                    {
                        rowNumToExec.Add(i);
                        dataGridView.Rows[i].Selected = false;
                        dataGridView.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                        Console.WriteLine("Check row number : " + i);
                    }
                }
                
                m_plugin.BeforeExecuteButAfterSelection?.Invoke(rowNumToExec, dataGridView.Rows, this, worker);

                foreach (int i in rowNumToExec)
                {
                    dataGridView.Rows[i].Selected = false;
                    dataGridView.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;

                    DataGridViewRow execution = dataGridView.Rows[i];
                    Console.WriteLine("Test Num : " + i);

                    GetFocusUI(i);// get focus
                    
                    // Send datarow
                    m_plugin.ToExecute = execution;

                    while (m_plugin.Busy)
                    {
                        //do nothing
                    }
                    
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;

                        // clear highlights
                        foreach (int j in rowNumToExec)
                            dataGridView.Rows[j].DefaultCellStyle.BackColor = Color.Empty;

                        break;
                    }

                    //Get datarow
                    DataGridViewRow executionResult = m_plugin.ToExecute;
                    dataGridView.Rows[i].DefaultCellStyle.BackColor = Color.Empty;

                    this.Invoke(new MethodInvoker(delegate
                    {
                        int j = 0;
                        foreach (DataGridViewColumn dc in dataGridView.Columns)
                        {
                            if (executionResult.Cells[j].Value != null)
                            {
                                Regex regex = new Regex(@"#ERROR");
                                Match match = regex.Match(executionResult.Cells[j].Value.ToString());
                                if (match.Success)
                                    executionResult.Cells[j].Style.BackColor = Color.Red;
                                else
                                    executionResult.Cells[j].Style.BackColor = Color.Empty;

                                if (executionResult.Cells[j].Value.ToString() == "NG1")
                                {
                                    executionResult.Cells[j].Style.BackColor = Color.Red;
                                }
                                else if (executionResult.Cells[j].Value.ToString() == "OK")
                                {
                                    executionResult.Cells[j].Style.BackColor = Color.Empty;
                                }
                            }

                            dataGridView.Rows[i].Cells[j].Value = executionResult.Cells[j].Value;
                            j++;
                        }
                    }));
                }

                e.Result = "execution";
            }
            // Load function
            else if (e.Argument.ToString() == "load")
            {
                string file_name = load_file_name;
                try
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        // Load the file.

                        dataGridView.Columns.Clear();
                        dataGridView.DataSource = null;
                        BindingSource dataSource = new BindingSource(ExcelBind.GetResult, null);
                        dataGridView.DataSource = dataSource;

                        if (dataGridView.DataSource == null)
                            return;

                        int index = 65;
                        char headerText;
                        foreach (DataGridViewColumn col in dataGridView.Columns)
                        {
                            headerText = (char)index;
                            col.HeaderText = headerText.ToString();
                            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            DataGridViewAutoFilterColumnHeaderCell tempCol =
                                new DataGridViewAutoFilterColumnHeaderCell(col.HeaderCell);
                            tempCol.AutomaticSortingEnabled = false;
                            col.HeaderCell = tempCol;

                            index++;
                        }

                        for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                        {
                            //Regex regex = new Regex(@"\d+");
                            Regex regex = new Regex(@"no|No|NO");
                            Match match = regex.Match(dataGridView.Rows[i].Cells[0].Value.ToString());
                            if (match.Success)
                            {
                                dataGridView.Rows[i].DefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
                                dataGridView.Rows[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                dataGridView.Rows[i].DefaultCellStyle.BackColor = Color.SkyBlue;
                            }
                        }

                        // only once creation
                        imageCol = new DataGridViewImageColumn();
                        dataGridView.Columns.Add(imageCol);

                        // Add the file to the MRU list.
                        _MruLoader.AddFile(file_name);
                    }));
                }

                catch (Exception ex)
                {
                    // Remove the file from the MRU list.
                    _MruLoader.RemoveFile(file_name);

                    // Tell the user what happened.
                    MessageBox.Show(ex.Message);
                }
                e.Result = "load";
            }
        }
        #endregion

        private void TopForm_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = "Welcome to FMRD SQA Automation Tool v2016! Please choose any of the testcases you wish to perform";
        }
    }


    public static class StatusBar
    {
        public static ToolStripStatusLabel toolLabel;

        public static void UpdateStatusBar_MouseEnter(object sender, EventArgs e)
        {
            var control = sender as Control;

            if (control == null)
                toolLabel.Text = "";
            else if (toolLabel.Text == null)
                toolLabel.Text = "";
            else if (control.Tag == null)
                toolLabel.Text = control.Text;
            else
                toolLabel.Text = control.Tag.ToString();
        }

        public static void UpdateStatusBar_MouseLeave(object sender, EventArgs e)
        {
            toolLabel.Text = "";
        }

        public static void UpdateStatusBar_MouseEvents(this Control control)
        {
            control.MouseEnter += UpdateStatusBar_MouseEnter;
            control.MouseLeave += UpdateStatusBar_MouseLeave;
        }
    }
}
