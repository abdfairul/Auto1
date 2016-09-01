using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PluginContracts;
using System.Windows.Forms;
using System.Data;
using SAAL;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;
using System.ComponentModel;
using System.Threading;

namespace Test
{
    public class EDID : IPlugin
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

        private ComboBox cbPanel;
        private ComboBox cbBrand;
        private ComboBox cbHDMI;
        private string ActiveSheet;
        private TextBox tbEDID;
        private string[] ArrEDID = new string[256];
        private Form popupBox;
        private DataGridView dgvSpec;
        private DataGridView dgvActual;
        

        private OpenFileDialog openFileDialog = new OpenFileDialog();
        private mainUI.ExcelBinding ExcelBind = new mainUI.ExcelBinding();
        private DataSet dsEDID = new DataSet();

        // Constructor
        public EDID()
        {
            // Add equipment lists for this application
            EquipmentList = new List<string>();
            EquipmentList.Add("QD882");

            PopulateUI();

            ClickEventFlag = dataGridView_CellClick;
        }

        #region Public function to be exposed
        // Passing Name for plugin contract
        public string Name
        {
            get
            {
                return "EDID Test";
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

            #region Load EDID Spec - Populate UI
            RibbonPanel LoadEDIDSpec = new RibbonPanel();
            LoadEDIDSpec.Text = "Load EDID Spec";
            ribbonEquipmentSetting.Panels.Add(LoadEDIDSpec);

            Panel pnlLoadEDID = new Panel();
            pnlLoadEDID.Size = new Size(100, 55);
            pnlLoadEDID.AutoScroll = true;

            cbPanel = new ComboBox();
            cbPanel.Items.Add("WXGA panel");
            cbPanel.Items.Add("Full HD panel");
            cbPanel.Items.Add("UHD panel");
            cbPanel.Items.Add("UHD panel - Laser");
            cbPanel.Size = new Size(80, 20);
            pnlLoadEDID.Controls.Add(cbPanel);

            cbBrand = new ComboBox();
            cbBrand.Items.Add("PHILIPS");
            cbBrand.Items.Add("SANYO");
            cbBrand.Items.Add("MAGNAVOX");
            cbBrand.Size = new Size(80, 20);
            cbBrand.Location = new Point(pnlLoadEDID.Controls[pnlLoadEDID.Controls.Count - 1].Location.X,
                pnlLoadEDID.Controls[pnlLoadEDID.Controls.Count - 1].Location.Y + cbBrand.Height);
            pnlLoadEDID.Controls.Add(cbBrand);

            cbHDMI = new ComboBox();
            cbHDMI.Items.Add("HDMI1");
            cbHDMI.Items.Add("HDMI2");
            cbHDMI.Items.Add("HDMI3");
            cbHDMI.Items.Add("HDMI4");
            cbHDMI.Items.Add("HDMI5");
            cbHDMI.Size = new Size(80, 20);
            cbHDMI.Location = new Point(pnlLoadEDID.Controls[pnlLoadEDID.Controls.Count - 1].Location.X,
                pnlLoadEDID.Controls[pnlLoadEDID.Controls.Count - 1].Location.Y + cbHDMI.Height);
            pnlLoadEDID.Controls.Add(cbHDMI);

            RibbonHost pnlLoadEDIDHost = new RibbonHost();
            pnlLoadEDIDHost.HostedControl = pnlLoadEDID;
            LoadEDIDSpec.Items.Add(pnlLoadEDIDHost);

            Button btLoadEDID = new Button();
            btLoadEDID.Text = "Load";
            btLoadEDID.Size = new Size(55, 22);
            btLoadEDID.FlatStyle = FlatStyle.Popup;
            btLoadEDID.BackColor = Color.LightBlue;
            btLoadEDID.Click += new EventHandler(btLoadEDID_Click);
            RibbonHost btLoadEDIDHost = new RibbonHost();
            btLoadEDIDHost.HostedControl = btLoadEDID;
            LoadEDIDSpec.Items.Add(btLoadEDIDHost);

            Button btRefEDID = new Button();
            btRefEDID.Text = "Refresh";
            btRefEDID.Size = new Size(55, 22);
            btRefEDID.FlatStyle = FlatStyle.Popup;
            btRefEDID.BackColor = Color.LightBlue;
            btRefEDID.Click += new EventHandler(btRefEDID_Click);
            RibbonHost btRefEDIDHost = new RibbonHost();
            btRefEDIDHost.HostedControl = btRefEDID;
            LoadEDIDSpec.Items.Add(btRefEDIDHost);

            Button btViewEDID = new Button();
            btViewEDID.Text = "View Data";
            btViewEDID.Size = new Size(55, 47);
            btViewEDID.FlatStyle = FlatStyle.Popup;
            btViewEDID.BackColor = Color.LightBlue;
            btViewEDID.Click += new EventHandler(btViewEDID_Click);
            RibbonHost btViewEDIDHost = new RibbonHost();
            btViewEDIDHost.HostedControl = btViewEDID;
            LoadEDIDSpec.Items.Add(btViewEDIDHost);

            dgvSpec = new DataGridView();
            dgvSpec.Size = new Size(500, 300);
            dgvSpec.Location = new Point(20, 20);
            dgvSpec.ColumnCount = 16;
            dgvSpec.ColumnHeadersVisible = true;
            dgvSpec.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvSpec.RowHeadersVisible = false;

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle();
            columnHeaderStyle.BackColor = Color.Beige;
            columnHeaderStyle.Font = new Font("Verdana", 10, FontStyle.Bold);
            columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvSpec.ColumnHeadersDefaultCellStyle = columnHeaderStyle;

            for (int i = 0; i < 16; i++)
            {
                dgvSpec.Columns[i].Name = string.Format("{0:X}", i);
                dgvSpec.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            Label lblSpec = new Label();
            lblSpec.Text = "Spec EDID Data";
            //lblSpec.Font = new Font("Verdana", 10, FontStyle.Bold);
            lblSpec.Location = new Point(dgvSpec.Location.X + 200, dgvSpec.Location.Y - 15);

            dgvActual = new DataGridView();
            dgvActual.Size = new Size(500, 300);
            dgvActual.Location = new Point(dgvSpec.Location.X + dgvSpec.Size.Width + 50, dgvSpec.Location.Y);

            Label lblActual = new Label();
            lblActual.Text = "Actual EDID Data";
            //lblActual.Font = new Font("Verdana", 10, FontStyle.Bold);
            lblActual.Location = new Point(dgvActual.Location.X + 200, dgvActual.Location.Y - 15);

            popupBox = new Form();
            popupBox.MaximizeBox = false;
            popupBox.ControlBox = true;
            popupBox.FormBorderStyle = FormBorderStyle.FixedSingle;
            popupBox.Size = new System.Drawing.Size(1100, 400);
            popupBox.FormClosing += new FormClosingEventHandler(popupBox_FormClosing);
            popupBox.StartPosition = FormStartPosition.CenterScreen;
            popupBox.Controls.Add(dgvActual);
            popupBox.Controls.Add(lblActual);
            popupBox.Controls.Add(dgvSpec);
            popupBox.Controls.Add(lblSpec);

            /*TextBox tbEDID = new TextBox();
            tbEDID.Size = new Size(50, 20);
            RibbonHost tbEDIDHost = new RibbonHost();
            tbEDIDHost.HostedControl = tbEDID;
            LoadEDIDSpec.Items.Add(tbEDIDHost);*/
            #endregion
        }

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

        #region Get EDID Value
        private void btLoadEDID_Click(object sender, EventArgs e)
        {
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            openFileDialog.Filter = "All Excel Files|*.xls;*.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FileName = "";
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                dsEDID.Clear();
                ExcelBind.SendFileName = openFileDialog.FileName;
                //dsEDID = ExcelBind.GetResultAllSheets;
            }
        }

        private void btRefEDID_Click(object sender, EventArgs e)
        {
            GetEDIDVal(dsEDID);
        }
        
        private void btViewEDID_Click(object sender, EventArgs e)
        {          
            dgvSpec_update();

            popupBox.BringToFront();
            popupBox.WindowState = FormWindowState.Minimized;
            popupBox.Show();
            popupBox.WindowState = FormWindowState.Normal;
        }

        private void dgvSpec_update()
        {
            if (ArrEDID != null)
            {
                dgvSpec.Rows.Clear();
                for (int i = 0; i < 16; i++)
                {
                    dgvSpec.Rows.Add();
                    for (int j = 0; j < 16; j++)
                    {
                        dgvSpec.Rows[i].Cells[j].Value = ArrEDID[j + i * 16];
                    }
                }
            }
        }

        // Close picture box
        private void popupBox_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            popupBox.Hide();
        }

        // Get EDID value from dataset
        private void GetEDIDVal(DataSet dataSet)
        {
            ActiveSheet = "";
            string columnNo = "";
            int cntRow = 0;
            int cntCol = 0;
            int cntColTemp = 0;
            bool checkPanel = false;
            bool checkBrand = false;
            bool checkHDMI = false;
            bool contentFlag = false;
            
            foreach (DataTable table in dataSet.Tables)
            {
                //Console.WriteLine("Print Table Name : " + table.TableName);
                checkPanel = false;
                checkBrand = false;
                checkHDMI = false;
                if (table.TableName.Contains("HDMI"))
                {
                    if (contentFlag)
                        break;

                    cntRow = 0;
                    foreach (DataRow row in table.Rows)
                    {
                        if (contentFlag)
                        {
                            if (cntRow > 255)
                                break;

                            if (row[columnNo].ToString() == "0")
                                row[columnNo] = "00";

                            ArrEDID[cntRow] = Convert.ToString(row[columnNo]);
                            Console.WriteLine("Content [" + cntRow.ToString() + "] = " + ArrEDID[cntRow]);                            
                            cntRow++;
                        }
                        else
                        {
                            cntCol = 0;
                            foreach (DataColumn column in table.Columns)
                            {
                                cntCol++;
                                if (cbPanel.SelectedItem.ToString() == "UHD panel" && row[column].ToString().Contains("UHD") && !row[column].ToString().Contains("Laser") ||
                                    cbPanel.SelectedItem.ToString() == "UHD panel - Laser" && row[column].ToString().Contains("Laser") ||
                                    cbPanel.SelectedItem.ToString() == "Full HD panel" && row[column].ToString().Contains("Full HD") ||
                                    cbPanel.SelectedItem.ToString() == "Full HD panel" && row[column].ToString().Contains("FHD") ||
                                    cbPanel.SelectedItem.ToString() == "HD+ panel" && row[column].ToString().Contains("HD+") ||
                                    cbPanel.SelectedItem.ToString() == "WXGA panel" && row[column].ToString().Contains("WXGA"))
                                {
                                    cntColTemp = cntCol;
                                    checkPanel = true;
                                    Console.WriteLine("checkPanel = " + checkPanel + " ::: " + row[column].ToString());
                                }

                                if (cbBrand.SelectedItem.ToString() == "PHILIPS" && row[column].ToString().Contains("PHILIPS") ||
                                    cbBrand.SelectedItem.ToString() == "SANYO" && row[column].ToString().Contains("SANYO") ||
                                    cbBrand.SelectedItem.ToString() == "MAGNAVOX" && row[column].ToString().Contains("MAGNAVOX"))
                                {
                                    checkBrand = true;
                                    Console.WriteLine("checkBrand = " + checkBrand + " ::: " + row[column].ToString());
                                }

                                if (cbHDMI.SelectedItem.ToString() == "HDMI1" && row[column].ToString().Contains("HDMI1") ||
                                    cbHDMI.SelectedItem.ToString() == "HDMI2" && row[column].ToString().Contains("HDMI2") ||
                                    cbHDMI.SelectedItem.ToString() == "HDMI3" && row[column].ToString().Contains("HDMI3") ||
                                    cbHDMI.SelectedItem.ToString() == "HDMI4" && row[column].ToString().Contains("HDMI4") ||
                                    cbHDMI.SelectedItem.ToString() == "HDMI5" && row[column].ToString().Contains("HDMI5"))
                                {
                                    if (cntCol >= cntColTemp)
                                        checkHDMI = true;
                                    Console.WriteLine("checkHDMI = " + checkHDMI + " ::: " + row[column].ToString());
                                }

                                if (checkPanel && checkBrand && checkHDMI)
                                {
                                    ActiveSheet = table.TableName;
                                    columnNo = column.ToString();
                                    Console.WriteLine("Table Name = " + ActiveSheet + ", Column = " + columnNo);
                                    
                                    contentFlag = true;
                                    break;
                                }
                            }                  
                        }                        
                    }
                }                
            }
            if(dgvSpec != null)
                dgvSpec_update();
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
