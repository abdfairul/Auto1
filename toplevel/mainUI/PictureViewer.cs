using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace mainUI
{
    public class PictureViewer : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private string[] folderFile = null;
        private int selected = 0;
        private int begin = 0;
        private int end = 0;
        private System.Windows.Forms.Timer timer1;
        private System.ComponentModel.IContainer components;
        private Label currentImageName;
        private string folderPath;
        private Regex searchPattern;

        public PictureViewer(string path, string pattern = ".+") //match everything by default
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            folderPath = path;  
            searchPattern = new Regex(pattern);

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.currentImageName = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.AutoScroll = true;
            this.panel1.BackColor = System.Drawing.Color.Black;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.MinimumSize = new System.Drawing.Size(425, 250);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(685, 503);
            this.panel1.TabIndex = 0;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(683, 501);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Location = new System.Drawing.Point(8, 509);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(128, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "<< Previous";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(548, 509);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(128, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "Next >>";
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.ShowNewFolderButton = false;
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // currentImageName
            // 
            this.currentImageName.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.currentImageName.AutoSize = true;
            this.currentImageName.Location = new System.Drawing.Point(139, 514);
            this.currentImageName.Name = "currentImageName";
            this.currentImageName.Size = new System.Drawing.Size(118, 13);
            this.currentImageName.TabIndex = 4;
            this.currentImageName.Text = "currentImageName";
            // 
            // PictureViewer
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(684, 542);
            this.Controls.Add(this.currentImageName);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(700, 580);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(700, 580);
            this.Name = "PictureViewer";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vision";
            this.Load += new System.EventHandler(this.PictureViewer_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private void load_Folder(string myPath)
        {
            string[] part = null;

            part = Directory.GetFiles(myPath, "*.png").Where(x => searchPattern.IsMatch(x)).ToArray();

            if (part.Length < 1)
            {                                                                               
                MessageBox.Show("No images found!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            folderFile = new string[part.Length];

            Array.Copy(part, 0, folderFile, 0, part.Length);

            selected = 0;
            begin = 0;
            end = folderFile.Length;

            showImage(folderFile[selected]);

            button1.Enabled = true;
            button3.Enabled = true;
        }

        private void showImage(string path)
        {
            Image imgtemp = Image.FromFile(path);
            pictureBox1.Width = imgtemp.Width / 2;
            pictureBox1.Height = imgtemp.Height / 2;
            pictureBox1.Image = imgtemp;
            currentImageName.Text = Path.GetFileName(path);
        }

        private void prevImage()
        {
            if (selected == 0)
            {
                selected = folderFile.Length - 1;
                showImage(folderFile[selected]);
            }
            else
            {
                selected = selected - 1;
                showImage(folderFile[selected]);
            }
        }

        private void nextImage()
        {
            if (selected == folderFile.Length - 1)
            {
                selected = 0;
                showImage(folderFile[selected]);
            }
            else
            {
                selected = selected + 1;
                showImage(folderFile[selected]);
            }
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            prevImage();
        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            nextImage();
        }

        private void timer1_Tick(object sender, System.EventArgs e)
        {
            nextImage();
        }

        private void PictureViewer_Load(object sender, System.EventArgs e)
        {
            button1.Enabled = false;
            button3.Enabled = false;

            load_Folder(folderPath);
        }
    }
}
