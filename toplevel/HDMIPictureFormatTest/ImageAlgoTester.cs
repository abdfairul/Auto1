using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.CvEnum;
using SAAL;

namespace Test
{
    public partial class ImageAlgoTester : Form
    {
        private Mat Image1;
        private Mat Image2;
        public ImageAlgoTester()
        {
            InitializeComponent();
        }

        private void loadImage1_Click(object sender, EventArgs e)
        {
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";

            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                image1tb.Text = openFileDialog1.FileName;
                if (Image1 != null) Image1.Dispose();
                Image1 = new Mat(image1tb.Text, LoadImageType.Color);
                pictureBox1_1.Image = Image1.Bitmap;
            }
        }

        private void loadImage2_Click(object sender, EventArgs e)
        {
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";

            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                image2tb.Text = openFileDialog1.FileName;
                if(Image2 != null) Image2.Dispose();

                Image2 = new Mat(image2tb.Text, LoadImageType.Color);
                pictureBox2_1.Image = Image2.Bitmap;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var p1 = Convert.ToInt16(param1tb.Text);
            var p2 = Convert.ToInt16(param2tb.Text);
            var p3 = Convert.ToInt16(param3tb.Text);
            var p4 = Convert.ToInt16(param4tb.Text);

            var bPercent = 0.0;
            var ori = new Bitmap(640, 480);
            
            if (comboBox1.SelectedItem.ToString() == "Global Binary")
            {
                //pictureBox1_2.Image = ImageHandler.GetBinarialImage2(ref bPercent, ref ori, image1tb.Text, p1, p2);
                //label1black.Text = bPercent.ToString();
                //pictureBox2_2.Image = ImageHandler.GetBinarialImage2(ref bPercent, ref ori, image2tb.Text, p1, p2);
                //label2black.Text = bPercent.ToString();

            }
            else if (comboBox1.SelectedItem.ToString() == "Adaptive Binary")
            {
                //pictureBox1_2.Image = ImageHandler.GetBinarialImage(ref bPercent, ref ori, image1tb.Text, p1, p2, p3);
                //label1black.Text = bPercent.ToString();
                //pictureBox2_2.Image = ImageHandler.GetBinarialImage(ref bPercent, ref ori, image2tb.Text, p1, p2, p3);
                //label2black.Text = bPercent.ToString();
            }
            else if (comboBox1.SelectedItem.ToString() == "InRange")
            {
                //pictureBox1_2.Image = ImageHandler.GetBinarialImage3(ref bPercent, ref ori, image1tb.Text, p1, p2, p3, p4);
                //label1black.Text = bPercent.ToString();
                //pictureBox2_2.Image = ImageHandler.GetBinarialImage3(ref bPercent, ref ori, image2tb.Text, p1, p2, p3, p4);
                //label2black.Text = bPercent.ToString();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var combo = sender as ComboBox;
            if (combo.SelectedItem.ToString() == "Global Binary")
            {
                param1tb.Visible = true;
                p1label.Visible = true;
                param2tb.Visible = true;
                p2label.Visible = true;
                param3tb.Visible = false;
                p3label.Visible = false;
                param4tb.Visible = false;
                p4label.Visible = false;
            }
            else if (combo.SelectedItem.ToString() == "Adaptive Binary")
            {
                param1tb.Visible = true;
                p1label.Visible = true;
                param2tb.Visible = true;
                p2label.Visible = true;
                param3tb.Visible = true;
                p3label.Visible = true;
                param4tb.Visible = false;
                p4label.Visible = false;
            }
            else if (combo.SelectedItem.ToString() == "InRange")
            {
                param1tb.Visible = true;
                p1label.Visible = true;
                param2tb.Visible = true;
                p2label.Visible = true;
                param3tb.Visible = true;
                p3label.Visible = true;
                param4tb.Visible = true;
                p4label.Visible = true;
            }

        }

    }
}
