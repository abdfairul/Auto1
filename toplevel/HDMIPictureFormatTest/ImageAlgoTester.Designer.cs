namespace Test
{
    partial class ImageAlgoTester
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImageAlgoTester));
            this.pictureBox1_1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2_1 = new System.Windows.Forms.PictureBox();
            this.pictureBox1_2 = new System.Windows.Forms.PictureBox();
            this.pictureBox2_2 = new System.Windows.Forms.PictureBox();
            this.loadImage1 = new System.Windows.Forms.Button();
            this.loadImage2 = new System.Windows.Forms.Button();
            this.image1tb = new System.Windows.Forms.TextBox();
            this.image2tb = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.p1label = new System.Windows.Forms.Label();
            this.param1tb = new System.Windows.Forms.TextBox();
            this.param2tb = new System.Windows.Forms.TextBox();
            this.p2label = new System.Windows.Forms.Label();
            this.param3tb = new System.Windows.Forms.TextBox();
            this.p3label = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label1black = new System.Windows.Forms.Label();
            this.label2black = new System.Windows.Forms.Label();
            this.param4tb = new System.Windows.Forms.TextBox();
            this.p4label = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1_1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2_1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1_2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2_2)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1_1
            // 
            this.pictureBox1_1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1_1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1_1.Image")));
            this.pictureBox1_1.Location = new System.Drawing.Point(12, 33);
            this.pictureBox1_1.Name = "pictureBox1_1";
            this.pictureBox1_1.Size = new System.Drawing.Size(289, 256);
            this.pictureBox1_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1_1.TabIndex = 0;
            this.pictureBox1_1.TabStop = false;
            // 
            // pictureBox2_1
            // 
            this.pictureBox2_1.Location = new System.Drawing.Point(335, 33);
            this.pictureBox2_1.Name = "pictureBox2_1";
            this.pictureBox2_1.Size = new System.Drawing.Size(289, 256);
            this.pictureBox2_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2_1.TabIndex = 1;
            this.pictureBox2_1.TabStop = false;
            // 
            // pictureBox1_2
            // 
            this.pictureBox1_2.Location = new System.Drawing.Point(12, 316);
            this.pictureBox1_2.Name = "pictureBox1_2";
            this.pictureBox1_2.Size = new System.Drawing.Size(289, 256);
            this.pictureBox1_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1_2.TabIndex = 2;
            this.pictureBox1_2.TabStop = false;
            // 
            // pictureBox2_2
            // 
            this.pictureBox2_2.Location = new System.Drawing.Point(335, 318);
            this.pictureBox2_2.Name = "pictureBox2_2";
            this.pictureBox2_2.Size = new System.Drawing.Size(289, 256);
            this.pictureBox2_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2_2.TabIndex = 3;
            this.pictureBox2_2.TabStop = false;
            // 
            // loadImage1
            // 
            this.loadImage1.Location = new System.Drawing.Point(226, 4);
            this.loadImage1.Name = "loadImage1";
            this.loadImage1.Size = new System.Drawing.Size(75, 23);
            this.loadImage1.TabIndex = 4;
            this.loadImage1.Text = "button1";
            this.loadImage1.UseVisualStyleBackColor = true;
            this.loadImage1.Click += new System.EventHandler(this.loadImage1_Click);
            // 
            // loadImage2
            // 
            this.loadImage2.Location = new System.Drawing.Point(548, 3);
            this.loadImage2.Name = "loadImage2";
            this.loadImage2.Size = new System.Drawing.Size(75, 23);
            this.loadImage2.TabIndex = 5;
            this.loadImage2.Text = "button2";
            this.loadImage2.UseVisualStyleBackColor = true;
            this.loadImage2.Click += new System.EventHandler(this.loadImage2_Click);
            // 
            // image1tb
            // 
            this.image1tb.Location = new System.Drawing.Point(12, 5);
            this.image1tb.Name = "image1tb";
            this.image1tb.Size = new System.Drawing.Size(208, 20);
            this.image1tb.TabIndex = 6;
            // 
            // image2tb
            // 
            this.image2tb.Location = new System.Drawing.Point(335, 5);
            this.image2tb.Name = "image2tb";
            this.image2tb.Size = new System.Drawing.Size(207, 20);
            this.image2tb.TabIndex = 7;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Adaptive Binary",
            "Global Binary",
            "InRange"});
            this.comboBox1.Location = new System.Drawing.Point(647, 55);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(108, 21);
            this.comboBox1.TabIndex = 8;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(644, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Method:";
            // 
            // p1label
            // 
            this.p1label.AutoSize = true;
            this.p1label.Location = new System.Drawing.Point(644, 90);
            this.p1label.Name = "p1label";
            this.p1label.Size = new System.Drawing.Size(46, 13);
            this.p1label.TabIndex = 10;
            this.p1label.Text = "Param1:";
            // 
            // param1tb
            // 
            this.param1tb.Location = new System.Drawing.Point(696, 87);
            this.param1tb.Name = "param1tb";
            this.param1tb.Size = new System.Drawing.Size(59, 20);
            this.param1tb.TabIndex = 12;
            this.param1tb.Text = "0";
            // 
            // param2tb
            // 
            this.param2tb.Location = new System.Drawing.Point(696, 113);
            this.param2tb.Name = "param2tb";
            this.param2tb.Size = new System.Drawing.Size(59, 20);
            this.param2tb.TabIndex = 14;
            this.param2tb.Text = "255";
            // 
            // p2label
            // 
            this.p2label.AutoSize = true;
            this.p2label.Location = new System.Drawing.Point(644, 116);
            this.p2label.Name = "p2label";
            this.p2label.Size = new System.Drawing.Size(46, 13);
            this.p2label.TabIndex = 13;
            this.p2label.Text = "Param2:";
            // 
            // param3tb
            // 
            this.param3tb.Location = new System.Drawing.Point(696, 139);
            this.param3tb.Name = "param3tb";
            this.param3tb.Size = new System.Drawing.Size(59, 20);
            this.param3tb.TabIndex = 16;
            this.param3tb.Text = "0";
            // 
            // p3label
            // 
            this.p3label.AutoSize = true;
            this.p3label.Location = new System.Drawing.Point(644, 142);
            this.p3label.Name = "p3label";
            this.p3label.Size = new System.Drawing.Size(46, 13);
            this.p3label.TabIndex = 15;
            this.p3label.Text = "Param3:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(647, 186);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 17;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 300);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Black %:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(332, 300);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Black %:";
            // 
            // label1black
            // 
            this.label1black.AutoSize = true;
            this.label1black.Location = new System.Drawing.Point(66, 300);
            this.label1black.Name = "label1black";
            this.label1black.Size = new System.Drawing.Size(35, 13);
            this.label1black.TabIndex = 20;
            this.label1black.Text = "label7";
            this.label1black.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2black
            // 
            this.label2black.AutoSize = true;
            this.label2black.Location = new System.Drawing.Point(386, 300);
            this.label2black.Name = "label2black";
            this.label2black.Size = new System.Drawing.Size(35, 13);
            this.label2black.TabIndex = 21;
            this.label2black.Text = "label8";
            this.label2black.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // param4tb
            // 
            this.param4tb.Location = new System.Drawing.Point(696, 165);
            this.param4tb.Name = "param4tb";
            this.param4tb.Size = new System.Drawing.Size(59, 20);
            this.param4tb.TabIndex = 23;
            this.param4tb.Text = "0";
            // 
            // p4label
            // 
            this.p4label.AutoSize = true;
            this.p4label.Location = new System.Drawing.Point(644, 168);
            this.p4label.Name = "p4label";
            this.p4label.Size = new System.Drawing.Size(46, 13);
            this.p4label.TabIndex = 22;
            this.p4label.Text = "Param4:";
            // 
            // ImageAlgoTester
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(847, 431);
            this.Controls.Add(this.param4tb);
            this.Controls.Add(this.p4label);
            this.Controls.Add(this.label2black);
            this.Controls.Add(this.label1black);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.param3tb);
            this.Controls.Add(this.p3label);
            this.Controls.Add(this.param2tb);
            this.Controls.Add(this.p2label);
            this.Controls.Add(this.param1tb);
            this.Controls.Add(this.p1label);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.image2tb);
            this.Controls.Add(this.image1tb);
            this.Controls.Add(this.loadImage2);
            this.Controls.Add(this.loadImage1);
            this.Controls.Add(this.pictureBox2_2);
            this.Controls.Add(this.pictureBox1_2);
            this.Controls.Add(this.pictureBox2_1);
            this.Controls.Add(this.pictureBox1_1);
            this.Name = "ImageAlgoTester";
            this.Text = "ImageAlgoTester";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1_1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2_1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1_2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2_2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1_1;
        private System.Windows.Forms.PictureBox pictureBox2_1;
        private System.Windows.Forms.PictureBox pictureBox1_2;
        private System.Windows.Forms.PictureBox pictureBox2_2;
        private System.Windows.Forms.Button loadImage1;
        private System.Windows.Forms.Button loadImage2;
        private System.Windows.Forms.TextBox image1tb;
        private System.Windows.Forms.TextBox image2tb;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label p1label;
        private System.Windows.Forms.TextBox param1tb;
        private System.Windows.Forms.TextBox param2tb;
        private System.Windows.Forms.Label p2label;
        private System.Windows.Forms.TextBox param3tb;
        private System.Windows.Forms.Label p3label;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label1black;
        private System.Windows.Forms.Label label2black;
        private System.Windows.Forms.TextBox param4tb;
        private System.Windows.Forms.Label p4label;
    }
}