namespace Test
{
    partial class ProgressBar
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.total_time = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.refPictureBox = new System.Windows.Forms.PictureBox();
            this.camPictureBox = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.refPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.camPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(32, 420);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(905, 27);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 370);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(217, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Estimated time remaining (total):";
            // 
            // total_time
            // 
            this.total_time.AutoSize = true;
            this.total_time.Location = new System.Drawing.Point(243, 370);
            this.total_time.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.total_time.Name = "total_time";
            this.total_time.Size = new System.Drawing.Size(32, 15);
            this.total_time.TabIndex = 4;
            this.total_time.Text = ".....";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(437, 453);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 27);
            this.button1.TabIndex = 5;
            this.button1.Text = "Stop";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 390);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 15);
            this.label2.TabIndex = 6;
            // 
            // refPictureBox
            // 
            this.refPictureBox.Location = new System.Drawing.Point(32, 67);
            this.refPictureBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.refPictureBox.Name = "refPictureBox";
            this.refPictureBox.Size = new System.Drawing.Size(427, 277);
            this.refPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.refPictureBox.TabIndex = 7;
            this.refPictureBox.TabStop = false;
            // 
            // camPictureBox
            // 
            this.camPictureBox.Location = new System.Drawing.Point(511, 67);
            this.camPictureBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.camPictureBox.Name = "camPictureBox";
            this.camPictureBox.Size = new System.Drawing.Size(427, 277);
            this.camPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.camPictureBox.TabIndex = 8;
            this.camPictureBox.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(507, 45);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(131, 15);
            this.label3.TabIndex = 9;
            this.label3.Text = "Camera Livestream";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(28, 45);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(116, 15);
            this.label4.TabIndex = 10;
            this.label4.Text = "Reference image";
            // 
            // ProgressBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(980, 495);
            this.ControlBox = false;
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.camPictureBox);
            this.Controls.Add(this.refPictureBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.total_time);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "ProgressBar";
            this.Text = "Executing test..";
            ((System.ComponentModel.ISupportInitialize)(this.refPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.camPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label total_time;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.PictureBox refPictureBox;
        public System.Windows.Forms.PictureBox camPictureBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}