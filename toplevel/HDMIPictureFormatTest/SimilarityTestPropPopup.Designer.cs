namespace Test
{
    partial class SimilarityTestPropPopup
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.simThresPSNR = new System.Windows.Forms.TextBox();
            this.simThresMSSIM = new System.Windows.Forms.TextBox();
            this.blurThreshold = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.simPeriodInt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.simPeriodTotal = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.label1.Location = new System.Drawing.Point(3, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Similarity + Blurry Test";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.simThresPSNR);
            this.groupBox1.Controls.Add(this.simThresMSSIM);
            this.groupBox1.Controls.Add(this.blurThreshold);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.groupBox1.Location = new System.Drawing.Point(6, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(164, 112);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Threshold";
            // 
            // simThresPSNR
            // 
            this.simThresPSNR.Location = new System.Drawing.Point(53, 82);
            this.simThresPSNR.Name = "simThresPSNR";
            this.simThresPSNR.Size = new System.Drawing.Size(42, 20);
            this.simThresPSNR.TabIndex = 11;
            // 
            // simThresMSSIM
            // 
            this.simThresMSSIM.Location = new System.Drawing.Point(112, 82);
            this.simThresMSSIM.Name = "simThresMSSIM";
            this.simThresMSSIM.Size = new System.Drawing.Size(42, 20);
            this.simThresMSSIM.TabIndex = 10;
            // 
            // blurThreshold
            // 
            this.blurThreshold.Location = new System.Drawing.Point(112, 23);
            this.blurThreshold.Name = "blurThreshold";
            this.blurThreshold.Size = new System.Drawing.Size(42, 20);
            this.blurThreshold.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 51);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 26);
            this.label3.TabIndex = 8;
            this.label3.Text = "Similarity \r\n(PSNR, MSSIM):";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Blur : ";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.simPeriodInt);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.simPeriodTotal);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.groupBox2.Location = new System.Drawing.Point(6, 150);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(164, 80);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Similarity Test Period";
            // 
            // simPeriodInt
            // 
            this.simPeriodInt.Location = new System.Drawing.Point(112, 48);
            this.simPeriodInt.Name = "simPeriodInt";
            this.simPeriodInt.Size = new System.Drawing.Size(42, 20);
            this.simPeriodInt.TabIndex = 14;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 51);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Interval (s):";
            // 
            // simPeriodTotal
            // 
            this.simPeriodTotal.Location = new System.Drawing.Point(112, 17);
            this.simPeriodTotal.Name = "simPeriodTotal";
            this.simPeriodTotal.Size = new System.Drawing.Size(42, 20);
            this.simPeriodTotal.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Total time (s):";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.button1.Location = new System.Drawing.Point(118, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(52, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "[Reset]";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // SimilarityTestPropPopup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.Location = new System.Drawing.Point(20, 0);
            this.Name = "SimilarityTestPropPopup";
            this.Size = new System.Drawing.Size(182, 241);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.TextBox simThresPSNR;
        public System.Windows.Forms.TextBox simThresMSSIM;
        public System.Windows.Forms.TextBox blurThreshold;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.TextBox simPeriodInt;
        private System.Windows.Forms.Label label5;
        public System.Windows.Forms.TextBox simPeriodTotal;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.Button button1;
    }
}
