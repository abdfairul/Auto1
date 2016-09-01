namespace Test
{
    partial class DeviceSettingPopup
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
            this.UARTradioButton = new System.Windows.Forms.RadioButton();
            this.LANradioButton = new System.Windows.Forms.RadioButton();
            this.UARTtextBox = new System.Windows.Forms.TextBox();
            this.LANtextBox = new System.Windows.Forms.TextBox();
            this.saveButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "DevName";
            // 
            // UARTradioButton
            // 
            this.UARTradioButton.AutoSize = true;
            this.UARTradioButton.Checked = true;
            this.UARTradioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.UARTradioButton.Location = new System.Drawing.Point(7, 24);
            this.UARTradioButton.Name = "UARTradioButton";
            this.UARTradioButton.Size = new System.Drawing.Size(54, 17);
            this.UARTradioButton.TabIndex = 1;
            this.UARTradioButton.TabStop = true;
            this.UARTradioButton.Text = "UART";
            this.UARTradioButton.UseVisualStyleBackColor = true;
            this.UARTradioButton.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // LANradioButton
            // 
            this.LANradioButton.AutoSize = true;
            this.LANradioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.LANradioButton.Location = new System.Drawing.Point(7, 51);
            this.LANradioButton.Name = "LANradioButton";
            this.LANradioButton.Size = new System.Drawing.Size(45, 17);
            this.LANradioButton.TabIndex = 2;
            this.LANradioButton.TabStop = true;
            this.LANradioButton.Text = "LAN";
            this.LANradioButton.UseVisualStyleBackColor = true;
            this.LANradioButton.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // UARTtextBox
            // 
            this.UARTtextBox.Location = new System.Drawing.Point(69, 23);
            this.UARTtextBox.Name = "UARTtextBox";
            this.UARTtextBox.ReadOnly = true;
            this.UARTtextBox.Size = new System.Drawing.Size(100, 20);
            this.UARTtextBox.TabIndex = 3;
            this.UARTtextBox.Text = "COM1";
            // 
            // LANtextBox
            // 
            this.LANtextBox.Location = new System.Drawing.Point(69, 50);
            this.LANtextBox.Name = "LANtextBox";
            this.LANtextBox.ReadOnly = true;
            this.LANtextBox.Size = new System.Drawing.Size(100, 20);
            this.LANtextBox.TabIndex = 4;
            this.LANtextBox.Text = "192.168.1.1";
            // 
            // saveButton
            // 
            this.saveButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveButton.Location = new System.Drawing.Point(114, 76);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(55, 23);
            this.saveButton.TabIndex = 5;
            this.saveButton.Tag = "Save the setting and exit";
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // DeviceSettingPopup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.LANtextBox);
            this.Controls.Add(this.UARTtextBox);
            this.Controls.Add(this.LANradioButton);
            this.Controls.Add(this.UARTradioButton);
            this.Controls.Add(this.label1);
            this.Location = new System.Drawing.Point(20, 0);
            this.Name = "DeviceSettingPopup";
            this.Size = new System.Drawing.Size(177, 106);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton UARTradioButton;
        private System.Windows.Forms.RadioButton LANradioButton;
        private System.Windows.Forms.TextBox UARTtextBox;
        private System.Windows.Forms.TextBox LANtextBox;
        private System.Windows.Forms.Button saveButton;
    }
}
