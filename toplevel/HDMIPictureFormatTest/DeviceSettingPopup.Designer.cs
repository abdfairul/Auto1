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
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 5);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "DevName";
            // 
            // UARTradioButton
            // 
            this.UARTradioButton.AutoSize = true;
            this.UARTradioButton.Checked = true;
            this.UARTradioButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.UARTradioButton.Location = new System.Drawing.Point(9, 28);
            this.UARTradioButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.UARTradioButton.Name = "UARTradioButton";
            this.UARTradioButton.Size = new System.Drawing.Size(64, 19);
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
            this.LANradioButton.Location = new System.Drawing.Point(9, 59);
            this.LANradioButton.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.LANradioButton.Name = "LANradioButton";
            this.LANradioButton.Size = new System.Drawing.Size(53, 19);
            this.LANradioButton.TabIndex = 2;
            this.LANradioButton.TabStop = true;
            this.LANradioButton.Text = "LAN";
            this.LANradioButton.UseVisualStyleBackColor = true;
            this.LANradioButton.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // UARTtextBox
            // 
            this.UARTtextBox.Location = new System.Drawing.Point(92, 27);
            this.UARTtextBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.UARTtextBox.Name = "UARTtextBox";
            this.UARTtextBox.ReadOnly = true;
            this.UARTtextBox.Size = new System.Drawing.Size(132, 25);
            this.UARTtextBox.TabIndex = 3;
            this.UARTtextBox.Text = "COM1";
            // 
            // LANtextBox
            // 
            this.LANtextBox.Location = new System.Drawing.Point(92, 58);
            this.LANtextBox.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.LANtextBox.Name = "LANtextBox";
            this.LANtextBox.ReadOnly = true;
            this.LANtextBox.Size = new System.Drawing.Size(132, 25);
            this.LANtextBox.TabIndex = 4;
            this.LANtextBox.Text = "192.168.1.1";
            // 
            // DeviceSettingPopup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.LANtextBox);
            this.Controls.Add(this.UARTtextBox);
            this.Controls.Add(this.LANradioButton);
            this.Controls.Add(this.UARTradioButton);
            this.Controls.Add(this.label1);
            this.Location = new System.Drawing.Point(20, 0);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "DeviceSettingPopup";
            this.Size = new System.Drawing.Size(236, 96);
            this.VisibleChanged += new System.EventHandler(this.Popup_VisibleChanged);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton UARTradioButton;
        private System.Windows.Forms.RadioButton LANradioButton;
        private System.Windows.Forms.TextBox UARTtextBox;
        private System.Windows.Forms.TextBox LANtextBox;
    }
}
