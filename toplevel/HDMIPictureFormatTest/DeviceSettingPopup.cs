using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class DeviceSettingPopup : UserControl
    {
        public string lanValue = "";
        public string UARTValue = "";
        public bool isUART = true ;

        /// <summary>
        /// device name, isUART, UART or IP
        /// </summary>
        public Dictionary<string, KeyValuePair<bool, string>> switchDictionary;

        public DeviceSettingPopup()
        {
            InitializeComponent();

            isUART = UARTradioButton.Checked;

            switchDictionary = new Dictionary<string, KeyValuePair<bool, string>>();

            switchDictionary.Add("TG45", new KeyValuePair<bool, string>());
            switchDictionary.Add("TG59", new KeyValuePair<bool, string>());
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            UpdateValue().Hide();
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            var radioButton = sender as RadioButton;
            isUART = radioButton.Text == "UART" && radioButton.Checked;  
            LANtextBox.ReadOnly = !isUART;

            UpdateValue();
        }

        private Control UpdateValue()
        {
            lanValue = LANtextBox.Text;
            UARTValue = UARTtextBox.Text;  

            var value = isUART ? UARTValue : lanValue;

            var keyPair = new KeyValuePair<bool, string>(isUART, value);
            var dialog = ((Control)this).Parent;

            if (switchDictionary.ContainsKey(dialog.Text))
                switchDictionary[dialog.Text] = keyPair;

            return dialog;
        }
    }
}
