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
        public enum UARTorLAN
        {
            UART,
            LAN
        }
        
        public struct Switcher
        {
            public Switcher(UARTorLAN i, string port, string ip)
            {
                IsUarTorLan = i;
                comPort = port;
                ipAddress = ip;
            } 

            public UARTorLAN IsUarTorLan;
            public string comPort;
            public string ipAddress;
        }


        public string lanValue = "";
        public string UARTValue = "";
        

        /// <summary>
        /// device name, isUART, comPort or IP
        /// </summary>
        public Dictionary<string, Switcher> switchDictionary;

        public DeviceSettingPopup()
        {
            InitializeComponent();
            
            switchDictionary = new Dictionary<string, Switcher>();

            switchDictionary.Add("TG45", new Switcher(UARTorLAN.UART,"COM1", "192.168.1.1"));
            switchDictionary.Add("TG59", new Switcher(UARTorLAN.UART, "COM1", "192.168.1.1"));
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            var radioButton = sender as RadioButton;
            //1. check current label
            //2. update value based on active label

            if (radioButton.Checked)
            {
                LANtextBox.ReadOnly = radioButton.Text == "UART";
                UpdateCurrentLabelValue();
            }
            
        }

        public void UpdateCurrentLabelValue()
        {
            lanValue = LANtextBox.Text;
            UARTValue = UARTtextBox.Text;

            var uartorlan = UARTradioButton.Checked ? UARTorLAN.UART : UARTorLAN.LAN;

            switchDictionary[label1.Text] = new Switcher(uartorlan, UARTValue, lanValue);
        }

        private void Popup_VisibleChanged(object sender, EventArgs e)
        {
            var control = sender as DeviceSettingPopup;
            
            if (control.Visible)
            {
                var value = switchDictionary[control.label1.Text];

                UARTradioButton.CheckedChanged -= radioButton_CheckedChanged;
                LANradioButton.CheckedChanged -= radioButton_CheckedChanged;

                UARTradioButton.Checked = value.IsUarTorLan == UARTorLAN.UART;
                LANradioButton.Checked = !UARTradioButton.Checked;
                LANtextBox.ReadOnly = !LANradioButton.Checked;

                UARTradioButton.CheckedChanged += radioButton_CheckedChanged;
                LANradioButton.CheckedChanged += radioButton_CheckedChanged;

                UARTtextBox.Text = value.comPort;
                LANtextBox.Text = value.ipAddress;
            }
              
        }
    }
}
