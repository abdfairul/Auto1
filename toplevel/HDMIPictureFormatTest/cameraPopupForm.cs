using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    public partial class cameraPopupForm : Form
    {
        public delegate void PanelItem_ClickEvent(object sender, EventArgs e);
        PanelItem_ClickEvent m_event;

        public cameraPopupForm(PanelItem_ClickEvent i_event)
        {
            InitializeComponent();
            m_event = i_event;
        }

        private void Options_MouseClick(object sender, MouseEventArgs e)
        {
            panel1.Visible = !panel1.Visible;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            panel1.Hide();
        }

        private void cameraPopupForm_Click(object sender, EventArgs e)
        {
            panel1.Hide();
        }

        private void panelItem_CheckedChanged(object sender, EventArgs e)
        {
            m_event(sender, e);
        }
    }
}
