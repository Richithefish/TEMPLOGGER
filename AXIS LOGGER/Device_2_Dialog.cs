using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AXIS_LOGGER
{
    public partial class Device_2_Dialog : Form
    {
        public Device_2_Dialog()
        {
            InitializeComponent();
        }

        private void Device_2_Dialog_Load(object sender, EventArgs e)
        {
            webBrowser_Device2.Navigate(MAIN_FORM.IPnr2);
        }
    }
}
