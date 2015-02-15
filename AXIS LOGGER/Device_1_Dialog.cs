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
    public partial class Device_1_Dialog : Form
    {
        public Device_1_Dialog()
        {
            InitializeComponent();
        }

        private void Device_1_Dialog_Load(object sender, EventArgs e)
        {
            webBrowser_Device1.Navigate(MAIN_FORM.IPnr1);
        }
    }
}
