using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AXIS_LOGGER
{
    static class Program
    {
        public static MAIN_FORM MAIN_FORM
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MAIN_FORM());
        }
    }
}
