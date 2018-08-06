using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.ComponentModel;

namespace EzyCal
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AppDomain.CurrentDomain.ProcessExit += new EventHandler (OnProcessExit);

            Boolean notRunning;

            // Check if already running
            using (Mutex mutex = new Mutex(true, "EzyCal", out notRunning))
            {
                if (notRunning)
                {
                    Application.Run(new EzyCal());
                }
                else
                {
                    MessageBox.Show("EzyCal already running!");
                    return;
                }
            }
        }

        static void OnProcessExit(object sender, EventArgs e)
        {
            try
            {
                foreach (Process proc in Process.GetProcessesByName("CtrComm"))
                {
                    proc.Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
