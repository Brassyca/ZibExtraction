using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Zibs
{
    namespace ZibExtraction
    {
        static class Program
        {
            /// <summary>
            /// The main entry point for the application.
            /// </summary>
            [STAThread]
            static void Main()
            {
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new ZibExtraction());
            }

            static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
            {
                MessageBox.Show((e.ExceptionObject as Exception).Message +
                    "\r\nHet programma wordt afgesloten", "Unhandled exception", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Application.Exit();
            }
            
        }
    }
}
