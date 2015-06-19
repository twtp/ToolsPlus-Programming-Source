using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PriceTwoSpyService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static int Main(string[] args)
        {
            if (args.GetLength(0) == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new guiFrm(0));
            }
            else
            {
                if (args[0] == "/scheduledtask")
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new guiFrm(1));
                }
                else if (args[0] == "/service")
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new guiFrm(2));
                }
                else
                {
                    MessageBox.Show("Unknown command line parameter. Starting under Manual Mode.");
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new guiFrm(0));
                }
            }

            return 0;
        }
    }
}
