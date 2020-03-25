using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Kod_H_Excelden_Veri_Okuma
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
            Application.Run(new Form1());
        }
    }
}
