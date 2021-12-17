using System;
using System.Windows.Forms;

namespace WordKiller
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main(string[] str)
        {
            if (str.Length > 0 && str[0] == "FileAssociation" && FileAssociation.IsRunAsAdmin())
            {
                FileAssociation.Associate("WordKiller", null);
                System.Environment.Exit(0);
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new CustomInterface(new string[] { "FileAssociation" }));
        }
    }
}
