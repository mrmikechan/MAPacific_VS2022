using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;

namespace MAPacificReportUtility
{
    public static class Program
    {
        delegate void SimpleDelegate();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // we have to ensure that we run only one instance of
            // MAPacifReportyUtility app
            // in case of the second instance,
            // we will focus original one

            string appName = "Event_" + Assembly.GetExecutingAssembly().FullName.Replace(' ', '_');

            if (ProcessChecker.IsOnlyProcess(appName))
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);


                try
                {
                     Application.Run(new MAPacificReportUtility.forms.ReportUtilityMainForm());

                }
                catch (Exception e) { }
                finally
                {
                    UserSettings.Save();
                }
            }
        }
    }
}