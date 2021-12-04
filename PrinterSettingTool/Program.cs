using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace PrinterSettingTool
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>

        [STAThread]
        static void Main()
        {
            if (!ProcessExists())
            {
                Application.Run(new Form1());
            }
        }


        public static bool ProcessExists()
        {
            Process thisProcess = Process.GetCurrentProcess();
            Process[] Processes = Process.GetProcessesByName(thisProcess.ProcessName);
            int thisProcessId = thisProcess.Id;

            foreach (Process aProcess in Processes)
            {
                if (aProcess.Id != thisProcessId)
                {
                    NativeMethods.ShowWindow(aProcess.MainWindowHandle, (int)SW.SHOWNORMAL);
                    NativeMethods.SetForegroundWindow(aProcess.MainWindowHandle);
                    return true;
                }
            }
            return false;
        }
    }
}