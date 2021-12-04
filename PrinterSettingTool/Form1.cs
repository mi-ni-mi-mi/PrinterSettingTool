using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;

namespace App
{
    public partial class Form1 : Form
    {
        private Form form1 { get; set; }

        public Form1()
        {
            InitializeComponent();
            WindowState = FormWindowState.Minimized;

        }
        protected override CreateParams CreateParams
        {//タスクマネージャー非表示
            [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle = 0x80; //WS_EX_TOOLWINDOW
                return cp;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            form1 = this;
            UpdatePrinterList();
            listBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {//プリンタ追加
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments = "printui.dll PrintUIEntry /il";
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {//プリンタのプロパティ
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments = "printui.dll PrintUIEntry /p /n " + listBox1.SelectedItems[0].ToString();
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {//プリンタ削除
            //まず、プリンタの情報を取得する
            PRINTER_INFO_2 pinfo = GetPrinterInfo(listBox1.SelectedItems[0].ToString());
            MessageBox.Show(pinfo.pPortName);
            //次に、プリンタのIPアドレスを取得する
            string ipAddress;
            GetPrinterIPAddress(pinfo.pPortName, out ipAddress);
            MessageBox.Show(ipAddress);
            //次に、プリンタポートを削除する
            Process proc1 = new Process();
            proc1.StartInfo.FileName = @"cscript";
            proc1.StartInfo.WorkingDirectory = @"C:\Windows\System32\Printing_Admin_Scripts\ja-JP\";
            proc1.StartInfo.Arguments = "prnport.vbs -d -h ";
            proc1.StartInfo.Verb = "RunAs";
            proc1.Start();
            proc1.WaitForExit();
            MessageBox.Show(proc1.ExitCode.ToString());
            proc1.Close();
            //最後に、プリンタを削除する
            Process proc2 = new Process();
            proc2.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc2.StartInfo.Arguments = "printui.dll PrintUIEntry /dl /n " + listBox1.SelectedItems[0].ToString();
            proc2.StartInfo.Verb = "RunAs";
            proc2.Start();
            proc2.WaitForExit();
            proc2.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {//通常使うプリンタ
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments = "printui.dll PrintUIEntry /y /n " + listBox1.SelectedItems[0].ToString();
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {//プリンタ一覧更新
            UpdatePrinterList();
        }

        private void button6_Click(object sender, EventArgs e)
        {//終了
            form1.WindowState = FormWindowState.Minimized;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
                e.Cancel = true;
            form1.WindowState = FormWindowState.Minimized;
        }




        /// <summary>
        /// 以下は関数です。
        /// </summary>
        public void UpdatePrinterList()
        {
            listBox1.Items.Clear();
            ManagementObjectSearcher mos = new ManagementObjectSearcher("Select * from Win32_Printer");
            ManagementObjectCollection moc = mos.Get();
            foreach (ManagementObject mo in moc)
            {
                listBox1.Items.Add(mo["Name"]);
                if ((((uint)mo["Attributes"]) & 4) == 4) //ATTRIBUTE_DEFAULT
                {
                    label2.Text = "通常使うプリンタ：" + mo["Name"];
                }
                mo.Dispose();
            }
            moc.Dispose();
            mos.Dispose();
        }

        public static PRINTER_INFO_2 GetPrinterInfo(string printerName)
        {
            IntPtr hPrinter;
            if (!NativeMethods.OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            IntPtr pPrinterInfo = IntPtr.Zero;
            try
            {
                int needed;
                NativeMethods.GetPrinter(hPrinter, 2, IntPtr.Zero, 0, out needed);
                if (needed <= 0) throw new Exception("失敗しました。");
                pPrinterInfo = Marshal.AllocHGlobal(needed);

                int temp;
                if (!NativeMethods.GetPrinter(hPrinter, 2, pPrinterInfo, needed, out temp))
                    throw new Win32Exception(Marshal.GetLastWin32Error());

                PRINTER_INFO_2 printerInfo = (PRINTER_INFO_2)Marshal.PtrToStructure(pPrinterInfo, typeof(PRINTER_INFO_2));
                return printerInfo;
            }
            finally
            {
                NativeMethods.ClosePrinter(hPrinter);
                Marshal.FreeHGlobal(pPrinterInfo);
            }
        }

        private void GetPrinterIPAddress(String pportName, out String pipAddress)
        {
            pipAddress = "";
            string query = string.Format("SELECT * FROM Win32_TCPIPPrinterPort WHERE Name = '{0}'", pportName);
            ManagementObjectSearcher mos = new ManagementObjectSearcher(query);
            ManagementObjectCollection moc = mos.Get();

            foreach (ManagementObject mo in moc)
            {
                pipAddress = mo["HostAddress"].ToString();
                mo.Dispose();
                break;
            }

            moc.Dispose();
            mos.Dispose();
        }
    }
}
