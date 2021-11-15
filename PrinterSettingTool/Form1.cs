using System;
using System.Management;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace PrinterSetting
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            updatePrinteList();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        public void updatePrinteList()
        {
            listBox1.Items.Clear();
            // Win32_OperatingSystemクラスを作成する
            using (ManagementClass managementClass = new ManagementClass("Win32_OperatingSystem"))
            {
                managementClass.Get(); // Win32_OperatingSystemオブジェクトを取得
                managementClass.Scope.Options.EnablePrivileges = true; // 権限を有効化

                // WMIのオブジェクトのコレクションを取得する
                using (ManagementObjectCollection managementObjectCollection
                    = managementClass.GetInstances())
                {
                    foreach (ManagementObject managementObject in managementObjectCollection)
                    {
                        System.Management.ManagementObjectSearcher mos =
                            new System.Management.ManagementObjectSearcher(
                                "Select * from Win32_Printer");
                        System.Management.ManagementObjectCollection moc =
                            mos.Get();

                        //プリンタを列挙
                        foreach (System.Management.ManagementObject mo in moc)
                        {
                            listBox1.Items.Add(mo["Name"]);
                            Console.WriteLine(mo["Name"]);
                            mo.Dispose();
                        }
                        mos.Dispose();
                        moc.Dispose();
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // プリンタ追加
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments = "printui.dll PrintUIEntry /il";
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // プリンタ削除
            for (int i = 0; i < listBox1.SelectedItems.Count; i++)
            {
                Console.WriteLine(listBox1.SelectedItems[i].ToString());
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            // 通常使うプリンタに設定
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments =
                "printui.dll PrintUIEntry /y /n "
                + listBox1.SelectedItems[0].ToString();
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            // プロパティ設定
            Process proc = new Process();
            proc.StartInfo.FileName = @"C:\Windows\System32\rundll32.exe";
            proc.StartInfo.Arguments =
                "printui.dll PrintUIEntry /p /n "
                + listBox1.SelectedItems[0].ToString();
            proc.StartInfo.Verb = "RunAs";
            proc.Start();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            // プリンタ一覧更新
            updatePrinteList();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            // 終了
            Application.Exit();
        }
    }
}
