using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace ExcelCheck
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Boolean createdNew;
            System.Threading.Mutex instance = new System.Threading.Mutex(true, "ArresterSerialPort", out createdNew);
            if (createdNew)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                string path = Application.StartupPath + "/";
                string dllFileName = "主数据收集模板配置.xlsx";
                if (!File.Exists(path + dllFileName))
                {
                    FileStream fs = new FileStream(path + dllFileName, FileMode.CreateNew, FileAccess.Write);
                    byte[] buffer = ExcelCheck.Properties.Resources.主数据收集模板配置;
                    fs.Write(buffer, 0, buffer.Length);
                    fs.Close();
                }
                Application.Run(new ExcelEdit());
            }
            else
            {
                MessageBox.Show("已经启动了一个程序，请先退出", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
        }
    }
}
