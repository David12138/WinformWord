using System;
using System.Windows.Forms;

namespace WinFormsApp
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();//启动程序中可视样式
            Application.SetCompatibleTextRenderingDefault(false);//将属性设置为默认
            Application.Run(new Form1());//运行启动的窗体
        }
    }
}
