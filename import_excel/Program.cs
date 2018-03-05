using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace import_excel
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //决定哪个源文件才是启动项
            Application.Run(new Form1());
        }
    }
}
