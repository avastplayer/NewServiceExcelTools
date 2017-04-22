using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace NewServiceExcelTool
{
    public class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new NewServiceExcelTools.Form1());
        }

        public static _Workbook OpenWorkBook(Microsoft.Office.Interop.Excel.Application xls,string filename)
        {
            _Workbook book = xls.Workbooks.Open(filename);
            xls.Visible = false;//设置Excel后台运行
            xls.DisplayAlerts = false;//设置不显示确认修改提示
            return book;
        }

        public static void CloseWorkBook(Microsoft.Office.Interop.Excel.Application xls, _Workbook book,string filename)
        {
            book.Save();//保存
            book.Close(false);//关闭打开的表
            xls.Quit();//Excel程序退出
            //sheet,book,xls设置为null，防止内存泄露
            book = null;
            xls = null;
            GC.Collect();//系统回收资源
        }

        public static _Worksheet Sheet1(_Workbook book)
        {
            try
            {
                return (_Worksheet)book.Worksheets.get_Item(1);//获得第index个sheet，准备读取
            }
            catch (Exception ex)//不存在就退出
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        } 
    } 
}
