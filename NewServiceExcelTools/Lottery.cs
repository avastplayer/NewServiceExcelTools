using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewServiceExcelTools
{
    class Lottery : NewServiceExcelTool.Program
    {
        private string zgDate;
        private string filename;//文件夹名
        private int xiangou_avtivity_id;//限购活动id
        private int choujiang_avtivity_id;//新服抽奖活动id

        public string Filename { get => filename; set => filename = value; }
        public int Xiangou_avtivity_id {
            get => xiangou_avtivity_id;
            set
            {
                if (value >= 166001 && value <= 166004)
                    xiangou_avtivity_id = value;
            }
                
        }
        public int Choujiang_avtivity_id { get => choujiang_avtivity_id; set => choujiang_avtivity_id = value; }
        public string ZgDate { get => zgDate; set => zgDate = value; }

        public Lottery()//构造函数
        {
            Filename = @"E:\新服及抽奖工具\data";
            Xiangou_avtivity_id = 166001;
            choujiang_avtivity_id = 152001;
            int todayWeek = (int) DateTime.Now.DayOfWeek;//获取配置时间是周几
            if (todayWeek < 5)
            {
                ZgDate = DateTime.Now.AddDays(5-todayWeek).ToShortDateString();//同一周配置需后延5-X天
            }
            else
            {
                ZgDate = DateTime.Now.AddDays(12- todayWeek).ToShortDateString();//前一周配置需后延12-X天
            }
        }

        static void DingshiHuodong()
        {
            Lottery lottery = new Lottery();
            string filename = lottery.Filename + @"d定时活动配置表.xlsx";

            Application xls = new Application(); ;
            _Workbook book = OpenWorkBook(xls, filename);
            _Worksheet sheet = Sheet1(book);

            int row = sheet.UsedRange.Row;//最后一行
            int xiangou_avtivity_row = sheet.Range[sheet.Cells[2,1],sheet.Cells[row,2]].Find(lottery.Xiangou_avtivity_id).Row;//限购最后一行
            int choujiang_avtivity_row = sheet.Range[sheet.Cells[2, 1], sheet.Cells[row, 2]].Find(lottery.Choujiang_avtivity_id).Row;//新服抽奖最后一行

            switch (lottery.Xiangou_avtivity_id)//更改限购开启时间
            {
                case 166001:
                    sheet.Cells[xiangou_avtivity_row-3, 6] = lottery.ZgDate + " 09：40：00";
                    break;
                case 166002:
                    sheet.Cells[xiangou_avtivity_row-2, 6] = lottery.ZgDate + " 09：40：00";
                    break;
                case 166003:
                    sheet.Cells[xiangou_avtivity_row-1, 6] = lottery.ZgDate + " 09：40：00";
                    break;
                case 166004:
                    sheet.Cells[xiangou_avtivity_row, 6] = lottery.ZgDate + " 09：40：00";
                    break;
            }
            sheet.Range[]



        }
        static void ShixiaoxingDaoju()
        {
        }
        static void ZahuoBiao()
        {
        }
        static void JiangliBiao()
        {
        }
        static void HuodongDingshiNPC()
        {
        }
        static void LibaoGongnengPeizhibiao()
        {
        }
        static void WupinLeixingku()
        {
        }
        static void YuanbaoShangchengPeizhibiao()
        {
        }
        static void ChengxuneiZifuchuan()
        {
        }
        static void KehuduanTishi()
        {
        }
        static void NPCFuwuPeizhiFujiabiao()
        {
        }
        static void NPCFuwuZongbiao()
        {
        }
        static void ShuangjieFudaiPeizhibiao()
        {
        }
        static void XianhuabangChengweiPeizhibiao()
        {
        }
        static void TongyongDuihuanFuwuPeizhibiao()
        {
        }
        static void NPCMaimaiYuanbaoWupinbiao()
        {
        }
        static void NPCMaimaiWupinbiao()
        {
        }
        static void ChongzhisongYuanbaopiaoDaojuPeizhi()
        {
        }
        static void JiehunXianhuadengDingshiHuodong()
        {
        }
        static void ShuangjieFudaiPeizhiFujiabiao()
        {
        }
        static void XinfuChoujiangPeizhi()
        {
        }
        static void DingshiHuodongPeizhiFujiabiao()
        {
        }
        static void GuoqiKeduihuandeShixiaoxingDaoju()
        {
        }
    }
}
