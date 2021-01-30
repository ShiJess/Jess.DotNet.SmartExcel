using Jess.SmartExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {


            SmartExcel excel = new SmartExcel();
            excel.CreateFile(@"C:\sjshiSourceCode\SmartExcel\ConsoleApp1\test.xls");
            excel.PrintGridLines = false;

            double height = 1.5;

            excel.SetMargin(MarginTypes.TopMargin, height);
            excel.SetMargin(MarginTypes.BottomMargin, height);
            excel.SetMargin(MarginTypes.LeftMargin, height);
            excel.SetMargin(MarginTypes.RightMargin, height);

            string font = "Arial";
            short fontsize = 12;
            excel.SetFont(font, fontsize, FontFormatting.Italic);

            excel.SetColumnWidth(1, 2, 10);
            byte b1 = 2, b2 = 12;
            short s3 = 18;
            excel.SetColumnWidth(b1, b2, s3);

            string header = "头";
            string footer = "角";
            excel.SetHeader(header);
            excel.SetFooter(footer);

            int row = 1, col = 1, cellformat = 0;
            object title = "没有使用任何EXCEL组件，直接写成了一个EXCEL文件，cool吧？！";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            col = 2;
            title = "abcd";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            excel.CloseFile();


        }
    }
}
