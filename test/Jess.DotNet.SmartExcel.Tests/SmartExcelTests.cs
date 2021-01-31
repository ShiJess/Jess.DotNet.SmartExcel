using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;

namespace Jess.DotNet.SmartExcel.Tests
{
    public class SmartExcelTests
    {
        [Fact]
        public void Test()
        {
            SmartExcel excel = new SmartExcel();
            //SmartExcel excel = new SmartExcel(Encoding.GetEncoding(936));

            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xls");
            excel.CreateFile(path);
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
            string footer = "脚";
            excel.SetHeader(header);
            excel.SetFooter(footer);

            int row = 1, col = 1, cellformat = 0;
            object title = "生成Biff2格式Excel";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            col = 2;
            title = "abcd";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            excel.CloseFile();
        }
    }
}
