using System;

namespace Jess.SmartExcel
{
    public class SmartExcelException : ApplicationException
    {
        public SmartExcelException() : base("请首先调用CreateFile方法!")
        {

        }
    }
}