using System;

namespace Jess.DotNet.SmartExcel
{
    public class SmartExcelException : ApplicationException
    {
        public SmartExcelException() : base("请首先调用CreateFile方法!")
        {

        }
    }
}