using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class MyTest : TestBase
    {
        /// <summary>
        /// 
        /// </summary>
        static ExcelPackage _pck;

        /// <summary>
        /// 
        /// </summary>
        public MyTest()
        {
            _pck = new ExcelPackage($@"D:\Simple\庆元县(331126)低效用地调查成果1\表1-1 庆元县(331126)低效用地调查区基本信息清单.xlsx");
            InitBase();
        }
        [TestMethod]
        public void StartTest()
        {
            var w = _pck.Workbook;

            var sheets = w.Worksheets;
            if (sheets == null)
            {

            }
        }

    }
}
