using System;
using ExcelMerge.Provider;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var dir =@"F:\work\note\2018\04";
            var dp = new DataProvider();
            var list = dp.GetList(dir);
            var book = dp.Merge(list);
            dp.Export(book,$"{dir}\\test001.xlsx");
        }
    }
}
