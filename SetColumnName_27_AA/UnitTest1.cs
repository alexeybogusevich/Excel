using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApplication;

namespace SetColumnName_27_AA
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //arrange
            Form1 testForm = new Form1();
            string expected = "AA";
            int i = 26;

            //act
            string actual = testForm.SetColumnName(i);

            //assert
            Assert.AreEqual(expected, actual);
        }
    }
}
