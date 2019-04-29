using System;
using ExcelApplication;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SetCellName_00_A0
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //arrange
            string expected = "A0";

            Form1 testForm = new Form1();

            //act
            string actual = testForm.SetCellName(0, 0);

            //assert
            Assert.AreEqual(expected, actual);
        }
    }
}
