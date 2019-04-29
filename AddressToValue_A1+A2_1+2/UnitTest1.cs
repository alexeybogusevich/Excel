using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApplication;

namespace AddressToValue_A1_1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //arrange
            string expected = "1";
            Form1 testForm = new Form1();
            testForm.dictionary["A1"].Value = 1;
            testForm.dictionary["A2"].Exp = "A1";

            //act
            string actual = testForm.dictionary["A2"].AddressToValue(ref testForm.dictionary);

            //assert
            Assert.AreEqual(expected, actual);
        }
    }
}
