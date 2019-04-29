using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApplication;
using System.Collections.Generic;

namespace ReccurenceCheck_Reccurence_True
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //arrange
            bool expected = true;

            Form1 testForm = new Form1();
            testForm.dictionary["A1"].IDependOn.Add("A2");
            testForm.dictionary["A2"].IDependOn.Add("A3");
            testForm.dictionary["A3"].IDependOn.Add("A1");

            //act
            bool actual = testForm.RecurrenceCheck(testForm.dictionary["A3"], testForm.dictionary["A3"]);

            //assert
            Assert.AreEqual(expected, actual);
        }
    }
}
