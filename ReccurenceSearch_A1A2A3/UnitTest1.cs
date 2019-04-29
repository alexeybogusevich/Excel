using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApplication;
using System.Collections.Generic;

namespace ReccurenceSearch_A1A2A3
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //arrange
            string expected = "A1A2A3";
            Form1 testForm = new Form1();
            testForm.dictionary["A1"].DependentOnMe.Add("A2");
            testForm.dictionary["A2"].DependentOnMe.Add("A3");
            testForm.dictionary["A3"].DependentOnMe.Add("A1");
            List<string> visited = new List<string>();

            //act
            testForm.ReccurenceSearch("A1", ref visited);
            string actual = "";
            foreach (string s in visited)
                actual += s;

            //assert
            Assert.AreEqual(expected, actual);
        }
    }
}
