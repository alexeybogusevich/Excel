using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelApplication;
using System.Collections.Generic;

namespace ExcelApplicationTests
{
    [TestClass]
    public class UnitTest
    {
        [TestMethod]
        public void ReccurenceCheck_Reccurence_True()
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

        [TestMethod]
        public void AddressToValue_A1_1()
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

        [TestMethod]
        public void SetCellName_00_A0()
        {
            //arrange
            string expected = "A0";

            Form1 testForm = new Form1();

            //act
            string actual = testForm.SetCellName(0, 0);

            //assert
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void ReccurenceSearch_A1A2A3()
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

        [TestMethod]
        public void SetColumnName_26_AA()
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
