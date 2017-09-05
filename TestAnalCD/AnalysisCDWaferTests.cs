using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Collections;

namespace AnalysisCDWafer.Tests
{
    [TestClass]
    public class AnalysisCDWaferTests
    {
        [TestMethod]
        public void TestM_Mean_Regular()
        {
            double expected = 2.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { 1.0, 2.0, 3.0 };
            double actual = fileAnaliser.Mean(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Mean_WithMinus()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { -1.0, 0.0, 1.0 };
            double actual = fileAnaliser.Mean(list);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestM_Mean_Int16MaxSizeListZero()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i=0;i<Int16.MaxValue ;i++)
            {
                list.Add(0);
            }
            double actual = fileAnaliser.Mean(list);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestM_Mean_Int16MaxSizeListOne()
        {
            double expected = 1.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i = 0; i < Int16.MaxValue; i++)
            {
                list.Add(1);
            }
            double actual = fileAnaliser.Mean(list);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestM_Sigma_Regular()
        {
            double expected = 1.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { 1.0, 2.0, 3.0 };
            double actual = fileAnaliser.Sigma(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Sigma_WithMinus()
        {
            double expected = 1.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { -1.0, 0.0, 1.0 };
            double actual = fileAnaliser.Sigma(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Sigma_Int16MaxSizeZero()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i = 0; i < Int16.MaxValue; i++)
            {
                list.Add(0);
            }
            double actual = fileAnaliser.Sigma(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Sigma_Int16MaxSizeOne()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i = 0; i < Int16.MaxValue; i++)
            {
                list.Add(1);
            }
            double actual = fileAnaliser.Sigma(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Range_Regular()
        {
            double expected = 2.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { 1.0, 2.0, 3.0 };
            double actual = fileAnaliser.Range(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Range_WithMinus()
        {
            double expected = 2.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>() { -1.0, 0.0, 1.0 };
            double actual = fileAnaliser.Range(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Range_Int16MaxSizeZero()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i = 0; i < Int16.MaxValue; i++)
            {
                list.Add(0);
            }
            double actual = fileAnaliser.Range(list);
            Assert.AreEqual(expected, actual);

        }

        [TestMethod]
        public void TestM_Range_Int16MaxSizeOne()
        {
            double expected = 0.0;
            string path = "files/20170821_162000_ABB-PRODUCT-C5L5IDV_E720002_ABB-L.msr";
            FileAnalyiser fileAnaliser = new FileAnalyiser(path);
            List<double> list = new List<double>();
            for (int i = 0; i < Int16.MaxValue; i++)
            {
                list.Add(1);
            }
            double actual = fileAnaliser.Range(list);
            Assert.AreEqual(expected, actual);

        }

    }
}
