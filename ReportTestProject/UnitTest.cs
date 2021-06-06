namespace ExcelTemplateTestProject
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using System.Xml.Serialization;
    using ReportLib;
    using Xunit;

    public class UnitTest
    {
        [Fact]
        public void TestReportData()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(ReportLib.ReportData.Reports));
            ReportLib.ReportData.Reports resultingMessage = (ReportLib.ReportData.Reports)serializer.Deserialize(new XmlTextReader(".\\TestData\\Test.xml"));
        }

        [Fact]
        public void TestReportTemplate()
        {
            ExcelTemplate excelReportTemplate = new ReportLib.ExcelTemplate();
            excelReportTemplate.Load(".\\TestData\\TestReport.xlsx");
            var worksheet = excelReportTemplate.Read("F 20.04");
        }


        [Fact]
        public void TestReportBuilder()
        {
            ExcelTemplate excelReportTemplate = new ReportLib.ExcelTemplate();
            excelReportTemplate.Load(".\\TestData\\TestReport.xlsx");
            var worksheet = excelReportTemplate.Read("F 20.04");

            XmlSerializer serializer = new XmlSerializer(typeof(ReportLib.ReportData.Reports));
            ReportLib.ReportData.Reports reportData = (ReportLib.ReportData.Reports)serializer.Deserialize(new XmlTextReader(".\\TestData\\Test.xml"));

            var rb = new ReportLib.ReportBuilder
            {
                ReportTemplate = excelReportTemplate,
                ReportData = reportData.Items,
            };
            var data = rb.GetReport();
            System.IO.File.WriteAllBytes(".\\TestData\\TestReportOutput.xlsx", data);
        }
    }
}
