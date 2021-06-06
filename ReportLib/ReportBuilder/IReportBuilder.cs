using OfficeOpenXml;
using ReportLib.ReportData;
using System.Collections.Generic;

namespace ReportLib
{
    public interface IReportBuilder
    {
        IList<Report> ReportData { get; set; }
        IExcelTemplate ReportTemplate { get; set; }

        byte[] GetReport();
    }
}