using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using ReportLib;

namespace ReportAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {

        private readonly ILogger<ReportController> _logger;

        public ReportController(ILogger<ReportController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IActionResult Get()
        {
            ExcelTemplate excelReportTemplate = new ReportLib.ExcelTemplate();
            excelReportTemplate.Load(".\\Inputs\\TestReport.xlsx");
            var worksheet = excelReportTemplate.Read("F 20.04");

            XmlSerializer serializer = new XmlSerializer(typeof(ReportLib.ReportData.Reports));
            ReportLib.ReportData.Reports reportData = (ReportLib.ReportData.Reports)serializer.Deserialize(new XmlTextReader(".\\Inputs\\Test.xml"));

            var rb = new ReportLib.ReportBuilder
            {
                ReportTemplate = excelReportTemplate,
                ReportData = reportData.Items,
            };
            var data = rb.GetReport();

            using (var package = new ExcelPackage())
            {
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = "MyWorkbook.xlsx";
                return File(data, contentType, fileName);
            }
        }
    }
}
