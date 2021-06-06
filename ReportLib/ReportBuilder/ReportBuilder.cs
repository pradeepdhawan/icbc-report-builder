using OfficeOpenXml;
using ReportLib.ReportData;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


namespace ReportLib
{
    public class ReportBuilder : IReportBuilder
    {
        private List<ExcelWorksheet> worksheets;
        public IExcelTemplate ReportTemplate { get; set; }

        public IList<Report> ReportData { get; set; }

        public void Build()
        {
            worksheets = new List<ExcelWorksheet>();

            foreach (Report report in ReportData)
            {
                var worksheet = ReportTemplate.Read(report.Name);

                var cells = worksheet.Cells;
                var dictionary = cells
                    .GroupBy(c => new { c.Start.Row, c.Start.Column })
                    .ToDictionary(
                        rcg => new KeyValuePair<int, int>(rcg.Key.Row, rcg.Key.Column),
                        rcg => cells[rcg.Key.Row, rcg.Key.Column].Value == null ? String.Empty : cells[rcg.Key.Row, rcg.Key.Column].Value.ToString());

                

                foreach (ReportVal val in report.ReportVal)
                {
                    //find row that has value
                    //find column that has value
                    var matchingColumn = dictionary.Where(pair => pair.Value == val.ReportCol.ToString("000")).Select(pair => pair.Key);
                    //since we are matching column, column that comes first will be given preference
                    var maxc = matchingColumn.Max(kvp => kvp.Value);
                    var col = matchingColumn.Where(kvp => kvp.Value == maxc).Select(kvp => kvp.Value).First();
                    var matchingRow = dictionary.Where(pair => pair.Value== val.ReportRow.ToString("000")).Select(pair => pair.Key);
                    var maxr = matchingRow.Max(kvp => kvp.Key);
                    var row = matchingRow.Where(kvp => kvp.Key == maxr).Select(kvp => kvp.Key).First();
                    cells[row, col].Value = val.Val;
                    Console.WriteLine(cells[row, col].Address);
                    Console.WriteLine(cells[row, col].Value);
                    //set value
                }

                worksheets.Add(worksheet);
            }
        }

        public byte[] GetReport()
        {
            Build();


            using (var package = new ExcelPackage())
            {
                foreach (var ws in worksheets)
                {
                    package.Workbook.Worksheets.Add(ws.Name, ws);
                }
                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return excelData;
            }

            //package.Save();
        }

    }
}
