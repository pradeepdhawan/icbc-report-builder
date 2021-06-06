using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;


namespace ReportLib
{
    public class ExcelTemplate : IExcelTemplate
    {
        private ExcelWorkbook workbook { get; set; }
        public ExcelTemplate()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        public void Load(string filepath)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(filepath));
            workbook = package.Workbook;
        }

        public ExcelWorksheet Read(string sheetName)
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);
            if (worksheet == null)
            {
                throw new IndexOutOfRangeException("No sheet found with such name!");
            }
            return worksheet;
        }
    }

}
