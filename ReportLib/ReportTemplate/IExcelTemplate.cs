using OfficeOpenXml;

namespace ReportLib
{
    public interface IExcelTemplate
    {
        void Load(string filepath);
        OfficeOpenXml.ExcelWorksheet Read(string sheetName);
    }
}