using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.IO;


namespace TimeTracker.Helper
{
    public static class ExcelWriter
    {
        public static void WriteToExcel(string filePath, InfoToExcel infoToExcels)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(filePath);
            bool ExistFile = fileInfo.Exists;

            //  excel
            using (var package = ExistFile ? new ExcelPackage(fileInfo) : new ExcelPackage())
            {
                // add worksheet to workbook if it empty. If not get the first
                var worksheet = ExistFile ? package.Workbook.Worksheets[0] : package.Workbook.Worksheets.Add("TimeTracker");

                // fill titel if it emtpy
                if (worksheet.Cells[1, 1].Value == null)
                {
                    worksheet.Cells[1, 1].Value = "ProjektName";
                    worksheet.Cells[1, 2].Value = "Data";
                    worksheet.Cells[1, 3].Value = "Time";
                    worksheet.Cells[1, 4].Value = "Description";
                }

                // find first empty row
                int nextRow = worksheet.Dimension?.End.Row + 1 ?? 2;

                // fill data
                worksheet.Cells[nextRow, 1].Value = infoToExcels.ProjectName;
                worksheet.Cells[nextRow, 2].Value = infoToExcels.Data;
                worksheet.Cells[nextRow, 3].Value = infoToExcels.Time;
                worksheet.Cells[nextRow, 4].Value = infoToExcels.Description;

                worksheet.Cells.AutoFitColumns(0);//Autofit columns for all cells

                // Set some extended property values
                //package.Workbook.Properties.Company = "Davydov GmbH";

                // save file
                if (!ExistFile)
                {
                    package.SaveAs(fileInfo); // create file if not exist
                }
                else
                {
                    package.Save(); // update file if exist
                }
            }
        }
    }
}
