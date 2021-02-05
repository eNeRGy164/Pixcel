using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

var source = args[0];

var bitmap = new Bitmap(64, 64);
using var graphics = Graphics.FromImage(bitmap);
graphics.DrawImage(new Bitmap(source), 0, 0, 64, 64);

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using var package = new ExcelPackage(new FileInfo(source.Replace(Path.GetExtension(source), ".xlsx")));
var worksheet = package.Workbook.Worksheets["Pixcel"] ?? package.Workbook.Worksheets.Add("Pixcel");

for (int y = 1; y <= bitmap.Width; y++)
{
    worksheet.Column(y).Width = 5;

    for (int x = 1; x <= bitmap.Height; x++)
    {
        worksheet.Row(x).Height = 27.5;
        worksheet.Cells[x, y].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[x, y].Style.Fill.BackgroundColor.SetColor(bitmap.GetPixel(y - 1, x - 1));
    }
}

package.Save();