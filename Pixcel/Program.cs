using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

var source = args[0];
if (!File.Exists(source)) throw new FileNotFoundException($"Input image not found", source);

var sourceImage = CvInvoke.Imread(source);
var mat = new Mat();
CvInvoke.ResizeForFrame(sourceImage, mat, new Size(64, 64), Inter.Lanczos4, scaleDownOnly: true);
var image = mat.ToImage<Bgr, byte>();

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using var package = new ExcelPackage(new FileInfo(source.Replace(Path.GetExtension(source), ".xlsx")));
var worksheet = package.Workbook.Worksheets["Pixcel"] ?? package.Workbook.Worksheets.Add("Pixcel");

for (int y = 1; y <= image.Height; y++)
{
    worksheet.Row(y).Height = 27.5;

    for (int x = 1; x <= image.Width; x++)
    {
        worksheet.Column(x).Width = 5;
        worksheet.Cells[y, x].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[y, x].Style.Fill.BackgroundColor.SetColor(0, image.Data[y-1, x-1, 2], image.Data[y-1, x-1, 1], image.Data[y-1, x-1, 0]);
    }
}

package.Save();