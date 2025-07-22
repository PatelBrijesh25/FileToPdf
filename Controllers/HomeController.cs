using Microsoft.AspNetCore.Mvc;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace FileToPdfWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
            GlobalFontSettings.FontResolver = CustomFontResolver.Instance;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ConvertToPdf(IFormFile uploadedFile)
        {
            if (uploadedFile == null || uploadedFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid file.";
                return View("Index");
            }

            string ext = Path.GetExtension(uploadedFile.FileName).ToLower();
            string uploads = Path.Combine(_env.WebRootPath, "uploads");
            string outputDir = Path.Combine(_env.WebRootPath, "output");

            Directory.CreateDirectory(uploads);
            Directory.CreateDirectory(outputDir);

            string uniqueName = Guid.NewGuid().ToString() + Path.GetExtension(uploadedFile.FileName);
            string inputPath = Path.Combine(uploads, uniqueName);
            string outputFileName = Path.GetFileNameWithoutExtension(uploadedFile.FileName) + ".pdf";
            string outputPath = Path.Combine(outputDir, outputFileName);

            try
            {
                using (var stream = new FileStream(inputPath, FileMode.Create))
                {
                    uploadedFile.CopyTo(stream);
                }

                switch (ext)
                {
                    case ".txt":
                        ConvertTxtToPdf(inputPath, outputPath);
                        break;
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                        ConvertImageToPdf(inputPath, outputPath);
                        break;
                    case ".docx":
                        ConvertDocxToPdf(inputPath, outputPath);
                        break;
                    case ".xlsx":
                        ConvertXlsxToPdf(inputPath, outputPath);
                        break;
                    default:
                        ViewBag.Message = "Unsupported file type.";
                        return View("Index");
                }

                ViewBag.DownloadLink = "/output/" + outputFileName;
                return View("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View("Index");
            }
            finally
            {
                // Schedule deletion of the uploaded file after 5 minutes regardless of outcome
                _ = Task.Run(async () =>
                {
                    await Task.Delay(TimeSpan.FromMinutes(1));
                    try
                    {
                        if (System.IO.File.Exists(inputPath))
                            System.IO.File.Delete(inputPath);
                    }
                    catch { }
                });

                // Schedule deletion of output file after 5 minutes
                _ = Task.Run(async () =>
                {
                    await Task.Delay(TimeSpan.FromMinutes(1));
                    try
                    {
                        if (System.IO.File.Exists(outputPath))
                            System.IO.File.Delete(outputPath);
                    }
                    catch { }
                });
            }
        }

        private void ConvertTxtToPdf(string inputPath, string outputPath)
        {
            // Ensure font resolver is set once
            if (GlobalFontSettings.FontResolver == null)
            {
                GlobalFontSettings.FontResolver = CustomFontResolver.Instance;
            }

            var pdf = new PdfDocument();
            var page = pdf.AddPage();
            var gfx = XGraphics.FromPdfPage(page);

            // Just pass font name and size
            var font = new XFont("Arial", 12);

            string text = System.IO.File.ReadAllText(inputPath);
            gfx.DrawString(text, font, XBrushes.Black,
                new XRect(40, 40, page.Width - 80, page.Height - 80),
                XStringFormats.TopLeft);

            pdf.Save(outputPath);
        }


        private void ConvertImageToPdf(string inputPath, string outputPath)
        {
            var pdf = new PdfDocument();
            var image = XImage.FromFile(inputPath);
            var page = pdf.AddPage();
            page.Width = image.PixelWidth;
            page.Height = image.PixelHeight;
            var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawImage(image, 0, 0);
            pdf.Save(outputPath);
        }

        private void ConvertDocxToPdf(string inputPath, string outputPath)
        {
            var wordApp = new Word.Application();
            var wordDoc = wordApp.Documents.Open(inputPath);
            wordDoc.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
            wordDoc.Close(false);
            wordApp.Quit();
        }

        private void ConvertXlsxToPdf(string inputPath, string outputPath)
        {
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(inputPath);
            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath);
            workbook.Close(false);
            excelApp.Quit();
        }
    }
}
