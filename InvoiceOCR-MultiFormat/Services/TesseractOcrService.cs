using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using Tesseract;
using InvoiceOCR_MultiFormat.OCR;

namespace InvoiceOCR_MultiFormat.Services
{
    public class TesseractOcrService
    {
        private readonly string _tessDataPath;

        public TesseractOcrService()
        {
            // Path to tessdata folder
            _tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

            if (!Directory.Exists(_tessDataPath))
            {
                throw new DirectoryNotFoundException($"Tessdata folder not found at: {_tessDataPath}");
            }
        }

        public async Task<string> ExtractTextFromImageAsync(string imagePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using (var engine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default))
                    {
                        // Set page segmentation mode for better invoice parsing
                        engine.DefaultPageSegMode = PageSegMode.Auto;

                        using (var img = Pix.LoadFromFile(imagePath))
                        {
                            using (var page = engine.Process(img))
                            {
                                return page.GetText();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Tesseract OCR failed: {ex.Message}", ex);
                }
            });
        }

        public Task<string> ExtractTextFromPdfAsync(string pdfPath)
        {
            throw new NotSupportedException("PDF support requires additional library. Please convert PDF to image first, or we can add PDF conversion.");
        }
    }
}