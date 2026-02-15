using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace InvoiceOCR_MultiFormat.OCR
{
    public class OcrSpaceService
    {
        private readonly string _apiKey = "K86721916188957";
        private readonly string _apiUrl = "https://api.ocr.space/parse/image";

        public async Task<List<OcrResult>> ExtractTextAsync(string filePath)
        {
            var results = new List<OcrResult>();

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(3);

                    using (var content = new MultipartFormDataContent())
                    {
                        byte[] fileBytes = File.ReadAllBytes(filePath);
                        var fileContent = new ByteArrayContent(fileBytes);
                        content.Add(fileContent, "file", Path.GetFileName(filePath));

                        content.Add(new StringContent(_apiKey), "apikey");
                        content.Add(new StringContent("2"), "OCREngine");
                        content.Add(new StringContent("true"), "isTable");

                        var response = await client.PostAsync(_apiUrl, content);
                        var jsonResult = await response.Content.ReadAsStringAsync();

                        var ocrResponse = JsonConvert.DeserializeObject<OcrApiResponse>(jsonResult);

                        if (ocrResponse?.ParsedResults != null && ocrResponse.ParsedResults.Count > 0)
                        {
                            results.Add(new OcrResult
                            {
                                Text = ocrResponse.ParsedResults[0].ParsedText,
                                Confidence = 90
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"OCR Error: {ex.Message}");
            }

            return results;
        }

        private class OcrApiResponse
        {
            public List<OcrParsedResult> ParsedResults { get; set; }
        }

        private class OcrParsedResult
        {
            public string ParsedText { get; set; }
        }
    }
}