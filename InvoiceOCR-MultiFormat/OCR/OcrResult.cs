namespace InvoiceOCR_MultiFormat.OCR
{
    public class OcrResult
    {
        public string Text { get; set; } = string.Empty;
        public double Confidence { get; set; }
    }

    public class InvoiceLineItem
    {
        public string SrNo { get; set; } = string.Empty;
        public string ItemCode { get; set; } = string.Empty;
        public string ItemDescription { get; set; } = string.Empty;
        public string UOM { get; set; } = string.Empty;
        public string Quantity { get; set; } = string.Empty;
        public string UnitRate { get; set; } = string.Empty;
        public string TotalExclVAT { get; set; } = string.Empty;
        public string VATPercent { get; set; } = string.Empty;
        public string VATAmount { get; set; } = string.Empty;
        public string TotalInclVAT { get; set; } = string.Empty;
    }
}