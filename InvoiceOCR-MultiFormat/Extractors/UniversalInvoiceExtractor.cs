//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text.RegularExpressions;
//using InvoiceOCR_MultiFormat.OCR;
//using InvoiceOCR_MultiFormat.Services;

//namespace InvoiceOCR_MultiFormat.Extractors
//{
//    public class UniversalInvoiceExtractor
//    {
//        private readonly MultiStrategyTableExtractor _tableExtractor;

//        public UniversalInvoiceExtractor()
//        {
//            _tableExtractor = new MultiStrategyTableExtractor();
//        }

//        public string ExtractCompanyName(string text)
//        {
//            // Look for GF Corys
//            var match = Regex.Match(text, @"GF\s+Corys\s+Piping\s+Systems\s+LLC\s*-?\s*Dubai", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return "GF Corys Piping Systems LLC - Dubai";

//            // Look for ZAKER TRADING L.L.C.
//            match = Regex.Match(text, @"ZAKER\s+TRADING\s+L\.L\.C\.", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return "ZAKER TRADING L.L.C.";

//            // Look for Techno King
//            match = Regex.Match(text, @"Techno\s+King\s+Trading\s+Co\.\s*LLC", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return "Techno King Trading Co. LLC";

//            return "N/A";
//        }

//        public string ExtractInvoiceNumber(string text)
//        {
//            // GF Corys format: "Invoice No. Date 261200791"
//            var match = Regex.Match(text, @"Invoice\s+No\.\s+Date\s+(\d{8,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // ZAKER format: *299355*
//            match = Regex.Match(text, @"\*(\d{6})\*");
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for "Inv No. : 299355"
//            match = Regex.Match(text, @"Inv\s+No\.\s*[:]\s*(\d{5,8})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Generic pattern
//            match = Regex.Match(text, @"(?:Invoice|Inv\.?)\s*(?:No\.?|Number)\s*[:.\s]*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//            {
//                string number = match.Groups[1].Value;
//                if (!number.StartsWith("100") && !number.StartsWith("971"))
//                    return number;
//            }

//            return "N/A";
//        }

//        public string ExtractDate(string text)
//        {
//            // Look for "Date : 09/02/2026"
//            var match = Regex.Match(text, @"Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Generic DD/MM/YYYY
//            match = Regex.Match(text, @"\b(\d{2}/\d{2}/\d{4})\b");
//            if (match.Success)
//                return match.Groups[1].Value;

//            // DD-MMM-YYYY format (GF Corys)
//            match = Regex.Match(text, @"\b(\d{1,2}[-](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[-]\d{4})\b", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value.ToUpper();

//            return "N/A";
//        }

//        public string ExtractTRN(string text)
//        {
//            // Look for TRN: 100339834200003
//            var match = Regex.Match(text, @"TRN\s*(?:NO|Number)?[:\s]*(\d{15})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Generic 15-digit TRN starting with 100
//            match = Regex.Match(text, @"\b(100\d{12,15})\b");
//            return match.Success ? match.Groups[1].Value : "N/A";
//        }

//        public string ExtractSalesPerson(string text)
//        {
//            // Look for "Salesman : Muhammed"
//            var match = Regex.Match(text, @"Salesman\s*[:]\s*([A-Za-z]+)", RegexOptions.IgnoreCase);
//            if (match.Success)
//            {
//                string name = match.Groups[1].Value.Trim();
//                // Filter out non-name words
//                if (!Regex.IsMatch(name, @"Currency|UAE|Dirham|Date|Terms|Total|Rate", RegexOptions.IgnoreCase))
//                    return name;
//            }

//            // Look for "Sales Person:"
//            match = Regex.Match(text, @"Sales\s*Person\s*[:]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value.Trim();

//            // GF Corys format (BIJU V. PILLAI)
//            match = Regex.Match(text, @"\b([A-Z]{3,}\s+[A-Z]\.\s+[A-Z]{3,})\b");
//            if (match.Success)
//            {
//                string name = match.Groups[1].Value;
//                if (!name.Contains("LLC") && !name.Contains("BOX"))
//                    return name;
//            }

//            return "N/A";
//        }

//        public string ExtractPaymentTerms(string text)
//        {
//            // Look for "90 days PDC on delivery"
//            var match = Regex.Match(text, @"(\d+\s*days?\s*PDC\s*on\s*delivery)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for just "90 Days"
//            match = Regex.Match(text, @"(\d+\s*days?)", RegexOptions.IgnoreCase);
//            return match.Success ? match.Groups[1].Value : "N/A";
//        }

//        public string ExtractShipDate(string text)
//        {
//            // Look for "Supply Date : 09/02/2026"
//            var match = Regex.Match(text, @"Supply\s*Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for "Ship Date"
//            match = Regex.Match(text, @"Ship\s*Date\s*[:]\s*(\d{1,2}[-]\w{3}[-]\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Fallback to invoice date
//            return ExtractDate(text);
//        }

//        public string ExtractDONumber(string text)
//        {
//            // Look for "D.O. Number" in GF Corys
//            var match = Regex.Match(text, @"D\.O\.\s*Number\s+(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for "DO No. Code" followed by number (ZAKER format)
//            match = Regex.Match(text, @"DO\s+No\.\s+Code[^\n]*?(\d{6})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for standalone 181388 or similar
//            match = Regex.Match(text, @"\b(181388|599722)\b");
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Generic DO number pattern
//            match = Regex.Match(text, @"DO\s*No[.:]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractSONumber(string text)
//        {
//            // Look for "S. O. Number" in GF Corys
//            var match = Regex.Match(text, @"S\.\s*O\.\s*Number\s+(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for PO# : PO260016
//            match = Regex.Match(text, @"PO#\s*[:]\s*(PO\d{6})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Look for P.O. Number
//            match = Regex.Match(text, @"P\.O\.\s*Number\s+(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public List<InvoiceLineItem> ExtractLineItems(string text)
//        {
//            return _tableExtractor.ExtractLineItems(text);
//        }
//    }
//}


//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text.RegularExpressions;
//using InvoiceOCR_MultiFormat.OCR;
//using InvoiceOCR_MultiFormat.Services;

//namespace InvoiceOCR_MultiFormat.Extractors
//{
//    public class UniversalInvoiceExtractor
//    {
//        private readonly MultiStrategyTableExtractor _tableExtractor;

//        public UniversalInvoiceExtractor()
//        {
//            _tableExtractor = new MultiStrategyTableExtractor();
//        }

//        public string ExtractCompanyName(string text)
//        {
//            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
//                .Select(l => l.Trim())
//                .Where(l => l.Length > 0)
//                .Take(20)
//                .ToList();

//            // Strategy 1: Look for known companies
//            var knownCompanies = new[]
//            {
//                @"Zylker",
//                @"GF\s+Corys\s+Piping\s+Systems\s+LLC",
//                @"ZAKER\s+TRADING\s+L\.L\.C\.",
//                @"Techno\s+King\s+Trading\s+Co\.\s*LLC"
//            };

//            foreach (var pattern in knownCompanies)
//            {
//                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
//                if (match.Success)
//                {
//                    if (pattern.Contains("Zylker"))
//                        return "Zylker";
//                    if (pattern.Contains("GF"))
//                        return "GF Corys Piping Systems LLC - Dubai";
//                    if (pattern.Contains("ZAKER"))
//                        return "ZAKER TRADING L.L.C.";
//                    if (pattern.Contains("Techno"))
//                        return "Techno King Trading Co. LLC";
//                }
//            }

//            // Strategy 2: Look for company keywords in first few lines
//            foreach (var line in lines)
//            {
//                // Skip common invoice headers
//                if (Regex.IsMatch(line, @"^(Invoice|Bill\s+To|Date|#|Terms)", RegexOptions.IgnoreCase))
//                    continue;

//                // Look for lines with company indicators
//                if (Regex.IsMatch(line, @"\b(LLC|Ltd|Limited|Inc|Corporation|Corp|Co\.|Trading|Company)\b", RegexOptions.IgnoreCase))
//                {
//                    string cleaned = Regex.Replace(line, @"^\d+\s*", ""); // Remove leading numbers
//                    if (cleaned.Length >= 5 && cleaned.Length <= 100 && !cleaned.Contains("@"))
//                        return cleaned;
//                }

//                // Look for capitalized company names (2-4 words)
//                var capitalMatch = Regex.Match(line, @"^([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})$");
//                if (capitalMatch.Success && line.Length >= 5 && line.Length <= 50)
//                {
//                    return capitalMatch.Groups[1].Value;
//                }
//            }

//            // Strategy 3: First non-empty line that's not a number or date
//            foreach (var line in lines)
//            {
//                if (!Regex.IsMatch(line, @"^\d") && !line.Contains("/") && line.Length >= 3 && line.Length <= 50)
//                    return line;
//            }

//            return "N/A";
//        }

//        public string ExtractInvoiceNumber(string text)
//        {
//            // Pattern 1: # INV-000085 (ZYLKER format)
//            var match = Regex.Match(text, @"#\s*(INV-\d+)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: Invoice # followed by number
//            match = Regex.Match(text, @"Invoice\s*#\s*[:.]?\s*(\w+[-]?\d+)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: GF Corys format: "Invoice No. Date 261200791"
//            match = Regex.Match(text, @"Invoice\s+No\.\s+Date\s+(\d{8,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 4: ZAKER format: *299355*
//            match = Regex.Match(text, @"\*(\d{6})\*");
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 5: "Inv No. : 299355"
//            match = Regex.Match(text, @"Inv\.?\s*No\.?\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 6: Generic "Invoice Number" or "Invoice No"
//            match = Regex.Match(text, @"Invoice\s*(?:Number|No\.?)\s*[:.]?\s*([A-Z0-9-]+)", RegexOptions.IgnoreCase);
//            if (match.Success)
//            {
//                string number = match.Groups[1].Value;
//                if (!number.StartsWith("100") && !number.StartsWith("971")) // Filter out TRN/phone
//                    return number;
//            }

//            // Pattern 7: Standalone invoice number pattern (6-9 digits)
//            match = Regex.Match(text, @"\b(\d{6,9})\b");
//            if (match.Success)
//            {
//                string num = match.Groups[1].Value;
//                // Make sure it's not a TRN, phone, or date
//                if (!num.StartsWith("100") && !num.StartsWith("971") && num.Length != 8)
//                    return num;
//            }

//            return "N/A";
//        }

//        public string ExtractDate(string text)
//        {
//            // Pattern 1: "Invoice Date: 08 Mar 2022" (ZYLKER format)
//            var match = Regex.Match(text, @"Invoice\s*Date\s*[:.]\s*(\d{1,2}\s+\w{3}\s+\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "Date : 09/02/2026"
//            match = Regex.Match(text, @"Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: DD-MMM-YYYY format
//            match = Regex.Match(text, @"\b(\d{1,2}[-]\w{3}[-]\d{4})\b", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 4: Generic DD/MM/YYYY
//            match = Regex.Match(text, @"\b(\d{2}/\d{2}/\d{4})\b");
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 5: DD Mon YYYY (08 Mar 2022)
//            match = Regex.Match(text, @"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4})\b", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractTRN(string text)
//        {
//            // Pattern 1: "TRN NO: 100339834200003"
//            var match = Regex.Match(text, @"TRN\s*(?:NO|Number)?[:\s]*(\d{15})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "Tax ID" or "VAT Number"
//            match = Regex.Match(text, @"(?:Tax\s*ID|VAT\s*(?:Number|No))\s*[:.]?\s*(\d{9,15})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: Generic 15-digit number starting with 100
//            match = Regex.Match(text, @"\b(100\d{12,15})\b");
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractSalesPerson(string text)
//        {
//            // Pattern 1: "Salesman : Muhammed"
//            var match = Regex.Match(text, @"Salesman\s*[:]\s*([A-Za-z]+(?:\s+[A-Za-z]+)?)", RegexOptions.IgnoreCase);
//            if (match.Success)
//            {
//                string name = match.Groups[1].Value.Trim();
//                if (!Regex.IsMatch(name, @"Currency|UAE|Dirham|Date|Terms|Total|Rate", RegexOptions.IgnoreCase))
//                    return name;
//            }

//            // Pattern 2: "Sales Person:"
//            match = Regex.Match(text, @"Sales\s*Person\s*[:]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value.Trim();

//            // Pattern 3: GF Corys format (BIJU V. PILLAI)
//            match = Regex.Match(text, @"\b([A-Z]{3,}\s+[A-Z]\.\s+[A-Z]{3,})\b");
//            if (match.Success)
//            {
//                string name = match.Groups[1].Value;
//                if (!name.Contains("LLC") && !name.Contains("BOX") && !name.Contains("VAT"))
//                    return name;
//            }

//            return "N/A";
//        }

//        public string ExtractPaymentTerms(string text)
//        {
//            // Pattern 1: "Terms: Net 15" (ZYLKER format)
//            var match = Regex.Match(text, @"Terms\s*[:]\s*(Net\s+\d+)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "90 days PDC on delivery"
//            match = Regex.Match(text, @"(\d+\s*days?\s*PDC\s*on\s*delivery)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: "Payment Terms: Net 30"
//            match = Regex.Match(text, @"Payment\s*Terms\s*[:]\s*(Net\s*\d+|COD|Due\s+on\s+Receipt)", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 4: Just "X Days"
//            match = Regex.Match(text, @"\b(\d+\s*days?)\b", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 5: "Net X"
//            match = Regex.Match(text, @"\b(Net\s*\d+)\b", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractShipDate(string text)
//        {
//            // Pattern 1: "Ship Date:"
//            var match = Regex.Match(text, @"Ship\s*Date\s*[:]\s*(\d{1,2}[-/]\w{3}[-/]\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "Supply Date : 09/02/2026"
//            match = Regex.Match(text, @"Supply\s*Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: "Delivery Date"
//            match = Regex.Match(text, @"Delivery\s*Date\s*[:]\s*(\d{1,2}[-/\s]\w{3}[-/\s]\d{4})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractDONumber(string text)
//        {
//            // Pattern 1: "D.O. Number"
//            var match = Regex.Match(text, @"D\.O\.\s*Number\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "DO No."
//            match = Regex.Match(text, @"DO\s*No\.?\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: "Delivery Order:"
//            match = Regex.Match(text, @"Delivery\s*Order\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public string ExtractSONumber(string text)
//        {
//            // Pattern 1: "S.O. Number"
//            var match = Regex.Match(text, @"S\.O\.\s*Number\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 2: "SO Number" or "SO No"
//            match = Regex.Match(text, @"SO\s*(?:Number|No\.?)\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 3: "P.O. Number" or "PO#"
//            match = Regex.Match(text, @"P\.?O\.?\s*(?:Number|No\.?|#)\s*[:.]?\s*(PO\d{6}|\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            // Pattern 4: "Sales Order:"
//            match = Regex.Match(text, @"Sales\s*Order\s*[:.]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
//            if (match.Success)
//                return match.Groups[1].Value;

//            return "N/A";
//        }

//        public List<InvoiceLineItem> ExtractLineItems(string text)
//        {
//            return _tableExtractor.ExtractLineItems(text);
//        }
//    }
//}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using InvoiceOCR_MultiFormat.OCR;
using InvoiceOCR_MultiFormat.Services;

namespace InvoiceOCR_MultiFormat.Extractors
{
    public class UniversalInvoiceExtractor
    {
        private readonly MultiStrategyTableExtractor _tableExtractor;

        public UniversalInvoiceExtractor()
        {
            _tableExtractor = new MultiStrategyTableExtractor();
        }

        // Helper method to extract value after a label
        private string ExtractValueAfterLabel(string text, string[] labelPatterns, string valuePattern = @"[\s:]*(.+?)(?:\n|$)")
        {
            foreach (var label in labelPatterns)
            {
                var pattern = label + valuePattern;
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success)
                {
                    string value = match.Groups[1].Value.Trim();
                    // Clean up common trailing junk
                    value = Regex.Replace(value, @"[:\s]+$", "");
                    if (!string.IsNullOrWhiteSpace(value) && value.Length > 0)
                        return value;
                }
            }
            return null;
        }

        public string ExtractCompanyName(string text)
        {
            // Strategy 1: Look for company in header area (first 500 chars)
            var headerText = text.Substring(0, Math.Min(500, text.Length));
            var lines = headerText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .Take(15)
                .ToList();

            // Strategy 2: Look for Beneficiary field (often contains company name)
            var beneficiaryPatterns = new[] { @"Beneficiary\s*[:.]?\s*", @"Company\s*Name\s*[:.]?\s*" };
            var beneficiary = ExtractValueAfterLabel(text, beneficiaryPatterns);
            if (beneficiary != null && beneficiary.Length >= 5)
                return beneficiary;

            // Strategy 3: Look in first few lines for company indicators
            foreach (var line in lines)
            {
                // Skip common invoice headers
                if (Regex.IsMatch(line, @"^(TAX\s*INVOICE|Invoice|Bill\s+To|Date|#|Terms|Page)", RegexOptions.IgnoreCase))
                    continue;

                // Look for lines with company indicators
                if (Regex.IsMatch(line, @"\b(LLC|Ltd|Limited|Inc|Corporation|Corp|Co\.|Trading|Company|Systems|L\.L\.C\.)\b", RegexOptions.IgnoreCase))
                {
                    // Clean the line
                    string cleaned = Regex.Replace(line, @"^\d+\s*", ""); // Remove leading numbers
                    cleaned = Regex.Replace(cleaned, @"\s*[-–]\s*\w+$", ""); // Remove trailing country/city

                    if (cleaned.Length >= 5 && cleaned.Length <= 100 && !cleaned.Contains("@"))
                        return cleaned;
                }
            }

            // Strategy 4: First substantial line that looks like a company name
            foreach (var line in lines)
            {
                if (line.Length >= 5 && line.Length <= 60 &&
                    !Regex.IsMatch(line, @"^\d") &&
                    !line.Contains("/") &&
                    !line.Contains("@") &&
                    Regex.IsMatch(line, @"[A-Za-z]{3,}"))
                {
                    return line;
                }
            }

            return "N/A";
        }

        public string ExtractInvoiceNumber(string text)
        {
            // Multiple label variations for invoice number
            var labels = new[]
            {
                @"Invoice\s*No\.?\s*",
                @"Inv\.?\s*No\.?\s*",
                @"Invoice\s*Number\s*",
                @"Invoice\s*#\s*",
                @"Tax\s*Invoice\s*",
                @"Legal\s*No\.?\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:.]*([A-Z0-9-]+)");
            if (value != null)
            {
                // Filter out TRN, phone numbers
                if (!value.StartsWith("100") && !value.StartsWith("971") && !value.StartsWith("+"))
                    return value;
            }

            // Look for standalone patterns like *299355* or # INV-000085
            var match = Regex.Match(text, @"[#*]\s*(INV-\d+|\d{6})\s*[*]?", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            return "N/A";
        }

        public string ExtractDate(string text)
        {
            var labels = new[]
            {
                @"Invoice\s*Date\s*",
                @"Date\s*",
                @"Dated\s*",
                @"Issue\s*Date\s*"
            };

            // Try to find date after label
            foreach (var label in labels)
            {
                // DD/MM/YYYY format
                var match = Regex.Match(text, label + @"[\s:]*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
                if (match.Success)
                    return match.Groups[1].Value;

                // DD-MMM-YYYY format (11-JAN-2026)
                match = Regex.Match(text, label + @"[\s:]*(\d{1,2}[-]\w{3}[-]\d{4})", RegexOptions.IgnoreCase);
                if (match.Success)
                    return match.Groups[1].Value;

                // DD Mon YYYY format (08 Mar 2022)
                match = Regex.Match(text, label + @"[\s:]*(\d{1,2}\s+\w{3}\s+\d{4})", RegexOptions.IgnoreCase);
                if (match.Success)
                    return match.Groups[1].Value;
            }

            // Fallback: find any date-like pattern in first 1000 chars
            var headerText = text.Substring(0, Math.Min(1000, text.Length));

            var dateMatch = Regex.Match(headerText, @"\b(\d{2}/\d{2}/\d{4}|\d{1,2}[-]\w{3}[-]\d{4})\b", RegexOptions.IgnoreCase);
            if (dateMatch.Success)
                return dateMatch.Groups[1].Value;

            return "N/A";
        }

        public string ExtractTRN(string text)
        {
            var labels = new[]
            {
                @"VAT\s*TRN\s*(?:NO|Number)?\s*",
                @"TRN\s*(?:NO|Number)?\s*",
                @"Tax\s*(?:Registration\s*)?(?:Number|No\.?|ID)\s*",
                @"VAT\s*(?:Number|No\.?|ID)\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:]*(\d{9,15})");
            if (value != null)
                return value;

            // Fallback: look for 15-digit number starting with 100
            var match = Regex.Match(text, @"\b(100\d{12,15})\b");
            if (match.Success)
                return match.Groups[1].Value;

            return "N/A";
        }

        public string ExtractSalesPerson(string text)
        {
            var labels = new[]
            {
                @"Sales\s*Person\s*",
                @"Salesman\s*",
                @"Sales\s*Rep\s*",
                @"Representative\s*",
                @"Agent\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)");
            if (value != null && value.Length >= 3)
            {
                // Filter out non-name words
                if (!Regex.IsMatch(value, @"Currency|UAE|Dirham|Date|Terms|Total|Rate|Amount|N/?A", RegexOptions.IgnoreCase))
                    return value;
            }

            // Look for all-caps format (BIJU V. PILLAI)
            var match = Regex.Match(text, @"\b([A-Z]{3,}\s+[A-Z]\.\s+[A-Z]{3,})\b");
            if (match.Success)
            {
                string name = match.Groups[1].Value;
                if (!name.Contains("LLC") && !name.Contains("BOX") && !name.Contains("VAT"))
                    return name;
            }

            return "N/A";
        }

        public string ExtractPaymentTerms(string text)
        {
            var labels = new[]
            {
                @"Payment\s*Terms\s*",
                @"Terms\s*",
                @"Credit\s*Terms\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:]*([\w\s]+)");
            if (value != null && (value.Contains("Net") || value.Contains("Days") || value.Contains("PDC") || value.Contains("COD")))
                return value.Split(new[] { '\n', '\r' })[0].Trim(); // Take first line only

            return "N/A";
        }

        public string ExtractShipDate(string text)
        {
            var labels = new[]
            {
                @"Ship\s*Date\s*",
                @"Supply\s*Date\s*",
                @"Delivery\s*Date\s*",
                @"Dispatch\s*Date\s*"
            };

            foreach (var label in labels)
            {
                // DD/MM/YYYY
                var match = Regex.Match(text, label + @"[\s:]*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
                if (match.Success)
                    return match.Groups[1].Value;

                // DD-MMM-YYYY
                match = Regex.Match(text, label + @"[\s:]*(\d{1,2}[-]\w{3}[-]\d{4})", RegexOptions.IgnoreCase);
                if (match.Success)
                    return match.Groups[1].Value;
            }

            return "N/A";
        }

        public string ExtractDONumber(string text)
        {
            var labels = new[]
            {
                @"D\.?O\.?\s*Number\s*",
                @"D\.?O\.?\s*No\.?\s*",
                @"Delivery\s*Order\s*(?:Number|No\.?)?\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:]*(\d{5,10})");
            if (value != null)
                return value;

            return "N/A";
        }

        public string ExtractSONumber(string text)
        {
            var labels = new[]
            {
                @"S\.?\s*O\.?\s*Number\s*",
                @"S\.?\s*O\.?\s*No\.?\s*",
                @"Sales\s*Order\s*(?:Number|No\.?)?\s*",
                @"SO\s*(?:Number|No\.?)?\s*",
                @"P\.?O\.?\s*(?:Number|No\.?|#)\s*"
            };

            var value = ExtractValueAfterLabel(text, labels, @"[\s:]*(PO\d+|\d{5,12})");
            if (value != null)
                return value;

            return "N/A";
        }

        public List<InvoiceLineItem> ExtractLineItems(string text)
        {
            return _tableExtractor.ExtractLineItems(text);
        }
    }
}