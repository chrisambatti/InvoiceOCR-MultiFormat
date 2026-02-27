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

        public string ExtractCompanyName(string text)
        {
            // Look for ZAKER TRADING L.L.C.
            var match = Regex.Match(text, @"ZAKER\s+TRADING\s+L\.L\.C\.", RegexOptions.IgnoreCase);
            if (match.Success)
                return "ZAKER TRADING L.L.C.";

            // Look for Techno King
            match = Regex.Match(text, @"Techno\s+King\s+Trading\s+Co\.\s*LLC", RegexOptions.IgnoreCase);
            if (match.Success)
                return "Techno King Trading Co. LLC";

            // Look for GF Corys
            match = Regex.Match(text, @"GF\s+Corys\s+Piping\s+Systems\s+LLC", RegexOptions.IgnoreCase);
            if (match.Success)
                return "GF Corys Piping Systems LLC - Dubai";

            return "N/A";
        }

        public string ExtractInvoiceNumber(string text)
        {
            // Look for invoice number pattern: *299355*
            var match = Regex.Match(text, @"\*(\d{6})\*");
            if (match.Success)
                return match.Groups[1].Value;

            // Look for "Inv No. : 299355"
            match = Regex.Match(text, @"Inv\s+No\.\s*[:]\s*(\d{5,8})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Generic pattern
            match = Regex.Match(text, @"(?:Invoice|Inv\.?)\s*(?:No\.?|Number)\s*[:.\s]*(\d{5,10})", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string number = match.Groups[1].Value;
                if (!number.StartsWith("100") && !number.StartsWith("971"))
                    return number;
            }

            return "N/A";
        }

        public string ExtractDate(string text)
        {
            // Look for "Date : 09/02/2026"
            var match = Regex.Match(text, @"Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Generic DD/MM/YYYY
            match = Regex.Match(text, @"\b(\d{2}/\d{2}/\d{4})\b");
            if (match.Success)
                return match.Groups[1].Value;

            // DD-MMM-YYYY format
            match = Regex.Match(text, @"\b(\d{1,2}[-/](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[-/]\d{2,4})\b", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value.ToUpper();

            return "N/A";
        }

        public string ExtractTRN(string text)
        {
            // Look for TRN: 100339834200003
            var match = Regex.Match(text, @"TRN\s*(?:NO|Number)?[:\s]*(\d{15})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Generic 15-digit TRN starting with 100
            match = Regex.Match(text, @"\b(100\d{12,15})\b");
            return match.Success ? match.Groups[1].Value : "N/A";
        }

        public string ExtractSalesPerson(string text)
        {
            // Look for "Salesman : Muhammed"
            var match = Regex.Match(text, @"Salesman\s*[:]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string name = match.Groups[1].Value.Trim();
                // Filter out non-name words
                if (!Regex.IsMatch(name, @"Currency|UAE|Dirham|Date|Terms|Total", RegexOptions.IgnoreCase))
                    return name;
            }

            // Look for "Sales Person:"
            match = Regex.Match(text, @"Sales\s*Person\s*[:]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value.Trim();

            // GF Corys format (BIJU V. PILLAI)
            match = Regex.Match(text, @"\b([A-Z]{3,}\s+[A-Z]\.\s+[A-Z]{3,})\b");
            if (match.Success)
            {
                string name = match.Groups[1].Value;
                if (!name.Contains("LLC") && !name.Contains("BOX"))
                    return name;
            }

            return "N/A";
        }

        public string ExtractPaymentTerms(string text)
        {
            // Look for "90 days PDC on delivery"
            var match = Regex.Match(text, @"(\d+\s*days?\s*PDC\s*on\s*delivery)", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Look for just "90 days"
            match = Regex.Match(text, @"(\d+\s*days?)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : "N/A";
        }

        public string ExtractShipDate(string text)
        {
            // Look for "Supply Date : 09/02/2026"
            var match = Regex.Match(text, @"Supply\s*Date\s*[:]\s*(\d{2}/\d{2}/\d{4})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Fallback to invoice date
            return ExtractDate(text);
        }

        public string ExtractDONumber(string text)
        {
            // Look for "D.O. No. Code" followed by number
            var match = Regex.Match(text, @"D\.?O\.?\s*No\.?\s*Code\s*(\d{5,8})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Look for standalone DO number
            match = Regex.Match(text, @"DO\s*No[.:]?\s*(\d{5,10})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            return "N/A";
        }

        public string ExtractSONumber(string text)
        {
            // Look for PO# : PO260016
            var match = Regex.Match(text, @"PO#\s*[:]\s*(PO\d{6})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            // Look for S.O. Number
            match = Regex.Match(text, @"S\.?O\.?\s*(?:Number|No\.?)\s*[:.\s]*(\d{5,12})", RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;

            return "N/A";
        }

        public List<InvoiceLineItem> ExtractLineItems(string text)
        {
            return _tableExtractor.ExtractLineItems(text);
        }
    }
}