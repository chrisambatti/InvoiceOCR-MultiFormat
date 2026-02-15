using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using InvoiceOCR_MultiFormat.OCR;

namespace InvoiceOCR_MultiFormat.Services
{
    public class IntelligentTableExtractor
    {
        // Define possible column header variations
        private static readonly string[] DescriptionHeaders = new[]
        {
            "ITEM DESCRIPTION", "DESCRIPTION", "DESC", "ITEM", "PRODUCT",
            "PRODUCT DESCRIPTION", "PARTICULARS", "DETAILS"
        };

        private static readonly string[] ItemCodeHeaders = new[]
        {
            "ITEM CODE", "CODE", "ITEM NO", "PRODUCT CODE", "SKU",
            "PART NO", "PART NUMBER", "ITEM#"
        };

        private static readonly string[] QtyHeaders = new[]
        {
            "QTY", "QUANTITY", "QUAN", "QTY.", "NO", "NOS", "PCS"
        };

        private static readonly string[] UomHeaders = new[]
        {
            "UOM", "UNIT", "U/M", "UNITS", "UM"
        };

        private static readonly string[] RateHeaders = new[]
        {
            "UNIT RATE", "RATE", "UNIT PRICE", "PRICE", "RATE/UNIT",
            "UNIT RATE (AED)", "UNIT RATE(AED)"
        };

        private static readonly string[] AmountHeaders = new[]
        {
            "TOTAL (EXCL. VAT)", "TOTAL (EXCL VAT)", "AMOUNT", "TOTAL",
            "SUBTOTAL", "AMOUNT (AED)", "TOTAL(EXCL. VAT)"
        };

        private static readonly string[] VatPercentHeaders = new[]
        {
            "VAT %", "VAT%", "TAX %", "TAX%", "VAT"
        };

        private static readonly string[] VatAmountHeaders = new[]
        {
            "VAT AMOUNT", "VAT AMT", "TAX AMOUNT", "TAX",
            "VAT AMOUNT (AED)", "VAT AMT (AED)"
        };

        private static readonly string[] TotalInclHeaders = new[]
        {
            "TOTAL (INCL. VAT)", "TOTAL (INCL VAT)", "TOTAL INCL. VAT",
            "GRAND TOTAL", "NET TOTAL", "TOTAL(INCL. VAT)"
        };

        public class TableStructure
        {
            public int DescriptionCol { get; set; } = -1;
            public int ItemCodeCol { get; set; } = -1;
            public int QtyCol { get; set; } = -1;
            public int UomCol { get; set; } = -1;
            public int RateCol { get; set; } = -1;
            public int AmountCol { get; set; } = -1;
            public int VatPercentCol { get; set; } = -1;
            public int VatAmountCol { get; set; } = -1;
            public int TotalInclCol { get; set; } = -1;
            public int HeaderRowIndex { get; set; } = -1;
            public int TableEndIndex { get; set; } = -1;
        }

        public List<InvoiceLineItem> ExtractLineItems(string ocrText)
        {
            var lines = ocrText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            // Step 1: Find and analyze the table structure
            var tableStructure = AnalyzeTableStructure(lines);

            if (tableStructure.HeaderRowIndex == -1)
            {
                Console.WriteLine("⚠️ No table header found");
                return new List<InvoiceLineItem>();
            }

            Console.WriteLine($"✅ Table found at line {tableStructure.HeaderRowIndex}");
            Console.WriteLine($"📊 Columns detected: Desc={tableStructure.DescriptionCol}, Qty={tableStructure.QtyCol}, Rate={tableStructure.RateCol}");

            // Step 2: Extract rows based on table structure
            var items = ExtractRowsFromTable(lines, tableStructure);

            Console.WriteLine($"✅ Extracted {items.Count} line items");

            return items;
        }

        private TableStructure AnalyzeTableStructure(List<string> lines)
        {
            var structure = new TableStructure();

            // Find header row by looking for multiple column headers
            for (int i = 0; i < Math.Min(50, lines.Count); i++)
            {
                var line = lines[i].ToUpper();

                // Count how many known headers are in this line
                int headerMatches = 0;

                if (ContainsAny(line, DescriptionHeaders)) headerMatches++;
                if (ContainsAny(line, QtyHeaders)) headerMatches++;
                if (ContainsAny(line, RateHeaders)) headerMatches++;

                // If we find 3+ column headers, this is likely the header row
                if (headerMatches >= 3)
                {
                    structure.HeaderRowIndex = i;

                    // Now determine column positions based on header text positions
                    DetermineColumnPositions(line, structure);

                    // Find where table ends
                    structure.TableEndIndex = FindTableEnd(lines, i);

                    break;
                }
            }

            return structure;
        }

        private void DetermineColumnPositions(string headerLine, TableStructure structure)
        {
            var parts = headerLine.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            for (int col = 0; col < parts.Length; col++)
            {
                var part = parts[col].ToUpper();

                if (ContainsAny(part, DescriptionHeaders))
                    structure.DescriptionCol = col;
                else if (ContainsAny(part, ItemCodeHeaders))
                    structure.ItemCodeCol = col;
                else if (ContainsAny(part, QtyHeaders))
                    structure.QtyCol = col;
                else if (ContainsAny(part, UomHeaders))
                    structure.UomCol = col;
                else if (ContainsAny(part, RateHeaders))
                    structure.RateCol = col;
                else if (ContainsAny(part, AmountHeaders))
                    structure.AmountCol = col;
                else if (ContainsAny(part, VatPercentHeaders))
                    structure.VatPercentCol = col;
                else if (ContainsAny(part, VatAmountHeaders))
                    structure.VatAmountCol = col;
                else if (ContainsAny(part, TotalInclHeaders))
                    structure.TotalInclCol = col;
            }
        }

        private int FindTableEnd(List<string> lines, int startIndex)
        {
            for (int i = startIndex + 1; i < lines.Count; i++)
            {
                var line = lines[i].ToUpper();

                // Look for end markers
                if (Regex.IsMatch(line, @"^TOTAL NUMBER OF ITEMS|^FREIGHT|^MISCELLANEOUS|^EXCHANGE RATE|^TOTAL IN WORDS|^SUBTOTAL|^GRAND TOTAL", RegexOptions.IgnoreCase))
                {
                    return i;
                }
            }

            return lines.Count;
        }

        private List<InvoiceLineItem> ExtractRowsFromTable(List<string> lines, TableStructure structure)
        {
            var items = new List<InvoiceLineItem>();
            int counter = 1;

            for (int i = structure.HeaderRowIndex + 1; i < structure.TableEndIndex; i++)
            {
                var line = lines[i];

                // Skip obviously wrong lines
                if (IsNonItemRow(line))
                    continue;

                var item = ExtractItemFromRow(line, structure, counter);

                if (item != null)
                {
                    items.Add(item);
                    counter++;
                }
            }

            return items;
        }

        private InvoiceLineItem ExtractItemFromRow(string line, TableStructure structure, int counter)
        {
            // Try to split the line intelligently
            var parts = SplitLineIntelligently(line);

            // Extract numbers for numeric fields
            var numbers = Regex.Matches(line, @"\d+(?:\.\d+)?")
                .Cast<Match>()
                .Select(m => m.Value)
                .Where(n => {
                    if (double.TryParse(n, out double val))
                    {
                        // Filter out years, TRNs, very large numbers
                        if (val >= 2020 && val <= 2030) return false;
                        if (val > 10000000) return false;
                        return true;
                    }
                    return false;
                })
                .ToList();

            // Must have at least 2 numbers (qty and rate, or qty and total)
            if (numbers.Count < 2)
                return null;

            // Extract description (longest text segment without numbers)
            string description = ExtractDescription(parts, numbers);

            if (string.IsNullOrEmpty(description) || description.Length < 5)
                return null;

            // Extract item code if present
            string itemCode = ExtractItemCode(line);

            // Extract UOM
            string uom = ExtractUOM(line);

            // Extract VAT percentage
            string vatPercent = ExtractVATPercent(line);

            return new InvoiceLineItem
            {
                SrNo = counter.ToString(),
                ItemCode = itemCode,
                ItemDescription = description,
                UOM = uom,
                Quantity = numbers.Count > 0 ? numbers[0] : "",
                UnitRate = numbers.Count > 1 ? numbers[1] : "",
                TotalExclVAT = numbers.Count > 2 ? numbers[2] : "",
                VATPercent = vatPercent,
                VATAmount = numbers.Count > 3 ? numbers[3] : "",
                TotalInclVAT = numbers.Count > 4 ? numbers[4] : ""
            };
        }

        private string[] SplitLineIntelligently(string line)
        {
            // Split by multiple spaces (2+) or tabs
            return Regex.Split(line, @"\s{2,}|\t+")
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToArray();
        }

        private string ExtractDescription(string[] parts, List<string> numbers)
        {
            foreach (var part in parts)
            {
                // Description is usually the longest text segment
                if (part.Length > 10 && !numbers.Contains(part))
                {
                    // Check if it has letters
                    if (Regex.IsMatch(part, @"[A-Za-z]{5,}"))
                    {
                        return part.Trim();
                    }
                }
            }

            // Fallback: combine all text parts
            var textParts = parts.Where(p =>
                p.Length > 3 &&
                !numbers.Contains(p) &&
                Regex.IsMatch(p, @"[A-Za-z]")
            ).ToList();

            return textParts.Any() ? string.Join(" ", textParts) : "";
        }

        private string ExtractItemCode(string line)
        {
            // Common patterns for item codes
            var patterns = new[]
            {
                @"\b([A-Z]\d{6,12})\b",           // G665168000
                @"\b(\d{2,4}[A-Z]{2,4}\d{0,6})\b", // 70CB3X6
                @"\b([A-Z]{2,4}\d{4,8})\b"         // ABC12345
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(line, pattern);
                if (match.Success)
                {
                    return match.Groups[1].Value;
                }
            }

            return "";
        }

        private string ExtractUOM(string line)
        {
            var match = Regex.Match(line, @"\b(EA|EACH|PC|PCS|PIECES|BOX|BOXES|MTR|METER|UNIT|UNITS|SET|SETS|KG|TON)\b", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.ToUpper() : "";
        }

        private string ExtractVATPercent(string line)
        {
            var match = Regex.Match(line, @"(\d+(?:\.\d+)?)\s*%");
            return match.Success ? match.Groups[1].Value + "%" : "";
        }

        private bool IsNonItemRow(string line)
        {
            var upper = line.ToUpper();

            // Skip header repetitions
            if (ContainsAny(upper, DescriptionHeaders) && ContainsAny(upper, QtyHeaders))
                return true;

            // Skip address/contact lines
            if (Regex.IsMatch(upper, @"^(BILL TO|SHIP TO|CUSTOMER|EMAIL|PHONE|FAX|ADDRESS|ATTENTION|P\.O\.|@)", RegexOptions.IgnoreCase))
                return true;

            // Skip lines with only numbers (row numbers, page numbers)
            if (Regex.IsMatch(line, @"^\d{1,3}$"))
                return true;

            // Skip very short lines
            if (line.Length < 10)
                return true;

            return false;
        }

        private bool ContainsAny(string text, string[] terms)
        {
            return terms.Any(term => text.Contains(term, StringComparison.OrdinalIgnoreCase));
        }
    }
}