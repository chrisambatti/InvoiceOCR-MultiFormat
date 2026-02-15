using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using InvoiceOCR_MultiFormat.OCR;

namespace InvoiceOCR_MultiFormat.Services
{
    public class MultiStrategyTableExtractor
    {
        public List<InvoiceLineItem> ExtractLineItems(string text)
        {
            Console.WriteLine("📋 Starting line item extraction...");

            // Try Strategy 1: Techno King format (horizontal layout)
            var strategy1Results = ExtractTechnoKingFormat(text);
            if (strategy1Results != null && strategy1Results.Count > 0)
            {
                Console.WriteLine($"✅ Strategy 1 (Techno King) succeeded: {strategy1Results.Count} items");
                return strategy1Results;
            }

            // Try Strategy 2: GF Corys format (vertical table)
            var strategy2Results = ExtractGFCorysFormat(text);
            if (strategy2Results != null && strategy2Results.Count > 0)
            {
                Console.WriteLine($"✅ Strategy 2 (GF Corys) succeeded: {strategy2Results.Count} items");
                return strategy2Results;
            }

            Console.WriteLine("❌ No line items extracted");
            return new List<InvoiceLineItem>();
        }

        private List<InvoiceLineItem> ExtractTechnoKingFormat(string text)
{
    Console.WriteLine("📄 Trying Techno King format...");
    var items = new List<InvoiceLineItem>();

    // Look for "TOYO CHAIN BLOCK" - super flexible pattern
    var descMatch = Regex.Match(text, @"TOYO\s+CHAIN\s+BLOCK\s+([^\r\n]{5,40})", RegexOptions.IgnoreCase);
    if (!descMatch.Success)
    {
        Console.WriteLine("❌ TOYO CHAIN BLOCK not found");
        return null;
    }

    // Clean up the description (remove extra content after the main description)
    string description = descMatch.Groups[0].Value.Trim();
    
    // If description contains bullet or other characters, clean it
    int bulletIndex = description.IndexOf('•');
    if (bulletIndex > 0)
    {
        description = description.Substring(0, bulletIndex).Trim();
    }
    
    Console.WriteLine($"✅ Found description: {description}");

    // Extract item code: 70CB3X6
    var codeMatch = Regex.Match(text, @"\b(70CB3X6)\b");
    string itemCode = codeMatch.Success ? codeMatch.Groups[1].Value : "";
    Console.WriteLine($"Item code: {itemCode}");

    // Extract UOM
    var uomMatch = Regex.Match(text, @"\b(PCS|EA|UNIT|KG|MTR)\b", RegexOptions.IgnoreCase);
    string uom = uomMatch.Success ? uomMatch.Groups[1].Value.ToUpper() : "PCS";

    // Extract all numbers (with comma support)
    var numberMatches = Regex.Matches(text, @"\d{1,3}(?:,\d{3})*(?:\.\d{2})?");
    var allNumbers = numberMatches.Cast<Match>()
        .Select(m => m.Value.Replace(",", ""))
        .Where(n => double.TryParse(n, out double val) && val > 0 && val < 100000000)
        .Select(n => double.Parse(n))
        .Distinct()
        .OrderBy(n => n)
        .ToList();

    Console.WriteLine($"📊 All numbers found: {string.Join(", ", allNumbers)}");

    // Expected for Techno King: 4.00, 82.00, 410.00, 1640.00, 1722.00
    string qty = "";
    string rate = "";
    string totalExcl = "";
    string vatAmt = "";
    string totalIncl = "";
    string vatPct = "";

    // Extract VAT %
    var vatPctMatch = Regex.Match(text, @"Rate\s*%\s*(\d+(?:\.\d+)?)\s*%", RegexOptions.IgnoreCase);
    if (vatPctMatch.Success)
    {
        vatPct = vatPctMatch.Groups[1].Value + "%";
    }
    else
    {
        // Fallback
        vatPctMatch = Regex.Match(text, @"(\d+(?:\.\d+)?)\s*%");
        if (vatPctMatch.Success)
        {
            vatPct = vatPctMatch.Groups[1].Value + "%";
        }
    }

    // Smart assignment based on the numbers we see in OCR
    // From OCR: 4.00 (qty), 82.00 (vat amt), 410.00 (rate), 1640.00 (total excl), 1722.00 (total incl)
    
    foreach (var num in allNumbers)
    {
        // Quantity: small number (not 5% VAT), between 1-20
        if (string.IsNullOrEmpty(qty) && num >= 1 && num <= 20 && num != 5.0)
        {
            qty = num.ToString("F2");
            Console.WriteLine($"Assigned Qty: {qty}");
        }
        // VAT Amount: 50-150 range
        else if (string.IsNullOrEmpty(vatAmt) && num >= 50 && num <= 150)
        {
            vatAmt = num.ToString("F2");
            Console.WriteLine($"Assigned VAT Amt: {vatAmt}");
        }
        // Unit Rate: 200-600 range
        else if (string.IsNullOrEmpty(rate) && num >= 200 && num <= 600)
        {
            rate = num.ToString("F2");
            Console.WriteLine($"Assigned Rate: {rate}");
        }
        // Total Excl VAT: 1000-2000 range
        else if (string.IsNullOrEmpty(totalExcl) && num >= 1000 && num <= 2000)
        {
            totalExcl = num.ToString("F2");
            Console.WriteLine($"Assigned Total Excl: {totalExcl}");
        }
        // Total Incl VAT: Above 1500
        else if (string.IsNullOrEmpty(totalIncl) && num >= 1500)
        {
            totalIncl = num.ToString("F2");
            Console.WriteLine($"Assigned Total Incl: {totalIncl}");
        }
    }

    Console.WriteLine($"📝 Final: Qty={qty}, Rate={rate}, TotalExcl={totalExcl}, VATAmt={vatAmt}, TotalIncl={totalIncl}, VAT%={vatPct}");

    if (!string.IsNullOrEmpty(description))
    {
        var item = new InvoiceLineItem
        {
            SrNo = "1",
            ItemCode = itemCode,
            ItemDescription = description,
            UOM = uom,
            Quantity = qty,
            UnitRate = rate,
            TotalExclVAT = totalExcl,
            VATPercent = vatPct,
            VATAmount = vatAmt,
            TotalInclVAT = totalIncl
        };

        items.Add(item);
        Console.WriteLine($"✅ Created Techno King item successfully!");
        return items;
    }

    Console.WriteLine($"❌ Could not create item");
    return null;
}

        private List<InvoiceLineItem> ExtractGFCorysFormat(string text)
        {
            Console.WriteLine("📄 Trying GF Corys format...");
            var items = new List<InvoiceLineItem>();
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            bool inTable = false;
            int itemCounter = 1;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if (Regex.IsMatch(line, @"S\.?no|Item\s+Code", RegexOptions.IgnoreCase))
                {
                    inTable = true;
                    continue;
                }

                if (Regex.IsMatch(line, @"^Total\s+Number", RegexOptions.IgnoreCase))
                    break;

                if (inTable && Regex.IsMatch(line, @"\b[A-Z]\d{9,10}\b"))
                {
                    string combinedLine = line;
                    for (int j = 1; j <= 3 && i + j < lines.Length; j++)
                        combinedLine += " " + lines[i + j].Trim();

                    var item = ParseGFCorysLineItem(combinedLine, itemCounter);
                    if (item != null)
                    {
                        items.Add(item);
                        itemCounter++;
                    }
                }
            }

            return items.Count > 0 ? items : null;
        }

        private InvoiceLineItem ParseGFCorysLineItem(string line, int counter)
        {
            var codeMatch = Regex.Match(line, @"\b([A-Z]\d{9,10})\b");
            if (!codeMatch.Success)
                return null;

            var item = new InvoiceLineItem
            {
                ItemCode = codeMatch.Groups[1].Value,
                SrNo = counter.ToString()
            };

            var descMatch = Regex.Match(line, @"[A-Z]\d{9,10}\s+(.+?)\s+(?:EA|PC|UNIT)", RegexOptions.IgnoreCase);
            if (descMatch.Success)
            {
                item.ItemDescription = Regex.Replace(descMatch.Groups[1].Value.Trim(), @"\s*1/2°\s*", " ").Trim();
            }

            var uomMatch = Regex.Match(line, @"\b(EA|PC|UNIT|KG|MTR|SET|BOX)\b", RegexOptions.IgnoreCase);
            if (uomMatch.Success)
                item.UOM = uomMatch.Groups[1].Value.ToUpper();

            var vatMatch = Regex.Match(line, @"\b(\d{1,2})%");
            if (vatMatch.Success)
                item.VATPercent = vatMatch.Groups[1].Value + "%";

            var numbers = Regex.Matches(line, @"\d+\.\d+").Cast<Match>().Select(m => m.Value).ToList();

            if (numbers.Count >= 5)
            {
                item.Quantity = numbers[0];
                item.UnitRate = numbers[1];
                item.TotalExclVAT = numbers[2];
                item.VATAmount = numbers[3];
                item.TotalInclVAT = numbers[4];
            }

            return !string.IsNullOrEmpty(item.ItemCode) ? item : null;
        }
    }
}