using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using InvoiceOCR_MultiFormat.OCR;

namespace InvoiceOCR_MultiFormat.Extractors
{
    public class UniversalInvoiceExtractor
    {
        public string ExtractCompanyName(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            // Strategy 1: Look for company legal suffixes in first 30 lines
            for (int i = 0; i < Math.Min(30, lines.Length); i++)
            {
                var trimmed = lines[i].Trim();

                if (trimmed.Length < 5)
                    continue;

                // Skip pure numbers
                if (Regex.IsMatch(trimmed, @"^[\d\s,.\-]+$"))
                    continue;

                // Skip invoice keywords
                if (Regex.IsMatch(trimmed, @"^(Invoice|Tax\s*Invoice|Date|Total|TRN|VAT|Amount|Payment|Tel|Fax|Email|P\.O|Address|Bill\s*To|Ship\s*To|Page|Welding|Industrial)", RegexOptions.IgnoreCase))
                    continue;

                // Look for LLC, Ltd, etc.
                if (Regex.IsMatch(trimmed, @"\b(LLC|L\.L\.C|Ltd|Limited|Inc|Incorporated|Corp|Corporation|Co\.|Company|LTD|PLC|FZE|FZCO|FZ-LLC)\b", RegexOptions.IgnoreCase))
                {
                    // Extract just the company name (stop at TRN, Tel, numbers, commas)
                    var companyMatch = Regex.Match(trimmed, @"^([A-Z][^,\d]*?(?:LLC|L\.L\.C|Ltd|Limited|Inc|Corp|Corporation|Co\.|Company|FZE|FZCO|FZ-LLC))(?:\s|,|$)", RegexOptions.IgnoreCase);
                    if (companyMatch.Success)
                    {
                        string companyName = companyMatch.Groups[1].Value.Trim();
                        if (companyName.Length >= 10 && companyName.Length < 100)
                        {
                            return CleanCompanyName(companyName);
                        }
                    }

                    // Fallback: use line if short and no contact info
                    if (trimmed.Length < 80 && !Regex.IsMatch(trimmed, @"(Tel|Fax|Email|P\.O\.|Address|Box|\d{7,})", RegexOptions.IgnoreCase))
                    {
                        return CleanCompanyName(trimmed);
                    }
                }
            }

            // Strategy 2: Look for "Trading" keyword (common in UAE companies)
            for (int i = 0; i < Math.Min(30, lines.Length); i++)
            {
                var trimmed = lines[i].Trim();

                if (trimmed.Length < 10 || trimmed.Length > 100)
                    continue;

                if (Regex.IsMatch(trimmed, @"^[\d\s,.\-]+$"))
                    continue;

                // Check for Trading, Piping, etc.
                if (Regex.IsMatch(trimmed, @"\b(Trading|Enterprises|Group|Industries|International|Piping|ZAKER|Techno\s*King)\b", RegexOptions.IgnoreCase))
                {
                    // Skip if it's a keyword line
                    if (Regex.IsMatch(trimmed, @"^(Invoice|Date|Tel|Fax|Email|Bill|Ship|Welding|Industrial|Area)", RegexOptions.IgnoreCase))
                        continue;

                    // Extract up to first comma or large number
                    var cleanMatch = Regex.Match(trimmed, @"^([A-Z][^,]*?)(?:,|\d{5,}|$)");
                    if (cleanMatch.Success)
                    {
                        string name = cleanMatch.Groups[1].Value.Trim();
                        if (name.Length >= 10 && name.Length < 100)
                            return CleanCompanyName(name);
                    }
                }
            }

            // Strategy 3: First meaningful line starting with capital letter
            for (int i = 0; i < Math.Min(15, lines.Length); i++)
            {
                var trimmed = lines[i].Trim();

                // Must start with capital letter, be 15-100 chars
                if (!Regex.IsMatch(trimmed, @"^[A-Z+]") || trimmed.Length < 15 || trimmed.Length > 100)
                    continue;

                // Must have significant letters
                if (!Regex.IsMatch(trimmed, @"[A-Za-z]{8,}"))
                    continue;

                // Skip numbers-only
                if (Regex.IsMatch(trimmed, @"^[\d\s,.\-]+$"))
                    continue;

                // Skip common keywords
                if (Regex.IsMatch(trimmed, @"^(Invoice|Tax|Date|Page|TRN|VAT|Total|Amount|Tel|Fax|Email|P\.O|Address|Bill|Ship|Welding|Industrial)", RegexOptions.IgnoreCase))
                    continue;

                // This might be the company name
                var cleanMatch = Regex.Match(trimmed, @"^([A-Z][^,\d]*?)(?:,|\d{5,}|TRN|VAT|Tel|Fax|$)");
                if (cleanMatch.Success)
                {
                    string name = cleanMatch.Groups[1].Value.Trim();
                    if (name.Length >= 10 && name.Length < 100)
                        return CleanCompanyName(name);
                }
            }

            return "N/A";
        }

        private string CleanCompanyName(string name)
        {
            name = Regex.Replace(name, @"^\d+\s*", "");
            name = Regex.Replace(name, @"[|\\]", "");
            name = name.Trim();
            name = Regex.Replace(name, @"\s+", " ");
            name = Regex.Replace(name, @"[\-\s]+$", "");
            return name;
        }

        public string ExtractInvoiceNumber(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            // Strategy 1: Explicit labels
            for (int i = 0; i < Math.Min(50, lines.Length); i++)
            {
                var line = lines[i];

                var patterns = new[]
                {
                    @"Invoice\s*(?:No\.?|Number|#)[:\s]*([A-Z0-9\-]{3,20})",
                    @"Inv\.?\s*(?:No\.?|Number|#)[:\s]*([A-Z0-9\-]{3,20})",
                    @"Tax\s*Invoice\s*(?:No\.?|Number)?[:\s]*([A-Z0-9\-]{3,20})",
                };

                foreach (var pattern in patterns)
                {
                    var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        string number = match.Groups[1].Value.Trim();

                        if (!Regex.IsMatch(number, @"^(Date|Tax|Total|Amount|Ref|Number|No|Customer|Value|Salesman|Legal)$", RegexOptions.IgnoreCase) &&
                            number.Length >= 3 && number.Length <= 20 &&
                            !Regex.IsMatch(number, @"^\d{1,2}[-/]\d{1,2}"))
                        {
                            return number;
                        }
                    }
                }
            }

            // Strategy 2: Next line after label
            for (int i = 0; i < Math.Min(50, lines.Length - 1); i++)
            {
                var line = lines[i].Trim();

                if (Regex.IsMatch(line, @"^(Invoice\s*(?:No\.?|Number)|Inv\.?\s*(?:No\.?|Number)|Tax\s*Invoice)[:\s]*$", RegexOptions.IgnoreCase))
                {
                    if (i + 1 < lines.Length)
                    {
                        var nextLine = lines[i + 1].Trim();
                        var match = Regex.Match(nextLine, @"^([A-Z0-9][A-Z0-9\-]{2,19})");
                        if (match.Success)
                        {
                            string number = match.Groups[1].Value.Trim();
                            if (!Regex.IsMatch(number, @"^(Date|Tax|Total|Ref|Customer|Value|Legal|Salesman)$", RegexOptions.IgnoreCase))
                            {
                                return number;
                            }
                        }
                    }
                }
            }

            // Strategy 3: Standalone 6-digit numbers (common invoice format)
            for (int i = 0; i < Math.Min(30, lines.Length); i++)
            {
                var line = lines[i];

                // Look for 6-digit invoice number
                var match = Regex.Match(line, @"\b(\d{6})\b");
                if (match.Success)
                {
                    string number = match.Groups[1].Value;

                    // Make sure it's not part of a date
                    if (!Regex.IsMatch(line, @"\d{2}/\d{2}/\d{4}") &&
                        !Regex.IsMatch(line, @"\d{2}-\d{2}-\d{4}"))
                    {
                        // Make sure it's near invoice-related text
                        if (i > 0 && Regex.IsMatch(lines[i - 1], @"(Invoice|Inv|Tax)", RegexOptions.IgnoreCase))
                        {
                            return number;
                        }

                        // Or if the line itself mentions invoice
                        if (Regex.IsMatch(line, @"(Invoice|Inv)", RegexOptions.IgnoreCase))
                        {
                            return number;
                        }
                    }
                }
            }

            return "N/A";
        }

        public string ExtractDate(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < Math.Min(50, lines.Length); i++)
            {
                var line = lines[i];

                if (Regex.IsMatch(line, @"\b(Date|Invoice\s*Date|Dated|Inv\s*Date)(?:\s|:)", RegexOptions.IgnoreCase))
                {
                    var datePatterns = new[]
                    {
                        @"\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b",
                        @"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[,\s]+\d{4})\b",
                        @"\b(\d{1,2}[-/](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[-/]\d{2,4})\b",
                        @"\b(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b",
                    };

                    foreach (var pattern in datePatterns)
                    {
                        var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
                        if (match.Success)
                            return match.Groups[1].Value.Trim();
                    }

                    if (i + 1 < lines.Length)
                    {
                        foreach (var pattern in datePatterns)
                        {
                            var match = Regex.Match(lines[i + 1], pattern, RegexOptions.IgnoreCase);
                            if (match.Success)
                                return match.Groups[1].Value.Trim();
                        }
                    }
                }
            }

            foreach (var line in lines.Take(50))
            {
                var match = Regex.Match(line, @"\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b");
                if (match.Success)
                    return match.Groups[1].Value;
            }

            return "N/A";
        }

        public string ExtractTRN(string text)
        {
            var patterns = new[]
            {
                @"TRN[:\s#]*(\d{9,20})",
                @"VAT\s*TRN\s*(?:No\.?|Number)?[:\s]*(\d{9,20})",
                @"Tax\s*Registration\s*(?:No\.?|Number)[:\s]*(\d{9,20})",
                @"VAT\s*(?:No\.?|Number|Registration)[:\s]*(\d{9,20})",
                @"Tax\s*(?:ID|No\.?|Number)[:\s]*(\d{9,20})",
                @"\b(100\d{12,15})\b",
                @"\b(\d{15})\b",
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                foreach (Match match in matches)
                {
                    string number = match.Groups[1].Value.Trim();

                    if (number.Length >= 9 && number.Length <= 20)
                    {
                        return number;
                    }
                }
            }

            return "N/A";
        }

        public string ExtractSalesPerson(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            var labelPatterns = new[]
            {
                @"Salesman",
                @"Sales\s*Man",
                @"Sales\s*Person",
                @"Salesperson",
                @"Sales\s*Rep(?:resentative)?",
                @"Sales\s*Name",
                @"Sales\s*Agent",
                @"Sales\s*Executive",
            };

            int searchLines = Math.Min(lines.Length * 70 / 100, 120);

            for (int i = 0; i < searchLines; i++)
            {
                string line = lines[i];

                foreach (var labelPattern in labelPatterns)
                {
                    // Pattern 1: Inline
                    var inlinePattern = labelPattern + @"[:\s]+([A-Z][A-Za-z\s.]{2,50})";
                    var match = Regex.Match(line, inlinePattern, RegexOptions.IgnoreCase);

                    if (match.Success)
                    {
                        string name = match.Groups[1].Value.Trim();
                        name = Regex.Replace(name, @"\s+", " ");
                        name = Regex.Replace(name, @"[^A-Za-z\s.]+$", "");

                        if (name.Length >= 3 && name.Length <= 50 &&
                            Regex.IsMatch(name, @"^[A-Z][A-Za-z\s.]+$") &&
                            !Regex.IsMatch(name, @"^(Payment|Terms|Date|Ship|Invoice|Total|Customer|Number|TRN|Tax|Delivery|Rate|Amount|Qty|Price|UOM|Mohammed|Legal|Industrial)$", RegexOptions.IgnoreCase))
                        {
                            return name.Trim();
                        }
                    }

                    // Pattern 2: Next line
                    if (Regex.IsMatch(line, @"^\s*" + labelPattern + @"\s*:?\s*$", RegexOptions.IgnoreCase))
                    {
                        for (int j = 1; j <= 3 && i + j < lines.Length; j++)
                        {
                            string nextLine = lines[i + j].Trim();

                            if (string.IsNullOrWhiteSpace(nextLine) || nextLine.Length < 3)
                                continue;

                            if (Regex.IsMatch(nextLine, @"^[A-Z][A-Za-z\s.]{2,49}$") &&
                                !Regex.IsMatch(nextLine, @"^(Payment|Terms|Date|Invoice|Total|TRN|Tax|Delivery|Customer|Rate|Amount|Qty|Price|UOM|Ship|Mohammed|Legal|Industrial|Area)$", RegexOptions.IgnoreCase))
                            {
                                return nextLine;
                            }
                        }
                    }
                }
            }

            return "N/A";
        }

        public string ExtractPaymentTerms(string text)
        {
            var patterns = new[]
            {
                @"Payment\s*Terms[:\s]*([^\n\r]{3,50}?)(?:\n|\r|$)",
                @"Terms[:\s]*(\d+\s*(?:Days|day|Net))",
                @"(\d+\s*Days?)",
                @"(Net\s*\d+)",
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string terms = match.Groups[1].Value.Trim();
                    terms = Regex.Replace(terms, @"\s+", " ");

                    if (terms.Length >= 2 && terms.Length <= 50)
                        return terms;
                }
            }

            return "N/A";
        }

        public string ExtractShipDate(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            var labelPatterns = new[]
            {
                @"Ship\s*Date",
                @"Shipping\s*Date",
                @"Delivery\s*Date",
                @"Dispatch\s*Date",
            };

            var datePatterns = new[]
            {
                @"\b(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b",
                @"\b(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4})\b",
                @"\b(\d{1,2}[-/](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[-/]\d{2,4})\b",
            };

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                foreach (var labelPattern in labelPatterns)
                {
                    if (Regex.IsMatch(line, labelPattern, RegexOptions.IgnoreCase))
                    {
                        foreach (var datePattern in datePatterns)
                        {
                            var match = Regex.Match(line, datePattern, RegexOptions.IgnoreCase);
                            if (match.Success)
                                return match.Groups[1].Value.Trim();
                        }

                        if (i + 1 < lines.Length)
                        {
                            foreach (var datePattern in datePatterns)
                            {
                                var match = Regex.Match(lines[i + 1], datePattern, RegexOptions.IgnoreCase);
                                if (match.Success)
                                    return match.Groups[1].Value.Trim();
                            }
                        }
                    }
                }
            }

            return "N/A";
        }

        public string ExtractDONumber(string text)
        {
            var patterns = new[]
            {
                @"D\.?\s*O\.?\s*(?:No\.?|Number|#)[:\s]*(\d{5,12})",
                @"Delivery\s*Order\s*(?:No\.?|Number|#)?[:\s]*(\d{5,12})",
                @"DO\s*(?:No\.?|Number|#)[:\s]*(\d{5,12})",
                @"\bDO[:\s#-]*(\d{5,12})\b",
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string number = match.Groups[1].Value.Trim();
                    if (number.Length >= 5 && number.Length <= 12)
                        return number;
                }
            }

            return "N/A";
        }

        public string ExtractSONumber(string text)
        {
            var patterns = new[]
            {
                @"S\.?\s*O\.?\s*(?:No\.?|Number|#)[:\s]*(\d{5,12})",
                @"Sales\s*Order\s*(?:No\.?|Number|#)?[:\s]*(\d{5,12})",
                @"SO\s*(?:No\.?|Number|#)[:\s]*(\d{5,12})",
                @"\bSO[:\s#-]*(\d{5,12})\b",
                @"P\.?\s*O\.?\s*(?:No\.?|Number|#)[:\s]*(\d{5,12})",
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    string number = match.Groups[1].Value.Trim();
                    if (number.Length >= 5 && number.Length <= 12)
                        return number;
                }
            }

            return "N/A";
        }

        public List<InvoiceLineItem> ExtractLineItems(string text)
        {
            var items = new List<InvoiceLineItem>();
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            bool inTable = false;
            int itemCounter = 1;
            List<string> accumulatedLines = new List<string>();

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                // Detect table start
                if (Regex.IsMatch(line, @"S\.?\s*No|Item\s*Code|Code|Item\s*Description|Description|Product|Qty|Quantity|UOM|Unit|Rate|Amount|Price", RegexOptions.IgnoreCase))
                {
                    inTable = true;
                    continue;
                }

                // Detect table end
                if (Regex.IsMatch(line, @"^Total|^Sub\s*Total|^Grand\s*Total|^Amount\s*Due|^Balance|^Tax\s*Total|^Net\s*Total|^Freight|^Exchange\s*Rate|^Invoice\s*Value|^Total\s*Number", RegexOptions.IgnoreCase))
                {
                    break;
                }

                if (inTable && !string.IsNullOrWhiteSpace(line))
                {
                    // Skip obvious header/total lines
                    if (Regex.IsMatch(line, @"^(S\.?No|Code|Description|UOM|QTY|Rate|Amount|VAT|Total|Sr\.?\s*No)", RegexOptions.IgnoreCase))
                        continue;

                    // Try to parse this line as an item
                    var item = ParseLineItemMultiLine(line, lines, ref i, itemCounter);
                    if (item != null)
                    {
                        items.Add(item);
                        itemCounter++;
                    }
                }
            }

            return items;
        }

        private InvoiceLineItem ParseLineItemMultiLine(string currentLine, string[] allLines, ref int currentIndex, int counter)
        {
            // Skip if too short or no content
            if (string.IsNullOrWhiteSpace(currentLine) || currentLine.Length < 3)
                return null;

            // Combine current line with next 1-2 lines if they seem to be continuation
            string combinedLine = currentLine;
            int lookAhead = Math.Min(2, allLines.Length - currentIndex - 1);

            for (int j = 1; j <= lookAhead; j++)
            {
                string nextLine = allLines[currentIndex + j].Trim();

                // Stop if we hit a total or new item marker
                if (Regex.IsMatch(nextLine, @"^(Total|Sub|Grand|S\.?No|\d+\s+[A-Z]{3,})", RegexOptions.IgnoreCase))
                    break;

                // If next line has mostly numbers, it might be part of current item
                if (Regex.IsMatch(nextLine, @"\d") && !Regex.IsMatch(nextLine, @"^[A-Z]{5,}"))
                {
                    combinedLine += " " + nextLine;
                    currentIndex = currentIndex + j;
                }
                else
                {
                    break;
                }
            }

            return ParseGenericLineItem(combinedLine, counter);
        }

        private InvoiceLineItem ParseGenericLineItem(string line, int counter)
{
    if (string.IsNullOrWhiteSpace(line) || line.Length < 3)
        return null;

    // Skip header/total lines
    if (Regex.IsMatch(line, @"^(Total|Sub\s*Total|Grand|VAT|Tax|Freight|Misc|Exchange|S\.?No|Sl\.?No|Code|Description|UOM|QTY|Quantity|Rate|Amount|Price|Units)\b", RegexOptions.IgnoreCase))
        return null;

    var item = new InvoiceLineItem
    {
        SrNo = counter.ToString()
    };

    // === EXTRACT ITEM CODE ===
    // Pattern 1: G665168000 (10 digits starting with letter)
    // Pattern 2: 70CB3X6 (mixed alphanumeric)
    var codePatterns = new[]
    {
        @"\b([A-Z]\d{9,10})\b",                          // G665168000
        @"\b(\d{2}[A-Z]{2}\d[A-Z]\d)\b",                 // 70CB3X6
        @"\b([A-Z]{2,5}\d{1,10}[A-Z]{0,3})\b",          // General: ABC123XY
        @"\b([A-Z]\d{6,12})\b",                          // A123456789
        @"\b([A-Z]{1,3}-\d{3,10})\b",                   // A-12345
    };
    
    foreach (var pattern in codePatterns)
    {
        var match = Regex.Match(line, pattern);
        if (match.Success)
        {
            string code = match.Groups[1].Value;
            // Must be 4-15 characters
            if (code.Length >= 4 && code.Length <= 15)
            {
                // Not a TRN (those are 15 digits starting with 100)
                if (!code.StartsWith("100"))
                {
                    item.ItemCode = code;
                    break;
                }
            }
        }
    }

    // === EXTRACT DESCRIPTION ===
    var descPatterns = new[]
    {
        // "TOYO CHAIN BLOCK 3.0T X 6MTR" or "Duct Bend 90° CL D R300 UPVC 1 M/2""
        @"^([A-Za-z][A-Za-z0-9\s\-\(\)°/.:&""']+?)(?:\s+(?:EA|PC|PCS|KG|TON|BOX|SET|MTR|UNIT|METER)\b|\s+\d{2,}\.?\d*\s)",
        
        // All caps before UOM or numbers
        @"^([A-Z][A-Z0-9\s&\-°/.:""']+?)(?:\s+(?:EA|PC|PCS|KG|TON|BOX|SET|MTR|UNIT)\b)",
        
        // Mixed case description
        @"^([A-Za-z][A-Za-z0-9\s\-°/.:""']+?)(?:\s+\d{2,})",
    };
    
    foreach (var pattern in descPatterns)
    {
        var match = Regex.Match(line, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
        {
            string desc = match.Groups[1].Value.Trim();
            
            // Clean up
            desc = Regex.Replace(desc, @"\s+", " ");           // Remove extra spaces
            desc = Regex.Replace(desc, @"\s*[:;,]\s*$", "");  // Remove trailing punctuation
            desc = Regex.Replace(desc, @"^\d+\s+", "");       // Remove leading Sr.No
            
            // Remove item code if it's at the start
            if (!string.IsNullOrEmpty(item.ItemCode))
            {
                desc = desc.Replace(item.ItemCode, "").Trim();
            }
            
            // Must be at least 5 chars
            if (desc.Length >= 5)
            {
                item.ItemDescription = desc;
                break;
            }
        }
    }
    
    // Fallback: Extract text words before numbers
    if (string.IsNullOrEmpty(item.ItemDescription))
    {
        var words = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        var textWords = new List<string>();
        
        foreach (var word in words)
        {
            // Stop at pure numbers with 2+ digits (likely quantity)
            if (Regex.IsMatch(word, @"^\d{2,}\.?\d*$"))
                break;
                
            // Stop at UOM
            if (Regex.IsMatch(word, @"^(EA|PC|PCS|KG|TON|BOX|SET|MTR|UNIT|METER|PIECES)$", RegexOptions.IgnoreCase))
                break;
                
            // Skip item code
            if (!string.IsNullOrEmpty(item.ItemCode) && word == item.ItemCode)
                continue;
                
            // Include words with letters
            if (Regex.IsMatch(word, @"[A-Za-z]{2,}"))
            {
                textWords.Add(word);
            }
                
            if (textWords.Count >= 10)
                break;
        }
        
        if (textWords.Count > 0)
            item.ItemDescription = string.Join(" ", textWords);
    }

    // === EXTRACT UOM ===
    var uomMatch = Regex.Match(line, @"\b(EA|PC|PCS|PIECES|UNIT|UNITS|KG|KGS|KILOGRAM|MTR|METER|METRES|SET|SETS|BOX|BOXES|PACK|PACKS|NOS?|EACH|TON|TONS|TONNE)\b", RegexOptions.IgnoreCase);
    if (uomMatch.Success)
    {
        string uom = uomMatch.Groups[1].Value.ToUpper();
        // Normalize variations
        if (uom == "PCS" || uom == "PIECES") uom = "PC";
        if (uom == "METRE" || uom == "METRES" || uom == "METER") uom = "MTR";
        if (uom == "KGS" || uom == "KILOGRAM") uom = "KG";
        if (uom == "TONS" || uom == "TONNE") uom = "TON";
        if (uom == "PACKS") uom = "PACK";
        if (uom == "NO" || uom == "NOS") uom = "EACH";
        
        item.UOM = uom;
    }

    // === EXTRACT ALL NUMBERS ===
    // Match decimal numbers like 4.00, 410.00, 1640.00, 5.00, 82.00, 1722.00
    var numberMatches = Regex.Matches(line, @"\d+(?:\.\d+)?");
    var numbers = numberMatches.Cast<Match>().Select(m => m.Value).ToList();

    // Filter out very large numbers (TRNs, dates)
    numbers = numbers.Where(n => 
    {
        if (double.TryParse(n, out double val))
        {
            // Keep numbers under 1 million (reasonable for invoice amounts)
            return val < 1000000;
        }
        return true;
    }).ToList();

    // Remove leading serial numbers (1, 2, 3, etc. at start)
    if (numbers.Count > 0 && numbers[0].Length <= 2)
    {
        numbers.RemoveAt(0);
    }

    // === ASSIGN NUMBERS BASED ON COUNT ===
    // Both formats have similar structure:
    // Format 1: Qty, Rate, Amount, Amount, VAT%, VAT Amt, Total
    // Format 2: Qty, Rate, Total(ExclVAT), VAT%, VAT Amt, Total(InclVAT)
    
    if (numbers.Count >= 7)
    {
        // Full format with duplicates: Qty, Rate, Amt1, Amt2, VAT%, VAT Amt, Total
        item.Quantity = numbers[0];
        item.UnitRate = numbers[1];
        item.TotalExclVAT = numbers[2];  // or numbers[3]
        
        // Find VAT percentage
        var vatPctMatch = Regex.Match(line, @"(\d{1,2}(?:\.\d{1,2})?)%");
        if (vatPctMatch.Success)
        {
            item.VATPercent = vatPctMatch.Groups[1].Value + "%";
        }
        
        item.VATAmount = numbers[numbers.Count - 2];
        item.TotalInclVAT = numbers[numbers.Count - 1];
    }
    else if (numbers.Count == 6)
    {
        // Qty, Rate, Total(Excl), VAT%, VAT Amt, Total(Incl)
        item.Quantity = numbers[0];
        item.UnitRate = numbers[1];
        item.TotalExclVAT = numbers[2];
        
        // numbers[3] might be VAT% (like 5 or 5.00)
        var vatPctMatch = Regex.Match(line, @"(\d{1,2}(?:\.\d{1,2})?)%");
        if (vatPctMatch.Success)
        {
            item.VATPercent = vatPctMatch.Groups[1].Value + "%";
        }
        else if (double.TryParse(numbers[3], out double pct) && pct <= 20)
        {
            item.VATPercent = numbers[3] + "%";
        }
        
        item.VATAmount = numbers[4];
        item.TotalInclVAT = numbers[5];
    }
    else if (numbers.Count == 5)
    {
        // Qty, Rate, Total(Excl), VAT Amt, Total(Incl)
        item.Quantity = numbers[0];
        item.UnitRate = numbers[1];
        item.TotalExclVAT = numbers[2];
        item.VATAmount = numbers[3];
        item.TotalInclVAT = numbers[4];
    }
    else if (numbers.Count == 4)
    {
        // Qty, Rate, Total(Excl), Total(Incl)
        item.Quantity = numbers[0];
        item.UnitRate = numbers[1];
        item.TotalExclVAT = numbers[2];
        item.TotalInclVAT = numbers[3];
    }
    else if (numbers.Count == 3)
    {
        // Qty, Rate, Total
        item.Quantity = numbers[0];
        item.UnitRate = numbers[1];
        item.TotalExclVAT = numbers[2];
    }
    else if (numbers.Count == 2)
    {
        // Qty, Total
        item.Quantity = numbers[0];
        item.TotalExclVAT = numbers[1];
    }
    else if (numbers.Count == 1)
    {
        // Just total
        item.TotalExclVAT = numbers[0];
    }

    // === EXTRACT VAT PERCENTAGE (if not already set) ===
    if (string.IsNullOrEmpty(item.VATPercent))
    {
        var vatMatch = Regex.Match(line, @"(\d{1,2}(?:\.\d{1,2})?)%");
        if (vatMatch.Success)
            item.VATPercent = vatMatch.Groups[1].Value + "%";
    }

    // === VALIDATION ===
    // Return item only if we have meaningful data
    bool hasDescription = !string.IsNullOrEmpty(item.ItemDescription) && item.ItemDescription.Length >= 5;
    bool hasCode = !string.IsNullOrEmpty(item.ItemCode);
    bool hasQuantity = !string.IsNullOrEmpty(item.Quantity);
    
    // Must have at least description or code
    if (hasDescription || hasCode)
        return item;

    return null;
}
    }
}