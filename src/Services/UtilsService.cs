using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services.Interfaces;
using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace JsonToWord.Services
{
    public class UtilsService : IUtilsService
    {
        public int ConvertCmToDxa(double cm)
        {
            return (int)Math.Round(cm * 566.9291338582677);
        }

        public int ConvertDxaToPct(int dxa, int pageWidthDxa)
        {
            return (int)Math.Round((double)dxa / pageWidthDxa * 5000);
        }

        public int GetPageWidthDxa(MainDocumentPart mainPart)
        {
            var sectionProperties = mainPart.Document.Body.Elements<SectionProperties>().FirstOrDefault();
            if (sectionProperties != null)
            {
                var pageSize = sectionProperties.GetFirstChild<PageSize>();
                if (pageSize != null && pageSize.Width != null)
                {
                    return (int)pageSize.Width.Value;
                }
            }
            // Default to a standard page width if not found (e.g., 11906 DXA = 8.5 inches)
            return 11906;
        }

        public double ParseStringToDouble(string input)
        {
            // Regular expression to find numeric value (including decimals)
            var match = Regex.Match(input, @"-?\d+(\.\d+)?");

            if (match.Success)
            {
                // Try to parse the extracted numeric part
                if (double.TryParse(match.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out double result))
                {
                    return result;
                }
            }

            throw new FormatException($"Could not parse '{input}' into a float.");
        }
    }
}
