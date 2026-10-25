using System.Collections.Frozen;
using System.Data;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;

using DataTablePrettyPrinter;

using KON.OctoScan.NET;

using static KON.Liwest.ChannelFactory.Constants;

namespace KON.Liwest.ChannelFactory
{
    public static partial class Extensions
    {
        [GeneratedRegex(@"\d+(?:[.,]\d+)?")]
        private static partial Regex rxFrequencyValue();

        [GeneratedRegex(@"[ \t\r\n]+")]
        private static partial Regex rxWhitespace();

        [GeneratedRegex(@"[,;/|\s]+")]
        private static partial Regex rxLanguageSeparator();

        [GeneratedRegex(@"\b([01]?\d|2[0-3]):([0-5]\d)\b")]
        private static partial Regex rxTimeOfDay();

        [GeneratedRegex(@"<\s*br\s*/?\s*>|<\s*/\s*(?:p|div|li|tr)\s*>", RegexOptions.IgnoreCase)]
        private static partial Regex rxHtmlLineBreak();

        [GeneratedRegex(@"<[^>]*>")]
        private static partial Regex rxHtmlTag();

        [GeneratedRegex(@"&(?:#\d+|#x[0-9a-f]+|[a-z][a-z0-9]*);", RegexOptions.IgnoreCase)]
        private static partial Regex rxHtmlEntity();

        [GeneratedRegex(@"(?:\s*/\s*){2,}")]
        private static partial Regex rxRepeatedTextSeparator();

        [GeneratedRegex(@"^(?:\s*/\s*)+|(?:\s*/\s*)+$")]
        private static partial Regex rxOuterTextSeparator();

        /// <summary>
        /// ISO 639-1 (2 letters) and ISO 639-2/B, 639-2/T, 639-3 (3 letters) code -> ISO 639-2/T code (e.g. de, ger, deu -> deu)
        /// </summary>
        private static readonly Lazy<FrozenDictionary<string, string>> lfdLanguageCodes = new(() =>
        {
            var dCurrentLanguageCodes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var lCurrentLanguage in ISO639LanguageCollection.AllISO639Languages)
            {
                var strCurrentLanguageCode = GetThreeLetterLanguageCode(lCurrentLanguage.ISO6392Code) ?? lCurrentLanguage.ISO6393Code.ToLowerInvariant();

                dCurrentLanguageCodes.TryAdd(lCurrentLanguage.ISO6393Code, strCurrentLanguageCode);

                if (!string.IsNullOrEmpty(lCurrentLanguage.ISO6392Code))
                    dCurrentLanguageCodes.TryAdd(lCurrentLanguage.ISO6392Code, strCurrentLanguageCode);
            }

            return dCurrentLanguageCodes.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);
        });

        /// <summary>
        /// Extensions.ToBoolean() - converts a string to a boolean value (true/false) or null if the string is not recognized as a boolean value.
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        public static bool? ToBoolean(this string strLocalString)
        {
            return string.IsNullOrEmpty(strLocalString) switch
            {
                false when strLocalString.Equals("ja", StringComparison.OrdinalIgnoreCase) => true,
                false when strLocalString.Equals("nein", StringComparison.OrdinalIgnoreCase) => false,
                false when strLocalString.Equals("yes", StringComparison.OrdinalIgnoreCase) => true,
                false when strLocalString.Equals("no", StringComparison.OrdinalIgnoreCase) => false,
                false when strLocalString.Equals("wahr", StringComparison.OrdinalIgnoreCase) => true,
                false when strLocalString.Equals("falsch", StringComparison.OrdinalIgnoreCase) => false,
                false when strLocalString.Equals("true", StringComparison.OrdinalIgnoreCase) => true,
                false when strLocalString.Equals("false", StringComparison.OrdinalIgnoreCase) => false,
                false when strLocalString.Equals("1", StringComparison.OrdinalIgnoreCase) => true,
                false when strLocalString.Equals("0", StringComparison.OrdinalIgnoreCase) => false,
                false when bool.TryParse(strLocalString, out var bReturnValue) => bReturnValue,
                _ => null
            };
        }

        /// <summary>
        /// Extensions.ToInt32() - converts a string to an integer value or 0 if the string is not recognized as an integer value.
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        public static int ToInt32(this string? strLocalString)
        {
            return int.TryParse(strLocalString, out var result) ? result : 0;
        }

        /// <summary>
        /// Extensions.ToNullableInt32() - converts a string to an integer value or null if the string is not recognized as an integer value.
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        public static int? ToNullableInt32(this string? strLocalString)
        {
            return int.TryParse(strLocalString?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;
        }

        /// <summary>
        /// Extensions.ToUInt16OrDefault() - converts a nullable integer to an unsigned short value or 0 if the integer is null or out of range.
        /// </summary>
        /// <param name="iLocalValue"></param>
        /// <returns></returns>
        public static ushort ToUInt16OrDefault(this int? iLocalValue)
        {
            return iLocalValue is >= ushort.MinValue and <= ushort.MaxValue ? (ushort)iLocalValue.Value : (ushort)0;
        }

        /// <summary>
        /// Extensions.ToUInt32OrDefault() - converts a nullable integer to an unsigned integer value or 0 if the integer is null or negative.
        /// </summary>
        /// <param name="iLocalValue"></param>
        /// <returns></returns>
        public static uint ToUInt32OrDefault(this int? iLocalValue)
        {
            return iLocalValue is >= 0 ? (uint)iLocalValue.Value : 0;
        }

        /// <summary>
        /// Parses a frequency (e.g. "578000", "578", "578.00 MHz", "578,5") and normalizes it to kHz.
        /// </summary>
        public static int? ToFrequencyKHz(this string? strLocalString)
        {
            if (string.IsNullOrWhiteSpace(strLocalString))
                return null;

            var mCurrentMatch = rxFrequencyValue().Match(strLocalString);
            if (!mCurrentMatch.Success || !decimal.TryParse(mCurrentMatch.Value.Replace(',', '.'), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var decCurrentFrequency))
                return null;

            return decCurrentFrequency.ToFrequencyKHz();
        }

        /// <summary>
        /// Normalizes a frequency to kHz.
        /// </summary>
        public static int? ToFrequencyKHz(this int? iLocalFrequency)
        {
            return iLocalFrequency.HasValue ? ((decimal)iLocalFrequency.Value).ToFrequencyKHz() : null;
        }

        /// <summary>
        /// Normalizes a frequency to kHz - values below ciFrequencyMHzUpperLimit are interpreted as MHz, values above ciFrequencyHzLowerLimit as Hz.
        /// </summary>
        public static int? ToFrequencyKHz(this decimal decLocalFrequency)
        {
            if (decLocalFrequency <= 0)
                return null;

            if (decLocalFrequency < ciFrequencyMHzUpperLimit)
                decLocalFrequency *= 1000;
            else if (decLocalFrequency > ciFrequencyHzLowerLimit)
                decLocalFrequency /= 1000;

            return decLocalFrequency <= int.MaxValue ? (int)Math.Round(decLocalFrequency, MidpointRounding.AwayFromZero) : null;
        }

        /// <summary>
        /// Replaces line breaks and tabs by spaces, merges multiple spaces and trims the string.
        /// </summary>
        public static string TrimAdvanced(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? rxWhitespace().Replace(strLocalString, " ").Trim() : string.Empty;
        }

        /// <summary>
        /// Converts a language code (e.g. "DE", "ger", "deu") to the 3 letter ISO 639-2/T code ("deu") - unknown codes are kept in lower case.
        /// </summary>
        /// <returns>Language code or null if empty</returns>
        public static string? ToLanguageCode(this string? strLocalString)
        {
            var strCurrentLanguageCode = strLocalString?.Trim();

            if (string.IsNullOrEmpty(strCurrentLanguageCode))
                return null;

            return lfdLanguageCodes.Value.TryGetValue(strCurrentLanguageCode, out var strCurrentNormalizedLanguageCode) ? strCurrentNormalizedLanguageCode : strCurrentLanguageCode.ToLowerInvariant();
        }

        /// <summary>
        /// Converts a list of language codes (separated by comma, semicolon, slash, pipe or whitespace) to a comma separated list of 3 letter ISO 639-2/T codes.
        /// </summary>
        /// <returns>Language codes or null if empty</returns>
        public static string? ToLanguageCodeList(this string? strLocalString)
        {
            if (string.IsNullOrWhiteSpace(strLocalString))
                return null;

            var saCurrentLanguageCodes = rxLanguageSeparator().Split(strLocalString).Select(strCurrentLanguageCode => strCurrentLanguageCode.ToLanguageCode()).OfType<string>().Distinct().ToArray();
            return saCurrentLanguageCodes.Length > 0 ? string.Join(",", saCurrentLanguageCodes) : null;
        }

        /// <summary>
        /// Language code without real language (undetermined, miscellaneous, multiple, no linguistic content, reserved for local use qaa-qtz e.g. original sound)
        /// </summary>
        public static bool IsPlaceholderLanguageCode(this string strLocalLanguageCode)
        {
            return csaPlaceholderLanguageCodes.Contains(strLocalLanguageCode, StringComparer.OrdinalIgnoreCase) || (strLocalLanguageCode.Length == 3 && string.Compare(strLocalLanguageCode, "qaa", StringComparison.OrdinalIgnoreCase) >= 0 && string.Compare(strLocalLanguageCode, "qtz", StringComparison.OrdinalIgnoreCase) <= 0);
        }

        /// <summary>
        /// Extensions.GetThreeLetterLanguageCode() - converts a 2 letter ISO 639-1 code to a 3 letter ISO 639-2/T code (e.g. de -> deu) or null if the input is not a valid 2 letter code.
        /// </summary>
        /// <param name="strLocalTwoLetterLanguageCode"></param>
        /// <returns></returns>
        private static string? GetThreeLetterLanguageCode(string? strLocalTwoLetterLanguageCode)
        {
            if (string.IsNullOrEmpty(strLocalTwoLetterLanguageCode))
                return null;

            try
            {
                var strCurrentLanguageCode = CultureInfo.GetCultureInfo(strLocalTwoLetterLanguageCode).ThreeLetterISOLanguageName;
                return strCurrentLanguageCode is { Length: 3 } ? strCurrentLanguageCode.ToLowerInvariant() : null;
            }
            catch (CultureNotFoundException)
            {
                return null;
            }
        }

        /// <summary>
        /// Converts html (e.g. of the smartcard api) to plain text: line breaks (br, end of p/div/li/tr) become csTextLineSeparator, other tags are removed and entities decoded.
        /// Text without html is returned unchanged.
        /// </summary>
        public static string? ConvertHtmlToText(this string? strLocalString)
        {
            // Plain text (also "Film- & HD-Paket") stays unchanged - only tags or entities are converted
            if (string.IsNullOrEmpty(strLocalString) || (!strLocalString.Contains('<') && !rxHtmlEntity().IsMatch(strLocalString)))
                return strLocalString;

            var strCurrentText = rxHtmlLineBreak().Replace(strLocalString, csTextLineSeparator);
            strCurrentText = rxHtmlTag().Replace(strCurrentText, " ");
            strCurrentText = WebUtility.HtmlDecode(strCurrentText).TrimAdvanced();
            strCurrentText = rxRepeatedTextSeparator().Replace(strCurrentText, csTextLineSeparator);
            strCurrentText = rxOuterTextSeparator().Replace(strCurrentText, string.Empty).Trim();

            return strCurrentText.Length > 0 ? strCurrentText : null;
        }

        /// <summary>
        /// Extensions.ConvertUtcTimesToLocalTime() - converts the times of day of a text (e.g. "23:00", "1:00") given in UTC to the time zone of the system (daylight saving time of the current day).
        /// The rest of the text is kept, times keep their format (with / without leading zero).
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns>Text with local times</returns>
        public static string ConvertUtcTimesToLocalTime(this string strLocalString)
        {
            var dtCurrentUtcDate = DateTime.UtcNow.Date;

            return rxTimeOfDay().Replace(strLocalString, mCurrentMatch =>
            {
                var dtCurrentLocalTime = TimeZoneInfo.ConvertTimeFromUtc(dtCurrentUtcDate.AddHours(int.Parse(mCurrentMatch.Groups[1].Value, CultureInfo.InvariantCulture)).AddMinutes(int.Parse(mCurrentMatch.Groups[2].Value, CultureInfo.InvariantCulture)), TimeZoneInfo.Local);
                return dtCurrentLocalTime.ToString(mCurrentMatch.Groups[1].Value.StartsWith('0') ? "HH:mm" : "H:mm", CultureInfo.InvariantCulture);
            });
        }

        /// <summary>
        /// Extensions.CheckVariationForPrimary() - checks if the given string is a variation of a primary channel name and returns the primary channel name if it is, otherwise the given string unchanged.
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        public static string? CheckVariationForPrimary(this string? strLocalString)
        {
            return Global.GetPrimaryChannelName(strLocalString);
        }

        /// <summary>
        /// Extensions.ToBinary25CharArray() - converts a string to a 25 character array, padded with null characters and prefixed with the length of the string (or 25 if the string is longer than 25 characters).
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        public static char[] ToBinary25CharArray(this string? strLocalString)
        {
            if (!string.IsNullOrEmpty(strLocalString))
                return strLocalString.Substring(0, strLocalString.Length > 25 ? 25 : strLocalString.Length).PadRight(25, '\0').PadLeft(26, strLocalString.Length > 25 ? BitConverter.ToChar(BitConverter.GetBytes(25)) : BitConverter.ToChar(BitConverter.GetBytes(strLocalString.Length))).ToCharArray();

            return string.Empty.PadRight(25, '\0').PadLeft(26, BitConverter.ToChar(BitConverter.GetBytes(0))).ToCharArray();
        }

        /// <summary>
        /// Extensions.FromBinary25CharArrayToString() - converts a 25 character array (with optional leading size byte) to a string, trimmed of null characters.
        /// </summary>
        /// <param name="strLocalCharArray"></param>
        /// <param name="bLocalLeadingSizeByte"></param>
        /// <returns></returns>
        public static string FromBinary25CharArrayToString(this char[]? strLocalCharArray, bool bLocalLeadingSizeByte)
        {
            if (strLocalCharArray is { Length: > 0 })
            {
                if (bLocalLeadingSizeByte)
                {
                    int iCurrentStringLength = strLocalCharArray[0];
                    return new string(strLocalCharArray, 1, Math.Min(iCurrentStringLength, Math.Min(25, strLocalCharArray.Length - 1))).TrimEnd('\0');
                }

                return new string(strLocalCharArray).TrimEnd('\0');
            }

            return string.Empty;
        }

        /// <summary>
        /// Extensions.ToHexString4Byte() - converts an integer to a 4 byte hex string, padded with leading zeros and in lower case.
        /// </summary>
        /// <param name="strLocalValue"></param>
        /// <returns></returns>
        public static string ToHexString4Byte(this int strLocalValue)
        {
            return Convert.ToString(strLocalValue, 16).PadLeft(4, '0').ToLowerInvariant();
        }

        /// <summary>
        /// Extensions.SetColumnWidthToColumnCaption() - sets the width of each column in a DataTable to the length of its caption + 2, or 3 if the caption is empty.
        /// </summary>
        /// <param name="dtCurrentDataTable"></param>
        /// <returns></returns>
        public static DataTable SetColumnWidthToColumnCaption(this DataTable dtCurrentDataTable)
        {
            foreach (DataColumn dcCurrentDataColumn in dtCurrentDataTable.Columns)
            {
                if (dcCurrentDataColumn.ColumnName.Length > 0)
                    dcCurrentDataColumn.SetWidth(dcCurrentDataColumn.ColumnName.Length + 2);
                else
                    dcCurrentDataColumn.SetWidth(3);
            }

            return dtCurrentDataTable;
        }
    }
}