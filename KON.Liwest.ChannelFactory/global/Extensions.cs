using System.Collections.Frozen;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

using KON.OctoScan.NET;

using DataTablePrettyPrinter;
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

        public static int ToInt32(this string? strLocalString)
        {
            return int.TryParse(strLocalString, out var result) ? result : 0;
        }

        public static int? ToNullableInt32(this string? strLocalString)
        {
            return int.TryParse(strLocalString?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;
        }

        public static ushort ToUInt16OrDefault(this int? iLocalValue)
        {
            return iLocalValue is >= ushort.MinValue and <= ushort.MaxValue ? (ushort)iLocalValue.Value : (ushort)0;
        }

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

        public static string? CheckVariationForPrimary(this string? strLocalString)
        {
            return Global.GetPrimaryChannelName(strLocalString);
        }

        public static char[] ToBinary25CharArray(this string? strLocalString)
        {
            if (!string.IsNullOrEmpty(strLocalString))
                return strLocalString.Substring(0, strLocalString.Length > 25 ? 25 : strLocalString.Length).PadRight(25, '\0').PadLeft(26, strLocalString.Length > 25 ? BitConverter.ToChar(BitConverter.GetBytes(25)) : BitConverter.ToChar(BitConverter.GetBytes(strLocalString.Length))).ToCharArray();

            return string.Empty.PadRight(25, '\0').PadLeft(26, BitConverter.ToChar(BitConverter.GetBytes(0))).ToCharArray();
        }

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

        public static string ToHexString4Byte(this int strLocalValue)
        {
            return Convert.ToString(strLocalValue, 16).PadLeft(4, '0').ToLowerInvariant();
        }

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
