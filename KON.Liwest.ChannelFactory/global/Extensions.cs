using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

using DataTablePrettyPrinter;
using static KON.Liwest.ChannelFactory.Constants;

namespace KON.Liwest.ChannelFactory
{
    public static partial class Extensions
    {
        [GeneratedRegex(@"\d+(?:[.,]\d+)?")]
        private static partial Regex rxFrequencyValue();

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

        /// <summary>
        /// Parses a frequency (e.g. "578000", "578", "578.00 MHz", "578,5") and normalizes it to kHz.
        /// Values below ciFrequencyMHzUpperLimit are interpreted as MHz, values above ciFrequencyHzLowerLimit as Hz.
        /// </summary>
        public static int? ToFrequencyKHz(this string? strLocalString)
        {
            if (string.IsNullOrWhiteSpace(strLocalString))
                return null;

            var mCurrentMatch = rxFrequencyValue().Match(strLocalString);
            if (!mCurrentMatch.Success || !decimal.TryParse(mCurrentMatch.Value.Replace(',', '.'), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var decCurrentFrequency) || decCurrentFrequency <= 0)
                return null;

            if (decCurrentFrequency < ciFrequencyMHzUpperLimit)
                decCurrentFrequency *= 1000;
            else if (decCurrentFrequency > ciFrequencyHzLowerLimit)
                decCurrentFrequency /= 1000;

            return decCurrentFrequency <= int.MaxValue ? (int)Math.Round(decCurrentFrequency, MidpointRounding.AwayFromZero) : null;
        }

        public static bool IsNullOrEmptyValue(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            var oCurrentValue = drLocalDataRow[strLocalColumnName];
            return oCurrentValue == DBNull.Value || string.IsNullOrWhiteSpace(Convert.ToString(oCurrentValue, CultureInfo.InvariantCulture));
        }

        public static int? GetNullableInt32(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            return drLocalDataRow[strLocalColumnName] switch
            {
                DBNull => null,
                int iCurrentValue => iCurrentValue,
                var oCurrentValue => Convert.ToString(oCurrentValue, CultureInfo.InvariantCulture).ToNullableInt32()
            };
        }

        public static int GetInt32OrDefault(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            return drLocalDataRow.GetNullableInt32(strLocalColumnName) ?? 0;
        }

        public static ushort GetUInt16OrDefault(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            return drLocalDataRow.GetNullableInt32(strLocalColumnName) is { } iCurrentValue and >= ushort.MinValue and <= ushort.MaxValue ? (ushort)iCurrentValue : (ushort)0;
        }

        public static uint GetUInt32OrDefault(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            return drLocalDataRow.GetNullableInt32(strLocalColumnName) is { } iCurrentValue and >= 0 ? (uint)iCurrentValue : 0;
        }

        public static bool GetBooleanOrDefault(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            return drLocalDataRow[strLocalColumnName] switch
            {
                DBNull => false,
                bool bCurrentValue => bCurrentValue,
                var oCurrentValue => (Convert.ToString(oCurrentValue, CultureInfo.InvariantCulture) ?? string.Empty).ToBoolean() ?? false
            };
        }

        public static string GetStringOrEmpty(this DataRow drLocalDataRow, string strLocalColumnName)
        {
            var oCurrentValue = drLocalDataRow[strLocalColumnName];
            return oCurrentValue == DBNull.Value ? string.Empty : Convert.ToString(oCurrentValue, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        /// <summary>
        /// Same as CopyToDataTable(), but returns an empty clone of the template table instead of throwing on an empty sequence.
        /// </summary>
        public static DataTable CopyToDataTableOrEmpty(this IEnumerable<DataRow> iedrLocalDataRows, DataTable dtLocalTemplateDataTable)
        {
            var ldrCurrentDataRows = iedrLocalDataRows.ToList();
            return ldrCurrentDataRows.Count > 0 ? ldrCurrentDataRows.CopyToDataTable() : dtLocalTemplateDataTable.Clone();
        }

        public static string RepairEncoding(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? strLocalString.Replace("�", "ü") : string.Empty;
        }
        public static string TrimAdvanced(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? strLocalString.MergeNewLine().MergeTabs().MergeIndent().Replace("\r\n\t", " ").Replace("\r\n", " ").Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").MergeWhitespace().Trim() : string.Empty;
        }

        public static string MergeWhitespace(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? new Regex("[ ]{2,}", RegexOptions.None).Replace(strLocalString, " ") : string.Empty;
        }
        public static string MergeTabs(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? new Regex("[\t]{2,}", RegexOptions.None).Replace(strLocalString, "\t") : string.Empty;
        }
        public static string MergeNewLine(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? new Regex("[\r\n]{2,}", RegexOptions.None).Replace(strLocalString, "\r\n") : string.Empty;
        }
        public static string MergeIndent(this string? strLocalString)
        {
            return !string.IsNullOrEmpty(strLocalString) ? new Regex("[\r\n\t]{2,}", RegexOptions.None).Replace(strLocalString, "\r\n\t") : string.Empty;
        }

        public static string? CheckVariationForPrimary(this string? strLocalString)
        {
            if (Global.svdChannelVariationDictionary == null || strLocalString == null) 
                return strLocalString;
            
            foreach (var kvpCurrentKeyValuePair in Global.svdChannelVariationDictionary.Where(kvpCurrentKeyValuePair => kvpCurrentKeyValuePair.Value.AsEnumerable().Contains(strLocalString, StringComparer.OrdinalIgnoreCase)))
            {
                return kvpCurrentKeyValuePair.Key;
            }

            return strLocalString;
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
                    return new string(strLocalCharArray, 1, iCurrentStringLength > 25 ? 25 : iCurrentStringLength).TrimEnd('\0');
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