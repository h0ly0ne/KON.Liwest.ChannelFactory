using System.Text.RegularExpressions;

namespace KON.Liwest.ChannelFactory
{
    public static class Extensions
    {
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
            if (Global.svdStringVariationDictionary != null)
            {
                foreach (var kvpCurrentKeyValuePair in Global.svdStringVariationDictionary.Where(kvpCurrentKeyValuePair => kvpCurrentKeyValuePair.Value.AsEnumerable().Contains(strLocalString)))
                {
                    return kvpCurrentKeyValuePair.Key;
                }
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
    }
}