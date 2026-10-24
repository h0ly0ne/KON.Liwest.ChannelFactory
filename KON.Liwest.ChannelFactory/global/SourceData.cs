using System.Data;
using System.Globalization;
using System.Reflection;

using Newtonsoft.Json;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// One channel / service of the source data.
    /// Comments show which source fills the property: CL = channel list, SC = smartcard, BS = blindscan.
    /// </summary>
    [JsonConverter(typeof(NestedJsonPropertyConverter))]
    public class SourceDataItem
    {
        /// <summary>
        /// Public readable properties in declaration order - define the columns of the source data file, the console output and the excel export
        /// </summary>
        public static readonly PropertyInfo[] piaDataProperties = typeof(SourceDataItem).GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(piCurrentProperty => piCurrentProperty.CanRead).ToArray();

        /// <summary>
        /// Public read/write properties - used for merging items
        /// </summary>
        private static readonly PropertyInfo[] piaWritableDataProperties = piaDataProperties.Where(piCurrentProperty => piCurrentProperty.CanWrite).ToArray();

        private string? strChannelName;
        private string? strLanguage;
        private string? strCodec;
        private string? strPackage;
        private int? iChannelFrequency;

        [JsonProperty("program_number")]
        public int? ChannelNumber { get; set; }                         // CL, SC
        [JsonProperty("program_name")]
        public string? ChannelName                                      // CL, SC, BS
        {
            get => strChannelName;
            set
            {
                var strCurrentChannelName = value.TrimAdvanced().CheckVariationForPrimary();
                strChannelName = string.IsNullOrEmpty(strCurrentChannelName) ? null : strCurrentChannelName;
            }
        }
        [JsonConverter(typeof(FrequencyJsonConverter))]
        public int? ChannelFrequency                                    // CL, SC, BS - always in kHz
        {
            get => iChannelFrequency;
            set => iChannelFrequency = value.ToFrequencyKHz();
        }
        [JsonProperty("frequency.value")]
        private string? ChannelFrequencyAsString                        // <HIDDEN>
        {
            get => iChannelFrequency?.ToString();
            set => iChannelFrequency = value.ToFrequencyKHz();
        }
        [JsonProperty("package.nav_pack")]
        public string? Package                                          // CL, SC - html of the smartcard api converted to text (line breaks -> " / ")
        {
            get => strPackage;
            set => strPackage = value.ConvertHtmlToText();
        }
        public bool? Dolby { get; set; }                                // CL
        public bool? VoD { get; set; }                                  // CL
        public string? Language                                         // CL, SC, BS - 3 letter ISO 639-2/T code(s), comma separated - csUndeterminedLanguage if unknown
        {
            get => strLanguage ?? csUndeterminedLanguage;
            set => strLanguage = value.ToLanguageCodeList();
        }
        public bool? Stingray { get; set; }                             // CL
        [JsonProperty("authorization.authorized")]
        public bool? Authorized { get; set; }                           // SC
        public bool? Pairing                                            // SC
        {
            get => string.IsNullOrEmpty(PairingString) ? null : !PairingString.Contains(csSmartcardPairingNotRequired);
            set => PairingString = value switch
            {
                null => string.Empty,
                true => csSmartcardPairingRequired,
                false => csSmartcardPairingNotRequired
            };
        }
        [JsonProperty("pairing.text")]
        private string? PairingString { get; set; }                     // <HIDDEN>
        [JsonProperty("disruption.text")]
        public string? Issues { get; set; }                             // SC
        [JsonProperty("codec.value")]
        public string? Codec                                            // SC, BS - video codec (MPEG2, H264, HEVC, ...), audio codec of the first audio pid for radio
        {
            get => strCodec;
            set => strCodec = string.IsNullOrWhiteSpace(value) ? null : dCodecAliases.GetValueOrDefault(value.Trim(), value.Trim().ToUpperInvariant());
        }
        public string? ProviderName { get; set; }                       // BS
        public bool? Encrypted { get; set; }                            // BS - decision of the blindscan (measured scrambling bits, otherwise CA_descriptor / free_CA_mode)
        public string? CASystemIDs { get; set; }                        // BS - CA_system_IDs of the PMT CA_descriptors (hex, comma separated)
        public int? SID { get; set; }                                   // BS
        public int? TSID { get; set; }                                  // BS
        public int? ONID { get; set; }                                  // BS
        public int? PMT { get; set; }                                   // BS
        public int? PCRPID { get; set; }                                // BS
        public int? VPID { get; set; }                                  // BS
        public string? VPIDST { get; set; }                             // BS
        public string? APIDS { get; set; }                              // BS
        public string? APIDSST { get; set; }                            // BS
        public string? APIDSLANG { get; set; }                          // BS
        public string? APIDSCODEC { get; set; }                         // BS - codec per audio pid (MPEG, AAC, AC3, EAC3, AC4, DTS, TRUEHD, UNKNOWN)
        public int? SPID { get; set; }                                  // BS
        public int? TPID { get; set; }                                  // BS
        public int? RunningStatus { get; set; }                         // BS
        public string? Sources { get; set; }                            // CL, SC, BS - sources of the channel (ChannelList, SmartcardList, BlindscanList), comma separated

        /// <summary>
        /// Channel data comes from the smartcard - source data files without Sources (older versions) are recognized by the smartcard only fields Pairing / Authorized
        /// </summary>
        [JsonIgnore]
        internal bool IsFromSmartcard => Sources != null ? HasSource(FlagSource.SmartcardList) : Pairing.HasValue || Authorized.HasValue;

        /// <summary>
        /// SourceDataItem.HasSource()
        /// </summary>
        /// <param name="fsLocalSource"></param>
        /// <returns>bool</returns>
        internal bool HasSource(FlagSource fsLocalSource)
        {
            return (Sources ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Contains(fsLocalSource.ToString(), StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Service data (SID, TSID, ONID, frequency) required for channel database exports is available
        /// </summary>
        [JsonIgnore]
        internal bool HasServiceData => SID is > 0 && TSID.HasValue && ONID.HasValue && ChannelFrequency is > 0;

        /// <summary>
        /// Channel is a tv service (valid video pid)
        /// </summary>
        [JsonIgnore]
        internal bool IsVideoService => VPID is > 0;

        /// <summary>
        /// Channel contains a valid video or audio stream
        /// </summary>
        [JsonIgnore]
        internal bool HasValidStream => VPID is > 0 || (!string.IsNullOrWhiteSpace(APIDS) && APIDS != "0");

        /// <summary>
        /// Channel is encrypted - unknown encryption is treated as not encrypted
        /// </summary>
        [JsonIgnore]
        internal bool IsEncrypted => Encrypted ?? false;

        /// <summary>
        /// Dolby measured by the blindscan - at least one audio pid is AC-3, E-AC-3 or AC-4
        /// </summary>
        /// <returns>true / false, null if no audio codecs are known</returns>
        [JsonIgnore]
        internal bool? DolbyFromAudioCodecs
        {
            get
            {
                var saCurrentAudioCodecs = (APIDSCODEC ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                return saCurrentAudioCodecs.Length > 0 ? saCurrentAudioCodecs.Any(strCurrentAudioCodec => csaDolbyAudioCodecs.Contains(strCurrentAudioCodec, StringComparer.OrdinalIgnoreCase)) : null;
            }
        }

        /// <summary>
        /// Codec measured by the blindscan - video codec from the video stream type, for services without video the codec of the first audio pid
        /// </summary>
        /// <returns>Codec, null if not known</returns>
        [JsonIgnore]
        internal string? CodecFromBlindscan
        {
            get
            {
                if (VPID is > 0)
                    return byte.TryParse(VPIDST?.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var byCurrentVideoStreamType) ? dVideoStreamTypeCodecs.GetValueOrDefault(byCurrentVideoStreamType) : null;

                return (APIDSCODEC ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault(strCurrentAudioCodec => !string.Equals(strCurrentAudioCodec, csAudioCodecUnknown, StringComparison.OrdinalIgnoreCase))?.ToUpperInvariant();
            }
        }

        /// <summary>
        /// Language measured by the blindscan - language of the first audio pid with a real language (placeholders like und, mis or qaa are skipped)
        /// </summary>
        /// <returns>3 letter ISO 639-2/T code, null if no audio language is known</returns>
        [JsonIgnore]
        internal string? LanguageFromAudioLanguages => (APIDSLANG ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Where(strCurrentLanguageCode => !strCurrentLanguageCode.IsPlaceholderLanguageCode()).Select(strCurrentLanguageCode => strCurrentLanguageCode.ToLanguageCode()).FirstOrDefault();

        /// <summary>
        /// CA_system_IDs as numbers - invalid entries are skipped
        /// </summary>
        [JsonIgnore]
        internal int[] CASystemIDValues => (CASystemIDs ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(strCurrentCASystemID => int.TryParse(strCurrentCASystemID, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var iCurrentCASystemID) ? (int?)iCurrentCASystemID : null).OfType<int>().Distinct().ToArray();

        /// <summary>
        /// SourceDataItem.FillEmptyValuesFrom() - takes over every value that is still empty from the given item
        /// </summary>
        /// <param name="sdiLocalSourceDataItem"></param>
        public void FillEmptyValuesFrom(SourceDataItem sdiLocalSourceDataItem)
        {
            foreach (var piCurrentProperty in piaWritableDataProperties)
            {
                if (!IsEmptyValue(piCurrentProperty.GetValue(this)))
                    continue;

                var oCurrentValue = piCurrentProperty.GetValue(sdiLocalSourceDataItem);
                if (!IsEmptyValue(oCurrentValue))
                    piCurrentProperty.SetValue(this, oCurrentValue);
            }
        }

        /// <summary>
        /// SourceDataItem.ToDataTable() - table representation (e.g. for console output)
        /// </summary>
        /// <param name="iesdiLocalSourceDataItems"></param>
        /// <returns>DataTable</returns>
        public static DataTable ToDataTable(IEnumerable<SourceDataItem> iesdiLocalSourceDataItems)
        {
            var dtCurrentDataTable = new DataTable("SourceDataTable");

            foreach (var piCurrentProperty in piaDataProperties)
                dtCurrentDataTable.Columns.Add(new DataColumn(piCurrentProperty.Name, Nullable.GetUnderlyingType(piCurrentProperty.PropertyType) ?? piCurrentProperty.PropertyType) { AllowDBNull = true });

            foreach (var sdiCurrentSourceDataItem in iesdiLocalSourceDataItems)
                dtCurrentDataTable.Rows.Add(piaDataProperties.Select(piCurrentProperty => piCurrentProperty.GetValue(sdiCurrentSourceDataItem) ?? DBNull.Value).ToArray());

            return dtCurrentDataTable;
        }

        private static bool IsEmptyValue(object? oLocalValue)
        {
            return oLocalValue == null || (oLocalValue is string strCurrentValue && string.IsNullOrWhiteSpace(strCurrentValue));
        }
    }

    public class ChannelVariationDictionary : Dictionary<string, string[]>;

    public class ChannelRemovalDictionary
    {
        public string? ChannelName { get; set; }
        public string? ChannelFrequency { get; set; }
    }

    public class ChannelSortingDictionary
    {
        public string? SortNumber { get; set; }
        public string? ChannelName { get; set; }
        public string? ChannelFrequency { get; set; }
    }

    /// <summary>
    /// Key of a channel for matching (channel name without case, frequency in kHz)
    /// </summary>
    public readonly record struct ChannelKey(string strChannelName, int iChannelFrequency)
    {
        public static ChannelKey Create(string? strLocalChannelName, int? iLocalChannelFrequency) => new((strLocalChannelName ?? string.Empty).Trim().ToUpperInvariant(), iLocalChannelFrequency ?? 0);
    }
}