namespace KON.Liwest.ChannelFactory
{
    public static class Constants
    {
        public static readonly string cstrCurrentProjectName = "KON.Liwest.ChannelFactory";
        public static readonly string cstrCurrentProjectCopyright = "Copyright (C) 2024-2026 Oswald Oliver";

        public const string FIREWALL_RULE_APP_NAME = "KON.Liwest.ChannelFactory";

        public const float EXCEL_PIXEL_TO_WIDTH_CONSTANT = (float)(8.43 / 64);
        public const float EXCEL_CHARACTER_TO_WIDTH_CONSTANT = 8 * EXCEL_PIXEL_TO_WIDTH_CONSTANT;
        public const float EXCEL_CHARACTER_TO_WIDTH_CONSTANT_CF = 3 * EXCEL_CHARACTER_TO_WIDTH_CONSTANT;
        public const float EXCEL_MAXIMUM_COLUMN_WIDTH = 255;                // maximum column width supported by excel
        public const string csExcelWorksheetName = "Channels";              // name of the worksheet within the excel export

        public const string csCategory = "LIWEST";
        public const string csTransponderRootElementNameFormat = "LIWEST ({0:yyyy-MM-dd HH:mm})"; // root element name of transponder lists, {0} = export timestamp
        public const int ciTunerType = 0;                                   // 0 = dvb-c
        public const int ciModulationSystem = 1;                            // 1 = dvb-c
        public const int ciOrbitalPosition = 4000;                          // 4000 = default position for dvb-c
        public const int ciModulation = 5;                                  // 5 = qam256
        public const int ciSymbolRate = 6900;                               // 6900 = default symbolrate for dvb-c
        public const int ciFEC = 9;                                         // 9 = default fec for dvb-c
        public const int ciSpectralInversion = 1;                           // 1 = auto spectral inversion for dvb-c
        public const bool ciSatFeed = true;                                 // true = default satfeed for dvb-c
        public const string ciCountryCode = "AUT";                          // "AUT" = default country code for liwest
        public const string ciDefaultDVBCNamespace = "ffff0000";            // default namespace for dvb-c transponders
        public const int ciEnigmaTransponderInversion = 2;                  // 2 = auto inversion for enigma dvb-c transponders
        public const string csTextLineSeparator = " / ";                    // replacement for html line breaks (e.g. package of the smartcard api)
        public const string csUndeterminedLanguage = "und";                 // iso 639-2 code for "undetermined" - fallback for channels without language
        public static readonly string[] csaPlaceholderLanguageCodes = ["und", "mis", "mul", "zxx"]; // language codes without real language (besides qaa-qtz)
        public const string csDefaultAudioLanguage = "mis";                // iso 639-2 code for "uncoded languages"

        public const int ciFrequencyMHzUpperLimit = 1000;                   // frequency values below are interpreted as MHz
        public const int ciFrequencyHzLowerLimit = 10000000;                // frequency values above are interpreted as Hz

        public const string csDefaultDVBViewerChannelDatabaseID = "B2C2";   // default identifier for dvbviewer channel database
        public const int csDefaultDVBViewerChannelDatabaseVersionHi = 1;    // dvbviewer channel database major version
        public const int csDefaultDVBViewerChannelDatabaseVersionLo = 10;   // dvbviewer channel database minor version

        public const int ciEnigmaServiceTypeTV = 1;                                         // enigma service type for tv services
        public const int ciEnigmaServiceTypeRadio = 2;                                      // enigma service type for radio services
        public const int ciEnigmaUnknownCASystemID = 0x0000;                                // enigma CAID entry for encrypted services without known CA_system_ID
        public const string csEnigmaServiceFlags = "f:08";                                 // enigma service flags (keep service data, no automatic update)
        public const string csEnigmaBouquetNameTV = "TV";                                  // name of the enigma tv bouquet
        public const string csEnigmaBouquetFilenameTV = "userbouquet.favourites.tv";        // filename of the enigma tv bouquet
        public const string csEnigmaBouquetNameRadio = "Radio";                             // name of the enigma radio bouquet
        public const string csEnigmaBouquetFilenameRadio = "userbouquet.favourites.radio";  // filename of the enigma radio bouquet

        public const string csLiwestChannelListUrlTV = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csvAll&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=tv&tx_liwestchannels_search[type][digital]=digital";
        public const string csLiwestChannelListUrlRadio = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csv&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=radio&tx_liwestchannels_search[type][digital]=digital";
        public const string csLiwestChannelListDelimiter = ",";             // delimiter of the liwest channel list csv

        public const string csSmartcardApiKeyHeader = "x-api-key";
        public const string csSmartcardApiKey = "03feee3b052dd7a95addd58e105c10447c7a9d0670a232e9942a64b04bd38523";
        public const string csSmartcardApiCheckUrl = "https://signalcheck-api.liwest.at/api/v1/smartcards/check";
        public const string csSmartcardApiResultUrlFormat = "https://signalcheck-api.liwest.at/api/v1/smartcards/check/{0}?type={1}"; // {0} = smartcard id, {1} = type
        public const string csSmartcardApiTypeTV = "TV";
        public const string csSmartcardApiTypeRadio = "Radio";
        public const string csSmartcardApiStatusAccepted = "accepted";
        public const string csSmartcardApiStatusProcessing = "processing";
        public const string csSmartcardApiStatusCompleted = "completed";
        public const int ciSmartcardApiPollingIntervalMilliseconds = 1000;  // delay between two status requests
        public const int ciSmartcardApiPollingTimeoutSeconds = 120;         // maximum time to wait for a completed check
        public const string csSmartcardPairingRequired = "Pairing notwendig";
        public const string csSmartcardPairingNotRequired = "kein Pairing notwendig";

        public const string csAudioCodecMPEG = "MPEG";                      // audio codecs of KON.OctoScan.NET (lsAudioPacketIdentifiersCodec)
        public const string csAudioCodecAAC = "AAC";
        public const string csAudioCodecAC3 = "AC3";
        public const string csAudioCodecEAC3 = "EAC3";
        public const string csAudioCodecAC4 = "AC4";
        public const string csAudioCodecDTS = "DTS";
        public const string csAudioCodecTrueHD = "TRUEHD";
        public const string csAudioCodecUnknown = "UNKNOWN";

        public const string csVideoCodecMPEG1 = "MPEG1";                    // video codecs of the column Codec
        public const string csVideoCodecMPEG2 = "MPEG2";
        public const string csVideoCodecMPEG4Part2 = "MPEG4P2";
        public const string csVideoCodecH264 = "H264";
        public const string csVideoCodecHEVC = "HEVC";
        public const string csVideoCodecVVC = "VVC";
        public const string csVideoCodecAVS = "AVS";
        public const string csVideoCodecAVS2 = "AVS2";
        public const string csVideoCodecAVS3 = "AVS3";
        public const string csVideoCodecVC1 = "VC1";
        public const string csVideoCodecDirac = "DIRAC";

        /// <summary>
        /// Video stream type (PMT) -> video codec
        /// </summary>
        public static readonly IReadOnlyDictionary<byte, string> dVideoStreamTypeCodecs = new Dictionary<byte, string>
        {
            { 0x01, csVideoCodecMPEG1 },
            { 0x02, csVideoCodecMPEG2 },
            { 0x10, csVideoCodecMPEG4Part2 },
            { 0x1B, csVideoCodecH264 },
            { 0x24, csVideoCodecHEVC },
            { 0x33, csVideoCodecVVC },
            { 0x42, csVideoCodecAVS },
            { 0xD2, csVideoCodecAVS2 },
            { 0xD4, csVideoCodecAVS3 },
            { 0xEA, csVideoCodecVC1 },
            { 0xD1, csVideoCodecDirac }
        };

        /// <summary>
        /// Codec spellings of other sources (e.g. smartcard "MPEG4" = H.264) -> codec of the column Codec
        /// </summary>
        public static readonly IReadOnlyDictionary<string, string> dCodecAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "MPEG4", csVideoCodecH264 },
            { "MPEG-4", csVideoCodecH264 },
            { "AVC", csVideoCodecH264 },
            { "H.264", csVideoCodecH264 },
            { "MPEG-2", csVideoCodecMPEG2 },
            { "H.262", csVideoCodecMPEG2 },
            { "H.265", csVideoCodecHEVC },
            { "H.266", csVideoCodecVVC }
        };

        public static readonly string[] csaDolbyAudioCodecs = [csAudioCodecAC3, csAudioCodecEAC3, csAudioCodecAC4]; // audio codecs which mark a channel as dolby
        public const string csUninstallRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"; // uninstall entries of installed programs
        public const string csServicesRegistryKey = @"SYSTEM\CurrentControlSet\Services";                    // service configuration (ImagePath)
        public static readonly string[] csaDVBViewerProductNames = ["DVBViewer Pro", "DVBViewer Media Server", "DVBViewer Recording Service", "DVBViewer GE"]; // DisplayName prefixes of supported installations
        public const string csDVBViewerUserModeFile = "usermode.ini";       // defines the data directory of a DVBViewer installation
        public const string csDVBViewerDefaultRootDirectory = "CMUV";       // default parent of the data directory (Root of usermode.ini)
        public const string csDVBViewerChannelDatabaseFile = "channels.dat"; // channel database within the data directory
        public const string csDVBViewerTransponderDirectory = "Transponders"; // transponder lists within the data directory
        public const int ciDVBViewerProcessExitTimeoutSeconds = 30;         // time for a program to close regularly before it is terminated
        public const int ciDVBViewerServiceTimeoutSeconds = 60;             // time for a service to stop / start
        public const string csEnigmaSettingsDirectory = "/etc/enigma2";     // settings directory of the receiver (lamedb, bouquets)
        public const string csEnigmaChannelDatabaseFile = "lamedb";         // channel database of the receiver (lamedb5 next to it)
        public const string csEnigmaCablesFile = "/etc/tuxbox/cables.xml";  // transponder list of the receiver (symbolic links are resolved)
        public static readonly string[] csaEnigmaBouquetFilePatterns = ["bouquets.tv", "bouquets.radio", "userbouquet.*.tv", "userbouquet.*.radio", "blacklist", "whitelist"]; // files of an EnigmaBouquetsFiles export
        public const string csEnigmaImageVersionFile = "/etc/image-version"; // image information of the receiver (key=value)
        public static readonly string[] csaEnigmaImageNameKeys = ["comment", "distro", "displaydistro"];       // keys of the image name in csEnigmaImageVersionFile
        public static readonly string[] csaEnigmaImageVersionKeys = ["imageversion", "version", "displayversion"]; // keys of the image version in csEnigmaImageVersionFile
        public const int ciEnigmaSshPort = 22;                             // default port of the SSH server
        public const int ciEnigmaConnectTimeoutSeconds = 15;                // timeout for the SSH connection
        public const int ciEnigmaServiceTimeoutSeconds = 60;                // time for enigma2 to stop (init 4) / start (init 3)
        public const int ciEnigmaCommandTimeoutSeconds = 60;               // timeout for a command on the receiver
        public const string csExportDirectory = "Export";                   // export directory next to the executable (generated files with relative path)
        public const string csExportTimestampFormat = "yyyyMMdd-HHmmss";    // subdirectory per run (start of the run)
        public const string csEnigmaBackupDirectory = "EnigmaBackup";       // local backup directory of the receiver settings (within the export directory of the run)
        public const string csEnigmaInstallDirectorySuffix = "_install";    // local directory of the merged install files (next to the backup)
        public const string csEnigmaUploadSuffix = ".channelfactory.tmp";   // temporary file on the receiver during the upload
        public const string csEnigmaKnownHostsFile = "known_hosts";        // host keys of the receivers (in the local application data of the user)
        public const string csBackupTimestampFormat = "yyyyMMdd-HHmmss";    // timestamp of backup files (<file>.<timestamp>.bak)

        public const int ciBlindscanRetries = 3;                         // retries per transponder
        public const int ciBlindscanTunerLockGraceTime = 3000;              // grace time in ms for the tuner to lock
        public const string csBlindscanServerPort = "554";                  // rtsp port of the sat>ip server
        public const bool cbBlindscanIPv6 = false;                          // use ipv6 for the sat>ip connection
    }
}