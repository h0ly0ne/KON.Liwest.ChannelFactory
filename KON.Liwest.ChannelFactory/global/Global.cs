using System.Text;

using static KON.Liwest.ChannelFactory.Constants;

using Newtonsoft.Json;
using NLog;
using Pastel;

namespace KON.Liwest.ChannelFactory
{
    public static class Global
    {
        public static readonly Logger lCurrentLogger = LogManager.GetCurrentClassLogger();
        public static readonly HttpClient hcCurrentHttpClient = new();

        /// <summary>
        /// Variation (case-insensitive) -> primary channel name, built from the channel variation dictionary
        /// </summary>
        private static Dictionary<string, string> dChannelVariationLookup = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Set as soon as an error has been logged - results in a failure exit code
        /// </summary>
        public static bool bErrorOccurred { get; private set; }

        private const int NET_FW_IP_PROTOCOL_TCP = 6;
        private const int NET_FW_IP_PROTOCOL_UDP = 17;
        private const int NET_FW_RULE_DIR_IN = 1;
        private const int NET_FW_RULE_DIR_OUT = 2;
        private const int NET_FW_ACTION_ALLOW = 1;

        public enum ExitCode
        {
            Success = 0,
            InvalidArguments = 1,
            OperationFailed = 2,
            PrerequisitesNotMet = 3
        }
        public enum FlagMode
        {
            GenerateData = 0,
            ModifyData = 1,
            ExportData = 2,
            DumpData = 3,
            AutomaticData = 9
        }
        public enum FlagSubMode
        {
            None = 0,
            SubModeCustom1 = 1,
            SubModeCustom2 = 2,
            SubModeCustom3 = 3
        }
        public enum FlagSource
        {
            ChannelList = 0,
            SmartcardList = 1,
            BlindscanList = 2,
            CombinedList = 3,
            FileList = 4,
            CustomFile = 5
        }
        public enum FlagTarget
        {
            None = 0,
            ConsoleList = 1,
            FileList = 2
        }
        public enum FlagSourceCleanup
        {
            None = 0,
            RemoveInvalidSIDs = 1,
            GroupAndInvalidate = 2
        }
        public enum FlagExportMode
        {
            None = 0,
            ExcelChannelList = 1,
            DVBViewerTransponderList = 2,
            DVBViewerChannelDatabase = 3,
            EnigmaTransponderList = 4,
            EnigmaChannelDatabase = 5,
            EnigmaBouquetsFiles = 6
        }
        public enum FlagExportEnigmaFormat
        {
            None = 0,
            Enigma1Ver1 = 1,    // Enigma 1, version 1 - names and values as Krkadoni.EnigmaSettings.Enums.SettingsVersion
            Enigma1Ver2 = 2,    // Enigma 1, version 2
            Enigma2Ver3 = 3,    // Enigma 2, lamedb version 3
            Enigma2Ver4 = 4,    // Enigma 2, lamedb version 4 (default of current images)
            Enigma2Ver5 = 5     // Enigma 2, lamedb version 4 + lamedb5
        }
        public enum FlagDumpDataMode
        {
            None = 0,
            DVBViewerChannelDatabase = 1
        }

        /// <summary>
        /// Global.LogTrace()
        /// </summary>
        /// <param name="strLocalMessage"></param>
        public static void LogTrace(string strLocalMessage)
        {
            if (lCurrentLogger.IsTraceEnabled)
                lCurrentLogger.Trace(strLocalMessage.Pastel(ConsoleColor.Cyan));
        }

        /// <summary>
        /// Global.LogInfo()
        /// </summary>
        /// <param name="strLocalMessage"></param>
        public static void LogInfo(string strLocalMessage)
        {
            lCurrentLogger.Info(strLocalMessage.Pastel(ConsoleColor.Green));
        }

        /// <summary>
        /// Global.LogWarning()
        /// </summary>
        /// <param name="strLocalMessage"></param>
        public static void LogWarning(string strLocalMessage)
        {
            lCurrentLogger.Warn(strLocalMessage.Pastel(ConsoleColor.Yellow));
        }

        /// <summary>
        /// Global.LogError() - also marks the run as failed
        /// </summary>
        /// <param name="strLocalMessage"></param>
        public static void LogError(string strLocalMessage)
        {
            bErrorOccurred = true;
            lCurrentLogger.Error(strLocalMessage.Pastel(ConsoleColor.Red));
        }

        /// <summary>
        /// Global.SetChannelVariationDictionary()
        /// </summary>
        /// <param name="cvdLocalChannelVariationDictionary"></param>
        public static void SetChannelVariationDictionary(ChannelVariationDictionary? cvdLocalChannelVariationDictionary)
        {
            LogTrace("Global.SetChannelVariationDictionary()");

            var dCurrentChannelVariationLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // First primary name wins if a variation is listed more than once
            foreach (var kvpCurrentKeyValuePair in cvdLocalChannelVariationDictionary ?? [])
            {
                foreach (var strCurrentVariation in kvpCurrentKeyValuePair.Value ?? [])
                {
                    if (strCurrentVariation != null)
                        dCurrentChannelVariationLookup.TryAdd(strCurrentVariation, kvpCurrentKeyValuePair.Key);
                }
            }

            dChannelVariationLookup = dCurrentChannelVariationLookup;
        }

        /// <summary>
        /// Global.GetPrimaryChannelName()
        /// </summary>
        /// <param name="strLocalChannelName"></param>
        /// <returns>Primary channel name if the given name is a known variation, otherwise the given name</returns>
        public static string? GetPrimaryChannelName(string? strLocalChannelName)
        {
            return strLocalChannelName != null && dChannelVariationLookup.TryGetValue(strLocalChannelName, out var strCurrentPrimaryChannelName) ? strCurrentPrimaryChannelName : strLocalChannelName;
        }

        /// <summary>
        /// Global.LoadJsonFile()
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strLocalFilename"></param>
        /// <param name="strLocalDescription"></param>
        /// <returns>Deserialized content or default on error (error is logged)</returns>
        public static T? LoadJsonFile<T>(string strLocalFilename, string strLocalDescription)
        {
            LogTrace("Global.LoadJsonFile()");

            try
            {
                var tCurrentContent = JsonConvert.DeserializeObject<T>(File.ReadAllText(strLocalFilename, Encoding.UTF8));

                if (tCurrentContent == null)
                    LogError($"-» {strLocalDescription} '{strLocalFilename}' is empty");

                return tCurrentContent;
            }
            catch (Exception excCurrentException) when (excCurrentException is JsonException or IOException or UnauthorizedAccessException)
            {
                LogError($"-» Error on loading {strLocalDescription} '{strLocalFilename}': {excCurrentException.Message}");
                return default;
            }
        }

        /// <summary>
        ///     Global.GetOutputFilePath()
        ///     All output files are resolved relative to the current working directory (absolute paths are kept), missing directories are created.
        /// </summary>
        /// <param name="strLocalFilename"></param>
        /// <returns>Absolute path of the output file</returns>
        public static string GetOutputFilePath(string strLocalFilename)
        {
            var strCurrentFilePath = Path.GetFullPath(strLocalFilename);
            var strCurrentDirectory = Path.GetDirectoryName(strCurrentFilePath);

            if (!string.IsNullOrEmpty(strCurrentDirectory))
                Directory.CreateDirectory(strCurrentDirectory);

            return strCurrentFilePath;
        }

        /// <summary>
        /// Global.CreateFirewallRule()
        /// </summary>
        /// <returns>bool</returns>
        public static bool CreateFirewallRule()
        {
            LogTrace("Global.CreateFirewallRule()");

            if (!OperatingSystem.IsWindows())
            {
                LogWarning("-» No windows operating system - cannot create firewall rules.");
                return false;
            }

            var tCurrentFwPolicyType = Type.GetTypeFromProgID("HNetCfg.FwPolicy2", false);
            var tCurrentFwRuleType = Type.GetTypeFromProgID("HNetCfg.FWRule", false);

            if (tCurrentFwPolicyType == null || tCurrentFwRuleType == null)
                return false;

            (string strName, int iProtocol, int iDirection)[] taCurrentRules =
            [
                (FIREWALL_RULE_APP_NAME + "_OUT_TCP", NET_FW_IP_PROTOCOL_TCP, NET_FW_RULE_DIR_OUT),
                (FIREWALL_RULE_APP_NAME + "_IN_TCP", NET_FW_IP_PROTOCOL_TCP, NET_FW_RULE_DIR_IN),
                (FIREWALL_RULE_APP_NAME + "_OUT_UDP", NET_FW_IP_PROTOCOL_UDP, NET_FW_RULE_DIR_OUT),
                (FIREWALL_RULE_APP_NAME + "_IN_UDP", NET_FW_IP_PROTOCOL_UDP, NET_FW_RULE_DIR_IN)
            ];

            try
            {
                dynamic dCurrentFwPolicy = Activator.CreateInstance(tCurrentFwPolicyType)!;
                var dCurrentFwRules = dCurrentFwPolicy.Rules;

                if (taCurrentRules.All(tCurrentRule => FirewallRuleExists(dCurrentFwRules, tCurrentRule.strName)))
                    return true;

                foreach (var tCurrentRule in taCurrentRules)
                {
                    dCurrentFwRules.Remove(tCurrentRule.strName);

                    dynamic dCurrentFwRule = Activator.CreateInstance(tCurrentFwRuleType)!;
                    dCurrentFwRule.Name = tCurrentRule.strName;
                    dCurrentFwRule.Description = tCurrentRule.strName;
                    dCurrentFwRule.Protocol = tCurrentRule.iProtocol;
                    dCurrentFwRule.Action = NET_FW_ACTION_ALLOW;
                    dCurrentFwRule.Direction = tCurrentRule.iDirection;
                    dCurrentFwRule.ApplicationName = Environment.ProcessPath;
                    dCurrentFwRule.Enabled = true;
                    dCurrentFwRule.InterfaceTypes = "All";
                    dCurrentFwRules.Add(dCurrentFwRule);
                }

                return true;
            }
            catch (Exception excCurrentException)
            {
                LogError($"-» Error on creating firewall rules: {excCurrentException.Message}.");
                return false;
            }
        }

        /// <summary>
        /// Global.FirewallRuleExists()
        /// </summary>
        /// <param name="dLocalFwRules"></param>
        /// <param name="strLocalRuleName"></param>
        /// <returns>bool</returns>
        private static bool FirewallRuleExists(dynamic dLocalFwRules, string strLocalRuleName)
        {
            LogTrace("Global.FirewallRuleExists()");

            try
            {
                var dCurrentFwRule = dLocalFwRules.Item(strLocalRuleName);
                return (bool)dCurrentFwRule.Enabled && string.Equals((string)dCurrentFwRule.ApplicationName, Environment.ProcessPath, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
