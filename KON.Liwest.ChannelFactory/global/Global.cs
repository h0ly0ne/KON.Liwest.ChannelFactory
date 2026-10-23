using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using KON.OctoScan.NET;
using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.DVBViewer;

using CsvHelper;
using CsvHelper.Configuration;
using DataTablePrettyPrinter;
using Krkadoni.EnigmaSettings;
using NanoXLSX;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using Pastel;
using NullLogger = Krkadoni.EnigmaSettings.NullLogger;

namespace KON.Liwest.ChannelFactory
{
    public static class Global
    {
        public static readonly Logger lCurrentLogger = LogManager.GetCurrentClassLogger();

        public static HttpClient hcCurrentHttpClient = new();
        public static DataTable dtSourceDataTable = new("SourceDataTable");
        public static ChannelVariationDictionary? svdChannelVariationDictionary = new();

        [JsonConverter(typeof(NestedJsonPropertyConverter))]
        public class SourceDataTable
        {
            [JsonProperty("program_number")]
            public int? ChannelNumber { get; set; }                         // CL, SC
            [JsonProperty("program_name")]
            public string? ChannelName                                      // CL, SC, BS
            {
                get => string.IsNullOrEmpty(ChannelNameAsString) ? null : ChannelNameAsString.TrimAdvanced().CheckVariationForPrimary();
                set => ChannelNameAsString = Convert.ToString(value).TrimAdvanced().CheckVariationForPrimary();
            }
            private string? ChannelNameAsString { get; set; }               // <HIDDEN>
            public int? ChannelFrequency                                    // CL, SC, BS
            {
                get => ChannelFrequencyAsString.ToFrequencyKHz();
                set => ChannelFrequencyAsString = Convert.ToString(value);
            }
            [JsonProperty("frequency.value")]
            private string? ChannelFrequencyAsString { get; set; }          // <HIDDEN>
            [JsonProperty("package.nav_pack")]
            public string? Package { get; set; }                            // CL, SC
            public bool? Dolby { get; set; }                                // CL  
            public bool? VoD { get; set; }                                  // CL  
            public string? Language { get; set; }                           // CL  
            public bool? Stingray { get; set; }                             // CL  
            public bool? Free                                               // SC
            {
                get
                {
                    if (string.IsNullOrEmpty(Package))
                        return null;

                    try
                    {
                        if (Package.Contains(csSmartcardPackageFreeMarker))
                            return true;

                        return null;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            [JsonProperty("authorization.authorized")]
            public bool? Authorized { get; set; }                           // SC
            public bool? Pairing                                            // SC
            {
                get
                {
                    if (string.IsNullOrEmpty(PairingString))
                        return null;

                    try
                    {
                        return !PairingString.Contains(csSmartcardPairingNotRequired);
                    }
                    catch
                    {
                        return null;
                    }
                }
                set => PairingString = value == null ? PairingString = string.Empty : (bool)value ? csSmartcardPairingRequired : csSmartcardPairingNotRequired;
            }
            [JsonProperty("pairing.text")]
            private string? PairingString { get; set; }                     // <HIDDEN>
            [JsonProperty("disruption.text")]
            public string? Issues { get; set; }                             // SC  
            [JsonProperty("codec.value")]
            public string? Codec { get; set; }                              // SC  
            public string? ProviderName { get; set; }                       // BS  
            public bool? CAM { get; set; }                                  // BS  
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
            public int? SPID { get; set; }                                  // BS
            public int? TPID { get; set; }                                  // BS
            public int? RunningStatus { get; set; }                         // BS
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

        private const int NET_FW_IP_PROTOCOL_TCP = 6;
        private const int NET_FW_IP_PROTOCOL_UDP = 17;
        private const int NET_FW_RULE_DIR_IN = 1;
        private const int NET_FW_RULE_DIR_OUT = 2;
        private const int NET_FW_ACTION_ALLOW = 1;

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
            Engima2Ver1 = 1, 
            Engima2Ver2 = 2, 
            Engima2Ver3 = 3, 
            Engima2Ver4 = 4, 
            Engima2Ver5 = 5
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
        /// Global.CreateFirewallRule()
        /// </summary>
        /// <returns>bool</returns>
        public static bool CreateFirewallRule()
        {
            LogTrace("Global.CreateFirewallRule()");

            if (!OperatingSystem.IsWindows())
            {
                lCurrentLogger.Warn("-» No windows operating system - cannot create firewall rules.".Pastel(ConsoleColor.Yellow));
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
                lCurrentLogger.Error($"-» Error on creating firewall rules: {excCurrentException.Message}.".Pastel(ConsoleColor.Red));
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

        /// <summary>
        ///     Global.InitDataContainers()
        /// </summary>
        /// <param name="dtLocalDataTable"></param>
        public static void InitDataContainers(ref DataTable dtLocalDataTable)
        {
            dtLocalDataTable.Columns.Clear();

            var piaCurrentPropertyInfos = typeof(SourceDataTable).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var piCurrentPropertyInfo in piaCurrentPropertyInfos)
            {
                if (!piCurrentPropertyInfo.CanRead)
                    continue;

                dtLocalDataTable.Columns.Add(new DataColumn(piCurrentPropertyInfo.Name, Nullable.GetUnderlyingType(piCurrentPropertyInfo.PropertyType) ?? piCurrentPropertyInfo.PropertyType) { AllowDBNull = true });
            }
        }

        /// <summary>
        /// Global.GetSourceDataFromChannelList()
        /// </summary>
        public static void GetSourceDataFromChannelList()
        {
            DataTable dtLocalProviderChannels = new("SourceData");
            InitDataContainers(ref dtLocalProviderChannels);

            lCurrentLogger.Trace("Global.GetSourceDataFromChannelList()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Get source data from channel list".Pastel(ConsoleColor.Green));

            var piaCurrentPropertyInfo = typeof(SourceDataTable).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var oaCurrentWebRequests = new[]
            {
                (Url: csLiwestChannelListUrlTV, IsRadio: false),
                (Url: csLiwestChannelListUrlRadio, IsRadio: true)
            };

            foreach (var oCurrentWebRequest in oaCurrentWebRequests)
            {
                DataTable dtCurrentRequestDataTable;

                try
                {
                    using var hrmHttpResponseMessage = hcCurrentHttpClient.GetAsync(oCurrentWebRequest.Url).ConfigureAwait(false).GetAwaiter().GetResult();
                    hrmHttpResponseMessage.EnsureSuccessStatusCode();

                    using var sCurrentStream = hrmHttpResponseMessage.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                    dtCurrentRequestDataTable = ConvertLiwestChannelListCSVtoDataTable(sCurrentStream, true, true);
                }
                catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or CsvHelperException)
                {
                    lCurrentLogger.Error($"-» Error on getting {(oCurrentWebRequest.IsRadio ? "radio" : "tv")} channel list: {excCurrentException.Message}".Pastel(ConsoleColor.Red));
                    continue;
                }

                if (dtCurrentRequestDataTable.Columns.Count < (oCurrentWebRequest.IsRadio ? 4 : 7))
                {
                    lCurrentLogger.Error($"-» Unexpected format of {(oCurrentWebRequest.IsRadio ? "radio" : "tv")} channel list ({dtCurrentRequestDataTable.Columns.Count} columns)".Pastel(ConsoleColor.Red));
                    continue;
                }

                foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                {
                    var sdtCurrentSourceDataTableItem = new SourceDataTable
                    {
                        ChannelNumber = Convert.ToString(drCurrentRequestDataRow[0]).ToNullableInt32(),
                        ChannelName = Convert.ToString(drCurrentRequestDataRow[1]),
                        ChannelFrequency = Convert.ToString(drCurrentRequestDataRow[3]).ToFrequencyKHz()
                    };

                    if (!oCurrentWebRequest.IsRadio)
                    {
                        sdtCurrentSourceDataTableItem.Package = Convert.ToString(drCurrentRequestDataRow[2])?.TrimAdvanced();
                        sdtCurrentSourceDataTableItem.Dolby = Convert.ToString(drCurrentRequestDataRow[4])?.ToBoolean();
                        sdtCurrentSourceDataTableItem.VoD = Convert.ToString(drCurrentRequestDataRow[5])?.ToBoolean();
                        sdtCurrentSourceDataTableItem.Language = Convert.ToString(drCurrentRequestDataRow[6])?.TrimAdvanced();
                    }
                    else
                        sdtCurrentSourceDataTableItem.Stingray = Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean();

                    var sdtCurrentSourceDataTableRow = new object?[piaCurrentPropertyInfo.Length];

                    for (var iCurrentSourceDataTableColumn = 0; iCurrentSourceDataTableColumn < piaCurrentPropertyInfo.Length; iCurrentSourceDataTableColumn++)
                    {
                        sdtCurrentSourceDataTableRow[iCurrentSourceDataTableColumn] = piaCurrentPropertyInfo[iCurrentSourceDataTableColumn].GetValue(sdtCurrentSourceDataTableItem) ?? DBNull.Value;
                    }

                    dtLocalProviderChannels.Rows.Add(sdtCurrentSourceDataTableRow);
                }
            }

            dtSourceDataTable.Merge(dtLocalProviderChannels);
        }

        /// <summary>
        ///     Global.GetSourceDataFromSmartcard()
        /// </summary>
        /// <param name="strLocalSmartcardId"></param>
        public static void GetSourceDataFromSmartcard(string? strLocalSmartcardId)
        {
            if (string.IsNullOrEmpty(strLocalSmartcardId))
                return;

            DataTable dtLocalProviderChannels = new("SourceData");
            InitDataContainers(ref dtLocalProviderChannels);

            lCurrentLogger.Trace("Global.GetSourceDataFromSmartcard()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Get source data from smartcard".Pastel(ConsoleColor.Green));

            try
            {
                var joCurrentResponse = SendSmartcardApiRequest(HttpMethod.Post, csSmartcardApiCheckUrl, new FormUrlEncodedContent(new Dictionary<string, string> { { "smartcard_id", strLocalSmartcardId } }));
                var strCurrentStatus = joCurrentResponse?["status"]?.ToString();

                if (strCurrentStatus != csSmartcardApiStatusAccepted)
                {
                    lCurrentLogger.Error($"-» Smartcard check has not been accepted (status: '{strCurrentStatus ?? "<none>"}')".Pastel(ConsoleColor.Red));
                    return;
                }
            }
            catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or JsonException or InvalidOperationException)
            {
                lCurrentLogger.Error($"-» Error on requesting smartcard check: {excCurrentException.Message}".Pastel(ConsoleColor.Red));
                return;
            }

            foreach (var strCurrentType in new[] { csSmartcardApiTypeTV, csSmartcardApiTypeRadio })
            {
                var lsdtCurrentSourceDataTableList = GetSmartcardPrograms(strLocalSmartcardId, strCurrentType);
                if (lsdtCurrentSourceDataTableList == null)
                    continue;

                var piaCurrentPropertyInfo = typeof(SourceDataTable).GetProperties();
                foreach (var sdtiCurrentSourceDataTableItem in lsdtCurrentSourceDataTableList)
                {
                    var drCurrentDataRow = dtLocalProviderChannels.NewRow();
                    foreach (var piCurrentPropertyInfo in piaCurrentPropertyInfo)
                    {
                        if (piCurrentPropertyInfo.CanRead && dtLocalProviderChannels.Columns.Contains(piCurrentPropertyInfo.Name))
                            drCurrentDataRow[piCurrentPropertyInfo.Name] = piCurrentPropertyInfo.GetValue(sdtiCurrentSourceDataTableItem) ?? DBNull.Value;
                    }
                    dtLocalProviderChannels.Rows.Add(drCurrentDataRow);
                }
            }

            dtSourceDataTable.Merge(dtLocalProviderChannels);
        }

        /// <summary>
        ///     Global.GetSmartcardPrograms()
        /// </summary>
        /// <param name="strLocalSmartcardId"></param>
        /// <param name="strLocalType"></param>
        /// <returns>List of programs or null on error / timeout</returns>
        private static List<SourceDataTable>? GetSmartcardPrograms(string strLocalSmartcardId, string strLocalType)
        {
            LogTrace("Global.GetSmartcardPrograms()");

            var strCurrentUrl = string.Format(csSmartcardApiResultUrlFormat, Uri.EscapeDataString(strLocalSmartcardId), strLocalType);
            var swCurrentStopwatch = Stopwatch.StartNew();

            try
            {
                while (true)
                {
                    var joCurrentResponse = SendSmartcardApiRequest(HttpMethod.Get, strCurrentUrl);
                    var strCurrentStatus = joCurrentResponse?["data"]?["status"]?.ToString();

                    if (strCurrentStatus == csSmartcardApiStatusCompleted)
                        return joCurrentResponse!["data"]!["programs"]?.ToObject<List<SourceDataTable>>() ?? [];

                    if (strCurrentStatus != csSmartcardApiStatusProcessing)
                    {
                        lCurrentLogger.Error($"-» Smartcard check for type '{strLocalType}' returned unexpected status '{strCurrentStatus ?? "<none>"}'".Pastel(ConsoleColor.Red));
                        return null;
                    }

                    if (swCurrentStopwatch.Elapsed >= TimeSpan.FromSeconds(ciSmartcardApiPollingTimeoutSeconds))
                    {
                        lCurrentLogger.Error($"-» Smartcard check for type '{strLocalType}' has not been completed within {ciSmartcardApiPollingTimeoutSeconds} seconds".Pastel(ConsoleColor.Red));
                        return null;
                    }

                    Thread.Sleep(TimeSpan.FromMilliseconds(ciSmartcardApiPollingIntervalMilliseconds));
                }
            }
            catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or JsonException or InvalidOperationException)
            {
                lCurrentLogger.Error($"-» Error on getting smartcard programs for type '{strLocalType}': {excCurrentException.Message}".Pastel(ConsoleColor.Red));
                return null;
            }
        }

        /// <summary>
        ///     Global.SendSmartcardApiRequest()
        /// </summary>
        /// <param name="hmLocalHttpMethod"></param>
        /// <param name="strLocalUrl"></param>
        /// <param name="hcLocalHttpContent"></param>
        /// <returns>JObject or null on empty response</returns>
        private static JObject? SendSmartcardApiRequest(HttpMethod hmLocalHttpMethod, string strLocalUrl, HttpContent? hcLocalHttpContent = null)
        {
            using var hrqCurrentHttpRequestMessage = new HttpRequestMessage(hmLocalHttpMethod, strLocalUrl);
            hrqCurrentHttpRequestMessage.Headers.Add(csSmartcardApiKeyHeader, csSmartcardApiKey);
            hrqCurrentHttpRequestMessage.Content = hcLocalHttpContent;

            using var hrmCurrentHttpResponseMessage = hcCurrentHttpClient.SendAsync(hrqCurrentHttpRequestMessage).ConfigureAwait(false).GetAwaiter().GetResult();
            var strCurrentContent = hrmCurrentHttpResponseMessage.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();

            // The api returns its error details (e.g. maintenance window, validation) as json body - pass them on instead of the bare status code
            if (!hrmCurrentHttpResponseMessage.IsSuccessStatusCode)
            {
                string? strCurrentApiMessage = null;

                try
                {
                    var joCurrentErrorResponse = string.IsNullOrWhiteSpace(strCurrentContent) ? null : JObject.Parse(strCurrentContent);
                    strCurrentApiMessage = joCurrentErrorResponse?["message"]?.ToString() ?? joCurrentErrorResponse?["error"]?.ToString();
                }
                catch (JsonException)
                {
                    // Body is no json - fall back to the status code only
                }

                throw new HttpRequestException($"HTTP {(int)hrmCurrentHttpResponseMessage.StatusCode} ({hrmCurrentHttpResponseMessage.ReasonPhrase}){(string.IsNullOrEmpty(strCurrentApiMessage) ? string.Empty : $" - {strCurrentApiMessage}")}", null, hrmCurrentHttpResponseMessage.StatusCode);
            }

            return string.IsNullOrWhiteSpace(strCurrentContent) ? null : JObject.Parse(strCurrentContent);
        }

        /// <summary>
        ///     Global.GetSourceDataFromBlindscan()
        /// </summary>
        /// <param name="strLocalServerIP"></param>
        /// <param name="strLocalFrequencies"></param>
        public static void GetSourceDataFromBlindscan(string? strLocalServerIP, string? strLocalFrequencies)
        {
            if (string.IsNullOrEmpty(strLocalServerIP) || string.IsNullOrEmpty(strLocalFrequencies))
                return;

            DataTable dtLocalProviderChannels = new("SourceData");
            InitDataContainers(ref dtLocalProviderChannels);

            lCurrentLogger.Trace("Global.GetSourceDataFromBlindscan()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Get source data from blindscan (octoscan)".Pastel(ConsoleColor.Green));

            DataTable dtCurrentRequestDataTable;
            using (var osiCurrentOSScanIP = new OSScanIP())
            {
                osiCurrentOSScanIP.Init(strLocalServerIP);

                foreach (var iCurrentFrequency in OctoScan.NET.Global.ParseFrequencies(strLocalFrequencies!))
                    osiCurrentOSScanIP.AddTransponderInfo(OctoScan.NET.Global.CreateTransponderInfo(iCurrentFrequency, ciModulationSystem, ciSymbolRate, ciModulation));

                osiCurrentOSScanIP.Scan(ciBlindscanRetries, ciBlindscanTunerLockGraceTime, csBlindscanServerPort, cbBlindscanIPv6);
                OctoScan.NET.Global.bDone = true;

                dtCurrentRequestDataTable = osiCurrentOSScanIP.ExportServicesToDataTable();
            }

            if (dtCurrentRequestDataTable is { Rows.Count: > 0 })
            {
                var piaCurrentPropertyInfo = typeof(SourceDataTable).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                var sdtCurrentSourceDataTableRow = new object?[piaCurrentPropertyInfo.Length];

                foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                {
                    var sdtiCurrentSourceDataTableItem = new SourceDataTable
                    {
                        ChannelName = Convert.ToString(drCurrentRequestDataRow[2]),
                        ChannelFrequency = Convert.ToString(drCurrentRequestDataRow[0]).ToFrequencyKHz(),
                        ProviderName = Convert.ToString(drCurrentRequestDataRow[1]).TrimAdvanced(),
                        CAM = Convert.ToString(drCurrentRequestDataRow[4])?.ToBoolean(),
                        SID = Convert.ToString(drCurrentRequestDataRow[3]).ToInt32(),
                        TSID = Convert.ToString(drCurrentRequestDataRow[6]).ToInt32(),
                        ONID = Convert.ToString(drCurrentRequestDataRow[7]).ToInt32(),
                        PMT = Convert.ToString(drCurrentRequestDataRow[8]).ToInt32(),
                        PCRPID = Convert.ToString(drCurrentRequestDataRow[9]).ToInt32(),
                        VPID = Convert.ToString(drCurrentRequestDataRow[10]).ToInt32(),
                        VPIDST = Convert.ToString(drCurrentRequestDataRow[11]).TrimAdvanced(),
                        APIDS = Convert.ToString(drCurrentRequestDataRow[12]).TrimAdvanced(),
                        APIDSST = Convert.ToString(drCurrentRequestDataRow[13]).TrimAdvanced(),
                        APIDSLANG = Convert.ToString(drCurrentRequestDataRow[14]).TrimAdvanced(),
                        SPID = Convert.ToString(drCurrentRequestDataRow[15]).ToInt32(),
                        TPID = Convert.ToString(drCurrentRequestDataRow[16]).ToInt32(),
                        RunningStatus = Convert.ToString(drCurrentRequestDataRow[5]).ToInt32()
                    };

                    for (var iCurrentPropertyInfoCounter = 0; iCurrentPropertyInfoCounter < piaCurrentPropertyInfo.Length; iCurrentPropertyInfoCounter++)
                    {
                        sdtCurrentSourceDataTableRow[iCurrentPropertyInfoCounter] = piaCurrentPropertyInfo[iCurrentPropertyInfoCounter].GetValue(sdtiCurrentSourceDataTableItem) ?? DBNull.Value;
                    }

                    dtLocalProviderChannels.Rows.Add(sdtCurrentSourceDataTableRow);
                }
            }

            dtSourceDataTable.Merge(dtLocalProviderChannels);
        }

        /// <summary>
        ///     Global.SaveSourceDataToConsole()
        /// </summary>
        public static void SaveSourceDataToConsole()
        {
            lCurrentLogger.Trace("Global.SaveSourceDataToConsole()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Save source data to console".Pastel(ConsoleColor.Green));

            Console.WriteLine(dtSourceDataTable.SetColumnWidthToColumnCaption().ToPrettyPrintedString());
        }

        /// <summary>
        ///     TODO - Test changes after revamp! - Global.SaveSourceDataToFile()
        /// </summary>
        /// <param name="strLocalSourceDataFilename"></param>
        public static void SaveSourceDataToFile(string? strLocalSourceDataFilename)
        {
            if (!string.IsNullOrEmpty(strLocalSourceDataFilename))
            {
                lCurrentLogger.Trace("Global.SaveSourceDataToFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Save source data to file".Pastel(ConsoleColor.Green));

                File.WriteAllText(GetOutputFilePath(strLocalSourceDataFilename), JsonConvert.SerializeObject(dtSourceDataTable, Formatting.Indented), Encoding.UTF8);
            }
        }

        /// <summary>
        ///     TODO - Test changes after revamp! - Global.LoadSourceDataFromFile()
        /// </summary>
        /// <param name="strLocalSourceDataFilename"></param>
        public static void LoadSourceDataFromFile(string? strLocalSourceDataFilename)
        {
            if (!string.IsNullOrEmpty(strLocalSourceDataFilename))
            {
                lCurrentLogger.Trace("Global.LoadSourceDataFromFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Load source data from file".Pastel(ConsoleColor.Green));

                var lsdtCurrentListSourceDataTable = JsonConvert.DeserializeObject<List<SourceDataTable>>(File.ReadAllText(strLocalSourceDataFilename, Encoding.UTF8));
                if (lsdtCurrentListSourceDataTable is { Count: > 0 })
                {
                    var piaSourceDataTablePropertyInfo = typeof(SourceDataTable).GetProperties();

                    foreach (var sdtCurrentSourceDataTable in lsdtCurrentListSourceDataTable)
                    {
                        var drNewDataRow = dtSourceDataTable.NewRow();
                        foreach (var piSourceDataTablePropertyInfo in piaSourceDataTablePropertyInfo)
                        {
                            if (piSourceDataTablePropertyInfo.GetValue(sdtCurrentSourceDataTable) == null)
                                drNewDataRow[piSourceDataTablePropertyInfo.Name] = DBNull.Value;
                            else
                            {
                                if (piSourceDataTablePropertyInfo.Name == "ChannelName")
                                    drNewDataRow[piSourceDataTablePropertyInfo.Name] = Convert.ToString(piSourceDataTablePropertyInfo.GetValue(sdtCurrentSourceDataTable));
                                else
                                    drNewDataRow[piSourceDataTablePropertyInfo.Name] = piSourceDataTablePropertyInfo.GetValue(sdtCurrentSourceDataTable);
                            }
                        }
                        dtSourceDataTable.Rows.Add(drNewDataRow);
                    }
                }
            }
        }

        /// <summary>
        ///     Global.LoadChannelVariationDictionaryFromFile()
        /// </summary>
        /// <param name="strLocalChannelVariationDictionaryFilename"></param>
        public static void LoadChannelVariationDictionaryFromFile(string? strLocalChannelVariationDictionaryFilename)
        {
            if (!string.IsNullOrEmpty(strLocalChannelVariationDictionaryFilename))
            {
                lCurrentLogger.Trace("Global.LoadChannelVariationDictionaryFromFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Load string variation dictionary from file".Pastel(ConsoleColor.Green));

                svdChannelVariationDictionary = JsonConvert.DeserializeObject<ChannelVariationDictionary>(File.ReadAllText(strLocalChannelVariationDictionaryFilename, Encoding.UTF8));
            }
        }

        /// <summary>
        ///     Global.ExportSourceDataToDVBViewerTransponderFile()
        /// </summary>
        /// <param name="strLocalDVBViewerTransponderFilename"></param>
        public static void ExportSourceDataToDVBViewerTransponderFile(string? strLocalDVBViewerTransponderFilename)
        {
            if (!string.IsNullOrEmpty(strLocalDVBViewerTransponderFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToDVBViewerTransponderFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to DVBViewer transponder file".Pastel(ConsoleColor.Green));

                var strTransponderRootElementName = string.Format(csTransponderRootElementNameFormat, DateTime.Now);

                var dtLocalSourceDataTable = dtSourceDataTable.DefaultView.ToTable(true, "ChannelFrequency");
                dtLocalSourceDataTable = dtLocalSourceDataTable.AsEnumerable().Where(drCurrentDataRow => drCurrentDataRow.GetNullableInt32("ChannelFrequency") is > 0).CopyToDataTableOrEmpty(dtLocalSourceDataTable);
                dtLocalSourceDataTable.DefaultView.Sort = "ChannelFrequency";
                dtLocalSourceDataTable = dtLocalSourceDataTable.DefaultView.ToTable();

                using var swCurrentStreamWriter = new StreamWriter(File.Open(GetOutputFilePath(strLocalDVBViewerTransponderFilename), FileMode.Create));
                swCurrentStreamWriter.WriteLine("[SATTYPE]");
                swCurrentStreamWriter.WriteLine("1={0}", ciOrbitalPosition);
                swCurrentStreamWriter.WriteLine("2={0}", strTransponderRootElementName);
                swCurrentStreamWriter.WriteLine(string.Empty);
                swCurrentStreamWriter.WriteLine("[DVB]");
                swCurrentStreamWriter.WriteLine("0=" + dtLocalSourceDataTable.Rows.Count);

                var iCurrentCounter = 1;
                foreach (DataRow drCurrentDataRow in dtLocalSourceDataTable.Rows)
                {
                    swCurrentStreamWriter.WriteLine("{0}={1},{2},{3},|%{4}", iCurrentCounter, drCurrentDataRow["ChannelFrequency"], ciModulation, ciSymbolRate, ciSpectralInversion);
                    iCurrentCounter++;
                }
            }
        }

        /// <summary>
        ///     Global.ExportSourceDataToDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static void ExportSourceDataToDVBViewerChannelDatabase(string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename)) 
                return;

            lCurrentLogger.Trace("Global.ExportSourceDataToDVBViewerChannelDatabase()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Export source data to DVBViewer channel database file".Pastel(ConsoleColor.Green));

            if (dtSourceDataTable.Rows.Count == 0) 
                return;

            var ldrCurrentExportableDataRows = GetExportableServiceRows("DVBViewer channel database");
            if (ldrCurrentExportableDataRows.Count == 0)
                return;

            using BinaryWriter bwCurrentBinaryWriter = new BinaryWriter(File.Open(GetOutputFilePath(strLocalDVBViewerChannelDatabaseFilename), FileMode.Create));
            var cdhCurrentChannelDatabaseHeader = new ChannelDatabaseHeader
            {
                IDLength = (byte)csDefaultDVBViewerChannelDatabaseID.Length,
                ID = csDefaultDVBViewerChannelDatabaseID.ToCharArray(),
                VersionHi = csDefaultDVBViewerChannelDatabaseVersionHi,
                VersionLo = csDefaultDVBViewerChannelDatabaseVersionLo
            };
            bwCurrentBinaryWriter.Write(Binarize(cdhCurrentChannelDatabaseHeader));

            foreach (var drCurrentDataRow in ldrCurrentExportableDataRows)
            {
                var strCurrentAPIDS = drCurrentDataRow.GetStringOrEmpty("APIDS");
                var strCurrentAPIDSStreamType = drCurrentDataRow.GetStringOrEmpty("APIDSST");
                var strCurrentAPIDSLanguage = drCurrentDataRow.GetStringOrEmpty("APIDSLANG");
                var saCurrentAPIDS = !string.IsNullOrEmpty(strCurrentAPIDS) ? strCurrentAPIDS.Split(',') : [string.Empty];
                var saCurrentAPIDSStreamType = !string.IsNullOrEmpty(strCurrentAPIDSStreamType) ? strCurrentAPIDSStreamType.Split(',') : [string.Empty];
                var saCurrentAPIDSLanguage = !string.IsNullOrEmpty(strCurrentAPIDSLanguage) ? strCurrentAPIDSLanguage.Split(',') : [string.Empty];
                var bParentAPID = true;

                for (var iCurrentCounter = 0; iCurrentCounter < saCurrentAPIDS.Length; iCurrentCounter++)
                {
                    var strCurrentAPID = saCurrentAPIDS[iCurrentCounter];
                    var strCurrentAPIDStreamType = iCurrentCounter < saCurrentAPIDSStreamType.Length ? saCurrentAPIDSStreamType[iCurrentCounter] : string.Empty;
                    var strCurrentAPIDLanguage = iCurrentCounter < saCurrentAPIDSLanguage.Length && !string.IsNullOrEmpty(saCurrentAPIDSLanguage[iCurrentCounter]) ? saCurrentAPIDSLanguage[iCurrentCounter] : csDefaultAudioLanguage;
                    var uiAudioPID = ushort.TryParse(strCurrentAPID.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var uiCurrentAudioPID) ? uiCurrentAudioPID : (ushort)0;

                    var cdcCurrentChannelDatabaseChannel = new ChannelDatabaseChannel
                    {
                        TunerData = new ChannelDatabaseTuner
                        {
                            TunerType = ciTunerType,
                            ChannelFlags = GenerateChannelFlags(drCurrentDataRow.GetBooleanOrDefault("CAM"), false, false, drCurrentDataRow.GetUInt16OrDefault("VPID") > 0, uiAudioPID > 0, !bParentAPID && uiAudioPID > 0),
                            Frequency = drCurrentDataRow.GetUInt32OrDefault("ChannelFrequency"),
                            Symbolrate = ciSymbolRate,
                            PMT_PID = drCurrentDataRow.GetUInt16OrDefault("PMT"),
                            ChannelAVFormat = GenerateChannelAVFormat(strCurrentAPIDStreamType, drCurrentDataRow.GetStringOrEmpty("VPIDST")),
                            ChannelNumber = drCurrentDataRow.GetUInt16OrDefault("ChannelNumber"),
                            PolarityOrModulation = ciModulation,
                            FEC = ciFEC,
                            OrbitalPosition = ciOrbitalPosition,
                            Language = strCurrentAPIDLanguage.ToCharArray(),
                            Audio_PID = uiAudioPID,
                            Video_PID = drCurrentDataRow.GetUInt16OrDefault("VPID"),
                            TransportStream_ID = drCurrentDataRow.GetUInt16OrDefault("TSID"),
                            Teletext_PID = drCurrentDataRow.GetUInt16OrDefault("TPID"),
                            OriginalNetwork_ID = drCurrentDataRow.GetUInt16OrDefault("ONID"),
                            Service_ID = drCurrentDataRow.GetUInt16OrDefault("SID"),
                            PCR_PID = drCurrentDataRow.GetUInt16OrDefault("PCRPID"),
                            ChannelGroup = 0,
                            SatModulation = 0,
                            SatModulationSystem = 0,
                            LNB_LOF = 0,
                            DiSEqCExt = 0,
                            DiSEqC = 0,
                            Tone = 0,
                            Reserved1 = 0,
                            Reserved2 = 0,
                            Reserved4 = 0,
                            Reserved6 = 0,
                            Reserved9 = 0
                        },
                        Root = csCategory.ToBinary25CharArray(),
                        ChannelName = drCurrentDataRow.GetStringOrEmpty("ChannelName").ToBinary25CharArray(),
                        Category = csCategory.ToBinary25CharArray(),
                        Encrypted = drCurrentDataRow.GetBooleanOrDefault("CAM") ? (byte)1 : (byte)0,
                        Reserved = 0
                    };

                    bwCurrentBinaryWriter.Write(Binarize(cdcCurrentChannelDatabaseChannel));
                    bParentAPID = false;
                }
            }
        }

        /// <summary>
        ///     Global.DumpDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static unsafe void DumpDVBViewerChannelDatabase(string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (!string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename) && File.Exists(strLocalDVBViewerChannelDatabaseFilename))
            {
                lCurrentLogger.Trace("Global.DumpDVBViewerChannelDatabase()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Reading channel database file".Pastel(ConsoleColor.Green));

                using BinaryReader brCurrentBinaryReader = new BinaryReader(File.Open(strLocalDVBViewerChannelDatabaseFilename, FileMode.Open, FileAccess.Read, FileShare.Read));

                var baCurrentHeaderBuffer = brCurrentBinaryReader.ReadBytes(Marshal.SizeOf(typeof(ChannelDatabaseHeader)));
                if (baCurrentHeaderBuffer.Length < Marshal.SizeOf(typeof(ChannelDatabaseHeader)))
                {
                    lCurrentLogger.Error("-» Channel database file is too small to contain a valid header".Pastel(ConsoleColor.Red));
                    return;
                }

                fixed (byte* pCurrentHeaderBuffer = baCurrentHeaderBuffer)
                {
                    ChannelDatabaseHeader? dvbvHeader = Marshal.PtrToStructure<ChannelDatabaseHeader>(new IntPtr(pCurrentHeaderBuffer));
                    if (dvbvHeader != null)
                    {
                        lCurrentLogger.Info(("Header information of " + strLocalDVBViewerChannelDatabaseFilename).Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info("-----------------------------------------".Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info(("IDLength  : " + dvbvHeader.IDLength).Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info(("ID        : " + dvbvHeader.ID.FromBinary25CharArrayToString(false)).Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info(("VersionHi : " + dvbvHeader.VersionHi).Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info(("VersionLo : " + dvbvHeader.VersionLo).Pastel(ConsoleColor.Yellow));
                        lCurrentLogger.Info("-----------------------------------------".Pastel(ConsoleColor.Yellow));

                        var iCurrentChannelSize = Marshal.SizeOf(typeof(ChannelDatabaseChannel));
                        byte[] baCurrentChannelBuffer;
                        while ((baCurrentChannelBuffer = brCurrentBinaryReader.ReadBytes(iCurrentChannelSize)).Length > 0)
                        {
                            if (baCurrentChannelBuffer.Length < iCurrentChannelSize)
                            {
                                lCurrentLogger.Warn($"-» Ignored incomplete channel record at the end of the file ({baCurrentChannelBuffer.Length} of {iCurrentChannelSize} bytes)".Pastel(ConsoleColor.Yellow));
                                break;
                            }

                            fixed (byte* pCurrentChannelBuffer = baCurrentChannelBuffer)
                            {
                                ChannelDatabaseChannel? dvbvChannel = Marshal.PtrToStructure<ChannelDatabaseChannel>(new IntPtr(pCurrentChannelBuffer));

                                if (dvbvChannel != null)
                                {
                                    lCurrentLogger.Info(("Root         : " + dvbvChannel.Root?.FromBinary25CharArrayToString(true)).Pastel(ConsoleColor.Yellow));
                                    lCurrentLogger.Info(("Channel Name : " + dvbvChannel.ChannelName?.FromBinary25CharArrayToString(true)).Pastel(ConsoleColor.Yellow));
                                    lCurrentLogger.Info(("Channel Cat  : " + dvbvChannel.Category?.FromBinary25CharArrayToString(true)).Pastel(ConsoleColor.Yellow));
                                    lCurrentLogger.Info(("Service ID   : " + dvbvChannel.TunerData?.Service_ID).Pastel(ConsoleColor.Yellow));
                                    lCurrentLogger.Info("-----------------------------------------".Pastel(ConsoleColor.Yellow));
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        ///     Global.ExportSourceDataToExcelExportFile()
        /// </summary>
        /// <param name="strLocalExcelFilename"></param>
        public static void ExportSourceDataToExcelExportFile(string? strLocalExcelFilename)
        {
            if (!string.IsNullOrEmpty(strLocalExcelFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToExcelExportFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to excel export file".Pastel(ConsoleColor.Green));

                var dtLocalExportSourceData = dtSourceDataTable.Copy();

                if (dtLocalExportSourceData.Rows.Count > 0)
                {
                    var wbCurrentWorkbook = new Workbook(GetOutputFilePath(strLocalExcelFilename), csExcelWorksheetName);
                    var wsCurrentWorksheet = wbCurrentWorkbook.GetWorksheet(csExcelWorksheetName);

                    wbCurrentWorkbook.SetCurrentWorksheet(wsCurrentWorksheet);
                    wsCurrentWorksheet.SetCurrentRowNumber(0);

                    foreach (DataColumn dcCurrentDataColumn in dtLocalExportSourceData.Columns)
                    {
                        wsCurrentWorksheet.AddNextCell(dcCurrentDataColumn.ColumnName);
                    }

                    foreach (DataRow drCurrentDataRow in dtLocalExportSourceData.Rows)
                    {
                        wsCurrentWorksheet.GoToNextRow();

                        foreach (DataColumn dcCurrentDataColumn in dtLocalExportSourceData.Columns)
                        {
                            wsCurrentWorksheet.AddNextCell(drCurrentDataRow[dcCurrentDataColumn]);
                        }
                    }

                    float? fSenderFieldMaximumLength = 0;
                    float? fPaketFieldMaximumLength = 0;

                    var fSenderFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtLocalExportSourceData.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[1])) select Convert.ToString(drCurrentDataRow[1])?.Length).Prepend(0).Max();
                    if (fSenderFieldMaximumLengthCurrent > fSenderFieldMaximumLength)
                        fSenderFieldMaximumLength = fSenderFieldMaximumLengthCurrent;

                    var fPaketFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtLocalExportSourceData.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[2])) select Convert.ToString(drCurrentDataRow[2])?.Length).Prepend(0).Max();
                    if (fPaketFieldMaximumLengthCurrent > fPaketFieldMaximumLength)
                        fPaketFieldMaximumLength = fPaketFieldMaximumLengthCurrent;

                    for (var iCurrentColumnNumber = 0; iCurrentColumnNumber <= wsCurrentWorksheet.GetLastColumnNumber(); iCurrentColumnNumber++)
                    {
                        float? fCurrentMaximumLengthValuesWithHeader = Convert.ToSingle(Convert.ToString(wsCurrentWorksheet.GetCell(iCurrentColumnNumber, 0).Value)!.Length) * EXCEL_CHARACTER_TO_WIDTH_CONSTANT + EXCEL_CHARACTER_TO_WIDTH_CONSTANT_CF;

                        switch (iCurrentColumnNumber)
                        {
                            case 1:
                            {
                                if (fCurrentMaximumLengthValuesWithHeader < fSenderFieldMaximumLength)
                                    fCurrentMaximumLengthValuesWithHeader = fSenderFieldMaximumLength;
                                break;
                            }
                            case 2:
                            {
                                if (fCurrentMaximumLengthValuesWithHeader < fPaketFieldMaximumLength)
                                    fCurrentMaximumLengthValuesWithHeader = fPaketFieldMaximumLength;
                                break;
                            }
                        }

                        wsCurrentWorksheet.SetColumnWidth(GetExcelColumnName(iCurrentColumnNumber + 1), Convert.ToSingle(fCurrentMaximumLengthValuesWithHeader));
                    }

                    wsCurrentWorksheet.SetAutoFilter(0, wsCurrentWorksheet.GetLastColumnNumber());
                    wbCurrentWorkbook.Save();
                }
            }
        }

        /// <summary>
        ///     Global.ExportSourceDataToEnigmaGlobal()
        /// </summary>
        /// <param name="femLocalFlagExportMode"></param>
        /// <param name="strLocalFilename"></param>
        /// <param name="fefLocalExportFilenameFormat"></param>
        public static void ExportSourceDataToEnigmaSelector(FlagExportMode femLocalFlagExportMode, string? strLocalFilename, FlagExportEnigmaFormat? fefLocalExportFilenameFormat)
        {
            switch (femLocalFlagExportMode)
            {
                case FlagExportMode.EnigmaTransponderList:
                {
                    ExportSourceDataToEnigma(strLocalFilename, fefLocalExportFilenameFormat, true, false, false);

                    break;
                }
                case FlagExportMode.EnigmaChannelDatabase:
                {
                    ExportSourceDataToEnigma(strLocalFilename, fefLocalExportFilenameFormat, false, true, false);

                    break;
                }
                case FlagExportMode.EnigmaBouquetsFiles:
                {
                    ExportSourceDataToEnigma(strLocalFilename, fefLocalExportFilenameFormat, false, false, true);

                    break;
                }
            }
        }

        /// <summary>
        ///     Global.ExportSourceDataToEnigma()
        /// </summary>
        /// <param name="strLocalExportFilename"></param>
        /// <param name="fefLocalExportFilenameFormat"></param>
        /// <param name="bLocalExportTransponders"></param>
        /// <param name="bLocalExportChannelDatabase"></param>
        /// <param name="bLocalExportBouquetsFiles"></param>
        public static void ExportSourceDataToEnigma(string? strLocalExportFilename, FlagExportEnigmaFormat? fefLocalExportFilenameFormat, bool bLocalExportTransponders, bool bLocalExportChannelDatabase, bool bLocalExportBouquetsFiles)
        {
            lCurrentLogger.Trace("Global.ExportSourceDataToEnigma()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info($"-» Export source data to enigma Transponders:{bLocalExportTransponders} - ChannelDatabase:{bLocalExportChannelDatabase} - BouquetFiles:{bLocalExportBouquetsFiles}".Pastel(ConsoleColor.Green));

            if (dtSourceDataTable.Rows.Count == 0)
                return;

            // Bouquet files are saved without explicit filename into the current directory, all other files next to the given filename
            var strCurrentExportFilePath = string.IsNullOrEmpty(strLocalExportFilename) ? null : GetOutputFilePath(strLocalExportFilename);
            var strCurrentExportFolderAbsolutePath = strCurrentExportFilePath != null ? Path.GetDirectoryName(strCurrentExportFilePath)! : Directory.GetCurrentDirectory();
            var strCurrentExportFilename = strCurrentExportFilePath != null ? Path.GetFileName(strCurrentExportFilePath) : string.Empty;

            var sioCurrentSettingsIO = new SettingsIO { EditorName = cstrCurrentProjectName, Log = new NullLogger() };
            var isCurrentISettings = sioCurrentSettingsIO.Factory.InitNewSettings();
            isCurrentISettings.Log = sioCurrentSettingsIO.Log;
            isCurrentISettings.SettingsFileName = strCurrentExportFilename;
            isCurrentISettings.SettingsDirectory = strCurrentExportFolderAbsolutePath;

            isCurrentISettings.SettingsVersion = fefLocalExportFilenameFormat switch
            {
                FlagExportEnigmaFormat.Engima2Ver1 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma1Ver1,
                FlagExportEnigmaFormat.Engima2Ver2 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma1Ver2,
                FlagExportEnigmaFormat.Engima2Ver3 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver3,
                FlagExportEnigmaFormat.Engima2Ver4 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver4,
                FlagExportEnigmaFormat.Engima2Ver5 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver5,
                _ => isCurrentISettings.SettingsVersion
            };

            var ixcCurrentIXmlCable = sioCurrentSettingsIO.Factory.InitNewXmlCable();
            ixcCurrentIXmlCable.Name = string.Format(csTransponderRootElementNameFormat, DateTime.Now);
            ixcCurrentIXmlCable.SatFeed = Convert.ToString(ciSatFeed);
            ixcCurrentIXmlCable.Flags = Convert.ToString(ciFEC);
            ixcCurrentIXmlCable.CountryCode = ciCountryCode;

            // Transponders need a valid frequency, TSID and ONID
            var dtCurrentTransponderDataTable = dtSourceDataTable.DefaultView.ToTable(true, "ChannelFrequency", "TSID", "ONID");
            dtCurrentTransponderDataTable = dtCurrentTransponderDataTable.AsEnumerable().Where(drCurrentDataRow => drCurrentDataRow.GetNullableInt32("ChannelFrequency") is > 0 && !drCurrentDataRow.IsNullOrEmptyValue("TSID") && !drCurrentDataRow.IsNullOrEmptyValue("ONID")).CopyToDataTableOrEmpty(dtCurrentTransponderDataTable);
            dtCurrentTransponderDataTable.DefaultView.Sort = "ChannelFrequency";
            dtCurrentTransponderDataTable = dtCurrentTransponderDataTable.DefaultView.ToTable();

            foreach (DataRow drCurrentDataRow in dtCurrentTransponderDataTable.Rows)
            {
                var strCurrentFrequency = drCurrentDataRow.GetStringOrEmpty("ChannelFrequency");

                ixcCurrentIXmlCable.Transponders.Add(new XmlTransponder
                {
                    Frequency = strCurrentFrequency,
                    SymbolRate = Convert.ToString(ciSymbolRate * 1000),
                    FEC = Convert.ToString(ciFEC),
                    Modulation = Convert.ToString(ciModulation)
                });

                isCurrentISettings.Transponders.Add(sioCurrentSettingsIO.Factory.InitNewTransponderDVBC($"{ciDefaultDVBCNamespace}:{drCurrentDataRow.GetInt32OrDefault("TSID").ToHexString4Byte()}:{drCurrentDataRow.GetInt32OrDefault("ONID").ToHexString4Byte()}", $"\tc {strCurrentFrequency}:{ciSymbolRate * 1000}:{ciEnigmaTransponderInversion}:{ciModulation}:{ciFEC}:0:0"));
            }

            isCurrentISettings.Cables.Add(ixcCurrentIXmlCable);
            var ixcioCurrentIXmlCablesIO = sioCurrentSettingsIO.Factory.InitNewXmlCableIO(sioCurrentSettingsIO._fileHelper);

            if (bLocalExportTransponders && strCurrentExportFilePath != null)
                ixcioCurrentIXmlCablesIO.SaveCablesToFile(strCurrentExportFilePath, isCurrentISettings);

            if (!bLocalExportChannelDatabase && !bLocalExportBouquetsFiles)
                return;

            var ifbCurrentIFileBouquetTV = sioCurrentSettingsIO.Factory.InitNewFileBouquet(csEnigmaBouquetNameTV, csEnigmaBouquetFilenameTV);
            var ifbCurrentIFileBouquetRadio = sioCurrentSettingsIO.Factory.InitNewFileBouquet(csEnigmaBouquetNameRadio, csEnigmaBouquetFilenameRadio);

            foreach (var drCurrentDataRow in GetExportableServiceRows("enigma channel database / bouquets"))
            {
                var iCurrentVPID = drCurrentDataRow.GetInt32OrDefault("VPID");
                var iCurrentTPID = drCurrentDataRow.GetInt32OrDefault("TPID");
                var iCurrentPCRPID = drCurrentDataRow.GetInt32OrDefault("PCRPID");
                var bCurrentIsVideoService = iCurrentVPID > 0;
                var strCurrentTSID = drCurrentDataRow.GetInt32OrDefault("TSID").ToHexString4Byte();

                var strCurrentServiceReference = $"{drCurrentDataRow.GetInt32OrDefault("SID").ToHexString4Byte()}:{ciDefaultDVBCNamespace}:{strCurrentTSID}:{drCurrentDataRow.GetInt32OrDefault("ONID").ToHexString4Byte()}:{(bCurrentIsVideoService ? ciEnigmaServiceTypeTV : ciEnigmaServiceTypeRadio)}:{drCurrentDataRow.GetStringOrEmpty("ChannelNumber")}:0";
                var strCurrentServiceFlags = $"p:{drCurrentDataRow.GetStringOrEmpty("ProviderName")}" +
                                             (bCurrentIsVideoService ? $",c:00{iCurrentVPID.ToHexString4Byte()}" : string.Empty) +
                                             (iCurrentTPID > 0 ? $",c:02{iCurrentTPID.ToHexString4Byte()}" : string.Empty) +
                                             (iCurrentPCRPID > 0 ? $",c:03{iCurrentPCRPID.ToHexString4Byte()}" : string.Empty) +
                                             (bCurrentIsVideoService ? ",c:050001" : string.Empty) +
                                             (drCurrentDataRow.GetBooleanOrDefault("CAM") ? ",C:0000" : string.Empty) +
                                             $",{csEnigmaServiceFlags}";

                var isCurrentIService = sioCurrentSettingsIO.Factory.InitNewService(strCurrentServiceReference, drCurrentDataRow.GetStringOrEmpty("ChannelName"), strCurrentServiceFlags);
                isCurrentIService.Transponder = isCurrentISettings.Transponders.FirstOrDefault(x => x.TSID == strCurrentTSID);
                isCurrentISettings.Services.Add(isCurrentIService);

                if (bCurrentIsVideoService)
                    ifbCurrentIFileBouquetTV.BouquetItems.Add(new BouquetItemService(isCurrentIService));
                else
                    ifbCurrentIFileBouquetRadio.BouquetItems.Add(new BouquetItemService(isCurrentIService));
            }

            isCurrentISettings.Bouquets.Add(ifbCurrentIFileBouquetTV);
            isCurrentISettings.Bouquets.Add(ifbCurrentIFileBouquetRadio);

            sioCurrentSettingsIO.Save(strCurrentExportFolderAbsolutePath, isCurrentISettings, strCurrentExportFilename, bLocalExportChannelDatabase, false, bLocalExportTransponders, bLocalExportBouquetsFiles);
        }

        /// <summary>
        ///     Global.ModifyDataRemoveChannels()
        /// </summary>
        /// <param name="strLocalChannelRemovalFile"></param>
        public static void ModifyDataRemoveChannels(string? strLocalChannelRemovalFile)
        {
            if (!string.IsNullOrEmpty(strLocalChannelRemovalFile))
            {
                lCurrentLogger.Trace("Global.ModifyDataRemoveChannels()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Remove channels by channelremoval file".Pastel(ConsoleColor.Green));

                var crdCurrentChannelRemovalDictionary = JsonConvert.DeserializeObject<ChannelRemovalDictionary[]>(File.ReadAllText(strLocalChannelRemovalFile, Encoding.UTF8));
                List<DataRow> ldrCurrentRemovedChannels = [];

                if (dtSourceDataTable.Rows.Count > 0 && crdCurrentChannelRemovalDictionary?.Length > 0)
                {
                    ldrCurrentRemovedChannels.AddRange(from DataRow drCurrentDataRow in dtSourceDataTable.Rows let iecrdCurrentChannelRemovalDictionary = crdCurrentChannelRemovalDictionary.Where(x => x.ChannelName == Convert.ToString(drCurrentDataRow["ChannelName"]) && x.ChannelFrequency.ToFrequencyKHz() == drCurrentDataRow.GetNullableInt32("ChannelFrequency")) where iecrdCurrentChannelRemovalDictionary.Any() select drCurrentDataRow);

                    foreach (var drCurrentDataRow in ldrCurrentRemovedChannels)
                    {
                        dtSourceDataTable.Rows.Remove(drCurrentDataRow);
                    }
                }
            }
        }

        /// <summary>
        ///     Global.ModifyDataSortChannels()
        /// </summary>
        /// <param name="strLocalChannelSortingFile"></param>
        public static void ModifyDataSortChannels(string? strLocalChannelSortingFile)
        {
            if (!string.IsNullOrEmpty(strLocalChannelSortingFile))
            {
                lCurrentLogger.Trace("Global.ModifyDataSortChannels()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Sorting channels by channelsorting file".Pastel(ConsoleColor.Green));

                var csdCurrentChannelSortingDictionary = JsonConvert.DeserializeObject<ChannelSortingDictionary[]>(File.ReadAllText(strLocalChannelSortingFile, Encoding.UTF8));

                if (dtSourceDataTable.Rows.Count > 0 && csdCurrentChannelSortingDictionary?.Length > 0)
                {
                    foreach (DataRow drCurrentDataRow in dtSourceDataTable.Rows)
                    {
                        drCurrentDataRow["ChannelNumber"] = DBNull.Value;

                        var csdCurrentChannelSortingEntry = csdCurrentChannelSortingDictionary.FirstOrDefault(x => x.ChannelName == Convert.ToString(drCurrentDataRow["ChannelName"]) && x.ChannelFrequency.ToFrequencyKHz() == drCurrentDataRow.GetNullableInt32("ChannelFrequency"));
                        if (csdCurrentChannelSortingEntry?.SortNumber.ToNullableInt32() is { } iCurrentSortNumber)
                            drCurrentDataRow["ChannelNumber"] = iCurrentSortNumber;
                    }

                    // Unsorted channels (without sort number) are always placed in front of the sorted ones
                    dtSourceDataTable = dtSourceDataTable.AsEnumerable().OrderBy(r => r.IsNull("ChannelNumber") ? 0 : 1).ThenBy(r => r.Field<int?>("ChannelNumber")).ThenBy(r => r.Field<string>("ChannelName")).CopyToDataTableOrEmpty(dtSourceDataTable);
                }
            }
        }

        /// <summary>
        ///     Global.CleanupSourceData()
        /// </summary>
        /// <param name="fscLocalFlagSourceCleanup"></param>
        public static void CleanupSourceData(FlagSourceCleanup fscLocalFlagSourceCleanup)
        {
            lCurrentLogger.Trace("Global.CleanupSourceData()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by defined operation".Pastel(ConsoleColor.Green));

            switch (fscLocalFlagSourceCleanup)
            {
                case FlagSourceCleanup.RemoveInvalidSIDs:
                    CleanupSourceData_RemoveInvalidSIDs();
                    break;
                case FlagSourceCleanup.GroupAndInvalidate:
                    CleanupSourceData_GroupAndInvalidate();
                    CleanupSourceData_RemoveInvalidSIDs();
                    break;
            }
        }

        /// <summary>
        ///     Global.CleanupSourceData_RemoveInvalidSIDs()
        /// </summary>
        public static void CleanupSourceData_RemoveInvalidSIDs()
        {
            lCurrentLogger.Trace("Global.CleanupSourceData_RemoveInvalidSIDs()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by removing invalid SIDs".Pastel(ConsoleColor.Green));

            if (dtSourceDataTable is not { Rows.Count: > 0 }) 
                return;

            var iCurrentRowCount = dtSourceDataTable.Rows.Count;
            dtSourceDataTable = dtSourceDataTable.AsEnumerable().Where(x => x.Field<int?>("SID") != null).CopyToDataTableOrEmpty(dtSourceDataTable);

            if (dtSourceDataTable.Rows.Count == 0)
                lCurrentLogger.Warn($"-» All {iCurrentRowCount} channels have been removed as none of them contains a SID (blindscan data required)".Pastel(ConsoleColor.Yellow));
        }

        /// <summary>
        ///     Global.CleanupSourceData_GroupAndInvalidate()
        /// </summary>
        public static void CleanupSourceData_GroupAndInvalidate()
        {
            lCurrentLogger.Trace("Global.CleanupSourceData_GroupAndInvalidate()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by grouping and invalidating".Pastel(ConsoleColor.Green));
            
            if (dtSourceDataTable is not { Rows.Count: > 0 })
                return;

            var dtCurrentDataTable = dtSourceDataTable.Clone();
            var saColumnNames = dtSourceDataTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            var iegGroupedRows = dtSourceDataTable.AsEnumerable().GroupBy(drCurrentDataRow => new
            {
                ChannelName = Convert.ToString(drCurrentDataRow["ChannelName"])?.ToUpperInvariant().Trim(),
                ChannelFrequency = drCurrentDataRow["ChannelFrequency"] != DBNull.Value ? Convert.ToInt32(drCurrentDataRow["ChannelFrequency"]) : 0
            });

            foreach (var igCurrentGroupedRow in iegGroupedRows)
            {
                var drCurrentDataRow = dtCurrentDataTable.NewRow();
                foreach (var sCurrentColumnName in saColumnNames)
                {
                    object oFinalCellValue = DBNull.Value;
                    foreach (var drCurrentGroupedDataRow in igCurrentGroupedRow)
                    {
                        if (!drCurrentGroupedDataRow.Table.Columns.Contains(sCurrentColumnName)) 
                            continue;

                        var oCurrentCellValue= drCurrentGroupedDataRow[sCurrentColumnName];
                        if (oCurrentCellValue == DBNull.Value || (oCurrentCellValue is string s && string.IsNullOrWhiteSpace(s)))
                            continue;

                        oFinalCellValue = oCurrentCellValue;
                        break;
                    }

                    drCurrentDataRow[sCurrentColumnName] = oFinalCellValue;
                }

                var oCurrentVPIDValue = drCurrentDataRow["VPID"];
                var oCurrentAPIDValues = drCurrentDataRow["APIDS"];
                var iCurrentVPIDValue = (oCurrentVPIDValue != DBNull.Value) ? Convert.ToInt32(oCurrentVPIDValue) : 0;
                var strCurrentAPIDValues = (oCurrentAPIDValues != DBNull.Value) ? Convert.ToString(oCurrentAPIDValues) ?? string.Empty : string.Empty;
                var bHasValidVideo = iCurrentVPIDValue > 0;
                var bHasValidAudio = !string.IsNullOrEmpty(strCurrentAPIDValues) && strCurrentAPIDValues != "0" && !string.IsNullOrWhiteSpace(strCurrentAPIDValues);

                if (!bHasValidVideo && !bHasValidAudio)
                    continue;

                dtCurrentDataTable.Rows.Add(drCurrentDataRow);
            }

            dtSourceDataTable = dtCurrentDataTable;
        }

        /// <summary>
        ///     Global.ConvertLiwestChannelListCSVtoDataTable()
        /// </summary>
        /// <param name="sLocalStream"></param>
        /// <param name="bLocalParseHeader"></param>
        /// <param name="bLocalSkipEmptyLines"></param>
        public static DataTable ConvertLiwestChannelListCSVtoDataTable(Stream sLocalStream, bool bLocalParseHeader, bool bLocalSkipEmptyLines)
        {
            var ccCurrentCsvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = csLiwestChannelListDelimiter,
                HasHeaderRecord = bLocalParseHeader,
                IgnoreBlankLines = bLocalSkipEmptyLines,
                BadDataFound = null,
                MissingFieldFound = null
            };

            using var srCurrentStreamReader = new StreamReader(sLocalStream, Encoding.UTF8, true);
            using var crCurrentCsvReader = new CsvReader(srCurrentStreamReader, ccCurrentCsvConfiguration);
            using var cdrCurrentCsvDataReader = new CsvDataReader(crCurrentCsvReader);

            var dtCurrentDataTable = new DataTable();
            dtCurrentDataTable.Load(cdrCurrentCsvDataReader);

            return dtCurrentDataTable;
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
        ///     Global.GetExportableServiceRows()
        ///     Returns all rows which contain the service data (SID, TSID, ONID, frequency) required for channel database exports.
        /// </summary>
        /// <param name="strLocalExportName"></param>
        private static List<DataRow> GetExportableServiceRows(string strLocalExportName)
        {
            var ldrCurrentExportableDataRows = dtSourceDataTable.AsEnumerable().Where(drCurrentDataRow => drCurrentDataRow.GetNullableInt32("SID") is > 0 && !drCurrentDataRow.IsNullOrEmptyValue("TSID") && !drCurrentDataRow.IsNullOrEmptyValue("ONID") && drCurrentDataRow.GetNullableInt32("ChannelFrequency") is > 0).ToList();
            var iCurrentSkippedRows = dtSourceDataTable.Rows.Count - ldrCurrentExportableDataRows.Count;

            if (iCurrentSkippedRows > 0)
                lCurrentLogger.Warn($"-» Skipped {iCurrentSkippedRows} channel(s) without service data (SID/TSID/ONID/frequency) for {strLocalExportName} export".Pastel(ConsoleColor.Yellow));

            return ldrCurrentExportableDataRows;
        }

        /// <summary>
        ///    Global.GetExcelColumnName()
        /// </summary>
        /// <param name="iLocalColumnNumber"></param>
        private static string GetExcelColumnName(int iLocalColumnNumber)
        {
            if (iLocalColumnNumber <= 0)
                return string.Empty;

            Span<char> scCurrentCharacterBuffer = stackalloc char[8];
            var iCurrentPosition = scCurrentCharacterBuffer.Length;

            while (iLocalColumnNumber > 0)
            {
                var iCurrentModulo = (iLocalColumnNumber - 1) % 26;
                scCurrentCharacterBuffer[--iCurrentPosition] = (char)('A' + iCurrentModulo);
                iLocalColumnNumber = (iLocalColumnNumber - iCurrentModulo) / 26;
            }

            return new string(scCurrentCharacterBuffer[iCurrentPosition..]);
        }
    }
}