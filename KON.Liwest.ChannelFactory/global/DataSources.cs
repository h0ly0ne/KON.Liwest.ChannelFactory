using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text;

using KON.OctoScan.NET;
using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;

using CsvHelper;
using CsvHelper.Configuration;
using DataTablePrettyPrinter;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Fetching, loading and saving of source data
    /// </summary>
    public static class DataSources
    {
        /// <summary>
        /// DataSources.GetSourceDataFromChannelList()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        public static void GetSourceDataFromChannelList(List<SourceDataItem> lsdiLocalSourceData)
        {
            LogTrace("DataSources.GetSourceDataFromChannelList()");
            LogInfo("-» Get source data from channel list");

            var oaCurrentWebRequests = new[]
            {
                (Url: csLiwestChannelListUrlTV, IsRadio: false, MinimumColumns: 7),
                (Url: csLiwestChannelListUrlRadio, IsRadio: true, MinimumColumns: 4)
            };

            foreach (var oCurrentWebRequest in oaCurrentWebRequests)
            {
                var strCurrentListName = oCurrentWebRequest.IsRadio ? "radio" : "tv";
                List<string[]> lsaCurrentRecords;

                try
                {
                    using var hrmHttpResponseMessage = hcCurrentHttpClient.GetAsync(oCurrentWebRequest.Url).ConfigureAwait(false).GetAwaiter().GetResult();
                    hrmHttpResponseMessage.EnsureSuccessStatusCode();

                    using var sCurrentStream = hrmHttpResponseMessage.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                    lsaCurrentRecords = ReadLiwestChannelListCSV(sCurrentStream, oCurrentWebRequest.MinimumColumns);
                }
                catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or CsvHelperException or InvalidDataException)
                {
                    LogError($"-» Error on getting {strCurrentListName} channel list: {excCurrentException.Message}");
                    continue;
                }

                foreach (var saCurrentRecord in lsaCurrentRecords)
                {
                    var sdiCurrentSourceDataItem = new SourceDataItem
                    {
                        ChannelNumber = saCurrentRecord[0].ToNullableInt32(),
                        ChannelName = saCurrentRecord[1],
                        ChannelFrequency = saCurrentRecord[3].ToFrequencyKHz(),
                        Sources = nameof(FlagSource.ChannelList)
                    };

                    if (!oCurrentWebRequest.IsRadio)
                    {
                        sdiCurrentSourceDataItem.Package = saCurrentRecord[2].TrimAdvanced();
                        sdiCurrentSourceDataItem.Dolby = saCurrentRecord[4].ToBoolean();
                        sdiCurrentSourceDataItem.VoD = saCurrentRecord[5].ToBoolean();
                        sdiCurrentSourceDataItem.Language = saCurrentRecord[6].TrimAdvanced();
                    }
                    else
                        sdiCurrentSourceDataItem.Stingray = saCurrentRecord[2].ToBoolean();

                    lsdiLocalSourceData.Add(sdiCurrentSourceDataItem);
                }
            }
        }

        /// <summary>
        /// DataSources.GetSourceDataFromSmartcard()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalSmartcardId"></param>
        public static void GetSourceDataFromSmartcard(List<SourceDataItem> lsdiLocalSourceData, string? strLocalSmartcardId)
        {
            if (string.IsNullOrEmpty(strLocalSmartcardId))
                return;

            LogTrace("DataSources.GetSourceDataFromSmartcard()");
            LogInfo("-» Get source data from smartcard");

            try
            {
                var joCurrentResponse = SendSmartcardApiRequest(HttpMethod.Post, csSmartcardApiCheckUrl, new FormUrlEncodedContent(new Dictionary<string, string> { { "smartcard_id", strLocalSmartcardId } }));
                var strCurrentStatus = joCurrentResponse?["status"]?.ToString();

                if (strCurrentStatus != csSmartcardApiStatusAccepted)
                {
                    LogError($"-» Smartcard check has not been accepted (status: '{strCurrentStatus ?? "<none>"}')");
                    return;
                }
            }
            catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or JsonException or InvalidOperationException)
            {
                LogError($"-» Error on requesting smartcard check: {excCurrentException.Message}");
                return;
            }

            foreach (var strCurrentType in new[] { csSmartcardApiTypeTV, csSmartcardApiTypeRadio })
            {
                var lsdiCurrentPrograms = GetSmartcardPrograms(strLocalSmartcardId, strCurrentType);
                if (lsdiCurrentPrograms == null)
                    continue;

                lsdiCurrentPrograms.ForEach(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.Sources = nameof(FlagSource.SmartcardList));
                lsdiLocalSourceData.AddRange(lsdiCurrentPrograms);
            }
        }

        /// <summary>
        /// DataSources.GetSmartcardPrograms()
        /// </summary>
        /// <param name="strLocalSmartcardId"></param>
        /// <param name="strLocalType"></param>
        /// <returns>List of programs or null on error / timeout</returns>
        private static List<SourceDataItem>? GetSmartcardPrograms(string strLocalSmartcardId, string strLocalType)
        {
            LogTrace("DataSources.GetSmartcardPrograms()");

            var strCurrentUrl = string.Format(csSmartcardApiResultUrlFormat, Uri.EscapeDataString(strLocalSmartcardId), strLocalType);
            var swCurrentStopwatch = Stopwatch.StartNew();

            try
            {
                while (true)
                {
                    var joCurrentResponse = SendSmartcardApiRequest(HttpMethod.Get, strCurrentUrl);
                    var strCurrentStatus = joCurrentResponse?["data"]?["status"]?.ToString();

                    if (strCurrentStatus == csSmartcardApiStatusCompleted)
                        return joCurrentResponse!["data"]!["programs"]?.ToObject<List<SourceDataItem>>() ?? [];

                    if (strCurrentStatus != csSmartcardApiStatusProcessing)
                    {
                        LogError($"-» Smartcard check for type '{strLocalType}' returned unexpected status '{strCurrentStatus ?? "<none>"}'");
                        return null;
                    }

                    if (swCurrentStopwatch.Elapsed >= TimeSpan.FromSeconds(ciSmartcardApiPollingTimeoutSeconds))
                    {
                        LogError($"-» Smartcard check for type '{strLocalType}' has not been completed within {ciSmartcardApiPollingTimeoutSeconds} seconds");
                        return null;
                    }

                    Thread.Sleep(TimeSpan.FromMilliseconds(ciSmartcardApiPollingIntervalMilliseconds));
                }
            }
            catch (Exception excCurrentException) when (excCurrentException is HttpRequestException or TaskCanceledException or JsonException or InvalidOperationException)
            {
                LogError($"-» Error on getting smartcard programs for type '{strLocalType}': {excCurrentException.Message}");
                return null;
            }
        }

        /// <summary>
        /// DataSources.SendSmartcardApiRequest()
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
        /// DataSources.GetSourceDataFromBlindscan()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalServerIP"></param>
        /// <param name="strLocalFrequencies"></param>
        public static void GetSourceDataFromBlindscan(List<SourceDataItem> lsdiLocalSourceData, string? strLocalServerIP, string? strLocalFrequencies)
        {
            if (string.IsNullOrEmpty(strLocalServerIP) || string.IsNullOrEmpty(strLocalFrequencies))
                return;

            LogTrace("DataSources.GetSourceDataFromBlindscan()");
            LogInfo("-» Get source data from blindscan (octoscan)");

            DataTable dtCurrentRequestDataTable;
            using (var osiCurrentOSScanIP = new OSScanIP())
            {
                osiCurrentOSScanIP.Init(strLocalServerIP);

                foreach (var iCurrentFrequency in OctoScan.NET.Global.ParseFrequencies(strLocalFrequencies))
                    osiCurrentOSScanIP.AddTransponderInfo(OctoScan.NET.Global.CreateTransponderInfo(iCurrentFrequency, ciModulationSystem, ciSymbolRate, ciModulation));

                osiCurrentOSScanIP.Scan(ciBlindscanRetries, ciBlindscanTunerLockGraceTime, csBlindscanServerPort, cbBlindscanIPv6);
                OctoScan.NET.Global.bDone = true;

                dtCurrentRequestDataTable = osiCurrentOSScanIP.ExportServicesToDataTable();
            }

            // Columns are addressed by the names of OSScanIP.ExportServicesToDataTable()
            foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
            {
                // Columns missing in older KON.OctoScan.NET versions are read as empty
                string? GetValue(string strLocalColumnName) => dtCurrentRequestDataTable.Columns.Contains(strLocalColumnName) ? Convert.ToString(drCurrentRequestDataRow[strLocalColumnName], CultureInfo.InvariantCulture) : null;
                string? GetText(string strLocalColumnName) => GetValue(strLocalColumnName).TrimAdvanced() is { Length: > 0 } strCurrentText ? strCurrentText : null;

                var sdiCurrentSourceDataItem = new SourceDataItem
                {
                    ChannelName = GetValue("strName"),
                    ChannelFrequency = GetValue("iFrequency").ToFrequencyKHz(),
                    ProviderName = GetValue("strProviderName").TrimAdvanced(),
                    Encrypted = GetValue("bEncrypted")?.ToBoolean(),
                    CASystemIDs = GetText("liCASystemIDs"),
                    SID = GetValue("iServiceID").ToInt32(),
                    TSID = GetValue("iTransportStreamID").ToInt32(),
                    ONID = GetValue("iOriginalNetworkID").ToInt32(),
                    PMT = GetValue("iProgramMapTable").ToInt32(),
                    PCRPID = GetValue("iProgramClockReferencePacketIdentifier").ToInt32(),
                    VPID = GetValue("iVideoPacketIdentifier").ToInt32(),
                    VPIDST = GetValue("byVideoPacketIdentifierStreamType").TrimAdvanced(),
                    APIDS = GetValue("liAudioPacketIdentifiers").TrimAdvanced(),
                    APIDSST = GetValue("lbAudioPacketIdentifiersStreamType").TrimAdvanced(),
                    APIDSLANG = GetValue("lsAudioPacketIdentifiersLanguage").TrimAdvanced(),
                    APIDSCODEC = GetText("lsAudioPacketIdentifiersCodec"),
                    SPID = GetValue("iSubtitlePacketIdentifier").ToInt32(),
                    TPID = GetValue("iTeletextPacketIdentifier").ToInt32(),
                    RunningStatus = GetValue("iRunningStatus").ToInt32(),
                    Sources = nameof(FlagSource.BlindscanList)
                };

                // Dolby from the audio codecs - older KON.OctoScan.NET versions without codecs leave it empty
                sdiCurrentSourceDataItem.Dolby = sdiCurrentSourceDataItem.DolbyFromAudioCodecs;
                sdiCurrentSourceDataItem.Language = sdiCurrentSourceDataItem.LanguageFromAudioLanguages;
                sdiCurrentSourceDataItem.Codec = sdiCurrentSourceDataItem.CodecFromBlindscan;

                lsdiLocalSourceData.Add(sdiCurrentSourceDataItem);
            }
        }

        /// <summary>
        /// DataSources.LoadSourceDataFromFile()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalSourceDataFilename"></param>
        /// <returns>bool - false if the file could not be loaded</returns>
        public static bool LoadSourceDataFromFile(List<SourceDataItem> lsdiLocalSourceData, string? strLocalSourceDataFilename)
        {
            if (string.IsNullOrEmpty(strLocalSourceDataFilename))
                return false;

            LogTrace("DataSources.LoadSourceDataFromFile()");
            LogInfo("-» Load source data from file");

            var lsdiCurrentSourceData = LoadJsonFile<List<SourceDataItem>>(strLocalSourceDataFilename, "source data file");
            if (lsdiCurrentSourceData == null)
                return false;

            lsdiLocalSourceData.AddRange(lsdiCurrentSourceData.Where(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem != null));
            return true;
        }

        /// <summary>
        /// DataSources.SaveSourceDataToFile()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalSourceDataFilename"></param>
        public static void SaveSourceDataToFile(List<SourceDataItem> lsdiLocalSourceData, string? strLocalSourceDataFilename)
        {
            if (string.IsNullOrEmpty(strLocalSourceDataFilename))
                return;

            LogTrace("DataSources.SaveSourceDataToFile()");
            LogInfo("-» Save source data to file");

            File.WriteAllText(GetOutputFilePath(strLocalSourceDataFilename), JsonConvert.SerializeObject(lsdiLocalSourceData, Formatting.Indented), Encoding.UTF8);
        }

        /// <summary>
        /// DataSources.SaveSourceDataToConsole()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        public static void SaveSourceDataToConsole(List<SourceDataItem> lsdiLocalSourceData)
        {
            LogTrace("DataSources.SaveSourceDataToConsole()");
            LogInfo("-» Save source data to console");

            Console.WriteLine(SourceDataItem.ToDataTable(lsdiLocalSourceData).SetColumnWidthToColumnCaption().ToPrettyPrintedString());
        }

        /// <summary>
        /// DataSources.ReadLiwestChannelListCSV()
        /// </summary>
        /// <param name="sLocalStream"></param>
        /// <param name="iLocalMinimumColumns"></param>
        /// <returns>Records (without header) with at least iLocalMinimumColumns fields</returns>
        public static List<string[]> ReadLiwestChannelListCSV(Stream sLocalStream, int iLocalMinimumColumns)
        {
            var ccCurrentCsvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = csLiwestChannelListDelimiter,
                HasHeaderRecord = true,
                IgnoreBlankLines = true,
                BadDataFound = null,
                MissingFieldFound = null
            };

            using var srCurrentStreamReader = new StreamReader(sLocalStream, Encoding.UTF8, true);
            using var crCurrentCsvReader = new CsvReader(srCurrentStreamReader, ccCurrentCsvConfiguration);

            if (!crCurrentCsvReader.Read() || !crCurrentCsvReader.ReadHeader() || crCurrentCsvReader.HeaderRecord is not { } saCurrentHeaderRecord)
                throw new InvalidDataException("Channel list is empty");

            if (saCurrentHeaderRecord.Length < iLocalMinimumColumns)
                throw new InvalidDataException($"Unexpected format of channel list ({saCurrentHeaderRecord.Length} columns)");

            List<string[]> lsaCurrentRecords = [];

            while (crCurrentCsvReader.Read())
            {
                if (crCurrentCsvReader.Parser.Record is { } saCurrentRecord && saCurrentRecord.Length >= iLocalMinimumColumns)
                    lsaCurrentRecords.Add(saCurrentRecord);
            }

            return lsaCurrentRecords;
        }
    }
}
