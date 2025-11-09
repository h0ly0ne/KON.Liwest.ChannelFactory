using CsvHelper;
using CsvHelper.Configuration;
using KON.OctoScan.NET;
using Krkadoni.EnigmaSettings;
using Krkadoni.EnigmaSettings.Interfaces;
using NanoXLSX;
using NetFwTypeLib;
using Newtonsoft.Json;
using NLog;
using Pastel;

using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.DVBViewer;
using NullLogger = Krkadoni.EnigmaSettings.NullLogger;

namespace KON.Liwest.ChannelFactory
{
    public static class Global
    {
        public static readonly Logger lCurrentLogger = LogManager.GetCurrentClassLogger();

        public static HttpClient hcCurrentHttpClient = new();
        public static DataTable dtSourceDataTable = new("SourceDataTable");
        public static ChannelVariationDictionary? svdChannelVariationDictionary = new();

        public class SourceDataTable
        {
            public int? ChannelNumber { get; set; }          // CL, SC
            public string? ChannelName { get; set; }         // CL, SC, BS
            public int? ChannelFrequency { get; set; }       // CL, SC, BS
            public string? Package { get; set; }             // CL, SC
            public bool? Dolby { get; set; }                 // CL  
            public bool? VoD { get; set; }                   // CL  
            public string? Language { get; set; }            // CL  
            public bool? Stingray { get; set; }              // CL  
            public bool? Free { get; set; }                  // SC  
            public string? Pairing { get; set; }             // SC  
            public string? Issues { get; set; }              // SC  
            public string? Codec { get; set; }               // SC  
            public string? ProviderName { get; set; }        // BS  
            public bool? CAM { get; set; }                   // BS  
            public int? SID { get; set; }                    // BS  
            public int? TSID { get; set; }                   // BS  
            public int? ONID { get; set; }                   // BS
            public int? PMT { get; set; }                    // BS
            public int? PCRPID { get; set; }                 // BS
            public int? VPID { get; set; }                   // BS
            public string? VPIDST { get; set; }              // BS
            public string? APIDS { get; set; }               // BS
            public string? APIDSST { get; set; }             // BS
            public string? APIDSLANG { get; set; }           // BS
            public int? SPID { get; set; }                   // BS
            public int? TPID { get; set; }                   // BS
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

        public enum FlagMode
        {
            GenerateData = 0,
            ModifyData = 1,
            DumpDVBViewerChannelDatabase = 2
        }

        public enum FlagSource
        {
            ChannelList = 0,
            SmartcardList = 1,
            BlindscanList = 2,
            CombinedList = 3,
            FileList = 4
        }
        public enum FlagTarget
        {
            None = 0,
            FileList = 1
        }

        public enum FlagSourceCleanup
        {
            None = 0,
            RemoveInvalidSIDs = 1,
            GroupAndInvalidate = 2
        }

        public static void InitDataContainers(ref DataTable dtLocalDataTable)
        {
            dtLocalDataTable.Columns.Clear();
            dtLocalDataTable.Columns.Add(new DataColumn("ChannelNumber", typeof(int)) { AllowDBNull = true });     // CL, SC
            dtLocalDataTable.Columns.Add(new DataColumn("ChannelName", typeof(string)) { AllowDBNull = true });    // CL, SC, BS
            dtLocalDataTable.Columns.Add(new DataColumn("ChannelFrequency", typeof(int)) { AllowDBNull = true });  // CL, SC, BS
            dtLocalDataTable.Columns.Add(new DataColumn("Package", typeof(string)) { AllowDBNull = true });        // CL, SC
            dtLocalDataTable.Columns.Add(new DataColumn("Dolby", typeof(bool)) { AllowDBNull = true });            // CL
            dtLocalDataTable.Columns.Add(new DataColumn("VoD", typeof(bool)) { AllowDBNull = true });              // CL
            dtLocalDataTable.Columns.Add(new DataColumn("Language", typeof(string)) { AllowDBNull = true });       // CL
            dtLocalDataTable.Columns.Add(new DataColumn("Stingray", typeof(bool)) { AllowDBNull = true });         // CL
            dtLocalDataTable.Columns.Add(new DataColumn("Free", typeof(bool)) { AllowDBNull = true });             // SC
            dtLocalDataTable.Columns.Add(new DataColumn("Pairing", typeof(string)) { AllowDBNull = true });        // SC
            dtLocalDataTable.Columns.Add(new DataColumn("Issues", typeof(string)) { AllowDBNull = true });         // SC
            dtLocalDataTable.Columns.Add(new DataColumn("Codec", typeof(string)) { AllowDBNull = true });          // SC
            dtLocalDataTable.Columns.Add(new DataColumn("ProviderName", typeof(string)) { AllowDBNull = true });   // BS
            dtLocalDataTable.Columns.Add(new DataColumn("CAM", typeof(bool)) { AllowDBNull = true });              // BS
            dtLocalDataTable.Columns.Add(new DataColumn("SID", typeof(int)) { AllowDBNull = true });               // BS
            dtLocalDataTable.Columns.Add(new DataColumn("TSID", typeof(int)) { AllowDBNull = true });              // BS
            dtLocalDataTable.Columns.Add(new DataColumn("ONID", typeof(int)) { AllowDBNull = true });              // BS
            dtLocalDataTable.Columns.Add(new DataColumn("PMT", typeof(int)) { AllowDBNull = true });               // BS
            dtLocalDataTable.Columns.Add(new DataColumn("PCRPID", typeof(int)) { AllowDBNull = true });            // BS
            dtLocalDataTable.Columns.Add(new DataColumn("VPID", typeof(int)) { AllowDBNull = true });              // BS
            dtLocalDataTable.Columns.Add(new DataColumn("VPIDST", typeof(string)) { AllowDBNull = true });         // BS
            dtLocalDataTable.Columns.Add(new DataColumn("APIDS", typeof(string)) { AllowDBNull = true });          // BS
            dtLocalDataTable.Columns.Add(new DataColumn("APIDSST", typeof(string)) { AllowDBNull = true });        // BS
            dtLocalDataTable.Columns.Add(new DataColumn("APIDSLANG", typeof(string)) { AllowDBNull = true });      // BS
            dtLocalDataTable.Columns.Add(new DataColumn("SPID", typeof(int)) { AllowDBNull = true });              // BS
            dtLocalDataTable.Columns.Add(new DataColumn("TPID", typeof(int)) { AllowDBNull = true });              // BS
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

            var strWebRequestUrlTV = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csvAll&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=tv&tx_liwestchannels_search[type][digital]=digital";
            var strWebRequestUrlRadio = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csv&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=radio&tx_liwestchannels_search[type][digital]=digital";

            var hrmHttpResponseMessageTV = hcCurrentHttpClient.GetAsync(strWebRequestUrlTV).ConfigureAwait(false).GetAwaiter().GetResult();
            using (var sCurrentStream = hrmHttpResponseMessageTV.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            {
                var dtCurrentRequestDataTable = ConvertLiwestChannelListCSVtoDataTable(sCurrentStream, true, true);

                foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                {
                    var drModifiedDataRow = dtLocalProviderChannels.NewRow();
                    drModifiedDataRow["ChannelNumber"] = Convert.ToInt32(drCurrentRequestDataRow[0]);
                    drModifiedDataRow["ChannelName"] = Convert.ToString(drCurrentRequestDataRow[1])?.TrimAdvanced().CheckVariationForPrimary();
                    drModifiedDataRow["ChannelFrequency"] = Convert.ToInt32(drCurrentRequestDataRow[3]);
                    drModifiedDataRow["Package"] = Convert.ToString(drCurrentRequestDataRow[2])?.TrimAdvanced();
                    drModifiedDataRow["Dolby"] = Convert.ToString(drCurrentRequestDataRow[4])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[4])?.ToBoolean() : DBNull.Value;
                    drModifiedDataRow["VoD"] = Convert.ToString(drCurrentRequestDataRow[5])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[5])?.ToBoolean() : DBNull.Value;
                    drModifiedDataRow["Language"] = Convert.ToString(drCurrentRequestDataRow[6])?.TrimAdvanced();
                    drModifiedDataRow["Stingray"] = DBNull.Value;
                    dtLocalProviderChannels.Rows.Add(drModifiedDataRow);
                }
            }

            var hrmHttpResponseMessageRadio = hcCurrentHttpClient.GetAsync(strWebRequestUrlRadio).ConfigureAwait(false).GetAwaiter().GetResult();
            using (var sCurrentStream = hrmHttpResponseMessageRadio.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            {
                var dtCurrentRequestDataTable = ConvertLiwestChannelListCSVtoDataTable(sCurrentStream, true, true);

                foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                {
                    var drModifiedDataRow = dtLocalProviderChannels.NewRow();
                    drModifiedDataRow["ChannelNumber"] = Convert.ToInt32(drCurrentRequestDataRow[0]);
                    drModifiedDataRow["ChannelName"] = Convert.ToString(drCurrentRequestDataRow[1])?.TrimAdvanced().CheckVariationForPrimary();
                    drModifiedDataRow["ChannelFrequency"] = Convert.ToInt32(drCurrentRequestDataRow[3]);
                    drModifiedDataRow["Package"] = DBNull.Value;
                    drModifiedDataRow["Dolby"] = DBNull.Value;
                    drModifiedDataRow["VoD"] = DBNull.Value;
                    drModifiedDataRow["Language"] = DBNull.Value;
                    drModifiedDataRow["Stingray"] = Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() : DBNull.Value;
                    dtLocalProviderChannels.Rows.Add(drModifiedDataRow);
                }
            }

            dtSourceDataTable.Merge(dtLocalProviderChannels);
        }

        /// <summary>
        /// Global.GetSourceDataFromSmartcard()
        /// </summary>
        /// <param name="strLocalSmartcardId"></param>
        public static void GetSourceDataFromSmartcard(string? strLocalSmartcardId)
        {
            if (!string.IsNullOrEmpty(strLocalSmartcardId))
            {
                DataTable dtLocalProviderChannels = new("SourceData");
                InitDataContainers(ref dtLocalProviderChannels);

                lCurrentLogger.Trace("Global.GetSourceDataFromSmartcard()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Get source data from smartcard".Pastel(ConsoleColor.Green));

                var hrmCurrentHttpRequestMessage = new HttpRequestMessage(HttpMethod.Post, $"https://auskunft.liwest.at/index.php?scID={strLocalSmartcardId}#smartcard-abfrage");
                var mfdcCurrentMultipartFormDataContent = new MultipartFormDataContent { { new StringContent(strLocalSmartcardId, Encoding.UTF8, MediaTypeNames.Text.Plain), "scID" }, { new StringContent("&raquo; Jetzt anzeigen", Encoding.UTF8, MediaTypeNames.Text.Plain), "action" } };
                hrmCurrentHttpRequestMessage.Content = mfdcCurrentMultipartFormDataContent;

                var cdhvCurrentContentDispositionHeaderValue = new ContentDispositionHeaderValue("form-data");
                hrmCurrentHttpRequestMessage.Content.Headers.ContentDisposition = cdhvCurrentContentDispositionHeaderValue;
                var hrmCurrentHttpResponseMessage = hcCurrentHttpClient.PostAsync(hrmCurrentHttpRequestMessage.RequestUri?.ToString(), hrmCurrentHttpRequestMessage.Content).ConfigureAwait(false).GetAwaiter().GetResult();
                _ = hrmCurrentHttpResponseMessage.Content.ReadAsStringAsync().Result;

                Thread.Sleep(new TimeSpan(0, 0, 15));

                if (hrmCurrentHttpResponseMessage.IsSuccessStatusCode)
                {
                    var hrmHttpResponseMessageTV = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/sc_query_result.php?scID={strLocalSmartcardId}&pType=TV").ConfigureAwait(false).GetAwaiter().GetResult();
                    _ = hrmHttpResponseMessageTV.Content.ReadAsStringAsync().Result;

                    if (hrmHttpResponseMessageTV.IsSuccessStatusCode)
                    {
                        var hrmHttpResponseMessageTVFile = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/files/{strLocalSmartcardId}_{DateTime.Now:yyyyMMdd}.csv").ConfigureAwait(false).GetAwaiter().GetResult();
                        using var sCurrentStream = hrmHttpResponseMessageTVFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                        var dtCurrentRequestDataTable = new DataTable();
                        dtCurrentRequestDataTable.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream), new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ";" })));

                        foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                        {
                            var drModifiedDataRow = dtLocalProviderChannels.NewRow();
                            drModifiedDataRow["ChannelNumber"] = Convert.ToInt32(drCurrentRequestDataRow[0]);
                            drModifiedDataRow["ChannelName"] = Convert.ToString(drCurrentRequestDataRow[1]).RepairEncoding().TrimAdvanced().CheckVariationForPrimary();
                            drModifiedDataRow["ChannelFrequency"] = Convert.ToString(drCurrentRequestDataRow[6])?.Replace(".00 MHz", string.Empty).ToInt32() * 1000;
                            drModifiedDataRow["Free"] = Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() : DBNull.Value;
                            drModifiedDataRow["Pairing"] = Convert.ToString(drCurrentRequestDataRow[3]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Issues"] = Convert.ToString(drCurrentRequestDataRow[4]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Package"] = Convert.ToString(drCurrentRequestDataRow[5]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Codec"] = Convert.ToString(drCurrentRequestDataRow[7]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            dtLocalProviderChannels.Rows.Add(drModifiedDataRow);
                        }
                    }

                    var hrmHttpResponseMessageRADIO = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/sc_query_result.php?scID={strLocalSmartcardId}&pType=Radio").ConfigureAwait(false).GetAwaiter().GetResult();
                    _ = hrmHttpResponseMessageRADIO.Content.ReadAsStringAsync().Result;

                    if (hrmHttpResponseMessageRADIO.IsSuccessStatusCode)
                    {
                        var hrmHttpResponseMessageRADIOFile = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/files/{strLocalSmartcardId}_{DateTime.Now:yyyyMMdd}.csv").ConfigureAwait(false).GetAwaiter().GetResult();
                        using var sCurrentStream = hrmHttpResponseMessageRADIOFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                        var dtCurrentRequestDataTable = new DataTable();
                        dtCurrentRequestDataTable.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream), new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ";" })));

                        foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                        {
                            var drModifiedDataRow = dtLocalProviderChannels.NewRow();
                            drModifiedDataRow["ChannelNumber"] = Convert.ToInt32(drCurrentRequestDataRow[0]);
                            drModifiedDataRow["ChannelName"] = Convert.ToString(drCurrentRequestDataRow[1]).RepairEncoding().TrimAdvanced().CheckVariationForPrimary();
                            drModifiedDataRow["ChannelFrequency"] = Convert.ToString(drCurrentRequestDataRow[6])?.Replace(".00 MHz", string.Empty).ToInt32() * 1000;
                            drModifiedDataRow["Free"] = Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[2])?.ToBoolean() : DBNull.Value;
                            drModifiedDataRow["Pairing"] = Convert.ToString(drCurrentRequestDataRow[3]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Issues"] = Convert.ToString(drCurrentRequestDataRow[4]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Package"] = Convert.ToString(drCurrentRequestDataRow[5]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            drModifiedDataRow["Codec"] = Convert.ToString(drCurrentRequestDataRow[7]).RepairEncoding().TrimAdvanced().Replace("'", string.Empty);
                            dtLocalProviderChannels.Rows.Add(drModifiedDataRow);
                        }
                    }
                }

                dtSourceDataTable.Merge(dtLocalProviderChannels);
            }
        }

        /// <summary>
        /// Global.GetSourceDataFromBlindscan()
        /// </summary>
        /// <param name="strLocalServerIP"></param>
        /// <param name="strLocalFrequencies"></param>
        public static void GetSourceDataFromBlindscan(string? strLocalServerIP, string? strLocalFrequencies)
        {
            if (!string.IsNullOrEmpty(strLocalServerIP) && !string.IsNullOrEmpty(strLocalFrequencies))
            {
                DataTable dtLocalProviderChannels = new("SourceData");
                InitDataContainers(ref dtLocalProviderChannels);

                lCurrentLogger.Trace("Global.GetSourceDataFromBlindscan()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Get source data from blindscan (octoscan)".Pastel(ConsoleColor.Green));

                using var osiCurrentOSScanIP = new OSScanIP();
                osiCurrentOSScanIP.Init(strLocalServerIP);

                var saFrequencies = strLocalFrequencies.Split(',');
                var lsFrequencies = new List<string>();

                foreach (var strFrequency in saFrequencies)
                {
                    if (strFrequency.Contains('-') && strFrequency[^2] == ':')
                    {
                        var saFrequenciesRange = strFrequency.Split('-');

                        if (saFrequenciesRange.Length == 2)
                        {
                            var iFrequenciesRangeStart = Convert.ToInt32(saFrequenciesRange[0]);
                            var iFrequenciesRangeEnd = Convert.ToInt32(saFrequenciesRange[1].Split(':')[0]);
                            var iFrequenciesRangeStep = Convert.ToInt32(saFrequenciesRange[1].Split(':')[1]);

                            if (iFrequenciesRangeEnd > iFrequenciesRangeStart && (iFrequenciesRangeEnd - iFrequenciesRangeStart) % iFrequenciesRangeStep == 0)
                            {
                                for (var iCurrentFrequency = iFrequenciesRangeStart; iCurrentFrequency <= iFrequenciesRangeEnd; iCurrentFrequency += iFrequenciesRangeStep)
                                {
                                    lsFrequencies.Add(Convert.ToString(iCurrentFrequency));
                                }
                            }
                            else
                            {
                                if (iFrequenciesRangeStart == iFrequenciesRangeEnd)
                                    lsFrequencies.Add(Convert.ToString(iFrequenciesRangeStart));
                                else
                                    lCurrentLogger.Warn("-» Invalid frequency start/end/step".Pastel(ConsoleColor.Yellow));
                            }
                        }
                    }
                    else
                    {
                        lsFrequencies.Add(strFrequency);
                    }
                }

                foreach (var lsFrequency in lsFrequencies)
                {
                    var otiCurrentOSTransponderInfo = new OSTransponderInfo
                    {
                        iFrequency = Convert.ToInt32(lsFrequency),
                        iModulationSystem = ciModulationSystem,
                        iSymbolRate = ciSymbolRate,
                        iModulationType = ciModulation,
                        bUseNetworkInformationTable = true
                    };

                    osiCurrentOSScanIP.AddTransponderInfo(otiCurrentOSTransponderInfo);
                }

                osiCurrentOSScanIP.Scan(lLocalTimeout: 10, iLocalScanRetries: 1, iLocalTransponderRetries: 3);

                var dtCurrentRequestDataTable = osiCurrentOSScanIP.ExportServicesToDataTable();
                foreach (DataRow drCurrentRequestDataRow in dtCurrentRequestDataTable.Rows)
                {
                    var drModifiedDataRow = dtLocalProviderChannels.NewRow();
                    drModifiedDataRow["ChannelName"] = Convert.ToString(drCurrentRequestDataRow[2]).TrimAdvanced().CheckVariationForPrimary();
                    drModifiedDataRow["ChannelFrequency"] = Convert.ToString(drCurrentRequestDataRow[0]).ToInt32() * 1000;
                    drModifiedDataRow["ProviderName"] = Convert.ToString(drCurrentRequestDataRow[1]).TrimAdvanced();
                    drModifiedDataRow["CAM"] = Convert.ToString(drCurrentRequestDataRow[3])?.ToBoolean() != null ? Convert.ToString(drCurrentRequestDataRow[3])?.ToBoolean() : DBNull.Value;
                    drModifiedDataRow["SID"] = Convert.ToString(drCurrentRequestDataRow[4]).ToInt32();
                    drModifiedDataRow["TSID"] = Convert.ToString(drCurrentRequestDataRow[5]).ToInt32();
                    drModifiedDataRow["ONID"] = Convert.ToString(drCurrentRequestDataRow[6]).ToInt32();
                    drModifiedDataRow["PMT"] = Convert.ToString(drCurrentRequestDataRow[7]).ToInt32();
                    drModifiedDataRow["PCRPID"] = Convert.ToString(drCurrentRequestDataRow[8]).ToInt32();
                    drModifiedDataRow["VPID"] = Convert.ToString(drCurrentRequestDataRow[9]).ToInt32();
                    drModifiedDataRow["VPIDST"] = Convert.ToString(drCurrentRequestDataRow[10]).TrimAdvanced();
                    drModifiedDataRow["APIDS"] = Convert.ToString(drCurrentRequestDataRow[11]).TrimAdvanced();
                    drModifiedDataRow["APIDSST"] = Convert.ToString(drCurrentRequestDataRow[12]).TrimAdvanced();
                    drModifiedDataRow["APIDSLANG"] = Convert.ToString(drCurrentRequestDataRow[13]).TrimAdvanced();
                    drModifiedDataRow["SPID"] = Convert.ToString(drCurrentRequestDataRow[14]).ToInt32();
                    drModifiedDataRow["TPID"] = Convert.ToString(drCurrentRequestDataRow[15]).ToInt32();
                    dtLocalProviderChannels.Rows.Add(drModifiedDataRow);
                }

                dtSourceDataTable.Merge(dtLocalProviderChannels);
            }
        }

        /// <summary>
        /// Global.SaveSourceDataToFile()
        /// </summary>
        /// <param name="strLocalSourceDataFilename"></param>
        public static void SaveSourceDataToFile(string? strLocalSourceDataFilename)
        {
            if (!string.IsNullOrEmpty(strLocalSourceDataFilename))
            {
                lCurrentLogger.Trace("Global.SaveSourceDataToFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Save source data to file".Pastel(ConsoleColor.Green));

                File.WriteAllText(strLocalSourceDataFilename, JsonConvert.SerializeObject(dtSourceDataTable, Formatting.Indented), Encoding.UTF8);
            }
        }

        /// <summary>
        /// Global.LoadSourceDataFromFile()
        /// </summary>
        /// <param name="strLocalSourceDataFilename"></param>
        public static void LoadSourceDataFromFile(string? strLocalSourceDataFilename)
        {
            if (!string.IsNullOrEmpty(strLocalSourceDataFilename))
            {
                lCurrentLogger.Trace("Global.LoadSourceDataFromFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Load source data from file".Pastel(ConsoleColor.Green));

                List<SourceDataTable>? lsdtCurrentListSourceDataTable = JsonConvert.DeserializeObject<List<SourceDataTable>>(File.ReadAllText(strLocalSourceDataFilename, Encoding.UTF8));
                if (lsdtCurrentListSourceDataTable is { Count: > 0 })
                {
                    PropertyInfo[] piaSourceDataTablePropertyInfo = typeof(SourceDataTable).GetProperties();

                    foreach (SourceDataTable sdtCurrentSourceDataTable in lsdtCurrentListSourceDataTable)
                    {
                        DataRow drNewDataRow = dtSourceDataTable.NewRow();
                        foreach (var piSourceDataTablePropertyInfo in piaSourceDataTablePropertyInfo)
                        {
                            if (piSourceDataTablePropertyInfo.GetValue(sdtCurrentSourceDataTable) == null)
                                drNewDataRow[piSourceDataTablePropertyInfo.Name] = DBNull.Value;
                            else
                            {
                                if (piSourceDataTablePropertyInfo.Name == "ChannelName")
                                    drNewDataRow[piSourceDataTablePropertyInfo.Name] = Convert.ToString(piSourceDataTablePropertyInfo.GetValue(sdtCurrentSourceDataTable)).TrimAdvanced().CheckVariationForPrimary();
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
        /// Global.LoadChannelVariationDictionaryFromFile()
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
        /// Global.ExportSourceDataToDVBViewerTransponderFile()
        /// </summary>
        /// <param name="strLocalDVBViewerTransponderFilename"></param>
        public static void ExportSourceDataToDVBViewerTransponderFile(string? strLocalDVBViewerTransponderFilename)
        {
            if (!string.IsNullOrEmpty(strLocalDVBViewerTransponderFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToDVBViewerTransponderFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to DVBViewer transponder file".Pastel(ConsoleColor.Green));

                var strTransponderRootElementName = $"LIWEST ({DateTime.Now:yyyy-MM-dd HH:mm})";

                var dtLocalSourceDataTable = dtSourceDataTable.DefaultView.ToTable(true, "ChannelFrequency");
                dtLocalSourceDataTable.DefaultView.Sort = "ChannelFrequency";
                dtLocalSourceDataTable = dtLocalSourceDataTable.DefaultView.ToTable();

                using var swCurrentStreamWriter = new StreamWriter(File.Open(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName) ?? string.Empty, strLocalDVBViewerTransponderFilename), FileMode.Create));
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
        /// Global.ExportSourceDataToDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static void ExportSourceDataToDVBViewerChannelDatabase(string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (!string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToDVBViewerChannelDatabase()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to DVBViewer channel database file".Pastel(ConsoleColor.Green));

                var dtLocalExportSourceData = dtSourceDataTable.Copy();

                using BinaryWriter bwCurrentBinaryWriter = new BinaryWriter(File.Open(strLocalDVBViewerChannelDatabaseFilename, FileMode.Create));
                var cdhCurrentChannelDatabaseHeader = new ChannelDatabaseHeader { IDLength = (byte)csDefaultDVBViewerChannelDatabaseID.Length, ID = csDefaultDVBViewerChannelDatabaseID.ToCharArray(), VersionHi = csDefaultDVBViewerChannelDatabaseVersionHi, VersionLo = csDefaultDVBViewerChannelDatabaseVersionLo };
                bwCurrentBinaryWriter.Write(Binarize(cdhCurrentChannelDatabaseHeader));

                if (dtLocalExportSourceData.Rows.Count > 0)
                {
                    foreach (DataRow drCurrentDataRow in dtLocalExportSourceData.Rows)
                    {
                        var strCurrentAPIDS = Convert.ToString(drCurrentDataRow["APIDS"]);
                        var strCurrentAPIDSST = Convert.ToString(drCurrentDataRow["APIDSST"]);
                        var strCurrentAPIDSLANG = Convert.ToString(drCurrentDataRow["APIDSLANG"]);

                        ChannelDatabaseChannel cdcCurrentChannelDatabaseChannel;
                        if (!string.IsNullOrEmpty(strCurrentAPIDS) && !string.IsNullOrEmpty(strCurrentAPIDSST))
                        {
                            var bParentAPID = true;
                            var iCurrentAPIDCounter = 0;
                            foreach (var iCurrentAPID in strCurrentAPIDS.Split(','))
                            {
                                var strCurrentAPIDC = strCurrentAPIDSST.Split(',')[iCurrentAPIDCounter];

                                cdcCurrentChannelDatabaseChannel = new ChannelDatabaseChannel
                                {
                                    TunerData = new ChannelDatabaseTuner
                                    {
                                        TunerType = ciTunerType,
                                        ChannelFlags = GenerateChannelFlags(Convert.ToBoolean(drCurrentDataRow["CAM"]), false, false, Convert.ToUInt16(drCurrentDataRow["VPID"]) > 0, Convert.ToUInt16(iCurrentAPID) > 0, !bParentAPID),
                                        Frequency = Convert.ToUInt32(drCurrentDataRow["ChannelFrequency"]),
                                        Symbolrate = ciSymbolRate,
                                        PMT_PID = Convert.ToUInt16(drCurrentDataRow["PMT"]),
                                        //ChannelAVFormat = GenerateChannelAVFormat(drCurrentDataRow["Dolby"] != DBNull.Value && Convert.ToBoolean(drCurrentDataRow["Dolby"]) ? AVFormat.AUDIO_AC3 : AVFormat.AUDIO_MPEG, Convert.ToString(drCurrentDataRow["Codec"]) == "MPEG4" ? AVFormat.VIDEO_H264 : AVFormat.VIDEO_MPEG2),
                                        ChannelAVFormat = GenerateChannelAVFormat(strCurrentAPIDC, Convert.ToString(drCurrentDataRow["VPIDST"]) ?? string.Empty),
                                        ChannelNumber = drCurrentDataRow["ChannelNumber"] != DBNull.Value ? Convert.ToUInt16(drCurrentDataRow["ChannelNumber"]) : Convert.ToUInt16(0),
                                        PolarityOrModulation = ciModulation,
                                        FEC = ciFEC,
                                        OrbitalPosition = ciOrbitalPosition,
                                        Language = "mis".ToCharArray(),
                                        Audio_PID = Convert.ToUInt16(iCurrentAPID),
                                        Video_PID = Convert.ToUInt16(drCurrentDataRow["VPID"]),
                                        TransportStream_ID = Convert.ToUInt16(drCurrentDataRow["TSID"]),
                                        Teletext_PID = Convert.ToUInt16(drCurrentDataRow["TPID"]),
                                        OriginalNetwork_ID = Convert.ToUInt16(drCurrentDataRow["ONID"]),
                                        Service_ID = Convert.ToUInt16(drCurrentDataRow["SID"]),
                                        PCR_PID = Convert.ToUInt16(drCurrentDataRow["PCRPID"]),
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
                                    ChannelName = drCurrentDataRow["ChannelName"].ToString().ToBinary25CharArray(),
                                    Category = csCategory.ToBinary25CharArray(),
                                    Encrypted = Convert.ToBoolean(drCurrentDataRow["CAM"]) ? (byte)1 : (byte)0,
                                    Reserved = 0
                                };
                                bwCurrentBinaryWriter.Write(Binarize(cdcCurrentChannelDatabaseChannel));

                                bParentAPID = false;
                                iCurrentAPIDCounter++;
                            }
                        }
                        else
                        {
                            cdcCurrentChannelDatabaseChannel = new ChannelDatabaseChannel
                            {
                                TunerData = new ChannelDatabaseTuner
                                {
                                    TunerType = ciTunerType,
                                    ChannelFlags = GenerateChannelFlags(Convert.ToBoolean(drCurrentDataRow["CAM"]), false, false, Convert.ToUInt16(drCurrentDataRow["VPID"]) > 0, false, false),
                                    Frequency = Convert.ToUInt32(drCurrentDataRow["ChannelFrequency"]),
                                    Symbolrate = ciSymbolRate,
                                    PMT_PID = Convert.ToUInt16(drCurrentDataRow["PMT"]),
                                    //ChannelAVFormat = GenerateChannelAVFormat(drCurrentDataRow["Dolby"] != DBNull.Value && Convert.ToBoolean(drCurrentDataRow["Dolby"]) ? AVFormat.AUDIO_AC3 : AVFormat.AUDIO_MPEG, Convert.ToString(drCurrentDataRow["Codec"]) == "MPEG4" ? AVFormat.VIDEO_H264 : AVFormat.VIDEO_MPEG2),
                                    ChannelAVFormat = GenerateChannelAVFormat(string.Empty, Convert.ToString(drCurrentDataRow["VPIDST"]) ?? string.Empty),
                                    FEC = ciFEC,
                                    ChannelNumber = drCurrentDataRow["ChannelNumber"] != DBNull.Value ? Convert.ToUInt16(drCurrentDataRow["ChannelNumber"]) : Convert.ToUInt16(0),
                                    PolarityOrModulation = ciModulation,
                                    OrbitalPosition = ciOrbitalPosition,
                                    Language = "mis".ToCharArray(),
                                    Video_PID = Convert.ToUInt16(drCurrentDataRow["VPID"]),
                                    TransportStream_ID = Convert.ToUInt16(drCurrentDataRow["TSID"]),
                                    Teletext_PID = Convert.ToUInt16(drCurrentDataRow["TPID"]),
                                    OriginalNetwork_ID = Convert.ToUInt16(drCurrentDataRow["ONID"]),
                                    Service_ID = Convert.ToUInt16(drCurrentDataRow["SID"]),
                                    PCR_PID = Convert.ToUInt16(drCurrentDataRow["PCRPID"]),
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
                                ChannelName = drCurrentDataRow["ChannelName"].ToString().ToBinary25CharArray(),
                                Category = csCategory.ToBinary25CharArray(),
                                Encrypted = Convert.ToByte(Convert.ToBoolean(drCurrentDataRow["CAM"])),
                                Reserved = 0
                            };
                            bwCurrentBinaryWriter.Write(Binarize(cdcCurrentChannelDatabaseChannel));
                        }
                    }
                }
            }

            
        }

        /// <summary>
        /// Global.DumpDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static unsafe void DumpDVBViewerChannelDatabase(string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (!string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename) && File.Exists(strLocalDVBViewerChannelDatabaseFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToDVBViewerChannelDatabase()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Reading channel database file".Pastel(ConsoleColor.Green));

                using BinaryReader brCurrentBinaryReader = new BinaryReader(File.Open(strLocalDVBViewerChannelDatabaseFilename, FileMode.Open));

                fixed (byte* pCurrentHeaderBuffer = brCurrentBinaryReader.ReadBytes(Marshal.SizeOf(typeof(ChannelDatabaseHeader))))
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

                        byte[] baCurrentChannelBuffer;
                        do
                        {
                            baCurrentChannelBuffer = brCurrentBinaryReader.ReadBytes(Marshal.SizeOf(typeof(ChannelDatabaseChannel)));
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
                        } while (baCurrentChannelBuffer.Length > 0);
                    }
                }
            }
        }

        /// <summary>
        /// Global.ExportSourceDataToDVBViewerTransponderFile()
        /// </summary>
        /// <param name="strLocalExcelFilename"></param>
        public static void ExportSourceDataToExcelExportFile(string? strLocalExcelFilename)
        {
            if (!string.IsNullOrEmpty(strLocalExcelFilename))
            {
                lCurrentLogger.Trace("Global.ExportSourceDataToExcelExportFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to excel export file".Pastel(ConsoleColor.Green));

                DataTable dtLocalExportSourceData = dtSourceDataTable.Copy();

                if (dtLocalExportSourceData.Rows.Count > 0)
                {
                    var wbCurrentWorkbook = new Workbook(strLocalExcelFilename, "Channels");
                    var wsCurrentWorksheet = wbCurrentWorkbook.GetWorksheet("Channels");

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

                    var fCorrectionFactorForColumnFilter = 3 * EXCEL_CHARACTER_TO_WIDTH_CONSTANT;
                    float fSenderFieldMaximumLength = 0;
                    float fPaketFieldMaximumLength = 0;

                    var fSenderFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtLocalExportSourceData.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[1])) select Convert.ToString(drCurrentDataRow[1]).Length).Prepend(0).Max();
                    if (fSenderFieldMaximumLengthCurrent > fSenderFieldMaximumLength)
                        fSenderFieldMaximumLength = fSenderFieldMaximumLengthCurrent;

                    var fPaketFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtLocalExportSourceData.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[2])) select Convert.ToString(drCurrentDataRow[2]).Length).Prepend(0).Max();
                    if (fPaketFieldMaximumLengthCurrent > fPaketFieldMaximumLength)
                        fPaketFieldMaximumLength = fPaketFieldMaximumLengthCurrent;

                    for (var iCurrentColumnNumber = 0; iCurrentColumnNumber <= wsCurrentWorksheet.GetLastColumnNumber(); iCurrentColumnNumber++)
                    {
                        var fCurrentMaximumLengthValuesWithHeader = Convert.ToSingle(Convert.ToString(wsCurrentWorksheet.GetCell(iCurrentColumnNumber, 0).Value)!.Length) * EXCEL_CHARACTER_TO_WIDTH_CONSTANT + fCorrectionFactorForColumnFilter;

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

                        wsCurrentWorksheet.SetColumnWidth(iCurrentColumnNumber, fCurrentMaximumLengthValuesWithHeader);
                    }

                    wsCurrentWorksheet.SetAutoFilter(0, wsCurrentWorksheet.GetLastColumnNumber());
                    wbCurrentWorkbook.Save();
                }
            }
        }

        /// <summary>
        /// Global.ExportSourceDataToEnigmaDBFileTransponders()
        /// </summary>
        /// <param name="sioLocalSettingsIO"></param>
        /// <param name="isLocalISettings"></param>
        public static void ExportSourceDataToEnigmaDBFileTransponders(ref SettingsIO sioLocalSettingsIO, ref ISettings isLocalISettings)
        {
            lCurrentLogger.Trace("Global.ExportSourceDataToEnigmaDBFileTransponders()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Export source data to enigma database transponders cables.xml file".Pastel(ConsoleColor.Green));

            var strTransponderRootElementName = $"LIWEST ({DateTime.Now:yyyy-MM-dd HH:mm})";
            XmlCable xcCurrentXMLCable = new XmlCable { Name = strTransponderRootElementName, SatFeed = Convert.ToString(ciSatFeed), Flags = Convert.ToString(ciFEC), CountryCode = ciCountryCode };

            var dtLocalSourceDataTable = dtSourceDataTable.DefaultView.ToTable(true, "ChannelFrequency", "TSID", "ONID");
            dtLocalSourceDataTable.DefaultView.Sort = "ChannelFrequency";
            dtLocalSourceDataTable = dtLocalSourceDataTable.DefaultView.ToTable();

            foreach (DataRow drCurrentDataRow in dtLocalSourceDataTable.Rows)
            {
                xcCurrentXMLCable.Transponders.Add(new XmlTransponder
                {
                    Frequency = Convert.ToString(drCurrentDataRow["ChannelFrequency"]),
                    SymbolRate = Convert.ToString(ciSymbolRate * 1000),
                    FEC = Convert.ToString(ciFEC),
                    Modulation = Convert.ToString(ciModulation)
                });

                isLocalISettings.Transponders.Add(sioLocalSettingsIO.Factory.InitNewTransponderDVBC($"{Convert.ToString("ffff0000")}:{Convert.ToInt32(drCurrentDataRow["TSID"]).ToHexString4Byte()}:{Convert.ToInt32(drCurrentDataRow["ONID"]).ToHexString4Byte()}", $"\tc {Convert.ToString(drCurrentDataRow["ChannelFrequency"])}:{Convert.ToString(ciSymbolRate * 1000)}:2:{Convert.ToString(ciModulation)}:{Convert.ToString(ciFEC)}:0:0"));
            }

            isLocalISettings.Cables.Add(xcCurrentXMLCable);
        }

        /// <summary>
        /// Global.ExportSourceDataToEnigmaDBFile()
        /// </summary>
        /// <param name="strLocalEnigmaDBFilename"></param>
        public static void ExportSourceDataToEnigmaDBFile(string? strLocalEnigmaDBFilename)
        {
            SettingsIO settingsIO = new SettingsIO { EditorName = "KON.Liwest.ChannelFactory", Log = new NullLogger() };

            if (!string.IsNullOrEmpty(strLocalEnigmaDBFilename))
            {
                string strLocalEnigmaDBFilenameFullPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName) ?? string.Empty, strLocalEnigmaDBFilename));
                ISettings settings = settingsIO.Load(strLocalEnigmaDBFilenameFullPath);
                settings.Log = new NullLogger();

                lCurrentLogger.Trace("Global.ExportSourceDataToEnigmaDBFile()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Export source data to enigma database file".Pastel(ConsoleColor.Green));

                DataTable dtLocalExportSourceData = dtSourceDataTable.Copy();

                if (dtLocalExportSourceData.Rows.Count > 0)
                {
                    settings.RemoveAllTransponders();
                    settings.RemoveAllSatellites();
                    settings.RemoveAllCables();

                    ExportSourceDataToEnigmaDBFileTransponders(ref settingsIO, ref settings);

                    IFileBouquet ifbCurrentIFileBouquetTV;
                    if (settings.Bouquets.Cast<IFileBouquet>().Any(x => x.FileName == "userbouquet.favourites.tv"))
                        ifbCurrentIFileBouquetTV = settings.Bouquets.Cast<IFileBouquet>().First(x => x.FileName == "userbouquet.favourites.tv");
                    else
                    {
                        ifbCurrentIFileBouquetTV = settingsIO.Factory.InitNewFileBouquet();
                        ifbCurrentIFileBouquetTV.Name = "TV";
                        ifbCurrentIFileBouquetTV.FileName = "userbouquet.favourites.tv";
                    }

                    IFileBouquet ifbCurrentIFileBouquetRadio;
                    if (settings.Bouquets.Cast<IFileBouquet>().Any(x => x.FileName == "userbouquet.favourites.radio"))
                        ifbCurrentIFileBouquetRadio = settings.Bouquets.Cast<IFileBouquet>().First(x => x.FileName == "userbouquet.favourites.radio");
                    else
                    {
                        ifbCurrentIFileBouquetRadio = settingsIO.Factory.InitNewFileBouquet();
                        ifbCurrentIFileBouquetRadio.Name = "Radio";
                        ifbCurrentIFileBouquetRadio.FileName = "userbouquet.favourites.radio";
                    }

                    foreach (DataRow drCurrentDataRow in dtLocalExportSourceData.Rows)
                    {
                        IService isCurrentIService = settingsIO.Factory.InitNewService($"{Convert.ToInt32(drCurrentDataRow["SID"]).ToHexString4Byte()}:{ciDefaultDVBCNamespace}:{Convert.ToInt32(drCurrentDataRow["TSID"]).ToHexString4Byte()}:{Convert.ToInt32(drCurrentDataRow["ONID"]).ToHexString4Byte()}:{(!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["VPID"])) && Convert.ToString(drCurrentDataRow["VPID"]) != "0" ? "1" : "2")}:{Convert.ToString(drCurrentDataRow["ChannelNumber"])}:0", Convert.ToString(drCurrentDataRow["ChannelName"]) ?? string.Empty, $"p:{Convert.ToString(drCurrentDataRow["ProviderName"]) ?? string.Empty}{(!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["VPID"])) && Convert.ToString(drCurrentDataRow["VPID"]) != "0" ? ",c:00" + Convert.ToInt32(drCurrentDataRow["VPID"]).ToHexString4Byte() : string.Empty)}{(!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["TPID"])) && Convert.ToString(drCurrentDataRow["TPID"]) != "0" ? ",c:02" + Convert.ToInt32(drCurrentDataRow["TPID"]).ToHexString4Byte() : string.Empty)}{(!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["PCRPID"])) && Convert.ToString(drCurrentDataRow["PCRPID"]) != "0" ? ",c:03" + Convert.ToInt32(drCurrentDataRow["PCRPID"]).ToHexString4Byte() : string.Empty)}{(!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["VPID"])) && Convert.ToString(drCurrentDataRow["VPID"]) != "0" ? ",c:050001" : string.Empty)}{(Convert.ToBoolean(drCurrentDataRow["CAM"]) ? ",C:0000" : string.Empty)},f:08");
                        settings.Services.Add(isCurrentIService);

                        if (!string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow["VPID"])) && Convert.ToString(drCurrentDataRow["VPID"]) != "0")
                            ifbCurrentIFileBouquetTV.BouquetItems.Add(new BouquetItemService(isCurrentIService));
                        else
                            ifbCurrentIFileBouquetRadio.BouquetItems.Add(new BouquetItemService(isCurrentIService));
                    }

                    if (settings.Bouquets.Cast<IFileBouquet>().Any(x => x.FileName == "userbouquet.favourites.tv"))
                        settings.Bouquets.Remove(ifbCurrentIFileBouquetTV);

                    if (settings.Bouquets.Cast<IFileBouquet>().Any(x => x.FileName == "userbouquet.favourites.radio"))
                        settings.Bouquets.Remove(ifbCurrentIFileBouquetRadio);

                    settings.Bouquets.Add(ifbCurrentIFileBouquetTV);
                    settings.Bouquets.Add(ifbCurrentIFileBouquetRadio);

                    settingsIO.Save(Path.Combine(Path.GetDirectoryName(strLocalEnigmaDBFilenameFullPath) ?? string.Empty), settings);
                }
            }
        }

        /// <summary>
        /// Global.ModifyDataRemoveChannels()
        /// </summary>
        /// <param name="strLocalChannelRemovalFile"></param>
        public static void ModifyDataRemoveChannels(string? strLocalChannelRemovalFile)
        {
            if (!string.IsNullOrEmpty(strLocalChannelRemovalFile))
            {
                lCurrentLogger.Trace("Global.ModifyDataRemoveChannels()".Pastel(ConsoleColor.Cyan));
                lCurrentLogger.Info("-» Remove channels by channelremoval file".Pastel(ConsoleColor.Green));

                var crdCurrentChannelRemovalDictionary = JsonConvert.DeserializeObject<ChannelRemovalDictionary[]>(File.ReadAllText(strLocalChannelRemovalFile, Encoding.UTF8));
                List<DataRow> ldrLocalRemovedChannels = [];

                if (dtSourceDataTable.Rows.Count > 0 && crdCurrentChannelRemovalDictionary?.Length > 0)
                {
                    ldrLocalRemovedChannels.AddRange(from DataRow drCurrentDataRow in dtSourceDataTable.Rows let iecrdCurrentChannelRemovalDictionary = crdCurrentChannelRemovalDictionary.Where(x => x.ChannelName == Convert.ToString(drCurrentDataRow["ChannelName"]) && x.ChannelFrequency == Convert.ToString(drCurrentDataRow["ChannelFrequency"])) where iecrdCurrentChannelRemovalDictionary.Any() select drCurrentDataRow);

                    foreach (var drCurrentDataRow in ldrLocalRemovedChannels)
                    {
                        dtSourceDataTable.Rows.Remove(drCurrentDataRow);
                    }
                }
            }
        }

        /// <summary>
        /// Global.ModifyDataSortChannels()
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

                        var iecsdCurrentChannelSortingDictionary = csdCurrentChannelSortingDictionary.Where(x => x.ChannelName == Convert.ToString(drCurrentDataRow["ChannelName"]) && x.ChannelFrequency == Convert.ToString(drCurrentDataRow["ChannelFrequency"]));
                        var currentChannelSortingDictionary = iecsdCurrentChannelSortingDictionary as ChannelSortingDictionary[] ?? iecsdCurrentChannelSortingDictionary.ToArray();
                        if (currentChannelSortingDictionary.Any())
                            drCurrentDataRow["ChannelNumber"] = currentChannelSortingDictionary.First().SortNumber;
                    }

                    dtSourceDataTable = dtSourceDataTable.AsEnumerable().OrderBy(r => r.Field<int?>("ChannelNumber")).ThenBy(r => r.Field<string>("ChannelName")).CopyToDataTable();
                }
            }
        }
        
        /// <summary>
        /// Global.CleanupSourceData()
        /// </summary>
        /// <param name="fscLocalFlagSourceCleanup"></param>
        /// <returns></returns>
        public static DataTable CleanupSourceData(FlagSourceCleanup fscLocalFlagSourceCleanup)
        {
            lCurrentLogger.Trace("Global.CleanupSourceData()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by defined operation".Pastel(ConsoleColor.Green));

            switch (fscLocalFlagSourceCleanup)
            {
                case FlagSourceCleanup.RemoveInvalidSIDs:
                {
                    return CleanupSourceData_RemoveInvalidSIDs(dtSourceDataTable);
                }
                case FlagSourceCleanup.GroupAndInvalidate:
                {
                    return CleanupSourceData_RemoveInvalidSIDs(CleanupSourceData_GroupAndInvalidate(dtSourceDataTable));
                }
            }

            return dtSourceDataTable;
        }

        /// <summary>
        /// Global.CleanupSourceData_RemoveInvalidSIDs()
        /// </summary>
        /// <param name="dtLocalDataTable"></param>
        /// <returns></returns>
        public static DataTable CleanupSourceData_RemoveInvalidSIDs(DataTable dtLocalDataTable)
        {
            lCurrentLogger.Trace("Global.CleanupSourceData_RemoveInvalidSIDs()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by removing invalid SIDs".Pastel(ConsoleColor.Green));

            return dtLocalDataTable.Rows.Count > 0 ? dtLocalDataTable.Copy().Rows.Cast<DataRow>().Where(r => r["SID"] != DBNull.Value).CopyToDataTable() : dtLocalDataTable;
        }

        /// <summary>
        /// Global.CleanupSourceData_GroupAndInvalidate()
        /// </summary>
        /// <param name="dtLocalDataTable"></param>
        /// <returns></returns>
        public static DataTable CleanupSourceData_GroupAndInvalidate(DataTable dtLocalDataTable)
        {
            lCurrentLogger.Trace("Global.CleanupSourceData_GroupAndInvalidate()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Cleanup sourcedata by grouping and invalidating".Pastel(ConsoleColor.Green));

            if (dtLocalDataTable.Rows.Count > 0)
            {
                IEnumerable<DataRow?> ercCurrentEnumerableRowCollection = dtLocalDataTable.Copy().AsEnumerable().GroupBy(gb => new { ChannelName = Convert.ToString(gb["ChannelName"])?.ToUpper(), ChannelFrequency = gb["ChannelFrequency"] }).Select(s =>
                {
                    DataRow drNewDataRow = dtLocalDataTable.Clone().NewRow();

                    drNewDataRow["ChannelNumber"] = s.Where(w => w["ChannelNumber"] != DBNull.Value).Select(ss => ss["ChannelNumber"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["ChannelName"] = s.Where(w => w["ChannelName"] != DBNull.Value).Select(ss => ss["ChannelName"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["ChannelFrequency"] = s.Where(w => w["ChannelFrequency"] != DBNull.Value).Select(ss => ss["ChannelFrequency"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Package"] = s.Where(w => w["Package"] != DBNull.Value).Select(ss => ss["Package"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Dolby"] = s.Where(w => w["Dolby"] != DBNull.Value).Select(ss => ss["Dolby"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["VoD"] = s.Where(w => w["VoD"] != DBNull.Value).Select(ss => ss["VoD"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Language"] = s.Where(w => w["Language"] != DBNull.Value).Select(ss => ss["Language"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Stingray"] = s.Where(w => w["Stingray"] != DBNull.Value).Select(ss => ss["Stingray"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Free"] = s.Where(w => w["Free"] != DBNull.Value).Select(ss => ss["Free"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Pairing"] = s.Where(w => w["Pairing"] != DBNull.Value).Select(ss => ss["Pairing"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Issues"] = s.Where(w => w["Issues"] != DBNull.Value).Select(ss => ss["Issues"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["Codec"] = s.Where(w => w["Codec"] != DBNull.Value).Select(ss => ss["Codec"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["ProviderName"] = s.Where(w => w["ProviderName"] != DBNull.Value).Select(ss => ss["ProviderName"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["CAM"] = s.Where(w => w["CAM"] != DBNull.Value).Select(ss => ss["CAM"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["SID"] = s.Where(w => w["SID"] != DBNull.Value).Select(ss => ss["SID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["TSID"] = s.Where(w => w["TSID"] != DBNull.Value).Select(ss => ss["TSID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["ONID"] = s.Where(w => w["ONID"] != DBNull.Value).Select(ss => ss["ONID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["PMT"] = s.Where(w => w["PMT"] != DBNull.Value).Select(ss => ss["PMT"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["PCRPID"] = s.Where(w => w["PCRPID"] != DBNull.Value).Select(ss => ss["PCRPID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["VPID"] = s.Where(w => w["VPID"] != DBNull.Value).Select(ss => ss["VPID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["VPIDST"] = s.Where(w => w["VPIDST"] != DBNull.Value).Select(ss => ss["VPIDST"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["APIDS"] = s.Where(w => w["APIDS"] != DBNull.Value).Select(ss => ss["APIDS"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["APIDSST"] = s.Where(w => w["APIDSST"] != DBNull.Value).Select(ss => ss["APIDSST"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["APIDSLANG"] = s.Where(w => w["APIDSLANG"] != DBNull.Value).Select(ss => ss["APIDSLANG"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["SPID"] = s.Where(w => w["SPID"] != DBNull.Value).Select(ss => ss["SPID"]).FirstOrDefault() ?? DBNull.Value;
                    drNewDataRow["TPID"] = s.Where(w => w["TPID"] != DBNull.Value).Select(ss => ss["TPID"]).FirstOrDefault() ?? DBNull.Value;

                    if (drNewDataRow["VPID"] == DBNull.Value && drNewDataRow["APIDS"] == DBNull.Value)
                        return null;
                    if (drNewDataRow["VPID"] == DBNull.Value && Convert.ToString(drNewDataRow["APIDS"]) == "0")
                        return null;
                    if (drNewDataRow["VPID"] == DBNull.Value && string.IsNullOrEmpty(Convert.ToString(drNewDataRow["APIDS"])))
                        return null;
                    if (Convert.ToString(drNewDataRow["VPID"]) == "0" && Convert.ToString(drNewDataRow["APIDS"]) == "0")
                        return null;
                    if (Convert.ToString(drNewDataRow["VPID"]) == "0" && drNewDataRow["APIDS"] == DBNull.Value)
                        return null;
                    if (Convert.ToString(drNewDataRow["VPID"]) == "0" && string.IsNullOrEmpty(Convert.ToString(drNewDataRow["APIDS"])))
                        return null;
                    if (string.IsNullOrEmpty(Convert.ToString(drNewDataRow["VPID"])) && string.IsNullOrEmpty(Convert.ToString(drNewDataRow["APIDS"])))
                        return null;
                    if (string.IsNullOrEmpty(Convert.ToString(drNewDataRow["VPID"])) && drNewDataRow["APIDS"] == DBNull.Value)
                        return null;
                    if (string.IsNullOrEmpty(Convert.ToString(drNewDataRow["VPID"])) && Convert.ToString(drNewDataRow["APIDS"]) == "0")
                        return null;

                    return drNewDataRow;
                }).Where(_ => true);

                DataTable dtLocalCleanupSourceData = ercCurrentEnumerableRowCollection.Cast<DataRow>().CopyToDataTable();
                return dtLocalCleanupSourceData.Rows.Count > 0 ? dtLocalCleanupSourceData : dtLocalDataTable;
            }

            return dtLocalDataTable;
        }

        /// <summary>
        /// Global.ConvertLiwestChannelListCSVtoDataTable()
        /// </summary>
        /// <param name="sLocalStream"></param>
        /// <param name="bLocalParseHeader"></param>
        /// <param name="bLocalSkipEmptyLines"></param>
        /// <returns></returns>
        public static DataTable ConvertLiwestChannelListCSVtoDataTable(Stream sLocalStream, bool bLocalParseHeader, bool bLocalSkipEmptyLines)
        {
            var dtCurrentDataTable = new DataTable();
            string[]? saHeaders = null;
            
            using StreamReader srcCurrentStreamReader = new(sLocalStream);

            while (!srcCurrentStreamReader.EndOfStream)
            {
                string? sCurrentRow;
                if (bLocalParseHeader && saHeaders == null)
                {
                    sCurrentRow = srcCurrentStreamReader.ReadLine();

                    if (!string.IsNullOrEmpty(sCurrentRow))
                    {
                        saHeaders = sCurrentRow.Split(',');

                        var iHeaderCounter = 0;
                        foreach (var sHeader in saHeaders)
                        {
                            if (iHeaderCounter == 0)
                                dtCurrentDataTable.Columns.Add("ProviderChannel", typeof(int));
                            else
                                dtCurrentDataTable.Columns.Add(sHeader);

                            iHeaderCounter++;
                        }
                    }
                    else
                    {
                        if (bLocalSkipEmptyLines)
                            continue;
                    }
                }

                sCurrentRow = srcCurrentStreamReader.ReadLine();

                if (!string.IsNullOrEmpty(sCurrentRow) || (string.IsNullOrEmpty(sCurrentRow) && !bLocalSkipEmptyLines))
                {
                    var saRows = sCurrentRow?.Split(',');
                    if (saRows != null)
                    {
                        if (!bLocalParseHeader && dtCurrentDataTable.Columns.Count < 1)
                        {
                            for (var i = 0; i < saRows.Length; i++)
                            {
                                dtCurrentDataTable.Columns.Add(Convert.ToString(i));
                            }
                        }

                        var drCurrentDataRow = dtCurrentDataTable.NewRow();
                        for (var i = 0; i < saRows.Length; i++)
                        {
                            drCurrentDataRow[i] = saRows[i];
                        }

                        dtCurrentDataTable.Rows.Add(drCurrentDataRow);
                    }
                }
            }

            return dtCurrentDataTable;
        }

        /// <summary>
        /// Global.CreateFirewallRule()
        /// </summary>
        /// <returns></returns>
        public static bool CreateFirewallRule()
        {
            lCurrentLogger.Trace("Global.CreateFirewallRule()".Pastel(ConsoleColor.Cyan));

            if (OperatingSystem.IsWindows())
            {
                var tCurrentFwPolicyType = Type.GetTypeFromProgID("HNetCfg.FwPolicy2", false);
                var tCurrentFwRuleType = Type.GetTypeFromProgID("HNetCfg.FWRule", false);

                if (tCurrentFwPolicyType != null && tCurrentFwRuleType != null)
                {
                    var oCurrentFwPolicyObject = Activator.CreateInstance(tCurrentFwPolicyType);
                    var oCurrentFwRuleObject = Activator.CreateInstance(tCurrentFwRuleType);

                    if (oCurrentFwPolicyObject != null && oCurrentFwRuleObject != null)
                    {
                        var infp2CurrentINetFwPolicy2 = (INetFwPolicy2)oCurrentFwPolicyObject;
                        var infrCurrentINetFwRules = infp2CurrentINetFwPolicy2.Rules;

                        if (infrCurrentINetFwRules != null)
                        {
                            try
                            {
                                infrCurrentINetFwRules.Remove(FIREWALL_RULE_APP_NAME + "_OUT_TCP");
                                infrCurrentINetFwRules.Remove(FIREWALL_RULE_APP_NAME + "_OUT_UDP");
                                infrCurrentINetFwRules.Remove(FIREWALL_RULE_APP_NAME + "_IN_TCP");
                                infrCurrentINetFwRules.Remove(FIREWALL_RULE_APP_NAME + "_IN_UDP");

                                var infrCurrentINetFwRule = (INetFwRule)Activator.CreateInstance(tCurrentFwRuleType)!;
                                infrCurrentINetFwRule.Name = FIREWALL_RULE_APP_NAME + "_OUT_TCP";
                                infrCurrentINetFwRule.Description = FIREWALL_RULE_APP_NAME + "_OUT_TCP";
                                infrCurrentINetFwRule.Protocol = 6;
                                infrCurrentINetFwRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                                infrCurrentINetFwRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                                infrCurrentINetFwRule.ApplicationName = Environment.ProcessPath;
                                infrCurrentINetFwRule.Enabled = true;
                                infrCurrentINetFwRule.InterfaceTypes = "All";
                                infrCurrentINetFwRules.Add(infrCurrentINetFwRule);

                                var infrCurrentINetFwRule2 = (INetFwRule)Activator.CreateInstance(tCurrentFwRuleType)!;
                                infrCurrentINetFwRule2.Name = FIREWALL_RULE_APP_NAME + "_IN_TCP";
                                infrCurrentINetFwRule2.Description = FIREWALL_RULE_APP_NAME + "_IN_TCP";
                                infrCurrentINetFwRule2.Protocol = 6;
                                infrCurrentINetFwRule2.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                                infrCurrentINetFwRule2.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                                infrCurrentINetFwRule2.ApplicationName = Environment.ProcessPath;
                                infrCurrentINetFwRule2.Enabled = true;
                                infrCurrentINetFwRule2.InterfaceTypes = "All";
                                infrCurrentINetFwRules.Add(infrCurrentINetFwRule2);

                                var infrCurrentINetFwRule3 = (INetFwRule)Activator.CreateInstance(tCurrentFwRuleType)!;
                                infrCurrentINetFwRule3.Name = FIREWALL_RULE_APP_NAME + "_OUT_UDP";
                                infrCurrentINetFwRule3.Description = FIREWALL_RULE_APP_NAME + "_OUT_UDP";
                                infrCurrentINetFwRule3.Protocol = 17;
                                infrCurrentINetFwRule3.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                                infrCurrentINetFwRule3.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
                                infrCurrentINetFwRule3.ApplicationName = Environment.ProcessPath;
                                infrCurrentINetFwRule3.Enabled = true;
                                infrCurrentINetFwRule3.InterfaceTypes = "All";
                                infrCurrentINetFwRules.Add(infrCurrentINetFwRule3);

                                var infrCurrentINetFwRule4 = (INetFwRule)Activator.CreateInstance(tCurrentFwRuleType)!;
                                infrCurrentINetFwRule4.Name = FIREWALL_RULE_APP_NAME + "_IN_UDP";
                                infrCurrentINetFwRule4.Description = FIREWALL_RULE_APP_NAME + "_IN_UDP";
                                infrCurrentINetFwRule4.Protocol = 17;
                                infrCurrentINetFwRule4.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
                                infrCurrentINetFwRule4.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                                infrCurrentINetFwRule4.ApplicationName = Environment.ProcessPath;
                                infrCurrentINetFwRule4.Enabled = true;
                                infrCurrentINetFwRule4.InterfaceTypes = "All";
                                infrCurrentINetFwRules.Add(infrCurrentINetFwRule4);

                                return true;
                            }
                            catch
                            {
                                return false;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}