using System.Data;
using System.Globalization;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text;

using CsvHelper;
using CsvHelper.Configuration;
using KON.OctoScan.NET;
using NanoXLSX;
using NetFwTypeLib;
using NLog;
using Pastel;

using static KON.Liwest.ChannelFactory.Constants;

namespace KON.Liwest.ChannelFactory
{
    public static class Global
    {
        public static readonly Logger lCurrentLogger = LogManager.GetCurrentClassLogger();

        public static HttpClient hcCurrentHttpClient = new();
        public static DataTable? dtProviderChannelsTV;
        public static DataTable? dtProviderChannelsRADIO;

        /// <summary>
        /// Generate Transponder Excel from Liwest channel list
        /// </summary>
        public static void FetchTransponderDataFromLiwestChannelList()
        {
            lCurrentLogger.Trace("Global.FetchTransponderDataFromLiwestChannelList()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Generate Transponder Excel from Liwest channel list".Pastel(ConsoleColor.Green));

            var strWebRequestUrlTV = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csvAll&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=tv&tx_liwestchannels_search[type][digital]=digital";
            var strWebRequestUrlRadio = "https://www.liwest.at/produkte/fernsehen-radio/kanalbelegung?tx_liwestchannels_search[action]=csv&tx_liwestchannels_search[controller]=Channel&tx_liwestchannels_search[product]=radio&tx_liwestchannels_search[type][digital]=digital";

            dtProviderChannelsTV = null;
            dtProviderChannelsRADIO = null;

            var hrmHttpResponseMessageTV = hcCurrentHttpClient.GetAsync(strWebRequestUrlTV).ConfigureAwait(false).GetAwaiter().GetResult();
            using (var sCurrentStream = hrmHttpResponseMessageTV.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            {
                dtProviderChannelsTV = ConvertLiwestChannelListCSVtoDataTable(sCurrentStream, true, true).Copy();
            }

            var hrmHttpResponseMessageRadio = hcCurrentHttpClient.GetAsync(strWebRequestUrlRadio).ConfigureAwait(false).GetAwaiter().GetResult();
            using (var sCurrentStream = hrmHttpResponseMessageRadio.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            {
                dtProviderChannelsRADIO = ConvertLiwestChannelListCSVtoDataTable(sCurrentStream, true, true).Copy();
            }

            DataTable dtProviderChannels = dtProviderChannelsTV.Copy();
            dtProviderChannels.Merge(dtProviderChannelsRADIO);
            dtProviderChannels.DefaultView.Sort = "Frequenz Digital ASC";
            dtProviderChannels = dtProviderChannels.DefaultView.ToTable();

            var strCurrentFilename = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_Liwest_ChannelList.xlsx";
            var wbCurrentWorkbook = new Workbook(strCurrentFilename, "Liwest_ChannelList");
            var wsCurrentWorksheet = wbCurrentWorkbook.GetWorksheet("Liwest_ChannelList");

            wbCurrentWorkbook.SetCurrentWorksheet(wsCurrentWorksheet);
            wsCurrentWorksheet.SetCurrentRowNumber(0);

            foreach (DataColumn dcCurrentDataColumn in dtProviderChannels.Columns)
            {
                wsCurrentWorksheet.AddNextCell(dcCurrentDataColumn.ColumnName);
            }

            foreach (DataRow drCurrentDataRow in dtProviderChannels.Rows)
            {
                wsCurrentWorksheet.GoToNextRow();

                foreach (DataColumn dcCurrentDataColumn in dtProviderChannels.Columns)
                {
                    wsCurrentWorksheet.AddNextCell(drCurrentDataRow[dcCurrentDataColumn]);
                }
            }

            var fCorrectionFactorForColumnFilter = 3 * EXCEL_CHARACTER_TO_WIDTH_CONSTANT;
            float fSenderFieldMaximumLength = 0;
            float fPaketFieldMaximumLength = 0;

            var fSenderFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtProviderChannels.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[1])) select Convert.ToString(drCurrentDataRow[1]).Length).Prepend(0).Max();
            if (fSenderFieldMaximumLengthCurrent > fSenderFieldMaximumLength)
                fSenderFieldMaximumLength = fSenderFieldMaximumLengthCurrent;

            var fPaketFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtProviderChannels.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[2])) select Convert.ToString(drCurrentDataRow[2]).Length).Prepend(0).Max();
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

        /// <summary>
        /// Generate Transponder Excel from Liwest smartcard data
        /// </summary>
        /// <param name="strLocalSmartcardID"></param>
        public static void FetchTransponderDataFromLiwestSmartcard(string? strLocalSmartcardID)
        {
            lCurrentLogger.Trace("Global.FetchTransponderDataFromLiwestSmartcard()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Generate Transponder Excel from Liwest smartcard data".Pastel(ConsoleColor.Green));

            dtProviderChannelsTV = null;
            dtProviderChannelsRADIO = null;

            var hrmCurrentHttpRequestMessage = new HttpRequestMessage(HttpMethod.Post, $"https://auskunft.liwest.at/index.php?scID={strLocalSmartcardID}#smartcard-abfrage");
            var mfdcCurrentMultipartFormDataContent = new MultipartFormDataContent { { new StringContent(strLocalSmartcardID ?? "01234567890", Encoding.UTF8, MediaTypeNames.Text.Plain), "scID" }, { new StringContent("&raquo; Jetzt anzeigen", Encoding.UTF8, MediaTypeNames.Text.Plain), "action" } };
            hrmCurrentHttpRequestMessage.Content = mfdcCurrentMultipartFormDataContent;

            var cdhvCurrentContentDispositionHeaderValue = new ContentDispositionHeaderValue("form-data");
            hrmCurrentHttpRequestMessage.Content.Headers.ContentDisposition = cdhvCurrentContentDispositionHeaderValue;
            var hrmCurrentHttpResponseMessage = hcCurrentHttpClient.PostAsync(hrmCurrentHttpRequestMessage.RequestUri?.ToString(), hrmCurrentHttpRequestMessage.Content).ConfigureAwait(false).GetAwaiter().GetResult();
            _ = hrmCurrentHttpResponseMessage.Content.ReadAsStringAsync().Result;

            Thread.Sleep(new TimeSpan(0, 0, 15));

            if (hrmCurrentHttpResponseMessage.IsSuccessStatusCode)
            {
                var hrmHttpResponseMessageTV = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/sc_query_result.php?scID={strLocalSmartcardID}&pType=TV").ConfigureAwait(false).GetAwaiter().GetResult();
                _ = hrmHttpResponseMessageTV.Content.ReadAsStringAsync().Result;

                if (hrmHttpResponseMessageTV.IsSuccessStatusCode)
                {
                    var hrmHttpResponseMessageTVFile = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/files/{strLocalSmartcardID}_{DateTime.Now:yyyyMMdd}.csv").ConfigureAwait(false).GetAwaiter().GetResult();
                    using var sCurrentStream = hrmHttpResponseMessageTVFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                    dtProviderChannelsTV = new DataTable();
                    dtProviderChannelsTV.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream, Encoding.Latin1), new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = Encoding.Latin1, Delimiter = ";" })));
                }

                var hrmHttpResponseMessageRADIO = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/sc_query_result.php?scID={strLocalSmartcardID}&pType=Radio").ConfigureAwait(false).GetAwaiter().GetResult();
                _ = hrmHttpResponseMessageRADIO.Content.ReadAsStringAsync().Result;

                if (hrmHttpResponseMessageRADIO.IsSuccessStatusCode)
                {
                    var hrmHttpResponseMessageRADIOFile = hcCurrentHttpClient.GetAsync($"https://auskunft.liwest.at/files/{strLocalSmartcardID}_{DateTime.Now:yyyyMMdd}.csv").ConfigureAwait(false).GetAwaiter().GetResult();
                    using var sCurrentStream = hrmHttpResponseMessageRADIOFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                    dtProviderChannelsRADIO = new DataTable();
                    dtProviderChannelsRADIO.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream, Encoding.Latin1), new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = Encoding.Latin1, Delimiter = ";" })));
                }
            }

            DataTable? dtProviderChannels = null;
            
            if (dtProviderChannelsTV != null)
                if (dtProviderChannels != null)
                    dtProviderChannels.Merge(dtProviderChannelsTV);
                else
                    dtProviderChannels = dtProviderChannelsTV.Copy();

            if (dtProviderChannelsRADIO != null)
                if (dtProviderChannels != null)
                    dtProviderChannels.Merge(dtProviderChannelsRADIO);
                else
                    dtProviderChannels = dtProviderChannelsRADIO.Copy();

            if (dtProviderChannels != null)
            {
                dtProviderChannels.DefaultView.Sort = "Nr. ASC";
                dtProviderChannels = dtProviderChannels.DefaultView.ToTable();
                dtProviderChannels.Columns.Remove("Pairing");
                dtProviderChannels.Columns.Remove("Stï¿½g");

                var strCurrentFilename = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_Liwest_Smartcard.xlsx";
                var wbCurrentWorkbook = new Workbook(strCurrentFilename, "Liwest_Smartcard");
                var wsCurrentWorksheet = wbCurrentWorkbook.GetWorksheet("Liwest_Smartcard");

                wbCurrentWorkbook.SetCurrentWorksheet(wsCurrentWorksheet);
                wsCurrentWorksheet.SetCurrentRowNumber(0);

                foreach (DataColumn dcCurrentDataColumn in dtProviderChannels.Columns)
                {
                    wsCurrentWorksheet.AddNextCell(dcCurrentDataColumn.ColumnName);
                }

                foreach (DataRow drCurrentDataRow in dtProviderChannels.Rows)
                {
                    wsCurrentWorksheet.GoToNextRow();

                    foreach (DataColumn dcCurrentDataColumn in dtProviderChannels.Columns)
                    {
                        wsCurrentWorksheet.AddNextCell(drCurrentDataRow[dcCurrentDataColumn]);
                    }
                }

                var fCorrectionFactorForColumnFilter = 3 * EXCEL_CHARACTER_TO_WIDTH_CONSTANT;
                float fProgrammFieldMaximumLength = 0;
                float fPaketFieldMaximumLength = 0;

                var fProgrammFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtProviderChannels.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[1])) select Convert.ToString(drCurrentDataRow[1]).Length).Prepend(0).Max();
                if (fProgrammFieldMaximumLengthCurrent > fProgrammFieldMaximumLength)
                    fProgrammFieldMaximumLength = fProgrammFieldMaximumLengthCurrent;

                var fPaketFieldMaximumLengthCurrent = (from DataRow drCurrentDataRow in dtProviderChannels.Rows where !string.IsNullOrEmpty(Convert.ToString(drCurrentDataRow[3])) select Convert.ToString(drCurrentDataRow[3]).Length).Prepend(0).Max();
                if (fPaketFieldMaximumLengthCurrent > fPaketFieldMaximumLength)
                    fPaketFieldMaximumLength = fPaketFieldMaximumLengthCurrent;

                for (var iCurrentColumnNumber = 0; iCurrentColumnNumber <= wsCurrentWorksheet.GetLastColumnNumber(); iCurrentColumnNumber++)
                {
                    var fCurrentMaximumLengthValuesWithHeader = Convert.ToSingle(Convert.ToString(wsCurrentWorksheet.GetCell(iCurrentColumnNumber, 0).Value)!.Length) * EXCEL_CHARACTER_TO_WIDTH_CONSTANT + fCorrectionFactorForColumnFilter;

                    switch (iCurrentColumnNumber)
                    {
                        case 1:
                        {
                            if (fCurrentMaximumLengthValuesWithHeader < fProgrammFieldMaximumLength)
                                fCurrentMaximumLengthValuesWithHeader = fProgrammFieldMaximumLength;
                            break;
                        }
                        case 3:
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

        /// <summary>
        /// Generate Transponder Excel from Liwest blindscan (octoscan)
        /// </summary>
        /// <param name="strLocalServerIP"></param>
        public static void FetchTransponderDataFromLiwestBlindscan(string? strLocalServerIP)
        {
            lCurrentLogger.Trace("Global.FetchTransponderDataFromLiwestBlindscan()".Pastel(ConsoleColor.Cyan));
            lCurrentLogger.Info("-» Generate Transponder Excel from Liwest blindscan (octoscan)".Pastel(ConsoleColor.Green));

            using var osiCurrentOSScanIP = new OSScanIP();
            osiCurrentOSScanIP.Init(strLocalServerIP);

            var saFrequencies = "50-954:8".Split(',');
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
                            {
                                lsFrequencies.Add(Convert.ToString(iFrequenciesRangeStart));
                            }
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
                    iModulationSystem = 1,
                    iSymbolRate = 6900,
                    iModulationType = 5,
                    bUseNetworkInformationTable = true
                };

                osiCurrentOSScanIP.AddTransponderInfo(otiCurrentOSTransponderInfo);
            }

            osiCurrentOSScanIP.Scan(5);
            osiCurrentOSScanIP.ExportServicesToExcel("Liwest_Blindscan");
        }

        public static void GenerateDVBViewerTransponderFile()
        {
            //Console.WriteLine("Generate: DVBViewerTransponder");

            //string strTransponderRootElementName = string.Format(@"LIWEST ({0})", DateTime.Now.ToString("yyyy-MM-dd"));

            //DataTable dtProviderChannels = null;
            //if (dtProviderChannelsTV != null)
            //    if (dtProviderChannels != null)
            //        dtProviderChannels.Merge(dtProviderChannelsTV);
            //    else
            //        dtProviderChannels = dtProviderChannelsTV.Copy();
            //if (dtProviderChannelsRADIO != null)
            //    if (dtProviderChannels != null)
            //        dtProviderChannels.Merge(dtProviderChannelsRADIO);
            //    else
            //        dtProviderChannels = dtProviderChannelsRADIO.Copy();
            //dtProviderChannels = dtProviderChannels.DefaultView.ToTable(true, "Frequenz Digital");
            //dtProviderChannels.DefaultView.Sort = "Frequenz Digital ASC";
            //dtProviderChannels = dtProviderChannels.DefaultView.ToTable();

            //Console.WriteLine("*-> Exporting provider channels and generate DVBViewer transponder file ...");

            //using (StreamWriter swCurrentStreamWriter = new StreamWriter(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName) ?? string.Empty, csDVBViewerTransponderFile), false))
            //{
            //    swCurrentStreamWriter.WriteLine(@"[SATTYPE]");
            //    swCurrentStreamWriter.WriteLine(@"1={0}", ciCableOrbitalPosition);
            //    swCurrentStreamWriter.WriteLine(@"2={0}", strTransponderRootElementName);
            //    swCurrentStreamWriter.WriteLine(string.Empty);
            //    swCurrentStreamWriter.WriteLine(@"[DVB]");
            //    swCurrentStreamWriter.WriteLine(@"0=" + dtProviderChannels.Rows.Count);

            //    int iCurrentCounter = 1;
            //    foreach (DataRow drCurrentDataRow in dtProviderChannels.Rows)
            //    {
            //        swCurrentStreamWriter.WriteLine(@"{0}={1},{2},{3}", iCurrentCounter, drCurrentDataRow["Frequenz Digital"], ciCablePolarity, ciCableSymbolrate);
            //        iCurrentCounter++;
            //    }
            //}
        }

        public static void GenerateExcelChannelSortFile()
        {
            //Console.WriteLine("Generate: ExcelSortFile");

            //string strChannelsRootElementName = string.Format(@"LIWEST ({0})", DateTime.Now.ToString("yyyy-MM-dd"));
            //DVBViewer dvCurrentDVBViewer = new DVBViewer(ciCableTunerType, strChannelsRootElementName, csCategory, ciCableOrbitalPosition, ciCablePolarity, ciCableSymbolrate, ciCableSatModulation, ciCableSubStreamID);

            //Console.WriteLine("*-> Convert DVBViewer channel file to DVBViewer channel data class ...");
            //dvCurrentDVBViewer.LoadDataFromINI(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName) ?? string.Empty, o.generateExcelChannelSortFilename), o.generateExcelChannelSortInputEncoding);

            //Console.WriteLine("*-> Save DVBViewer channel data class to Excel sort file ...");
            //dvCurrentDVBViewer.SaveToExcel(Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName) ?? string.Empty, csExcelChannelSortFile), strChannelsRootElementName);

            //Console.WriteLine("*-> Convert DVBViewer channel data class to DVBViewer channel datatable ...");
            //DataTable dtDVBViewerChannels = dvCurrentDVBViewer.DataTable().Copy();








            //DataTable dtProviderSmartcardChannelsTV = null;
            //DataTable dtProviderSmartcardChannelsRADIO = null;

            //string strSmartcardID = "01359096687";
            //var request = new HttpRequestMessage(HttpMethod.Post, string.Format("https://auskunft.liwest.at/index.php?scID={0}#smartcard-abfrage", strSmartcardID));
            //var content = new MultipartFormDataContent();
            //content.Add(new StringContent(strSmartcardID, Encoding.UTF8, MediaTypeNames.Text.Plain), "scID");
            //content.Add(new StringContent("&raquo; Jetzt anzeigen", Encoding.UTF8, MediaTypeNames.Text.Plain), "action");
            //request.Content = content;

            //var header = new ContentDispositionHeaderValue("form-data");
            //request.Content.Headers.ContentDisposition = header;
            //var response = hcCurrentHttpClient.PostAsync(request.RequestUri.ToString(), request.Content).ConfigureAwait(false).GetAwaiter().GetResult();
            //var result = response.Content.ReadAsStringAsync().Result;

            //if (response.IsSuccessStatusCode)
            //{
            //    var hrmHttpResponseMessageTV = hcCurrentHttpClient.GetAsync(string.Format("https://auskunft.liwest.at/sc_query_result.php?scID={0}&pType=TV", strSmartcardID)).ConfigureAwait(false).GetAwaiter().GetResult();
            //    var hrmHttpResponseMessageTVResult = hrmHttpResponseMessageTV.Content.ReadAsStringAsync().Result;

            //    if (hrmHttpResponseMessageTV.IsSuccessStatusCode)
            //    {
            //        var hrmHttpResponseMessageTVFile = hcCurrentHttpClient.GetAsync(string.Format("https://auskunft.liwest.at/files/{0}_{1}.csv", strSmartcardID, DateTime.Now.ToString("yyyyMMdd"))).ConfigureAwait(false).GetAwaiter().GetResult();
            //        using (var sCurrentStream = hrmHttpResponseMessageTVFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            //        {
            //            dtProviderSmartcardChannelsTV = new DataTable();
            //            dtProviderSmartcardChannelsTV.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream, Encoding.Latin1), new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = Encoding.Latin1, Delimiter = ";" })));
            //        }
            //    }

            //    var hrmHttpResponseMessageRADIO = hcCurrentHttpClient.GetAsync(string.Format("https://auskunft.liwest.at/sc_query_result.php?scID={0}&pType=Radio", strSmartcardID)).ConfigureAwait(false).GetAwaiter().GetResult();
            //    var hrmHttpResponseMessageRADIOResult = hrmHttpResponseMessageRADIO.Content.ReadAsStringAsync().Result;

            //    if (hrmHttpResponseMessageRADIO.IsSuccessStatusCode)
            //    {
            //        var hrmHttpResponseMessageRADIOFile = hcCurrentHttpClient.GetAsync(string.Format("https://auskunft.liwest.at/files/{0}_{1}.csv", strSmartcardID, DateTime.Now.ToString("yyyyMMdd"))).ConfigureAwait(false).GetAwaiter().GetResult();
            //        using (var sCurrentStream = hrmHttpResponseMessageRADIOFile.Content.ReadAsStreamAsync().ConfigureAwait(false).GetAwaiter().GetResult())
            //        {
            //            dtProviderSmartcardChannelsRADIO = new DataTable();
            //            dtProviderSmartcardChannelsRADIO.Load(new CsvDataReader(new CsvReader(new StreamReader(sCurrentStream, Encoding.Latin1), new CsvConfiguration(CultureInfo.InvariantCulture) { Encoding = Encoding.Latin1, Delimiter = ";" })));
            //        }
            //    }
            //}

            //DataTable dtProviderSmartcardChannels = null;
            //if (dtProviderSmartcardChannelsTV != null)
            //    if (dtProviderSmartcardChannels != null)
            //        dtProviderSmartcardChannels.Merge(dtProviderSmartcardChannelsTV);
            //    else
            //        dtProviderSmartcardChannels = dtProviderSmartcardChannelsTV.Copy();
            //if (dtProviderSmartcardChannelsRADIO != null)
            //    if (dtProviderSmartcardChannels != null)
            //        dtProviderSmartcardChannels.Merge(dtProviderSmartcardChannelsRADIO);
            //    else
            //        dtProviderSmartcardChannels = dtProviderSmartcardChannelsRADIO.Copy();
            //dtProviderSmartcardChannels.DefaultView.Sort = "Nr. ASC";
            //dtProviderSmartcardChannels = dtProviderSmartcardChannels.DefaultView.ToTable();
            //dtProviderSmartcardChannels.Columns.Remove("Pairing");
            //dtProviderSmartcardChannels.Columns.Remove("Stï¿½g");

            //// Does its job - but provider list is incorrect (regarding channel names!!!)
            ////Console.WriteLine("*-> Merge DVBViewer channel datatable with provider channel datatable ...");
            ////DataTable dtProviderChannels = null;
            ////if (dtProviderChannelsTV != null)
            ////    if (dtProviderChannels != null)
            ////        dtProviderChannels.Merge(dtProviderChannelsTV);
            ////    else
            ////        dtProviderChannels = dtProviderChannelsTV.Copy();
            ////if (dtProviderChannelsRADIO != null)
            ////    if (dtProviderChannels != null)
            ////        dtProviderChannels.Merge(dtProviderChannelsRADIO);
            ////    else
            ////        dtProviderChannels = dtProviderChannelsRADIO.Copy();
            ////dtProviderChannels.DefaultView.Sort = "ProviderChannel ASC";
            ////dtProviderChannels = dtProviderChannels.DefaultView.ToTable();

            //var results =
            //    (from drDVBViewerChannelsDataRow in dtDVBViewerChannels.AsEnumerable()
            //     join drProviderSmartcardChannelsDataRow in dtProviderSmartcardChannels.AsEnumerable()
            //     on new { a = drDVBViewerChannelsDataRow["Name"].ToString().ToLower().Trim() } equals new { a = drProviderSmartcardChannelsDataRow["Programm"].ToString().ToLower().Trim() } into dtMergedDataTableTemporary
            //     from drMergedDataTableTemporaryDataRow in dtMergedDataTableTemporary.DefaultIfEmpty()
            //     select new
            //     {
            //         Number = drDVBViewerChannelsDataRow.Field<int>("Number"),
            //         ProviderChannel = Convert.ToInt32((drMergedDataTableTemporaryDataRow == null) ? 0 : drMergedDataTableTemporaryDataRow.Field<string>("Nr.")),
            //         LCN = drDVBViewerChannelsDataRow.Field<int>("LCN"),
            //         TunerType = drDVBViewerChannelsDataRow.Field<int>("TunerType"),
            //         Root = drDVBViewerChannelsDataRow.Field<string>("Root"),
            //         Category = drDVBViewerChannelsDataRow.Field<string>("Category"),
            //         Name = drDVBViewerChannelsDataRow.Field<string>("Name"),
            //         ProviderName = (drMergedDataTableTemporaryDataRow == null) ? "" : drMergedDataTableTemporaryDataRow.Field<string>("Programm"),
            //         OrbitalPos = drDVBViewerChannelsDataRow.Field<int>("OrbitalPos"),
            //         NetworkID = drDVBViewerChannelsDataRow.Field<int>("NetworkID"),
            //         StreamID = drDVBViewerChannelsDataRow.Field<int>("StreamID"),
            //         SID = drDVBViewerChannelsDataRow.Field<int>("SID"),
            //         PMTPID = drDVBViewerChannelsDataRow.Field<int>("PMTPID"),
            //         VPID = drDVBViewerChannelsDataRow.Field<int>("VPID"),
            //         APID = drDVBViewerChannelsDataRow.Field<int>("APID"),
            //         PCRPID = drDVBViewerChannelsDataRow.Field<int>("PCRPID"),
            //         AC3 = drDVBViewerChannelsDataRow.Field<int>("AC3"),
            //         Language = drDVBViewerChannelsDataRow.Field<string>("Language"),
            //         Volume = drDVBViewerChannelsDataRow.Field<int>("Volume"),
            //         EPGFlag = drDVBViewerChannelsDataRow.Field<int>("EPGFlag"),
            //         TelePID = drDVBViewerChannelsDataRow.Field<int>("TelePID"),
            //         AudioChannel = drDVBViewerChannelsDataRow.Field<int>("AudioChannel"),
            //         Encrypted = drDVBViewerChannelsDataRow.Field<int>("Encrypted"),
            //         Group = drDVBViewerChannelsDataRow.Field<int>("Group"),
            //         Frequency = drDVBViewerChannelsDataRow.Field<int>("Frequency"),
            //         Polarity = drDVBViewerChannelsDataRow.Field<int>("Polarity"),
            //         Symbolrate = drDVBViewerChannelsDataRow.Field<int>("Symbolrate"),
            //         SatModulation = drDVBViewerChannelsDataRow.Field<int>("SatModulation"),
            //         SubStreamID = drDVBViewerChannelsDataRow.Field<int>("SubStreamID")
            //     }
            //    ).ToList();

            //DataTable dtMergedDataTable = Utility.LINQResultToDataTable(results);
        }

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
