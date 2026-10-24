using System.Globalization;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.DVBViewer;
using static KON.Liwest.ChannelFactory.Global;

using Krkadoni.EnigmaSettings;
using Krkadoni.EnigmaSettings.Interfaces;
using NanoXLSX;
using NullLogger = Krkadoni.EnigmaSettings.NullLogger;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Export of source data to excel, DVBViewer and Enigma files
    /// </summary>
    public static class DataExports
    {
        /// <summary>
        /// DataExports.ExportSourceDataToExcelExportFile()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalExcelFilename"></param>
        public static void ExportSourceDataToExcelExportFile(List<SourceDataItem> lsdiLocalSourceData, string? strLocalExcelFilename)
        {
            if (string.IsNullOrEmpty(strLocalExcelFilename))
                return;

            LogTrace("DataExports.ExportSourceDataToExcelExportFile()");
            LogInfo("-» Export source data to excel export file");

            if (lsdiLocalSourceData.Count == 0)
                return;

            var wbCurrentWorkbook = new Workbook(GetOutputFilePath(strLocalExcelFilename), csExcelWorksheetName);
            var wsCurrentWorksheet = wbCurrentWorkbook.GetWorksheet(csExcelWorksheetName);

            wbCurrentWorkbook.SetCurrentWorksheet(wsCurrentWorksheet);
            wsCurrentWorksheet.SetCurrentRowNumber(0);

            // Maximum text length per column (header and values) for the column width
            var iaCurrentMaximumLengths = SourceDataItem.piaDataProperties.Select(piCurrentProperty => piCurrentProperty.Name.Length).ToArray();

            foreach (var piCurrentProperty in SourceDataItem.piaDataProperties)
                wsCurrentWorksheet.AddNextCell(piCurrentProperty.Name);

            foreach (var sdiCurrentSourceDataItem in lsdiLocalSourceData)
            {
                wsCurrentWorksheet.GoToNextRow();

                for (var iCurrentColumnNumber = 0; iCurrentColumnNumber < SourceDataItem.piaDataProperties.Length; iCurrentColumnNumber++)
                {
                    var oCurrentValue = SourceDataItem.piaDataProperties[iCurrentColumnNumber].GetValue(sdiCurrentSourceDataItem);
                    wsCurrentWorksheet.AddNextCell(oCurrentValue);

                    iaCurrentMaximumLengths[iCurrentColumnNumber] = Math.Max(iaCurrentMaximumLengths[iCurrentColumnNumber], Convert.ToString(oCurrentValue, CultureInfo.InvariantCulture)?.Length ?? 0);
                }
            }

            for (var iCurrentColumnNumber = 0; iCurrentColumnNumber < iaCurrentMaximumLengths.Length; iCurrentColumnNumber++)
                wsCurrentWorksheet.SetColumnWidth(GetExcelColumnName(iCurrentColumnNumber + 1), Math.Min(iaCurrentMaximumLengths[iCurrentColumnNumber] * EXCEL_CHARACTER_TO_WIDTH_CONSTANT + EXCEL_CHARACTER_TO_WIDTH_CONSTANT_CF, EXCEL_MAXIMUM_COLUMN_WIDTH));

            wsCurrentWorksheet.SetAutoFilter(0, wsCurrentWorksheet.GetLastColumnNumber());
            wbCurrentWorkbook.Save();
        }

        /// <summary>
        /// DataExports.ExportSourceDataToDVBViewerTransponderFile()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalDVBViewerTransponderFilename"></param>
        public static void ExportSourceDataToDVBViewerTransponderFile(List<SourceDataItem> lsdiLocalSourceData, string? strLocalDVBViewerTransponderFilename)
        {
            if (string.IsNullOrEmpty(strLocalDVBViewerTransponderFilename))
                return;

            LogTrace("DataExports.ExportSourceDataToDVBViewerTransponderFile()");
            LogInfo("-» Export source data to DVBViewer transponder file");

            var iaCurrentFrequencies = lsdiLocalSourceData.Where(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelFrequency is > 0).Select(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelFrequency!.Value).Distinct().Order().ToArray();

            using var swCurrentStreamWriter = new StreamWriter(File.Open(GetOutputFilePath(strLocalDVBViewerTransponderFilename), FileMode.Create));
            swCurrentStreamWriter.WriteLine("[SATTYPE]");
            swCurrentStreamWriter.WriteLine("1={0}", ciOrbitalPosition);
            swCurrentStreamWriter.WriteLine("2={0}", string.Format(csTransponderRootElementNameFormat, DateTime.Now));
            swCurrentStreamWriter.WriteLine(string.Empty);
            swCurrentStreamWriter.WriteLine("[DVB]");
            swCurrentStreamWriter.WriteLine("0=" + iaCurrentFrequencies.Length);

            for (var iCurrentCounter = 0; iCurrentCounter < iaCurrentFrequencies.Length; iCurrentCounter++)
                swCurrentStreamWriter.WriteLine("{0}={1},{2},{3},|%{4}", iCurrentCounter + 1, iaCurrentFrequencies[iCurrentCounter], ciModulation, ciSymbolRate, ciSpectralInversion);
        }

        /// <summary>
        /// DataExports.ExportSourceDataToDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static void ExportSourceDataToDVBViewerChannelDatabase(List<SourceDataItem> lsdiLocalSourceData, string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename))
                return;

            LogTrace("DataExports.ExportSourceDataToDVBViewerChannelDatabase()");
            LogInfo("-» Export source data to DVBViewer channel database file");

            if (lsdiLocalSourceData.Count == 0)
                return;

            var lsdiCurrentExportableItems = GetExportableServiceItems(lsdiLocalSourceData, "DVBViewer channel database");
            if (lsdiCurrentExportableItems.Count == 0)
                return;

            using var bwCurrentBinaryWriter = new BinaryWriter(File.Open(GetOutputFilePath(strLocalDVBViewerChannelDatabaseFilename), FileMode.Create));
            var cdhCurrentChannelDatabaseHeader = new ChannelDatabaseHeader
            {
                IDLength = (byte)csDefaultDVBViewerChannelDatabaseID.Length,
                ID = csDefaultDVBViewerChannelDatabaseID.ToCharArray(),
                VersionHi = csDefaultDVBViewerChannelDatabaseVersionHi,
                VersionLo = csDefaultDVBViewerChannelDatabaseVersionLo
            };
            bwCurrentBinaryWriter.Write(Binarize(cdhCurrentChannelDatabaseHeader));

            foreach (var sdiCurrentSourceDataItem in lsdiCurrentExportableItems)
            {
                var saCurrentAPIDS = SplitList(sdiCurrentSourceDataItem.APIDS);
                var saCurrentAPIDSStreamType = SplitList(sdiCurrentSourceDataItem.APIDSST);
                var saCurrentAPIDSLanguage = SplitList(sdiCurrentSourceDataItem.APIDSLANG);
                var saCurrentAPIDSCodec = SplitList(sdiCurrentSourceDataItem.APIDSCODEC);
                var bCurrentIsEncrypted = sdiCurrentSourceDataItem.IsEncrypted;
                var uiCurrentVideoPID = sdiCurrentSourceDataItem.VPID.ToUInt16OrDefault();

                // One record per audio pid - the first one is the primary record, all following ones are flagged as following pid
                for (var iCurrentCounter = 0; iCurrentCounter < saCurrentAPIDS.Length; iCurrentCounter++)
                {
                    var strCurrentAPIDStreamType = iCurrentCounter < saCurrentAPIDSStreamType.Length ? saCurrentAPIDSStreamType[iCurrentCounter] : string.Empty;
                    var strCurrentAPIDLanguage = iCurrentCounter < saCurrentAPIDSLanguage.Length && !string.IsNullOrEmpty(saCurrentAPIDSLanguage[iCurrentCounter]) ? saCurrentAPIDSLanguage[iCurrentCounter] : csDefaultAudioLanguage;
                    var uiCurrentAudioPID = ushort.TryParse(saCurrentAPIDS[iCurrentCounter].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var uiCurrentParsedAudioPID) ? uiCurrentParsedAudioPID : (ushort)0;

                    var cdcCurrentChannelDatabaseChannel = new ChannelDatabaseChannel
                    {
                        TunerData = new ChannelDatabaseTuner
                        {
                            TunerType = ciTunerType,
                            ChannelFlags = GenerateChannelFlags(bCurrentIsEncrypted, false, false, uiCurrentVideoPID > 0, uiCurrentAudioPID > 0, iCurrentCounter > 0 && uiCurrentAudioPID > 0),
                            Frequency = sdiCurrentSourceDataItem.ChannelFrequency.ToUInt32OrDefault(),
                            Symbolrate = ciSymbolRate,
                            PMT_PID = sdiCurrentSourceDataItem.PMT.ToUInt16OrDefault(),
                            ChannelAVFormat = GenerateChannelAVFormat(strCurrentAPIDStreamType, sdiCurrentSourceDataItem.VPIDST ?? string.Empty, iCurrentCounter < saCurrentAPIDSCodec.Length ? saCurrentAPIDSCodec[iCurrentCounter] : null),
                            ChannelNumber = sdiCurrentSourceDataItem.ChannelNumber.ToUInt16OrDefault(),
                            PolarityOrModulation = ciModulation,
                            FEC = ciFEC,
                            OrbitalPosition = ciOrbitalPosition,
                            Language = strCurrentAPIDLanguage.ToCharArray(),
                            Audio_PID = uiCurrentAudioPID,
                            Video_PID = uiCurrentVideoPID,
                            TransportStream_ID = sdiCurrentSourceDataItem.TSID.ToUInt16OrDefault(),
                            Teletext_PID = sdiCurrentSourceDataItem.TPID.ToUInt16OrDefault(),
                            OriginalNetwork_ID = sdiCurrentSourceDataItem.ONID.ToUInt16OrDefault(),
                            Service_ID = sdiCurrentSourceDataItem.SID.ToUInt16OrDefault(),
                            PCR_PID = sdiCurrentSourceDataItem.PCRPID.ToUInt16OrDefault(),
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
                        ChannelName = sdiCurrentSourceDataItem.ChannelName.ToBinary25CharArray(),
                        Category = csCategory.ToBinary25CharArray(),
                        Encrypted = bCurrentIsEncrypted ? (byte)1 : (byte)0,
                        Reserved = 0
                    };

                    bwCurrentBinaryWriter.Write(Binarize(cdcCurrentChannelDatabaseChannel));
                }
            }
        }

        /// <summary>
        /// DataExports.ExportSourceDataToEnigmaSelector()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="femLocalFlagExportMode"></param>
        /// <param name="strLocalFilename"></param>
        /// <param name="fefLocalExportFilenameFormat"></param>
        public static void ExportSourceDataToEnigmaSelector(List<SourceDataItem> lsdiLocalSourceData, FlagExportMode femLocalFlagExportMode, string? strLocalFilename, FlagExportEnigmaFormat? fefLocalExportFilenameFormat)
        {
            switch (femLocalFlagExportMode)
            {
                case FlagExportMode.EnigmaTransponderList:
                    ExportSourceDataToEnigma(lsdiLocalSourceData, strLocalFilename, fefLocalExportFilenameFormat, true, false, false);
                    break;
                case FlagExportMode.EnigmaChannelDatabase:
                    ExportSourceDataToEnigma(lsdiLocalSourceData, strLocalFilename, fefLocalExportFilenameFormat, false, true, false);
                    break;
                case FlagExportMode.EnigmaBouquetsFiles:
                    ExportSourceDataToEnigma(lsdiLocalSourceData, strLocalFilename, fefLocalExportFilenameFormat, false, false, true);
                    break;
            }
        }

        /// <summary>
        /// DataExports.ExportSourceDataToEnigma()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalExportFilename"></param>
        /// <param name="fefLocalExportFilenameFormat"></param>
        /// <param name="bLocalExportTransponders"></param>
        /// <param name="bLocalExportChannelDatabase"></param>
        /// <param name="bLocalExportBouquetsFiles"></param>
        public static void ExportSourceDataToEnigma(List<SourceDataItem> lsdiLocalSourceData, string? strLocalExportFilename, FlagExportEnigmaFormat? fefLocalExportFilenameFormat, bool bLocalExportTransponders, bool bLocalExportChannelDatabase, bool bLocalExportBouquetsFiles)
        {
            LogTrace("DataExports.ExportSourceDataToEnigma()");
            LogInfo($"-» Export source data to enigma Transponders:{bLocalExportTransponders} - ChannelDatabase:{bLocalExportChannelDatabase} - BouquetFiles:{bLocalExportBouquetsFiles}");

            if (lsdiLocalSourceData.Count == 0)
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
                FlagExportEnigmaFormat.Enigma1Ver1 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma1Ver1,
                FlagExportEnigmaFormat.Enigma1Ver2 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma1Ver2,
                FlagExportEnigmaFormat.Enigma2Ver3 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver3,
                FlagExportEnigmaFormat.Enigma2Ver4 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver4,
                FlagExportEnigmaFormat.Enigma2Ver5 => Krkadoni.EnigmaSettings.Enums.SettingsVersion.Enigma2Ver5,
                _ => isCurrentISettings.SettingsVersion
            };

            var ixcCurrentIXmlCable = sioCurrentSettingsIO.Factory.InitNewXmlCable();
            ixcCurrentIXmlCable.Name = string.Format(csTransponderRootElementNameFormat, DateTime.Now);
            ixcCurrentIXmlCable.SatFeed = Convert.ToString(ciSatFeed);
            ixcCurrentIXmlCable.Flags = Convert.ToString(ciFEC);
            ixcCurrentIXmlCable.CountryCode = ciCountryCode;

            // Transponders need a valid frequency, TSID and ONID - services are assigned by TSID and ONID
            var dCurrentTransponders = new Dictionary<(int iTSID, int iONID), ITransponder>();
            var taCurrentTransponders = lsdiLocalSourceData
                .Where(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelFrequency is > 0 && sdiCurrentSourceDataItem.TSID.HasValue && sdiCurrentSourceDataItem.ONID.HasValue)
                .Select(sdiCurrentSourceDataItem => (iFrequency: sdiCurrentSourceDataItem.ChannelFrequency!.Value, iTSID: sdiCurrentSourceDataItem.TSID!.Value, iONID: sdiCurrentSourceDataItem.ONID!.Value))
                .Distinct()
                .OrderBy(tCurrentTransponder => tCurrentTransponder.iFrequency);

            foreach (var (iCurrentFrequency, iCurrentTSID, iCurrentONID) in taCurrentTransponders)
            {
                ixcCurrentIXmlCable.Transponders.Add(new XmlTransponder
                {
                    Frequency = Convert.ToString(iCurrentFrequency, CultureInfo.InvariantCulture),
                    SymbolRate = Convert.ToString(ciSymbolRate * 1000),
                    FEC = Convert.ToString(ciFEC),
                    Modulation = Convert.ToString(ciModulation)
                });

                var itCurrentTransponder = sioCurrentSettingsIO.Factory.InitNewTransponderDVBC($"{ciDefaultDVBCNamespace}:{iCurrentTSID.ToHexString4Byte()}:{iCurrentONID.ToHexString4Byte()}", $"\tc {iCurrentFrequency}:{ciSymbolRate * 1000}:{ciEnigmaTransponderInversion}:{ciModulation}:{ciFEC}:0:0");
                isCurrentISettings.Transponders.Add(itCurrentTransponder);
                dCurrentTransponders.TryAdd((iCurrentTSID, iCurrentONID), itCurrentTransponder);
            }

            isCurrentISettings.Cables.Add(ixcCurrentIXmlCable);

            if (bLocalExportTransponders && strCurrentExportFilePath != null)
                sioCurrentSettingsIO.Factory.InitNewXmlCableIO(sioCurrentSettingsIO.FileProvider).SaveCablesToFile(strCurrentExportFilePath, isCurrentISettings);

            if (!bLocalExportChannelDatabase && !bLocalExportBouquetsFiles)
                return;

            var ifbCurrentIFileBouquetTV = sioCurrentSettingsIO.Factory.InitNewFileBouquet(csEnigmaBouquetNameTV, csEnigmaBouquetFilenameTV);
            var ifbCurrentIFileBouquetRadio = sioCurrentSettingsIO.Factory.InitNewFileBouquet(csEnigmaBouquetNameRadio, csEnigmaBouquetFilenameRadio);

            foreach (var sdiCurrentSourceDataItem in GetExportableServiceItems(lsdiLocalSourceData, "enigma channel database / bouquets"))
            {
                var iCurrentVPID = sdiCurrentSourceDataItem.VPID ?? 0;
                var iCurrentTPID = sdiCurrentSourceDataItem.TPID ?? 0;
                var iCurrentPCRPID = sdiCurrentSourceDataItem.PCRPID ?? 0;
                var iCurrentTSID = sdiCurrentSourceDataItem.TSID!.Value;
                var iCurrentONID = sdiCurrentSourceDataItem.ONID!.Value;
                var bCurrentIsVideoService = sdiCurrentSourceDataItem.IsVideoService;

                var strCurrentServiceReference = $"{sdiCurrentSourceDataItem.SID!.Value.ToHexString4Byte()}:{ciDefaultDVBCNamespace}:{iCurrentTSID.ToHexString4Byte()}:{iCurrentONID.ToHexString4Byte()}:{(bCurrentIsVideoService ? ciEnigmaServiceTypeTV : ciEnigmaServiceTypeRadio)}:{sdiCurrentSourceDataItem.ChannelNumber?.ToString(CultureInfo.InvariantCulture)}:0";
                var strCurrentServiceFlags = $"p:{sdiCurrentSourceDataItem.ProviderName}" +
                                             (bCurrentIsVideoService ? $",c:00{iCurrentVPID.ToHexString4Byte()}" : string.Empty) +
                                             (iCurrentTPID > 0 ? $",c:02{iCurrentTPID.ToHexString4Byte()}" : string.Empty) +
                                             (iCurrentPCRPID > 0 ? $",c:03{iCurrentPCRPID.ToHexString4Byte()}" : string.Empty) +
                                             (bCurrentIsVideoService ? ",c:050001" : string.Empty) +
                                             GetEnigmaCAIDFlags(sdiCurrentSourceDataItem) +
                                             $",{csEnigmaServiceFlags}";

                var isCurrentIService = sioCurrentSettingsIO.Factory.InitNewService(strCurrentServiceReference, sdiCurrentSourceDataItem.ChannelName ?? string.Empty, strCurrentServiceFlags);
                isCurrentIService.Transponder = dCurrentTransponders.GetValueOrDefault((iCurrentTSID, iCurrentONID));
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
        /// DataExports.GetExportableServiceItems()
        /// Returns all items which contain the service data (SID, TSID, ONID, frequency) required for channel database exports.
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalExportName"></param>
        /// <returns>List of exportable items</returns>
        private static List<SourceDataItem> GetExportableServiceItems(List<SourceDataItem> lsdiLocalSourceData, string strLocalExportName)
        {
            var lsdiCurrentExportableItems = lsdiLocalSourceData.Where(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.HasServiceData).ToList();
            var iCurrentSkippedItems = lsdiLocalSourceData.Count - lsdiCurrentExportableItems.Count;

            if (iCurrentSkippedItems > 0)
                LogWarning($"-» Skipped {iCurrentSkippedItems} channel(s) without service data (SID/TSID/ONID/frequency) for {strLocalExportName} export");

            return lsdiCurrentExportableItems;
        }

        /// <summary>
        /// DataExports.GetEnigmaCAIDFlags() - one C:xxxx entry per CA_system_ID of an encrypted service, C:0000 if no CA_system_ID is known, none for unencrypted services
        /// </summary>
        /// <param name="sdiLocalSourceDataItem"></param>
        /// <returns>string</returns>
        private static string GetEnigmaCAIDFlags(SourceDataItem sdiLocalSourceDataItem)
        {
            if (!sdiLocalSourceDataItem.IsEncrypted)
                return string.Empty;

            var iaCurrentCASystemIDs = sdiLocalSourceDataItem.CASystemIDValues;
            return iaCurrentCASystemIDs.Length > 0 ? string.Concat(iaCurrentCASystemIDs.Select(iCurrentCASystemID => $",C:{iCurrentCASystemID.ToHexString4Byte()}")) : $",C:{ciEnigmaUnknownCASystemID.ToHexString4Byte()}";
        }

        /// <summary>
        /// DataExports.SplitList() - splits a comma separated list, an empty list results in one empty entry
        /// </summary>
        /// <param name="strLocalList"></param>
        /// <returns>string[]</returns>
        private static string[] SplitList(string? strLocalList)
        {
            return !string.IsNullOrEmpty(strLocalList) ? strLocalList.Split(',') : [string.Empty];
        }

        /// <summary>
        /// DataExports.GetExcelColumnName()
        /// </summary>
        /// <param name="iLocalColumnNumber"></param>
        /// <returns>Column name (A, B, ..., AA, ...)</returns>
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
