using CommandLine;
using CommandLine.Text;
using NLog;
using NLog.Layouts;
using Pastel;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;

using System.Text;

namespace KON.Liwest.ChannelFactory
{
    public class Program
    {
        public enum FlagMode
        {
            GenerateSourceData = 0,
            GenerateSourceDataAndSave = 1,
            //Automated                                   = 1,
            //FetchTransponderDataFromLiwestChannelList   = 2,
            //FetchTransponderDataFromLiwestSmartcard     = 3,
            //FetchTransponderDataFromLiwestBlindscan = 4,
            ExportSourceDataToExcelExportFile = 2,
            ExportSourceDataToDVBViewerTransponderFile = 3,
            ExportSourceDataToDVBViewerChannelDatabase = 4,
            DumpDVBViewerChannelDatabase = 5
            //GenerateExcelChannelSortFile = 6
        }
        public enum FlagSource
        {
            ChannelList = 0,
            Smartcard = 1,
            Blindscan = 2,
            All = 3,
            File = 4
        }

        //public enum FlagInputFileEncoding
        //{
        //    UTF8                                = 0x0,
        //    ISO88591                            = 0x1
        //}

        public enum FlagSourceDataCleanup
        {
            None = 0,
            RemoveInvalidSID = 1
        }

        public class Options
        {
            [Option(longName: "mode", Required = true, HelpText = "Factory mode to use")]
            public FlagMode cloMode { get; set; }

            [Option(longName: "source", Required = true, HelpText = "Factory source to use")]
            public FlagSource cloSource { get; set; }

            //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

            [Option(longName: "sourcedatafilename", Required = false, HelpText = "Filename for factory source data to load/save to")]
            public string? cloSourceDataFilename { get; set; }

            [Option(longName: "smartcardid", Required = false, HelpText = "Smartcard Identifier")]
            public string? cloSmartcardID { get; set; }

            [Option(longName: "serverip", Required = false, HelpText = "IP Address of SAT>IP server")]
            public string? cloServerIP { get; set; }

            [Option(longName: "frequencies", Required = false, HelpText = "Frequencies filer for blindscan")]
            public string? cloFrequencies { get; set; }

            [Option(longName: "excelexportfilename", Required = false, HelpText = "Filename for source data excel export to save to")]
            public string? cloExcelExportFilename { get; set; }

            [Option(longName: "dvbviewertransponderfilename", Required = false, HelpText = "Filename for dvbviewer transponder data to save to")]
            public string? cloDVBViewerTransponderFilename { get; set; }

            [Option(longName: "dvbviewerchanneldatabasefilename", Required = false, HelpText = "Filename for dvbviewer channel database to read from and save to")]
            public string? cloDVBViewerChannelDatabaseFilename { get; set; }

            [Option(longName: "stringvariationdictionaryfilename", Required = false, HelpText = "Filename for string variation dictionary to load/save to")]
            public string? cloStringVariationDictionaryFilename { get; set; }

            //[Option('t', longName: "enabletv", Required = false, HelpText = "Enable parsing of provider TV channels.")]
            //public bool bEnableTV { get; set; }

            //[Option('r', longName: "enableradio", Required = false, HelpText = "Enable parsing of provider RADIO channels.")]
            //public bool bEnableRADIO { get; set; }

            //[Option('s', longName: "silent", Required = false, HelpText = "Set program to silently exit.")]
            //public bool bSilent { get; set; }

            //[Option('c', longName: "inputfile", Required = false, HelpText = "Filename from which to generate Excel channel sort file.")]
            //public string strInputFile { get; set; }

            //[Option('d', longName: "inputfileencoding", Required = false, HelpText = "Encoding which is used in input file.")]
            //public FlagInputFileEncoding enumInputFileEncoding { get; set; }

            //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

            [Option(longName: "sourcedatacleanup", Required = false, HelpText = "Cleanup source data")]
            public FlagSourceDataCleanup? cloSourceDataCleanup { get; set; }

            //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

            [Option(longName: "verbose", Required = false, HelpText = "Set output to verbose level.")]
            public bool cloVerbose { get; set; }
        }

        static void Main(string[] saLocalArguments)
        {
            InitDataContainers(ref dtSourceDataTable);

            var pCurrentParser = new Parser(apsCurrentConfiguration => apsCurrentConfiguration.HelpWriter = null);
            var prCurrentParseResult = pCurrentParser.ParseArguments<Options>(saLocalArguments);

            prCurrentParseResult.WithParsed(delegate (Options oLocalOptions)
            {
                if (oLocalOptions.cloVerbose)
                    LogManager.Setup().LoadConfiguration(aslcbCurrentSetupLoadConfigurationBuilder => { aslcbCurrentSetupLoadConfigurationBuilder.ForLogger().FilterMinLevel(LogLevel.Trace).WriteToColoredConsole(new SimpleLayout("${message}")); });
                else
                    LogManager.Setup().LoadConfiguration(aslcbCurrentSetupLoadConfigurationBuilder => { aslcbCurrentSetupLoadConfigurationBuilder.ForLogger().FilterMinLevel(LogLevel.Info).WriteToColoredConsole(new SimpleLayout("${message}")); });

                Run(oLocalOptions);
            });

            prCurrentParseResult.WithNotParsed(delegate
            {
                LogManager.Setup().LoadConfiguration(aslcbCurrentSetupLoadConfigurationBuilder => { aslcbCurrentSetupLoadConfigurationBuilder.ForLogger().FilterMinLevel(LogLevel.Info).WriteToColoredConsole(new SimpleLayout("${message}")); });

                DisplayHelp(prCurrentParseResult);
            });
        }

        private static void DisplayHelp<T>(ParserResult<T> result)
        {
            var htCurrentHelpText = HelpText.AutoBuild(result, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = false;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = "KON.Liwest.ChannelFactory";
                htCurrentHelpText.Copyright = "Copyright (C) 2025 Oswald Oliver";

                htCurrentHelpText.AddPostOptionsLine($"* Example: Generate DVBViewer transponder file \"DVB-C_AT_LIWEST.ini\"");
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateDVBViewerTransponderFile -t -r");
                htCurrentHelpText.AddPostOptionsLine($"");
                //htCurrentHelpText.AddPostOptionsLine($"* Example: Generate Excel sort file \"ChannelListSort.xlsx\" with UTF8 decoding used");
                //htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateExcelChannelSortFile -t -r -c \"liwest (1900-01-01).ini\" -d UTF8");
                //htCurrentHelpText.AddPostOptionsLine("");
                //htCurrentHelpText.AddPostOptionsLine($"* Example: Generate Excel sort file \"ChannelListSort.xlsx\" with ISO-8859-1 decoding used");
                //htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateExcelChannelSortFile -t -r -c \"liwest (1900-01-01).ini\" -d ISO88591");
                //htCurrentHelpText.AddPostOptionsLine("");

                return HelpText.DefaultParsingErrorsHandler(result, htCurrentHelpText);
            }, e => e);

            lCurrentLogger.Info(htCurrentHelpText);

            #if DEBUG
                lCurrentLogger.Info($"Press any key to continue ...".Pastel(ConsoleColor.Green));
                Console.ReadKey();
            #endif
        }

        private static void Run(Options oLocalOptions)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            #if DEBUG
                if (OperatingSystem.IsWindows())
                    Console.BufferHeight = short.MaxValue - 1;
            #endif

            if (CreateFirewallRule())
            {
                LoadStringVariationDictionaryFromFile(strLocalStringVariationDictionaryFilename: oLocalOptions.cloStringVariationDictionaryFilename ?? csDefaultStringVariationDictionaryFilename);

                switch (oLocalOptions.cloMode)
                {
                    case FlagMode.GenerateSourceData:
                    case FlagMode.GenerateSourceDataAndSave:
                    {
                        switch (oLocalOptions.cloSource)
                        {
                            case FlagSource.ChannelList:
                            {
                                GetSourceDataFromChannelList();
                                break;
                            }
                            case FlagSource.Smartcard:
                            {
                                GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID ?? csDefaultSmartcardId);
                                break;
                                    }
                            case FlagSource.Blindscan:
                            {
                                GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies ?? csDefaultFrequencies);
                                break;
                            }
                            case FlagSource.File:
                            {
                                LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);
                                break;
                            }
                            case FlagSource.All:
                            {
                                GetSourceDataFromChannelList();
                                GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID ?? csDefaultSmartcardId);
                                GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies ?? csDefaultFrequencies);
                                break;
                            }
                        }

                        if (oLocalOptions.cloSourceDataCleanup != null)
                        {   
                            if (oLocalOptions.cloSourceDataCleanup != FlagSourceDataCleanup.None)
                            {
                                dtSourceDataTable = CleanupSourceData();

                                switch (oLocalOptions.cloSourceDataCleanup)
                                {
                                    case FlagSourceDataCleanup.RemoveInvalidSID:
                                    {
                                        dtSourceDataTable = CleanupSourceDataRemoveInvalidSID();
                                        break;
                                    }
                                }
                            }
                        }

                        if(oLocalOptions.cloMode == FlagMode.GenerateSourceDataAndSave)
                            SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);

                        break;
                    }
                    case FlagMode.ExportSourceDataToExcelExportFile:
                    {
                        LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);
                        ExportSourceDataToExcelExportFile(strLocalExcelExportFilename: oLocalOptions.cloExcelExportFilename ?? csDefaultExcelExportFilename);
                        break;
                    }
                    case FlagMode.ExportSourceDataToDVBViewerTransponderFile:
                    {
                        LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);
                        ExportSourceDataToDVBViewerTransponderFile(strLocalDVBViewerTransponderFilename: oLocalOptions.cloDVBViewerTransponderFilename ?? csDefaultDVBViewerTransponderFilename);
                        break;
                    }
                    case FlagMode.ExportSourceDataToDVBViewerChannelDatabase:
                    {
                        LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);
                        ExportSourceDataToDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloDVBViewerChannelDatabaseFilename ?? csDefaultDVBViewerChannelDatabaseFilename);
                        break;
                    }
                    case FlagMode.DumpDVBViewerChannelDatabase:
                    {
                        LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceDataFilename ?? csDefaultSourceDataFilename);
                        DumpDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloDVBViewerChannelDatabaseFilename ?? csDefaultDVBViewerChannelDatabaseFilename);
                        break;
                    }
                    //case FlagMode.GenerateExcelChannelSortFile:
                    //{
                    //    GenerateExcelChannelSortFile();
                    //    break;
                    //}
                }
            }
            else
            {
                lCurrentLogger.Error("Firewall rule has not been created as there are not sufficient permissions or wrong operation system used. Operation stopped!".Pastel(ConsoleColor.Red));
            }

            #if DEBUG
                lCurrentLogger.Info($"Press any key to continue ...".Pastel(ConsoleColor.Green));
                Console.ReadKey();
            #endif
        }
    }
}