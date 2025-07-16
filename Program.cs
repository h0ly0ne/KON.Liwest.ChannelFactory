using System.Text;

using CommandLine;
using CommandLine.Text;
using NLog;
using NLog.Layouts;
using Pastel;

using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    public class Program
    {
        public class Options
        {
            [Option(longName: "mode", Required = true, HelpText = "Factory mode to use")]
            public FlagMode cloMode { get; set; }
            [Option(longName: "source", Required = true, HelpText = "Factory source to use")]
            public FlagSource cloSource { get; set; }
            [Option(longName: "target", Required = true, HelpText = "Factory target to use")]
            public FlagTarget cloTarget { get; set; }

            [Option(longName: "smartcardid", HelpText = "Smartcard identifier")]
            public string? cloSmartcardID { get; set; }
            [Option(longName: "serverip", HelpText = "IP Address of SAT>IP server")]
            public string? cloServerIP { get; set; }
            [Option(longName: "frequencies", HelpText = "Frequencies filer for blindscan")]
            public string? cloFrequencies { get; set; }
            [Option(longName: "sourcefile", HelpText = "Filename for factory source to load from")]
            public string? cloSourceFile { get; set; }
            [Option(longName: "targetfile", HelpText = "Filename for factory target to save to")]
            public string? cloTargetFile { get; set; }

            [Option(longName: "channelvariationfile", HelpText = "Filename for channel variation dictionary to load from")]
            public string? cloChannelVariationFile { get; set; }
            [Option(longName: "sourcecleanup", HelpText = "Source cleanup operation mode")]
            public FlagSourceCleanup cloSourceCleanup { get; set; }
            [Option(longName: "channelremovalfile", HelpText = "Filename for channel removal dictionary to load from")]
            public string? cloChannelRemovalFile { get; set; }
            [Option(longName: "channelsortingfile", HelpText = "Filename for channel sorting dictionary to load from")]
            public string? cloChannelSortingFile { get; set; }

            [Option(longName: "exportdvbvtransponderfile", HelpText = "Filename for dvbviewer transponder data to save to")]
            public string? cloExportDVBViewerTransponderFile { get; set; }
            [Option(longName: "exportdvbvcdbfile", HelpText = "Filename for dvbviewer channel database to save to")]
            public string? cloExportDVBViewerChannelDatabaseFile { get; set; }
            [Option(longName: "exportexcelchannelfile", HelpText = "Filename for excel channel export to save to")]
            public string? cloExportExcelChannelFile { get; set; }
            [Option(longName: "exportenigmadbfile", HelpText = "Filename for enigma db channel export to save to")]
            public string? cloExportEnigmaDBFile{ get; set; }

            [Option(longName: "keypress", HelpText = "Wait for keypress after operation")]
            public bool cloKeyPress { get; set; }
            [Option(longName: "verbose", HelpText = "Set output to verbose level")]
            public bool cloVerbose { get; set; }
        }

        static void Main(string[] saLocalArguments)
        { 
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

        private static void DisplayHelp<T>(ParserResult<T> prLocalParseResult)
        {
            // TODO: Add new parameters and samples
            var htCurrentHelpText = HelpText.AutoBuild(prLocalParseResult, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = false;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = "KON.Liwest.ChannelFactory";
                htCurrentHelpText.Copyright = "Copyright (C) 2025 Oswald Oliver";

                //htCurrentHelpText.AddPostOptionsLine($"* Example: Generate DVBViewer transponder file \"DVB-C_AT_LIWEST.ini\"");
                //htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateDVBViewerTransponderFile -t -r");
                //htCurrentHelpText.AddPostOptionsLine($"");
                //htCurrentHelpText.AddPostOptionsLine($"* Example: Generate Excel sort file \"ChannelListSort.xlsx\" with UTF8 decoding used");
                //htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateExcelChannelSortFile -t -r -c \"liwest (1900-01-01).ini\" -d UTF8");
                //htCurrentHelpText.AddPostOptionsLine("");
                //htCurrentHelpText.AddPostOptionsLine($"* Example: Generate Excel sort file \"ChannelListSort.xlsx\" with ISO-8859-1 decoding used");
                //htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateExcelChannelSortFile -t -r -c \"liwest (1900-01-01).ini\" -d ISO88591");
                //htCurrentHelpText.AddPostOptionsLine("");

                return HelpText.DefaultParsingErrorsHandler(prLocalParseResult, htCurrentHelpText);
            }, e => e);

            lCurrentLogger.Info(htCurrentHelpText);
        }
        private static void DisplayHelpWithCustomCondition(string[] saLocalCustomConditions)
        {
            // TODO: Add new parameters and samples
            var prCurrentParseResult = new Parser(x => { x.HelpWriter = null; }).ParseArguments<Options>(["--help"]);
            var htCurrentHelpText = HelpText.AutoBuild(prCurrentParseResult, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = true;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = "KON.Liwest.ChannelFactory";
                htCurrentHelpText.Copyright = "Copyright (C) 2025 Oswald Oliver";

                if (saLocalCustomConditions.Length > 0)
                {
                    htCurrentHelpText.AddPreOptionsLine("ERROR(S):");

                    foreach (var strCurrentCustomCondition in saLocalCustomConditions)
                    {
                        if (!string.IsNullOrEmpty(strCurrentCustomCondition))
                            htCurrentHelpText.AddPreOptionsLine($"  * Custom option '{strCurrentCustomCondition}' not met or missing.");
                    }
                }

                return HelpText.DefaultParsingErrorsHandler(prCurrentParseResult, htCurrentHelpText);
            }, e => e);

            lCurrentLogger.Info(htCurrentHelpText);
        }

        private static void Run(Options oLocalOptions)
        {
            InitDataContainers(ref dtSourceDataTable);

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            if (OperatingSystem.IsWindows())
                Console.BufferHeight = short.MaxValue - 1;

            if (CreateFirewallRule())
            {
                string[] saLocalCustomCondition = [];

                // Check for custom conditions
                if (!string.IsNullOrEmpty(oLocalOptions.cloChannelVariationFile) && !File.Exists(oLocalOptions.cloChannelVariationFile))
                    saLocalCustomCondition = saLocalCustomCondition.Append("channelvariationfile").ToArray();

                switch (oLocalOptions.cloMode)
                {
                    case FlagMode.GenerateData:
                    {
                        if (oLocalOptions.cloSource is FlagSource.SmartcardList or FlagSource.CombinedList)
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloSmartcardID) ? saLocalCustomCondition.Append("smartcardid").ToArray() : saLocalCustomCondition;

                        if (oLocalOptions.cloSource is FlagSource.BlindscanList or FlagSource.CombinedList)
                        {
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloServerIP) ? saLocalCustomCondition.Append("serverip").ToArray() : saLocalCustomCondition;
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloFrequencies) ? saLocalCustomCondition.Append("frequencies").ToArray() : saLocalCustomCondition;
                        }

                        if (oLocalOptions.cloSource is FlagSource.FileList)
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile) ? saLocalCustomCondition.Append("sourcefile").ToArray() : saLocalCustomCondition;

                        if (oLocalOptions.cloTarget is FlagTarget.FileList)
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloTargetFile) ? saLocalCustomCondition.Append("targetfile").ToArray() : saLocalCustomCondition;

                        break;
                    }
                    case FlagMode.ModifyData:
                    {
                        if (oLocalOptions.cloSource is not FlagSource.FileList)
                            saLocalCustomCondition = saLocalCustomCondition.Append("source").ToArray();
                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("sourcefile").ToArray();
                        if (oLocalOptions.cloTarget is FlagTarget.FileList && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("targetfile").ToArray();
                        if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile) && !File.Exists(oLocalOptions.cloChannelRemovalFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("channelremovalfile").ToArray();
                        if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile) && !File.Exists(oLocalOptions.cloChannelSortingFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("channelsortingfile").ToArray();

                        break;
                    }
                    case FlagMode.DumpDVBViewerChannelDatabase:
                    {
                        if (oLocalOptions.cloSource is not FlagSource.FileList)
                            saLocalCustomCondition = saLocalCustomCondition.Append("source").ToArray();
                        if (oLocalOptions.cloTarget is not FlagTarget.None)
                            saLocalCustomCondition = saLocalCustomCondition.Append("target").ToArray();
                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("sourcefile").ToArray();

                        break;
                    }
                }

                // Show help dialog if custom conditions are not met
                if (saLocalCustomCondition.Length > 0)
                    DisplayHelpWithCustomCondition(saLocalCustomCondition);
                // Do operation if all custom conditions are met
                else
                {
                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelVariationFile))
                        LoadChannelVariationDictionaryFromFile(strLocalChannelVariationDictionaryFilename: oLocalOptions.cloChannelVariationFile);

                    switch (oLocalOptions.cloMode)
                    {
                        case FlagMode.GenerateData:
                        {
                            if (oLocalOptions.cloSource is FlagSource.ChannelList or FlagSource.CombinedList)
                                GetSourceDataFromChannelList();
                            if (oLocalOptions.cloSource is FlagSource.SmartcardList or FlagSource.CombinedList)
                                GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID);
                            if (oLocalOptions.cloSource is FlagSource.BlindscanList or FlagSource.CombinedList)
                                GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies);
                            if (oLocalOptions.cloSource is FlagSource.FileList)
                                LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);
                            if (oLocalOptions.cloSourceCleanup is not FlagSourceCleanup.None)
                                dtSourceDataTable = CleanupSourceData(oLocalOptions.cloSourceCleanup);
                            if (oLocalOptions.cloTarget is FlagTarget.FileList)
                                SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);
                            if (!string.IsNullOrEmpty(oLocalOptions.cloExportDVBViewerTransponderFile))
                                ExportSourceDataToDVBViewerTransponderFile(strLocalDVBViewerTransponderFilename: oLocalOptions.cloExportDVBViewerTransponderFile);
                            if (!string.IsNullOrEmpty(oLocalOptions.cloExportDVBViewerChannelDatabaseFile))
                                ExportSourceDataToDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloExportDVBViewerChannelDatabaseFile);
                            if (!string.IsNullOrEmpty(oLocalOptions.cloExportExcelChannelFile))
                                ExportSourceDataToExcelExportFile(strLocalExcelFilename: oLocalOptions.cloExportExcelChannelFile);
                            if (!string.IsNullOrEmpty(oLocalOptions.cloExportEnigmaDBFile))
                                ExportSourceDataToEnigmaDBFile(strLocalEnigmaDBFilename: oLocalOptions.cloExportEnigmaDBFile);

                            break;
                        }
                        case FlagMode.ModifyData:
                        {
                            LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);
                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);
                            if (oLocalOptions.cloTarget is FlagTarget.FileList)
                                SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                            break;
                        }
                        case FlagMode.DumpDVBViewerChannelDatabase:
                        {
                            DumpDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloSourceFile);

                            break;
                        }
                    }
                }
            }
            else
            {
                lCurrentLogger.Error("Firewall rule has not been created as there are not sufficient permissions or wrong operation system used. Operation stopped!".Pastel(ConsoleColor.Red));
            }

            if (oLocalOptions.cloKeyPress)
            {
                lCurrentLogger.Info("Press any key to continue ...".Pastel(ConsoleColor.Green));
                Console.ReadKey();
            }
        }
    }
}