using System.Drawing;
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
            [Option(longName: "target", HelpText = "Factory target to use")]
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

            [Option(longName: "exportmode", HelpText = "Export mode selection")]
            public FlagExportMode cloExportMode { get; set; }
            [Option(longName: "exportfilename", HelpText = "Filename for export operation to save to")]
            public string? cloExportFilename { get; set; }
            [Option(longName: "exportfilenameformat", HelpText = "Filename format for export operation to use to save to")]
            public FlagExportEnigmaFormat? cloExportFilenameFormat { get; set; }

            [Option(longName: "dumpdatamode", HelpText = "Dump data mode selection")]
            public FlagDumpDataMode cloDumpDataMode { get; set; }

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
                htCurrentHelpText.Heading = Constants.cstrCurrentProjectName.Pastel(Color.Green);
                htCurrentHelpText.Copyright = Constants.cstrCurrentProjectCopyright.Pastel(Color.LimeGreen);

                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Generate ChannelList and output to None".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData -source ChannelList -target None");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Generate SmartcardList and output to ConsoleList".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source SmartcardList --target ConsoleList --smartcardid 12345678900");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Generate BlindscanList and output to FileList".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source BlindscanList --target FileList --serverip 1.2.3.4");
                htCurrentHelpText.AddPostOptionsLine($"    --frequencies 386-786:8 --targetfile SourceData_BlindscanList.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Generate CombinedList and output to FileList".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source BlindscanList --target FileList --smartcardid 12345678900");
                htCurrentHelpText.AddPostOptionsLine($"    --serverip 1.2.3.4 --frequencies 386-786:8 --targetfile SourceData_CombinedList.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Output FileList to ConsoleList".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target ConsoleList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Copy FileList to new FileList".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --targetfile SourceData_FileList_Copy.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and remove invalid SIDs".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --sourcecleanup RemoveInvalidSIDs");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and remove invalid SIDs with variations".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelvariationfile .\\resources\\ChannelVariation.json --sourcecleanup RemoveInvalidSIDs");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and invalidate".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and invalidate with variations".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode GenerateData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelvariationfile .\\resources\\ChannelVariation.json --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and do channel removal".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ModifyData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelremovalfile .\\resources\\ChannelRemoval.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency and do channel sorting".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ModifyData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelsortingfile .\\resources\\ChannelSorting.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Check FileList for consistency with channel removal + sorting".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ModifyData --source FileList --target None --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelremovalfile .\\resources\\ChannelRemoval.json");
                htCurrentHelpText.AddPostOptionsLine($"    --channelsortingfile .\\resources\\ChannelSorting.json");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to Excel channel file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode ExcelChannelList --exportfilename channels.xlsx");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to DVBViewer transponder file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode DVBViewerTransponderList --exportfilename DVB-C_at_LIWEST.ini");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to DVBViewer channel database file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode DVBViewerChannelDatabase --exportfilename channels.dat");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to Enigma transponder file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode EnigmaTransponderList --exportfilename cables.xml");
                htCurrentHelpText.AddPostOptionsLine($"    --exportfilenameformat Engima2Ver4");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to Enigma database file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode EnigmaChannelDatabase --exportfilename lamedb");
                htCurrentHelpText.AddPostOptionsLine($"    --exportfilenameformat Engima2Ver4");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Export FileList to Enigma bouquets files".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode ExportData --source FileList --sourcefile SourceData_FileList.json");
                htCurrentHelpText.AddPostOptionsLine($"    --exportmode EnigmaBouquetsFiles --exportfilenameformat Engima2Ver4");
                htCurrentHelpText.AddPostOptionsLine($"");
                htCurrentHelpText.AddPostOptionsLine($"* e.g.: Dump CustomFile with DVBViewerChannelDatabase to console".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine($"    --mode DumpData --dumpdatamode DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine($"    --source CustomFile --sourcefile channels.dat");
                htCurrentHelpText.AddPostOptionsLine($"");

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
                htCurrentHelpText.Heading = Constants.cstrCurrentProjectName.Pastel(Color.Green);
                htCurrentHelpText.Copyright = Constants.cstrCurrentProjectCopyright.Pastel(Color.LimeGreen);

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
                    case FlagMode.ExportData:
                    {
                        if (oLocalOptions.cloSource is not FlagSource.FileList)
                            saLocalCustomCondition = saLocalCustomCondition.Append("source").ToArray();

                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = saLocalCustomCondition.Append("sourcefile").ToArray();

                        if (string.IsNullOrEmpty(oLocalOptions.cloExportFilename) && oLocalOptions.cloExportMode != FlagExportMode.EnigmaBouquetsFiles)
                            saLocalCustomCondition = saLocalCustomCondition.Append("exportfilename").ToArray();

                        if (oLocalOptions.cloExportFilenameFormat == null && (oLocalOptions.cloExportMode == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode == FlagExportMode.EnigmaBouquetsFiles))
                            saLocalCustomCondition = saLocalCustomCondition.Append("exportfilenameformat").ToArray();

                        if (oLocalOptions.cloExportFilenameFormat == FlagExportEnigmaFormat.None && (oLocalOptions.cloExportMode == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode == FlagExportMode.EnigmaBouquetsFiles))
                            saLocalCustomCondition = saLocalCustomCondition.Append("exportfilenameformat").ToArray();

                        break;
                    }
                    case FlagMode.DumpData:
                    {
                        if (oLocalOptions.cloSource is not FlagSource.CustomFile)
                            saLocalCustomCondition = saLocalCustomCondition.Append("source").ToArray();

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

                            if (oLocalOptions.cloTarget is FlagTarget.ConsoleList)
                                SaveSourceDataToConsole();
                            else if (oLocalOptions.cloTarget is FlagTarget.FileList)
                                SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                            break;
                        }
                        case FlagMode.ModifyData:
                        {
                            LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);

                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);

                            if (oLocalOptions.cloTarget is FlagTarget.ConsoleList)
                                SaveSourceDataToConsole();
                            else if (oLocalOptions.cloTarget is FlagTarget.FileList)
                                SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                            break;
                        }
                        case FlagMode.ExportData:
                        {
                            LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                            switch (oLocalOptions.cloExportMode)
                            {
                                case FlagExportMode.ExcelChannelList:
                                {
                                    if (!string.IsNullOrEmpty(oLocalOptions.cloExportFilename))
                                        ExportSourceDataToExcelExportFile(strLocalExcelFilename: oLocalOptions.cloExportFilename);

                                    break;
                                }
                                case FlagExportMode.DVBViewerTransponderList:
                                {
                                    ExportSourceDataToDVBViewerTransponderFile(strLocalDVBViewerTransponderFilename: oLocalOptions.cloExportFilename);

                                    break;
                                }
                                case FlagExportMode.DVBViewerChannelDatabase:
                                {
                                    ExportSourceDataToDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloExportFilename);

                                    break;
                                }
                                case FlagExportMode.EnigmaTransponderList or FlagExportMode.EnigmaChannelDatabase:
                                {
                                    ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode, strLocalFilename: oLocalOptions.cloExportFilename, fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat);

                                    break;
                                }
                                case FlagExportMode.EnigmaBouquetsFiles:
                                {
                                    ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode, strLocalFilename: string.Empty, fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat);

                                    break;
                                }
                            }

                            break;
                        }
                        case FlagMode.DumpData:
                        {
                            switch (oLocalOptions.cloDumpDataMode)
                            {
                                case FlagDumpDataMode.DVBViewerChannelDatabase:
                                {
                                    DumpDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloSourceFile);

                                    break;
                                }
                            }

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