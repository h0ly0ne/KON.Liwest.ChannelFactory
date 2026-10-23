using System.Diagnostics;
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
            [Option(longName: "mode", Required = true, HelpText = "Execution Mode or Operation to use")]
            public FlagMode cloMode { get; set; }
            
            [Option(longName: "submode", HelpText = "Execution Sub-Mode or Sub-Operation to use")]
            public FlagSubMode cloSubMode { get; set; }

            [Option(longName: "source", Required = true, HelpText = "Source to use")]
            public FlagSource cloSource { get; set; }
            
            [Option(longName: "target", HelpText = "Target to use")]
            public FlagTarget cloTarget { get; set; }

            [Option(longName: "smartcardid", HelpText = "Smartcard identifier")]
            public string? cloSmartcardID { get; set; }
            
            [Option(longName: "serverip", HelpText = "IP Address of SAT>IP server")]
            public string? cloServerIP { get; set; }
            
            [Option(longName: "frequencies", HelpText = "Frequencies in MHz, comma separated. Ranges as start-end:step.")]
            public string? cloFrequencies { get; set; }
            
            [Option(longName: "sourcefile", HelpText = "Filename for source to load from")]
            public string? cloSourceFile { get; set; }
            
            [Option(longName: "targetfile", HelpText = "Filename for target to save to")]
            public string? cloTargetFile { get; set; }

            [Option(longName: "channelvariationfile", HelpText = "Filename for channel variation dictionary to load from")]
            public string? cloChannelVariationFile { get; set; }
            
            [Option(longName: "sourcecleanup", HelpText = "Source cleanup operation mode")]
            public FlagSourceCleanup cloSourceCleanup { get; set; }
            
            [Option(longName: "channelremovalfile", HelpText = "Filename for channel removal dictionary to load from")]
            public string? cloChannelRemovalFile { get; set; }
            
            [Option(longName: "channelsortingfile", HelpText = "Filename for channel sorting dictionary to load from")]
            public string? cloChannelSortingFile { get; set; }

            [Option(longName: "exportmode", HelpText = "Export mode selection", Default = new FlagExportMode[0])]
            public IEnumerable<FlagExportMode> cloExportMode { get; set; } = [];

            [Option(longName: "exportfilename", HelpText = "Filename for export operation to save to", Default = new string[0])]
            public IEnumerable<string> cloExportFilename { get; set; } = [];

            [Option(longName: "exportfilenameformat", HelpText = "Filename format for export operation to use to save to", Default = new FlagExportEnigmaFormat[0])]
            public IEnumerable<FlagExportEnigmaFormat> cloExportFilenameFormat { get; set; } = [];

            [Option(longName: "dumpdatamode", HelpText = "Dump data mode selection")]
            public FlagDumpDataMode cloDumpDataMode { get; set; }

            [Option(longName: "verbose", HelpText = "Set output to verbose level")]
            public bool cloVerbose { get; set; }
        }

        /// <summary>
        /// Program.WaitForKeyPress()
        /// </summary>
        [Conditional("DEBUG")]
        private static void WaitForKeyPress()
        {
            LogTrace("Program.WaitForKeyPress()");

            if (Console.IsInputRedirected)
                return;

            lCurrentLogger.Info("Press any key to continue ...".Pastel(ConsoleColor.Green));
            Console.ReadKey();
        }

        /// <summary>
        /// Program.Main()
        /// </summary>
        /// <param name="saLocalArguments"></param>
        static void Main(string[] saLocalArguments)
        {
            LogTrace("Program.Main()");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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

        /// <summary>
        /// Program.DisplayHelp()
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="prLocalParseResult"></param>
        private static void DisplayHelp<T>(ParserResult<T> prLocalParseResult)
        {
            LogTrace("Program.DisplayHelp()");

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

        /// <summary>
        /// Program.DisplayHelpWithCustomCondition()
        /// </summary>
        /// <param name="saLocalCustomConditions"></param>
        private static void DisplayHelpWithCustomCondition(string[] saLocalCustomConditions)
        {
            LogTrace("Program.DisplayHelpWithCustomCondition()");

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

        /// <summary>
        /// Program.Run()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        private static void Run(Options oLocalOptions)
        {
            LogTrace("Program.Run()");

            #if DEBUG
            try
            {
                if (OperatingSystem.IsWindows())
                    Console.BufferHeight = short.MaxValue - 1;
            }
            catch (Exception) { /* Not supported by every terminal */ }
            #endif

            InitDataContainers(ref dtSourceDataTable);

            if (CreateFirewallRule())
            {
                string[] saLocalCustomCondition = [];

                /* Check for custom conditions */
                if (!string.IsNullOrEmpty(oLocalOptions.cloChannelVariationFile) && !File.Exists(oLocalOptions.cloChannelVariationFile))
                    saLocalCustomCondition = [.. saLocalCustomCondition, "channelvariationfile"];

                switch (oLocalOptions.cloMode)
                {
                    case FlagMode.GenerateData:
                    {
                        if (oLocalOptions.cloSource == FlagSource.SmartcardList || oLocalOptions.cloSource == FlagSource.CombinedList)
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloSmartcardID) ? [.. saLocalCustomCondition, "smartcardid"] : saLocalCustomCondition;

                        if (oLocalOptions.cloSource == FlagSource.BlindscanList || oLocalOptions.cloSource == FlagSource.CombinedList)
                        {
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloServerIP) ? [.. saLocalCustomCondition, "serverip"] : saLocalCustomCondition;
                            saLocalCustomCondition = string.IsNullOrEmpty(oLocalOptions.cloFrequencies) ? [.. saLocalCustomCondition, "frequencies"] : saLocalCustomCondition;
                        }

                        if (oLocalOptions.cloSource == FlagSource.FileList && (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) || !File.Exists(oLocalOptions.cloSourceFile)))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "sourcefile"];

                        if (oLocalOptions.cloTarget == FlagTarget.FileList && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "targetfile"];

                        break;
                    }
                    case FlagMode.ModifyData:
                    {
                        if (oLocalOptions.cloSource != FlagSource.FileList)
                            saLocalCustomCondition = [.. saLocalCustomCondition, "source"];

                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "sourcefile"];

                        if (oLocalOptions.cloTarget == FlagTarget.FileList && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "targetfile"];

                        if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile) && !File.Exists(oLocalOptions.cloChannelRemovalFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "channelremovalfile"];

                        if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile) && !File.Exists(oLocalOptions.cloChannelSortingFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "channelsortingfile"];

                        break;
                    }
                    case FlagMode.ExportData:
                    {
                        if (oLocalOptions.cloSource != FlagSource.FileList)
                            saLocalCustomCondition = [.. saLocalCustomCondition, "source"];

                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "sourcefile"];

                        if (!Array.TrueForAll([oLocalOptions.cloExportMode.Count(), oLocalOptions.cloExportFilename.Count(), oLocalOptions.cloExportFilenameFormat.Count()], vCurrentValue => oLocalOptions.cloExportMode.Count() == vCurrentValue))
                        {
                            var iMaxParameterCount = oLocalOptions.cloExportMode.Count();

                            if (iMaxParameterCount == 0)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportmode"];

                            if (oLocalOptions.cloExportFilename.Count() != iMaxParameterCount)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilename"];

                            if (oLocalOptions.cloExportFilenameFormat.Count() != iMaxParameterCount)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];
                        }
                        else
                        {
                            for (var iExportDataElement = 0; iExportDataElement < oLocalOptions.cloExportMode.Count(); iExportDataElement++)
                            {
                                if (string.IsNullOrEmpty(oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement)) && oLocalOptions.cloExportMode.ElementAt(iExportDataElement) != FlagExportMode.EnigmaBouquetsFiles)
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilename"];

                                if (oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement) == null && (oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaBouquetsFiles))
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];

                                if (oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement) == FlagExportEnigmaFormat.None && (oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaBouquetsFiles))
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];
                            }
                        }

                        break;
                    }
                    case FlagMode.DumpData:
                    {
                        if (oLocalOptions.cloSource != FlagSource.CustomFile)
                            saLocalCustomCondition = [.. saLocalCustomCondition, "source"];

                        if (string.IsNullOrEmpty(oLocalOptions.cloSourceFile) | !File.Exists(oLocalOptions.cloSourceFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "sourcefile"];

                        break;
                    }
                    case FlagMode.AutomaticData:
                    {
                        if (oLocalOptions.cloSubMode == FlagSubMode.None)
                            saLocalCustomCondition = [.. saLocalCustomCondition, "submode"];

                        if (oLocalOptions.cloSource != FlagSource.ChannelList && oLocalOptions.cloSource != FlagSource.SmartcardList && oLocalOptions.cloSource != FlagSource.BlindscanList && oLocalOptions.cloSource != FlagSource.CombinedList)
                            saLocalCustomCondition = [.. saLocalCustomCondition, "source"];

                        if (string.IsNullOrEmpty(oLocalOptions.cloSmartcardID))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "smartcardid"];

                        if (string.IsNullOrEmpty(oLocalOptions.cloServerIP))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "serverip"];
                            
                        if (string.IsNullOrEmpty(oLocalOptions.cloFrequencies))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "frequencies"];

                        if (oLocalOptions is { cloSubMode: FlagSubMode.SubModeCustom2, cloTarget: FlagTarget.None })
                            saLocalCustomCondition = [.. saLocalCustomCondition, "target"];

                        if (oLocalOptions is { cloSubMode: FlagSubMode.SubModeCustom2, cloTarget: FlagTarget.FileList } && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                            saLocalCustomCondition = [.. saLocalCustomCondition, "targetfile"];

                        if (!Array.TrueForAll([oLocalOptions.cloExportMode.Count(), oLocalOptions.cloExportFilename.Count(), oLocalOptions.cloExportFilenameFormat.Count()], vCurrentValue => oLocalOptions.cloExportMode.Count() == vCurrentValue))
                        {
                            var iMaxParameterCount = oLocalOptions.cloExportMode.Count();

                            if (iMaxParameterCount == 0)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportmode"];

                            if (oLocalOptions.cloExportFilename.Count() != iMaxParameterCount)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilename"];

                            if (oLocalOptions.cloExportFilenameFormat.Count() != iMaxParameterCount)
                                saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];
                        }
                        else
                        {
                            for (var iExportDataElement = 0; iExportDataElement < oLocalOptions.cloExportMode.Count(); iExportDataElement++)
                            {
                                if (string.IsNullOrEmpty(oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement)) && oLocalOptions.cloExportMode.ElementAt(iExportDataElement) != FlagExportMode.EnigmaBouquetsFiles)
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilename"];

                                if (oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement) == null && (oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaBouquetsFiles))
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];

                                if (oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement) == FlagExportEnigmaFormat.None && (oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaTransponderList || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaChannelDatabase || oLocalOptions.cloExportMode.ElementAt(iExportDataElement) == FlagExportMode.EnigmaBouquetsFiles))
                                    saLocalCustomCondition = [.. saLocalCustomCondition, "exportfilenameformat"];
                            }
                        }

                        break;
                    }
                }
                /* End-Check for custom conditions */

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

                            switch (oLocalOptions.cloSource)
                            {
                                case FlagSource.BlindscanList:
                                case FlagSource.CombinedList:
                                {
                                    GetSourceDataFromBlindscan(oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies);
                                    
                                    break;
                                }
                                case FlagSource.FileList:
                                {
                                    LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                                    break;
                                }
                            }

                            if (oLocalOptions.cloSourceCleanup != FlagSourceCleanup.None)
                                CleanupSourceData(oLocalOptions.cloSourceCleanup);

                            switch (oLocalOptions.cloTarget)
                            {
                                case FlagTarget.ConsoleList:
                                {
                                    SaveSourceDataToConsole();

                                    break;
                                }
                                case FlagTarget.FileList:
                                {
                                    SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                                    break;
                                }
                            }

                            break;
                        }
                        case FlagMode.ModifyData:
                        {
                            LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);

                            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);

                            switch (oLocalOptions.cloTarget)
                            {
                                case FlagTarget.ConsoleList:
                                {
                                    SaveSourceDataToConsole();

                                    break;
                                }
                                case FlagTarget.FileList:
                                {
                                    SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                                    break;
                                }
                            }

                            break;
                        }
                        case FlagMode.ExportData:
                        {
                            LoadSourceDataFromFile(strLocalSourceDataFilename: oLocalOptions.cloSourceFile);

                            for (var iExportDataElement = 0; iExportDataElement < oLocalOptions.cloExportMode.Count(); iExportDataElement++)
                            {
                                switch (oLocalOptions.cloExportMode.ElementAt(iExportDataElement))
                                {
                                    case FlagExportMode.ExcelChannelList:
                                    {
                                        if (!string.IsNullOrEmpty(oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement)))
                                            ExportSourceDataToExcelExportFile(strLocalExcelFilename: oLocalOptions.cloExportFilename.ElementAt(iExportDataElement));

                                        break;
                                    }
                                    case FlagExportMode.DVBViewerTransponderList:
                                    {
                                        ExportSourceDataToDVBViewerTransponderFile(strLocalDVBViewerTransponderFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement));

                                        break;
                                    }
                                    case FlagExportMode.DVBViewerChannelDatabase:
                                    {
                                        ExportSourceDataToDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement));

                                        break;
                                    }
                                    case FlagExportMode.EnigmaTransponderList:
                                    case FlagExportMode.EnigmaChannelDatabase:
                                    {
                                        ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode.ElementAt(iExportDataElement), strLocalFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement), fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement));

                                        break;
                                    }
                                    case FlagExportMode.EnigmaBouquetsFiles:
                                    {
                                        ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode.ElementAt(iExportDataElement), strLocalFilename: string.Empty, fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement));

                                        break;
                                    }
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
                        case FlagMode.AutomaticData:
                        {
                            switch (oLocalOptions.cloSubMode)
                            {
                                case FlagSubMode.SubModeCustom1:
                                {
                                    GetSourceDataFromChannelList();
                                    GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID);
                                    GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies);

                                    if (oLocalOptions.cloSourceCleanup != FlagSourceCleanup.None)
                                        CleanupSourceData(oLocalOptions.cloSourceCleanup);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                        ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                        ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);

                                    break;
                                }
                                case FlagSubMode.SubModeCustom2:
                                {
                                    GetSourceDataFromChannelList();
                                    GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID);
                                    GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies);

                                    if (oLocalOptions.cloSourceCleanup != FlagSourceCleanup.None)
                                        CleanupSourceData(oLocalOptions.cloSourceCleanup);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                        ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                        ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);

                                    switch (oLocalOptions.cloTarget)
                                    {
                                        case FlagTarget.ConsoleList:
                                        {
                                            SaveSourceDataToConsole();

                                            break;
                                        }
                                        case FlagTarget.FileList:
                                        {
                                            SaveSourceDataToFile(strLocalSourceDataFilename: oLocalOptions.cloTargetFile);

                                            break;
                                        }
                                    }

                                    break;
                                }
                                case FlagSubMode.SubModeCustom3:
                                {
                                    GetSourceDataFromChannelList();
                                    GetSourceDataFromSmartcard(strLocalSmartcardId: oLocalOptions.cloSmartcardID);
                                    GetSourceDataFromBlindscan(strLocalServerIP: oLocalOptions.cloServerIP, strLocalFrequencies: oLocalOptions.cloFrequencies);

                                    if (oLocalOptions.cloSourceCleanup != FlagSourceCleanup.None)
                                        CleanupSourceData(oLocalOptions.cloSourceCleanup);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile))
                                        ModifyDataRemoveChannels(strLocalChannelRemovalFile: oLocalOptions.cloChannelRemovalFile);

                                    if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile))
                                        ModifyDataSortChannels(strLocalChannelSortingFile: oLocalOptions.cloChannelSortingFile);

                                    for (var iExportDataElement = 0; iExportDataElement < oLocalOptions.cloExportMode.Count(); iExportDataElement++)
                                    {
                                        switch (oLocalOptions.cloExportMode.ElementAt(iExportDataElement))
                                        {
                                            case FlagExportMode.ExcelChannelList:
                                            {
                                                if (!string.IsNullOrEmpty(oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement)))
                                                    ExportSourceDataToExcelExportFile(strLocalExcelFilename: oLocalOptions.cloExportFilename.ElementAt(iExportDataElement));

                                                break;
                                            }
                                            case FlagExportMode.DVBViewerTransponderList:
                                            {
                                                ExportSourceDataToDVBViewerTransponderFile(strLocalDVBViewerTransponderFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement));

                                                break;
                                            }
                                            case FlagExportMode.DVBViewerChannelDatabase:
                                            {
                                                ExportSourceDataToDVBViewerChannelDatabase(strLocalDVBViewerChannelDatabaseFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement));

                                                break;
                                            }
                                            case FlagExportMode.EnigmaTransponderList:
                                            case FlagExportMode.EnigmaChannelDatabase:
                                            {
                                                ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode.ElementAt(iExportDataElement), strLocalFilename: oLocalOptions.cloExportFilename?.ElementAt(iExportDataElement), fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement));

                                                break;
                                            }
                                            case FlagExportMode.EnigmaBouquetsFiles:
                                            {
                                                ExportSourceDataToEnigmaSelector(femLocalFlagExportMode: oLocalOptions.cloExportMode.ElementAt(iExportDataElement), strLocalFilename: string.Empty, fefLocalExportFilenameFormat: oLocalOptions.cloExportFilenameFormat?.ElementAt(iExportDataElement));

                                                break;
                                            }
                                        }
                                    }

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
                lCurrentLogger.Error("-» Firewall rule has not been created as there are not sufficient permissions or wrong operation system used. Operation stopped!".Pastel(ConsoleColor.Red));
            }

            WaitForKeyPress();
        }
    }
}