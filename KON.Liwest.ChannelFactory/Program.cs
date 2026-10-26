using System.Diagnostics;
using System.Drawing;
using System.Text;

using CommandLine;
using CommandLine.Text;
using NLog;
using NLog.Layouts;
using Krkadoni.EnigmaSettings;
using Microsoft.Data.Sqlite;
using Pastel;

using static KON.Liwest.ChannelFactory.DataDumps;
using static KON.Liwest.ChannelFactory.DataExports;
using static KON.Liwest.ChannelFactory.DataModifications;
using static KON.Liwest.ChannelFactory.DataSources;
using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    public class Program
    {
        public class Options
        {
            [Option(longName: "mode", Required = true, HelpText = "Execution-Mode or Operation to use: GenerateData (collect source data), ModifyData (channel removal / sorting of a source data file), ExportData (export a source data file), DumpData (output an existing file), AutomaticData (collect, modify and output / export in one run).")]
            public FlagMode cloMode { get; set; }

            [Option(longName: "submode", HelpText = "Execution-Sub-Mode of AutomaticData: SubModeCustom1 (no output), SubModeCustom2 (output to --target), SubModeCustom3 (exports of --exportmode).")]
            public FlagSubMode cloSubMode { get; set; }

            [Option(longName: "source", Required = true, HelpText = "Source of the data: ChannelList (LIWEST channel list), SmartcardList (LIWEST signal check), BlindscanList (SAT>IP blindscan), CombinedList (all three), FileList (source data file), CustomFile (file for DumpData).")]
            public FlagSource cloSource { get; set; }

            [Option(longName: "target", HelpText = "Output of the source data: None, ConsoleList (table on the console), FileList (source data file --targetfile).")]
            public FlagTarget cloTarget { get; set; }

            [Option(longName: "smartcardid", HelpText = "Smartcard number for the LIWEST signal check (11 digits, starting with 01359 or 02766). Not available between 23:00 and 1:00 UTC.")]
            public string? cloSmartcardID { get; set; }

            [Option(longName: "serverip", HelpText = "IP Address of SAT>IP server for the blindscan.")]
            public string? cloServerIP { get; set; }

            [Option(longName: "frequencies", HelpText = "Frequencies in MHz for the blindscan, comma separated. Ranges as start-end:step.")]
            public string? cloFrequencies { get; set; }

            [Option(longName: "sourcefile", HelpText = "Source data file (json) to load from - for DumpData the file to output.")]
            public string? cloSourceFile { get; set; }

            [Option(longName: "targetfile", HelpText = "Source data file (json) to save to (--target FileList). Relative paths are placed in Export\\<start of the run> next to the executable.")]
            public string? cloTargetFile { get; set; }

            [Option(longName: "channelvariationfile", HelpText = "Channel variation file (json) - unifies the channel names of the different sources.")]
            public string? cloChannelVariationFile { get; set; }

            [Option(longName: "sourcecleanup", HelpText = "Cleanup of the source data: RemoveInvalidSIDs (remove channels without SID), GroupAndInvalidate (merge the sources per channel, remove channels without stream).")]
            public FlagSourceCleanup cloSourceCleanup { get; set; }

            [Option(longName: "channelremovalfile", HelpText = "Channel removal file (json) - channels (name and frequency) to remove.")]
            public string? cloChannelRemovalFile { get; set; }

            [Option(longName: "channelsortingfile", HelpText = "Channel sorting file (json) - sort number per channel (name and frequency), unsorted channels are placed in front.")]
            public string? cloChannelSortingFile { get; set; }

            [Option(longName: "exportmode", HelpText = "Exports, space separated (one entry per export): ExcelChannelList, DVBViewerTransponderList, DVBViewerChannelDatabase, EnigmaTransponderList, EnigmaChannelDatabase, EnigmaBouquetsFiles, PanasonicChannelList.")]
            public IEnumerable<FlagExportMode> cloExportMode { get; set; } = [];

            [Option(longName: "exportfilename", HelpText = "Filename per export, same order as --exportmode (\"\" for EnigmaBouquetsFiles, directory to create for PanasonicChannelList - e.g. hotel_setup). Relative paths are placed in Export\\<start of the run> next to the executable.")]
            public IEnumerable<string> cloExportFilename { get; set; } = [];

            [Option(longName: "exportfilenameformat", HelpText = "Enigma format per export, same order as --exportmode: Enigma1Ver1, Enigma1Ver2 (Enigma 1), Enigma2Ver3, Enigma2Ver4 (default of current images), Enigma2Ver5 (Enigma 2 lamedb version) - None for excel and DVBViewer exports.")]
            public IEnumerable<FlagExportEnigmaFormat> cloExportFilenameFormat { get; set; } = [];

            [Option(longName: "panasonictemplate", HelpText = "Template of the tv for PanasonicChannelList: model of a template in resources\\Panasonic next to the executable (e.g. TX-65FX780W) or the hotel_setup directory exported by the tv (usb channel list copy). The settings of the template are kept, only the channel list is replaced.")]
            public string? cloPanasonicTemplate { get; set; }

            [Option(longName: "installdvbviewerdata", HelpText = "Install the exported DVBViewer channel database and transponder list into all installed DVBViewer / DVBViewer Media Server instances (ExportData or AutomaticData SubModeCustom3). Running instances are stopped and restarted, existing files are backed up.")]
            public bool cloInstallDVBViewerData { get; set; }

            [Option(longName: "installenigmadata", HelpText = "Install the exported Enigma2 data into the receivers of --enigmahost via SSH (ExportData or AutomaticData SubModeCustom3). The settings of the receiver are backed up locally and merged: other bouquets (e.g. streaming), other cables and foreign lamedb entries are kept, enigma2 is stopped and restarted.")]
            public bool cloInstallEnigmaData { get; set; }

            [Option(longName: "enigmahost", HelpText = "Enigma2 receivers for --installenigmadata, space separated (host or host:port of the SSH server).")]
            public IEnumerable<string> cloEnigmaHost { get; set; } = [];

            [Option(longName: "enigmauser", Default = "root", HelpText = "SSH user of the Enigma2 receivers.")]
            public string cloEnigmaUser { get; set; } = "root";

            [Option(longName: "enigmapassword", HelpText = "SSH password of the Enigma2 receivers (required for --installenigmadata).")]
            public string? cloEnigmaPassword { get; set; }

            [Option(longName: "dryrun", HelpText = "Installation test without changes: finds the DVBViewer installations / connects to the Enigma2 receivers and backs up their settings, nothing is installed.")]
            public bool cloDryRun { get; set; }

            [Option(longName: "keepexports", HelpText = "Number of run directories in Export next to the executable to keep (including the current run) - older ones are removed. Without this option all run directories are kept.")]
            public int? cloKeepExports { get; set; }

            [Option(longName: "dumpdatamode", HelpText = "Format of the file to output with DumpData.")]
            public FlagDumpDataMode cloDumpDataMode { get; set; }

            [Option(longName: "verbose", HelpText = "Set output to verbose level.")]
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
        /// <returns>Exit code (0 = success, 1 = invalid arguments, 2 = operation failed, 3 = prerequisites not met)</returns>
        static int Main(string[] saLocalArguments)
        {
            LogTrace("Program.Main()");

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var pCurrentParser = new Parser(apsCurrentConfiguration => apsCurrentConfiguration.HelpWriter = null);
            var prCurrentParseResult = pCurrentParser.ParseArguments<Options>(saLocalArguments);
            var ecCurrentExitCode = ExitCode.InvalidArguments;

            prCurrentParseResult.WithParsed(delegate (Options oLocalOptions)
            {
                ConfigureLogging(oLocalOptions.cloVerbose ? LogLevel.Trace : LogLevel.Info);

                try
                {
                    ecCurrentExitCode = Run(oLocalOptions);
                }
                catch (Exception excCurrentException)
                {
                    LogError($"-» Unexpected error: {excCurrentException.Message}");
                    LogTrace(excCurrentException.ToString());
                    ecCurrentExitCode = ExitCode.OperationFailed;
                }
            });

            prCurrentParseResult.WithNotParsed(delegate (IEnumerable<Error> ieeLocalErrors)
            {
                ConfigureLogging(LogLevel.Info);

                DisplayHelp(prCurrentParseResult);

                var eaCurrentErrors = ieeLocalErrors.ToArray();
                ecCurrentExitCode = eaCurrentErrors.IsHelp() || eaCurrentErrors.IsVersion() ? ExitCode.Success : ExitCode.InvalidArguments;
            });

            WaitForKeyPress();

            return (int)ecCurrentExitCode;
        }

        /// <summary>
        /// Program.ConfigureLogging()
        /// </summary>
        /// <param name="llLocalMinimumLogLevel"></param>
        private static void ConfigureLogging(LogLevel llLocalMinimumLogLevel)
        {
            LogManager.Setup().LoadConfiguration(aslcbCurrentSetupLoadConfigurationBuilder => { aslcbCurrentSetupLoadConfigurationBuilder.ForLogger().FilterMinLevel(llLocalMinimumLogLevel).WriteToColoredConsole(new SimpleLayout("${message}")); });
        }

        /// <summary>
        /// Program.DisplayHelp()
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="prLocalParseResult"></param>
        private static void DisplayHelp<T>(ParserResult<T> prLocalParseResult)
        {
            LogTrace("Program.DisplayHelp()");

            var htCurrentHelpText = HelpText.AutoBuild(prLocalParseResult, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = false;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = Constants.cstrCurrentProjectName.Pastel(Color.Green);
                htCurrentHelpText.Copyright = Constants.cstrCurrentProjectCopyright.Pastel(Color.LimeGreen);

                htCurrentHelpText.AddPostOptionsLine("* e.g.: Channel list to the console".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source ChannelList --target ConsoleList");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Channel list to a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source ChannelList --target FileList");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_ChannelList.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Smartcard (LIWEST signal check) to a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source SmartcardList --target FileList --smartcardid 01359xxxxxx");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_SmartcardList.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Blindscan (SAT>IP) to a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source BlindscanList --target FileList --serverip 192.168.1.10");
                htCurrentHelpText.AddPostOptionsLine("    --frequencies 114,130-858:8 --targetfile SourceData_BlindscanList.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: All sources merged per channel to a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source CombinedList --target FileList --smartcardid 01359xxxxxx");
                htCurrentHelpText.AddPostOptionsLine("    --serverip 192.168.1.10 --frequencies 114,130-858:8 --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine("    --channelvariationfile .\\resources\\ChannelVariation.json");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_CombinedList.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Check a source data file and output it to the console".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source FileList --target ConsoleList");
                htCurrentHelpText.AddPostOptionsLine("    --sourcefile SourceData_CombinedList.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Remove channels without SID from a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source FileList --target FileList");
                htCurrentHelpText.AddPostOptionsLine("    --sourcefile SourceData_CombinedList.json --sourcecleanup RemoveInvalidSIDs");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_Cleaned.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Merge the sources of a source data file per channel".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source FileList --target FileList");
                htCurrentHelpText.AddPostOptionsLine("    --sourcefile SourceData_CombinedList.json --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine("    --channelvariationfile .\\resources\\ChannelVariation.json");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_Grouped.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Remove and sort channels of a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ModifyData --source FileList --target FileList --sourcefile SourceData_Grouped.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelremovalfile .\\resources\\ChannelRemoval.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelsortingfile .\\resources\\ChannelSorting.json");
                htCurrentHelpText.AddPostOptionsLine("    --targetfile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to an excel channel list".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode ExcelChannelList --exportfilename channels.xlsx --exportfilenameformat None");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to DVBViewer transponder list and channel database".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to DVBViewer and install into DVBViewer / Media Server".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None");
                htCurrentHelpText.AddPostOptionsLine("    --installdvbviewerdata");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Test the installation into DVBViewer without changes".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None");
                htCurrentHelpText.AddPostOptionsLine("    --installdvbviewerdata --dryrun");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Test the installation into an Enigma2 receiver (SSH, backup only)".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename cables.xml lamedb \"\"");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4");
                htCurrentHelpText.AddPostOptionsLine("    --installenigmadata --enigmahost 192.168.1.20 --enigmauser root");
                htCurrentHelpText.AddPostOptionsLine("    --enigmapassword xxxxxx --dryrun");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to Enigma2 and install into a receiver (SSH)".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename cables.xml lamedb \"\"");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4");
                htCurrentHelpText.AddPostOptionsLine("    --installenigmadata --enigmahost 192.168.1.20 --enigmauser root");
                htCurrentHelpText.AddPostOptionsLine("    --enigmapassword xxxxxx");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to Enigma2 transponders, channel database and bouquets".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename cables.xml lamedb \"\"");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to a Panasonic channel list (template of the model)".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode PanasonicChannelList --exportfilename hotel_setup --exportfilenameformat None");
                htCurrentHelpText.AddPostOptionsLine("    --panasonictemplate TX-65FX780W");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to a Panasonic channel list (template exported by the tv)".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode PanasonicChannelList --exportfilename hotel_setup --exportfilenameformat None");
                htCurrentHelpText.AddPostOptionsLine("    --panasonictemplate E:\\hotel_setup");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Export to all formats into the folder export".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode ExcelChannelList DVBViewerTransponderList DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    EnigmaTransponderList EnigmaChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilename export\\channels.xlsx export\\DVB-C_at_LIWEST.ini");
                htCurrentHelpText.AddPostOptionsLine("    export\\channels.dat export\\cables.xml export\\lamedb");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilenameformat None None None Enigma2Ver4 Enigma2Ver4");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Output a DVBViewer channel database to the console".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode DumpData --source CustomFile --dumpdatamode DVBViewerChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    --sourcefile channels.dat");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Automatic run without output (collect, cleanup, modify only)".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode AutomaticData --submode SubModeCustom1 --source CombinedList");
                htCurrentHelpText.AddPostOptionsLine("    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8");
                htCurrentHelpText.AddPostOptionsLine("    --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Automatic run with output to a source data file".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode AutomaticData --submode SubModeCustom2 --source CombinedList");
                htCurrentHelpText.AddPostOptionsLine("    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8");
                htCurrentHelpText.AddPostOptionsLine("    --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine("    --channelvariationfile .\\resources\\ChannelVariation.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelremovalfile .\\resources\\ChannelRemoval.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelsortingfile .\\resources\\ChannelSorting.json");
                htCurrentHelpText.AddPostOptionsLine("    --target FileList --targetfile SourceData_Automatic.json");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Automatic run with exports to excel, DVBViewer and Enigma2".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode AutomaticData --submode SubModeCustom3 --source CombinedList");
                htCurrentHelpText.AddPostOptionsLine("    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8");
                htCurrentHelpText.AddPostOptionsLine("    --sourcecleanup GroupAndInvalidate");
                htCurrentHelpText.AddPostOptionsLine("    --channelvariationfile .\\resources\\ChannelVariation.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelremovalfile .\\resources\\ChannelRemoval.json");
                htCurrentHelpText.AddPostOptionsLine("    --channelsortingfile .\\resources\\ChannelSorting.json");
                htCurrentHelpText.AddPostOptionsLine("    --exportmode ExcelChannelList DVBViewerChannelDatabase EnigmaChannelDatabase");
                htCurrentHelpText.AddPostOptionsLine("    EnigmaBouquetsFiles --exportfilename channels.xlsx channels.dat lamedb \"\"");
                htCurrentHelpText.AddPostOptionsLine("    --exportfilenameformat None None Enigma2Ver4 Enigma2Ver4");
                htCurrentHelpText.AddPostOptionsLine("");
                htCurrentHelpText.AddPostOptionsLine("* e.g.: Channel list with verbose (trace) output".Pastel(Color.Chocolate));
                htCurrentHelpText.AddPostOptionsLine("    --mode GenerateData --source ChannelList --target None --verbose");
                htCurrentHelpText.AddPostOptionsLine("");

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

            var prCurrentParseResult = new Parser(x => { x.HelpWriter = null; }).ParseArguments<Options>(["--help"]);
            var htCurrentHelpText = HelpText.AutoBuild(prCurrentParseResult, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = true;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = Constants.cstrCurrentProjectName.Pastel(Color.Green);
                htCurrentHelpText.Copyright = Constants.cstrCurrentProjectCopyright.Pastel(Color.LimeGreen);

                if (saLocalCustomConditions.Length <= 0)
                    return HelpText.DefaultParsingErrorsHandler(prCurrentParseResult, htCurrentHelpText);
                
                htCurrentHelpText.AddPreOptionsLine("ERROR(S):");

                foreach (var strCurrentCustomCondition in saLocalCustomConditions)
                {
                    if (!string.IsNullOrEmpty(strCurrentCustomCondition))
                        htCurrentHelpText.AddPreOptionsLine($"  * Custom option '{strCurrentCustomCondition}' not met or missing.");
                }

                return HelpText.DefaultParsingErrorsHandler(prCurrentParseResult, htCurrentHelpText);
            }, e => e);

            lCurrentLogger.Info(htCurrentHelpText);
        }

        /// <summary>
        /// Program.Run()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <returns>ExitCode</returns>
        private static ExitCode Run(Options oLocalOptions)
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

            if (!CreateFirewallRule())
            {
                LogError("-» Firewall rule has not been created as there are not sufficient permissions or wrong operation system used. Operation stopped!");
                return ExitCode.PrerequisitesNotMet;
            }

            // Show help dialog if custom conditions are not met
            var saCurrentCustomConditions = ValidateOptions(oLocalOptions);
            if (saCurrentCustomConditions.Count > 0)
            {
                DisplayHelpWithCustomCondition([.. saCurrentCustomConditions]);
                return ExitCode.InvalidArguments;
            }

            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelVariationFile) && !LoadChannelVariationDictionaryFromFile(oLocalOptions.cloChannelVariationFile))
                return ExitCode.OperationFailed;

            var lsdiCurrentSourceData = new List<SourceDataItem>();

            switch (oLocalOptions.cloMode)
            {
                case FlagMode.GenerateData:
                {
                    if (oLocalOptions.cloSource is FlagSource.ChannelList or FlagSource.CombinedList)
                        GetSourceDataFromChannelList(lsdiCurrentSourceData);

                    if (oLocalOptions.cloSource is FlagSource.SmartcardList or FlagSource.CombinedList)
                        GetSourceDataFromSmartcard(lsdiCurrentSourceData, oLocalOptions.cloSmartcardID);

                    if (oLocalOptions.cloSource is FlagSource.BlindscanList or FlagSource.CombinedList)
                        GetSourceDataFromBlindscan(lsdiCurrentSourceData, oLocalOptions.cloServerIP, oLocalOptions.cloFrequencies);

                    if (oLocalOptions.cloSource == FlagSource.FileList && !LoadSourceDataFromFile(lsdiCurrentSourceData, oLocalOptions.cloSourceFile))
                        break;

                    CleanupSourceData(lsdiCurrentSourceData, oLocalOptions.cloSourceCleanup);
                    WriteTarget(oLocalOptions, lsdiCurrentSourceData);

                    break;
                }
                case FlagMode.ModifyData:
                {
                    if (!LoadSourceDataFromFile(lsdiCurrentSourceData, oLocalOptions.cloSourceFile))
                        break;

                    ModifySourceData(oLocalOptions, lsdiCurrentSourceData);
                    WriteTarget(oLocalOptions, lsdiCurrentSourceData);

                    break;
                }
                case FlagMode.ExportData:
                {
                    if (!LoadSourceDataFromFile(lsdiCurrentSourceData, oLocalOptions.cloSourceFile))
                        break;

                    RunExportsAndInstall(oLocalOptions, lsdiCurrentSourceData);

                    break;
                }
                case FlagMode.DumpData:
                {
                    if (oLocalOptions.cloDumpDataMode == FlagDumpDataMode.DVBViewerChannelDatabase)
                        DumpDVBViewerChannelDatabase(oLocalOptions.cloSourceFile);

                    break;
                }
                case FlagMode.AutomaticData:
                {
                    GetSourceDataFromChannelList(lsdiCurrentSourceData);
                    GetSourceDataFromSmartcard(lsdiCurrentSourceData, oLocalOptions.cloSmartcardID);
                    GetSourceDataFromBlindscan(lsdiCurrentSourceData, oLocalOptions.cloServerIP, oLocalOptions.cloFrequencies);

                    CleanupSourceData(lsdiCurrentSourceData, oLocalOptions.cloSourceCleanup);
                    ModifySourceData(oLocalOptions, lsdiCurrentSourceData);

                    switch (oLocalOptions.cloSubMode)
                    {
                        case FlagSubMode.SubModeCustom1:
                            // Sources, cleanup, removal and sorting only - no output
                            break;
                        case FlagSubMode.SubModeCustom2:
                            WriteTarget(oLocalOptions, lsdiCurrentSourceData);
                            break;
                        case FlagSubMode.SubModeCustom3:
                            RunExportsAndInstall(oLocalOptions, lsdiCurrentSourceData);
                            break;
                    }

                    break;
                }
            }

            if (oLocalOptions.cloKeepExports is { } iCurrentKeepExports)
                CleanupExportDirectories(iCurrentKeepExports);

            return bErrorOccurred ? ExitCode.OperationFailed : ExitCode.Success;
        }

        /// <summary>
        /// Program.ValidateOptions() - checks the options which depend on mode, source, target and export selection
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <returns>Names of the options which are not met or missing</returns>
        private static List<string> ValidateOptions(Options oLocalOptions)
        {
            LogTrace("Program.ValidateOptions()");

            List<string> lsCurrentCustomConditions = [];

            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelVariationFile) && !File.Exists(oLocalOptions.cloChannelVariationFile))
                lsCurrentCustomConditions.Add("channelvariationfile");

            switch (oLocalOptions.cloMode)
            {
                case FlagMode.GenerateData:
                {
                    if (oLocalOptions.cloSource is FlagSource.SmartcardList or FlagSource.CombinedList && string.IsNullOrEmpty(oLocalOptions.cloSmartcardID))
                        lsCurrentCustomConditions.Add("smartcardid");

                    if (oLocalOptions.cloSource is FlagSource.BlindscanList or FlagSource.CombinedList)
                    {
                        if (string.IsNullOrEmpty(oLocalOptions.cloServerIP))
                            lsCurrentCustomConditions.Add("serverip");

                        if (string.IsNullOrEmpty(oLocalOptions.cloFrequencies))
                            lsCurrentCustomConditions.Add("frequencies");
                    }

                    if (oLocalOptions.cloSource == FlagSource.FileList && !IsExistingFile(oLocalOptions.cloSourceFile))
                        lsCurrentCustomConditions.Add("sourcefile");

                    ValidateTargetOptions(oLocalOptions, lsCurrentCustomConditions);
                    break;
                }
                case FlagMode.ModifyData:
                {
                    ValidateSourceFileOptions(oLocalOptions, FlagSource.FileList, lsCurrentCustomConditions);
                    ValidateTargetOptions(oLocalOptions, lsCurrentCustomConditions);
                    ValidateModifyOptions(oLocalOptions, lsCurrentCustomConditions);
                    break;
                }
                case FlagMode.ExportData:
                {
                    ValidateSourceFileOptions(oLocalOptions, FlagSource.FileList, lsCurrentCustomConditions);
                    ValidateExportOptions(oLocalOptions, lsCurrentCustomConditions);
                    break;
                }
                case FlagMode.DumpData:
                {
                    ValidateSourceFileOptions(oLocalOptions, FlagSource.CustomFile, lsCurrentCustomConditions);
                    break;
                }
                case FlagMode.AutomaticData:
                {
                    if (oLocalOptions.cloSubMode == FlagSubMode.None)
                        lsCurrentCustomConditions.Add("submode");

                    if (oLocalOptions.cloSource is not (FlagSource.ChannelList or FlagSource.SmartcardList or FlagSource.BlindscanList or FlagSource.CombinedList))
                        lsCurrentCustomConditions.Add("source");

                    if (string.IsNullOrEmpty(oLocalOptions.cloSmartcardID))
                        lsCurrentCustomConditions.Add("smartcardid");

                    if (string.IsNullOrEmpty(oLocalOptions.cloServerIP))
                        lsCurrentCustomConditions.Add("serverip");

                    if (string.IsNullOrEmpty(oLocalOptions.cloFrequencies))
                        lsCurrentCustomConditions.Add("frequencies");

                    if (oLocalOptions is { cloSubMode: FlagSubMode.SubModeCustom2, cloTarget: FlagTarget.None })
                        lsCurrentCustomConditions.Add("target");

                    if (oLocalOptions is { cloSubMode: FlagSubMode.SubModeCustom2, cloTarget: FlagTarget.FileList } && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                        lsCurrentCustomConditions.Add("targetfile");

                    ValidateExportOptions(oLocalOptions, lsCurrentCustomConditions);
                    break;
                }
            }

            // Installation of DVBViewer data only together with a DVBViewer export
            if (oLocalOptions.cloInstallDVBViewerData && (oLocalOptions is not ({ cloMode: FlagMode.ExportData } or { cloMode: FlagMode.AutomaticData, cloSubMode: FlagSubMode.SubModeCustom3 }) || !oLocalOptions.cloExportMode.Any(femCurrentExportMode => femCurrentExportMode is FlagExportMode.DVBViewerChannelDatabase or FlagExportMode.DVBViewerTransponderList)))
                lsCurrentCustomConditions.Add("installdvbviewerdata");

            // Installation of Enigma2 data only together with an Enigma2 export and a receiver
            if (oLocalOptions.cloInstallEnigmaData && (oLocalOptions is not ({ cloMode: FlagMode.ExportData } or { cloMode: FlagMode.AutomaticData, cloSubMode: FlagSubMode.SubModeCustom3 }) || !oLocalOptions.cloExportMode.Any(femCurrentExportMode => femCurrentExportMode is FlagExportMode.EnigmaChannelDatabase or FlagExportMode.EnigmaTransponderList or FlagExportMode.EnigmaBouquetsFiles)))
                lsCurrentCustomConditions.Add("installenigmadata");

            if (oLocalOptions.cloInstallEnigmaData && oLocalOptions.cloEnigmaHost.All(string.IsNullOrWhiteSpace))
                lsCurrentCustomConditions.Add("enigmahost");

            if (oLocalOptions.cloInstallEnigmaData && (string.IsNullOrEmpty(oLocalOptions.cloEnigmaUser) || string.IsNullOrEmpty(oLocalOptions.cloEnigmaPassword)))
                lsCurrentCustomConditions.Add(string.IsNullOrEmpty(oLocalOptions.cloEnigmaUser) ? "enigmauser" : "enigmapassword");

            if (oLocalOptions.cloKeepExports is < 1)
                lsCurrentCustomConditions.Add("keepexports");

            if (oLocalOptions.cloDryRun && oLocalOptions is { cloInstallDVBViewerData: false, cloInstallEnigmaData: false })
                lsCurrentCustomConditions.Add("dryrun");

            return lsCurrentCustomConditions;
        }

        /// <summary>
        /// Program.ValidateSourceFileOptions()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="fsLocalRequiredSource"></param>
        /// <param name="lsLocalCustomConditions"></param>
        private static void ValidateSourceFileOptions(Options oLocalOptions, FlagSource fsLocalRequiredSource, List<string> lsLocalCustomConditions)
        {
            if (oLocalOptions.cloSource != fsLocalRequiredSource)
                lsLocalCustomConditions.Add("source");

            if (!IsExistingFile(oLocalOptions.cloSourceFile))
                lsLocalCustomConditions.Add("sourcefile");
        }

        /// <summary>
        /// Program.ValidateTargetOptions()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsLocalCustomConditions"></param>
        private static void ValidateTargetOptions(Options oLocalOptions, List<string> lsLocalCustomConditions)
        {
            if (oLocalOptions.cloTarget == FlagTarget.FileList && string.IsNullOrEmpty(oLocalOptions.cloTargetFile))
                lsLocalCustomConditions.Add("targetfile");
        }

        /// <summary>
        /// Program.ValidateModifyOptions()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsLocalCustomConditions"></param>
        private static void ValidateModifyOptions(Options oLocalOptions, List<string> lsLocalCustomConditions)
        {
            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelRemovalFile) && !File.Exists(oLocalOptions.cloChannelRemovalFile))
                lsLocalCustomConditions.Add("channelremovalfile");

            if (!string.IsNullOrEmpty(oLocalOptions.cloChannelSortingFile) && !File.Exists(oLocalOptions.cloChannelSortingFile))
                lsLocalCustomConditions.Add("channelsortingfile");
        }

        /// <summary>
        /// Program.ValidateExportOptions() - export mode, filename and filename format are given as lists of the same length
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsLocalCustomConditions"></param>
        private static void ValidateExportOptions(Options oLocalOptions, List<string> lsLocalCustomConditions)
        {
            var femaCurrentExportModes = oLocalOptions.cloExportMode.ToArray();
            var saCurrentExportFilenames = oLocalOptions.cloExportFilename.ToArray();
            var fefaCurrentExportFilenameFormats = oLocalOptions.cloExportFilenameFormat.ToArray();

            if (saCurrentExportFilenames.Length != femaCurrentExportModes.Length || fefaCurrentExportFilenameFormats.Length != femaCurrentExportModes.Length)
            {
                if (femaCurrentExportModes.Length == 0)
                    lsLocalCustomConditions.Add("exportmode");

                if (saCurrentExportFilenames.Length != femaCurrentExportModes.Length)
                    lsLocalCustomConditions.Add("exportfilename");

                if (fefaCurrentExportFilenameFormats.Length != femaCurrentExportModes.Length)
                    lsLocalCustomConditions.Add("exportfilenameformat");

                return;
            }

            for (var iCurrentExportElement = 0; iCurrentExportElement < femaCurrentExportModes.Length; iCurrentExportElement++)
            {
                if (string.IsNullOrEmpty(saCurrentExportFilenames[iCurrentExportElement]) && femaCurrentExportModes[iCurrentExportElement] != FlagExportMode.EnigmaBouquetsFiles)
                    lsLocalCustomConditions.Add("exportfilename");

                if (fefaCurrentExportFilenameFormats[iCurrentExportElement] == FlagExportEnigmaFormat.None && femaCurrentExportModes[iCurrentExportElement] is FlagExportMode.EnigmaTransponderList or FlagExportMode.EnigmaChannelDatabase or FlagExportMode.EnigmaBouquetsFiles)
                    lsLocalCustomConditions.Add("exportfilenameformat");
            }

            // Panasonic channel list only with an existing template (model or hotel_setup directory)
            if (femaCurrentExportModes.Contains(FlagExportMode.PanasonicChannelList) && Panasonic.ResolveTemplateDirectory(oLocalOptions.cloPanasonicTemplate) == null)
                lsLocalCustomConditions.Add("panasonictemplate");
        }

        /// <summary>
        /// Program.IsExistingFile()
        /// </summary>
        /// <param name="strLocalFilename"></param>
        /// <returns>bool</returns>
        private static bool IsExistingFile(string? strLocalFilename)
        {
            return !string.IsNullOrEmpty(strLocalFilename) && File.Exists(strLocalFilename);
        }

        /// <summary>
        /// Program.ModifySourceData() - channel removal and channel sorting
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsdiLocalSourceData"></param>
        private static void ModifySourceData(Options oLocalOptions, List<SourceDataItem> lsdiLocalSourceData)
        {
            ModifyDataRemoveChannels(lsdiLocalSourceData, oLocalOptions.cloChannelRemovalFile);
            ModifyDataSortChannels(lsdiLocalSourceData, oLocalOptions.cloChannelSortingFile);
        }

        /// <summary>
        /// Program.WriteTarget()
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsdiLocalSourceData"></param>
        private static void WriteTarget(Options oLocalOptions, List<SourceDataItem> lsdiLocalSourceData)
        {
            switch (oLocalOptions.cloTarget)
            {
                case FlagTarget.ConsoleList:
                    SaveSourceDataToConsole(lsdiLocalSourceData);
                    break;
                case FlagTarget.FileList:
                    ExecuteFileOperation(FlagTarget.FileList, () => SaveSourceDataToFile(lsdiLocalSourceData, oLocalOptions.cloTargetFile));
                    break;
            }
        }

        /// <summary>
        /// Program.RunExportsAndInstall() - exports and installation of the exported data (--installdvbviewerdata, --installenigmadata)
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsdiLocalSourceData"></param>
        private static void RunExportsAndInstall(Options oLocalOptions, List<SourceDataItem> lsdiLocalSourceData)
        {
            var lerCurrentExportResults = RunExports(oLocalOptions, lsdiLocalSourceData);

            if (oLocalOptions.cloInstallDVBViewerData)
                InstallDVBViewerData(lerCurrentExportResults, oLocalOptions.cloDryRun);

            if (oLocalOptions.cloInstallEnigmaData)
                InstallEnigmaData(oLocalOptions, lerCurrentExportResults);
        }

        /// <summary>
        /// Program.RunExports() - every export runs on its own, a failing export does not stop the following ones
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lsdiLocalSourceData"></param>
        /// <returns>Successful exports</returns>
        private static List<ExportResult> RunExports(Options oLocalOptions, List<SourceDataItem> lsdiLocalSourceData)
        {
            var femaCurrentExportModes = oLocalOptions.cloExportMode.ToArray();
            var saCurrentExportFilenames = oLocalOptions.cloExportFilename.ToArray();
            var fefaCurrentExportFilenameFormats = oLocalOptions.cloExportFilenameFormat.ToArray();
            List<ExportResult> lerCurrentExportResults = [];

            for (var iCurrentExportElement = 0; iCurrentExportElement < femaCurrentExportModes.Length; iCurrentExportElement++)
            {
                var femCurrentExportMode = femaCurrentExportModes[iCurrentExportElement];
                var strCurrentExportFilename = iCurrentExportElement < saCurrentExportFilenames.Length ? saCurrentExportFilenames[iCurrentExportElement] : null;
                FlagExportEnigmaFormat? fefCurrentExportFilenameFormat = iCurrentExportElement < fefaCurrentExportFilenameFormats.Length ? fefaCurrentExportFilenameFormats[iCurrentExportElement] : null;
                var dtCurrentExportStart = DateTime.UtcNow.AddSeconds(-1);

                if (!ExecuteFileOperation(femCurrentExportMode, () => RunExport(lsdiLocalSourceData, femCurrentExportMode, strCurrentExportFilename, fefCurrentExportFilenameFormat, oLocalOptions.cloPanasonicTemplate)))
                    continue;

                // Bouquet files are written to the export directory, all other exports to their filename
                var strCurrentPath = femCurrentExportMode == FlagExportMode.EnigmaBouquetsFiles ? GetExportDirectory() : string.IsNullOrEmpty(strCurrentExportFilename) ? null : GetOutputFilePath(strCurrentExportFilename);
                if (strCurrentPath != null)
                    lerCurrentExportResults.Add(new ExportResult(femCurrentExportMode, strCurrentPath, dtCurrentExportStart));
            }

            return lerCurrentExportResults;
        }

        /// <summary>
        /// Program.InstallDVBViewerData() - installs the exported channel database / transponder list into all DVBViewer / DVBViewer Media Server installations
        /// </summary>
        /// <param name="lerLocalExportResults"></param>
        /// <param name="bLocalDryRun">Only show what would be installed</param>
        private static void InstallDVBViewerData(List<ExportResult> lerLocalExportResults, bool bLocalDryRun)
        {
            LogTrace("Program.InstallDVBViewerData()");
            LogInfo($"-» Install DVBViewer data{(bLocalDryRun ? " (dry run)" : string.Empty)}");

            if (!OperatingSystem.IsWindows())
            {
                LogError("-» DVBViewer data can only be installed on windows");
                return;
            }

            // Only files written by this run (an export without data writes no file)
            var ltCurrentFiles = lerLocalExportResults
                .Where(erCurrentExportResult => erCurrentExportResult.femExportMode is FlagExportMode.DVBViewerChannelDatabase or FlagExportMode.DVBViewerTransponderList && erCurrentExportResult.IsWrittenFile(erCurrentExportResult.strPath))
                .Select(erCurrentExportResult => (erCurrentExportResult.strPath, erCurrentExportResult.femExportMode == FlagExportMode.DVBViewerChannelDatabase ? Constants.csDVBViewerChannelDatabaseFile : Path.Combine(Constants.csDVBViewerTransponderDirectory, Path.GetFileName(erCurrentExportResult.strPath))))
                .ToList();

            if (ltCurrentFiles.Count == 0)
            {
                LogError("-» No DVBViewer data has been exported in this run - nothing to install");
                return;
            }

            var ldviCurrentInstances = DataInstallations.FindDVBViewerInstances();
            if (ldviCurrentInstances.Count == 0)
            {
                LogError("-» No DVBViewer / DVBViewer Media Server installation found");
                return;
            }

            foreach (var dviCurrentInstance in ldviCurrentInstances)
                DataInstallations.InstallDVBViewerData(dviCurrentInstance, ltCurrentFiles, bLocalDryRun);
        }

        /// <summary>
        /// Program.InstallEnigmaData() - installs the exported Enigma2 data into the receivers of --enigmahost (SSH / SFTP)
        /// </summary>
        /// <param name="oLocalOptions"></param>
        /// <param name="lerLocalExportResults"></param>
        private static void InstallEnigmaData(Options oLocalOptions, List<ExportResult> lerLocalExportResults)
        {
            LogTrace("Program.InstallEnigmaData()");
            LogInfo($"-» Install Enigma2 data{(oLocalOptions.cloDryRun ? " (dry run)" : string.Empty)}");

            List<(string strSourceFile, string strTargetFile)> ltCurrentFiles = [];

            foreach (var erCurrentExportResult in lerLocalExportResults)
            {
                switch (erCurrentExportResult.femExportMode)
                {
                    case FlagExportMode.EnigmaChannelDatabase:
                    {
                        if (erCurrentExportResult.IsWrittenFile(erCurrentExportResult.strPath))
                            ltCurrentFiles.Add((erCurrentExportResult.strPath, $"{Constants.csEnigmaSettingsDirectory}/{Constants.csEnigmaChannelDatabaseFile}"));

                        if (erCurrentExportResult.IsWrittenFile(erCurrentExportResult.strPath + "5"))
                            ltCurrentFiles.Add((erCurrentExportResult.strPath + "5", $"{Constants.csEnigmaSettingsDirectory}/{Constants.csEnigmaChannelDatabaseFile}5"));

                        break;
                    }
                    case FlagExportMode.EnigmaTransponderList:
                    {
                        if (erCurrentExportResult.IsWrittenFile(erCurrentExportResult.strPath))
                            ltCurrentFiles.Add((erCurrentExportResult.strPath, Constants.csEnigmaCablesFile));

                        break;
                    }
                    case FlagExportMode.EnigmaBouquetsFiles:
                    {
                        ltCurrentFiles.AddRange(from strCurrentFile in Constants.csaEnigmaBouquetFilePatterns.SelectMany(strCurrentPattern => Directory.GetFiles(erCurrentExportResult.strPath, strCurrentPattern)).Distinct(StringComparer.OrdinalIgnoreCase) where erCurrentExportResult.IsWrittenFile(strCurrentFile) select (strCurrentFile, $"{Constants.csEnigmaSettingsDirectory}/{Path.GetFileName(strCurrentFile)}"));

                        break;
                    }
                }
            }

            if (ltCurrentFiles.Count == 0)
            {
                LogError("-» No Enigma2 data has been exported in this run - nothing to install");
                return;
            }

            foreach (var strCurrentHost in oLocalOptions.cloEnigmaHost.Where(strCurrentHost => !string.IsNullOrWhiteSpace(strCurrentHost)))
                DataInstallations.InstallEnigmaData(strCurrentHost, oLocalOptions.cloEnigmaUser, oLocalOptions.cloEnigmaPassword!, ltCurrentFiles, oLocalOptions.cloDryRun);
        }

        /// <summary>
        /// Successful export
        /// </summary>
        /// <param name="femExportMode"></param>
        /// <param name="strPath">Exported file - for EnigmaBouquetsFiles the directory of the bouquet files</param>
        /// <param name="dtStartUtc">Start of the export</param>
        private record ExportResult(FlagExportMode femExportMode, string strPath, DateTime dtStartUtc)
        {
            /// <summary>
            /// File exists and has been written by this export
            /// </summary>
            public bool IsWrittenFile(string strLocalFile) => File.Exists(strLocalFile) && File.GetLastWriteTimeUtc(strLocalFile) >= dtStartUtc;
        }

        /// <summary>
        /// Program.RunExport()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="femLocalExportMode"></param>
        /// <param name="strLocalExportFilename"></param>
        /// <param name="fefLocalExportFilenameFormat"></param>
        /// <param name="strLocalPanasonicTemplate"></param>
        private static void RunExport(List<SourceDataItem> lsdiLocalSourceData, FlagExportMode femLocalExportMode, string? strLocalExportFilename, FlagExportEnigmaFormat? fefLocalExportFilenameFormat, string? strLocalPanasonicTemplate)
        {
            switch (femLocalExportMode)
            {
                case FlagExportMode.ExcelChannelList:
                {
                    ExportSourceDataToExcelExportFile(lsdiLocalSourceData, strLocalExportFilename);
                    
                    break;
                }
                case FlagExportMode.DVBViewerTransponderList:
                {
                    ExportSourceDataToDVBViewerTransponderFile(lsdiLocalSourceData, strLocalExportFilename);

                    break;
                }
                case FlagExportMode.DVBViewerChannelDatabase:
                {
                    ExportSourceDataToDVBViewerChannelDatabase(lsdiLocalSourceData, strLocalExportFilename);
                    
                    break;
                }
                case FlagExportMode.EnigmaTransponderList:
                case FlagExportMode.EnigmaChannelDatabase:
                {
                    ExportSourceDataToEnigmaSelector(lsdiLocalSourceData, femLocalExportMode, strLocalExportFilename, fefLocalExportFilenameFormat);
                    
                    break;
                }
                case FlagExportMode.EnigmaBouquetsFiles:
                {
                    ExportSourceDataToEnigmaSelector(lsdiLocalSourceData, femLocalExportMode, string.Empty, fefLocalExportFilenameFormat);

                    break;
                }
                case FlagExportMode.PanasonicChannelList:
                {
                    ExportSourceDataToPanasonicChannelList(lsdiLocalSourceData, strLocalExportFilename, strLocalPanasonicTemplate);

                    break;
                }
            }
        }

        /// <summary>
        /// Program.ExecuteFileOperation() - reports file errors (e.g. file in use, missing permissions) instead of terminating
        /// </summary>
        /// <param name="eLocalOperation">Operation (export mode or target) for the error message</param>
        /// <param name="aLocalOperation"></param>
        /// <returns>bool - false if the operation failed (error is logged)</returns>
        private static bool ExecuteFileOperation(Enum eLocalOperation, Action aLocalOperation)
        {
            try
            {
                aLocalOperation();
                return true;
            }
            catch (Exception excCurrentException) when (excCurrentException is IOException or UnauthorizedAccessException or InvalidDataException or SettingsException or SqliteException)
            {
                LogError($"-» Error on {eLocalOperation.GetType().Name}.{eLocalOperation}: {excCurrentException.Message}");
                return false;
            }
        }
    }
}