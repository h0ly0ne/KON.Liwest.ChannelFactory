using CommandLine;
using CommandLine.Text;
using NLog;
using NLog.Layouts;
using Pastel;

using KON.OctoScan.NET;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;
using System.Text;

namespace KON.Liwest.ChannelFactory
{
    public class Program
    {
        public enum FlagMode
        {
            None                                        = 0,
            Automated                                   = 1,
            FetchTransponderDataFromLiwestChannelList   = 2,
            FetchTransponderDataFromLiwestSmartcard     = 3,
            FetchTransponderDataFromLiwestBlindscan     = 4,
            GenerateDVBViewerTransponderFile            = 5,
            GenerateExcelChannelSortFile                = 6
        }

        public enum FlagInputFileEncoding
        {
            UTF8                                = 0x0,
            ISO88591                            = 0x1
        }

        public class Options
        {
            [Option(longName: "mode", Required = true, HelpText = "Factory mode to use.")]
            public FlagMode cloMode { get; set; }

            //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

            [Option(longName: "smartcardid", Required = false, HelpText = "Smartcard Identifier.")]
            public string? cloSmartcardID { get; set; }

            [Option(longName: "serverip", Required = false, HelpText = "IP Address of SAT>IP server")]
            public string? cloServerIP { get; set; }

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

            [Option(longName: "verbose", Required = false, HelpText = "Set output to verbose level.")]
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

        private static void DisplayHelp<T>(ParserResult<T> result)
        {
            var htCurrentHelpText = HelpText.AutoBuild(result, htCurrentHelpText =>
            {
                htCurrentHelpText.AddNewLineBetweenHelpSections = false;
                htCurrentHelpText.AdditionalNewLineAfterOption = false;
                htCurrentHelpText.MaximumDisplayWidth = 100;
                htCurrentHelpText.Heading = "KON.Liwest.ChannelFactory";
                htCurrentHelpText.Copyright = "Copyright (C) 2024 Oswald Oliver";

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
                switch (oLocalOptions.cloMode)
                {
                    case FlagMode.Automated:
                    {
                        FetchTransponderDataFromLiwestChannelList();

                        if (!string.IsNullOrEmpty(oLocalOptions.cloSmartcardID))
                            FetchTransponderDataFromLiwestSmartcard(oLocalOptions.cloSmartcardID);
                        
                        if (!string.IsNullOrEmpty(oLocalOptions.cloServerIP))
                            FetchTransponderDataFromLiwestBlindscan(oLocalOptions.cloServerIP);

                        break;
                    }
                    case FlagMode.FetchTransponderDataFromLiwestChannelList:
                    {
                        FetchTransponderDataFromLiwestChannelList();
                        break;
                    }
                    case FlagMode.FetchTransponderDataFromLiwestSmartcard:
                    {
                        if (!string.IsNullOrEmpty(oLocalOptions.cloSmartcardID))
                            FetchTransponderDataFromLiwestSmartcard(oLocalOptions.cloSmartcardID);
                        break;
                    }
                    case FlagMode.FetchTransponderDataFromLiwestBlindscan:
                    {
                        if (!string.IsNullOrEmpty(oLocalOptions.cloServerIP))
                            FetchTransponderDataFromLiwestBlindscan(oLocalOptions.cloServerIP);
                        break;
                    }
                    case FlagMode.GenerateDVBViewerTransponderFile:
                    {
                        GenerateDVBViewerTransponderFile();
                        break;
                    }
                    case FlagMode.GenerateExcelChannelSortFile:
                    {
                        GenerateExcelChannelSortFile();
                        break;
                    }
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