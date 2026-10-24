using System.Runtime.InteropServices;

using static KON.Liwest.ChannelFactory.DVBViewer;
using static KON.Liwest.ChannelFactory.Global;

using Pastel;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Console dumps of existing files
    /// </summary>
    public static class DataDumps
    {
        /// <summary>
        /// DataDumps.DumpDVBViewerChannelDatabase()
        /// </summary>
        /// <param name="strLocalDVBViewerChannelDatabaseFilename"></param>
        public static void DumpDVBViewerChannelDatabase(string? strLocalDVBViewerChannelDatabaseFilename)
        {
            if (string.IsNullOrEmpty(strLocalDVBViewerChannelDatabaseFilename) || !File.Exists(strLocalDVBViewerChannelDatabaseFilename))
                return;

            LogTrace("DataDumps.DumpDVBViewerChannelDatabase()");
            LogInfo("-» Reading channel database file");

            using var brCurrentBinaryReader = new BinaryReader(File.Open(strLocalDVBViewerChannelDatabaseFilename, FileMode.Open, FileAccess.Read, FileShare.Read));

            var iCurrentHeaderSize = Marshal.SizeOf<ChannelDatabaseHeader>();
            var baCurrentHeaderBuffer = brCurrentBinaryReader.ReadBytes(iCurrentHeaderSize);
            if (baCurrentHeaderBuffer.Length < iCurrentHeaderSize)
            {
                LogError("-» Channel database file is too small to contain a valid header");
                return;
            }

            var cdhCurrentHeader = Debinarize<ChannelDatabaseHeader>(baCurrentHeaderBuffer);
            WriteDumpLine("Header information of " + strLocalDVBViewerChannelDatabaseFilename);
            WriteDumpLine("-----------------------------------------");
            WriteDumpLine("IDLength  : " + cdhCurrentHeader.IDLength);
            WriteDumpLine("ID        : " + cdhCurrentHeader.ID.FromBinary25CharArrayToString(false));
            WriteDumpLine("VersionHi : " + cdhCurrentHeader.VersionHi);
            WriteDumpLine("VersionLo : " + cdhCurrentHeader.VersionLo);
            WriteDumpLine("-----------------------------------------");

            var iCurrentChannelSize = Marshal.SizeOf<ChannelDatabaseChannel>();
            byte[] baCurrentChannelBuffer;
            while ((baCurrentChannelBuffer = brCurrentBinaryReader.ReadBytes(iCurrentChannelSize)).Length > 0)
            {
                if (baCurrentChannelBuffer.Length < iCurrentChannelSize)
                {
                    LogWarning($"-» Ignored incomplete channel record at the end of the file ({baCurrentChannelBuffer.Length} of {iCurrentChannelSize} bytes)");
                    break;
                }

                var cdcCurrentChannel = Debinarize<ChannelDatabaseChannel>(baCurrentChannelBuffer);
                WriteDumpLine("Root         : " + cdcCurrentChannel.Root.FromBinary25CharArrayToString(true));
                WriteDumpLine("Channel Name : " + cdcCurrentChannel.ChannelName.FromBinary25CharArrayToString(true));
                WriteDumpLine("Channel Cat  : " + cdcCurrentChannel.Category.FromBinary25CharArrayToString(true));
                WriteDumpLine("Service ID   : " + cdcCurrentChannel.TunerData?.Service_ID);
                WriteDumpLine("-----------------------------------------");
            }
        }

        private static void WriteDumpLine(string strLocalLine)
        {
            lCurrentLogger.Info(strLocalLine.Pastel(ConsoleColor.Yellow));
        }
    }
}
