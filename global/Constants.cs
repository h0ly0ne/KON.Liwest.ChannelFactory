namespace KON.Liwest.ChannelFactory
{
    public static class Constants
    {
        public const string FIREWALL_RULE_APP_NAME = "KON.Liwest.ChannelFactory";

        public const float EXCEL_PIXEL_TO_WIDTH_CONSTANT = (float)(8.43 / 64);
        public const float EXCEL_CHARACTER_TO_WIDTH_CONSTANT = 8 * EXCEL_PIXEL_TO_WIDTH_CONSTANT;

        public const string csCategory = "LIWEST";
        public const int ciCableTunerType = 0;
        public const int ciCableOrbitalPosition = 4000;
        public const int ciCablePolarity = 5;
        public const int ciCableSymbolrate = 6900;
        public const int ciCableSatModulation = 0;
        public const int ciCableSubStreamID = -1;
        public const string csDVBViewerTransponderFile = "DVB-C_at_LIWEST.ini";
        public const string csExcelChannelSortFile = "ChannelListSort.xlsx";
        public const string csExcelChannelSortTemplate = "ChannelListSortTemplate.xlsx";

    }
}