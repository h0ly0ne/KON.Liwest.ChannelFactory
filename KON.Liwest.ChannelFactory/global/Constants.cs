namespace KON.Liwest.ChannelFactory
{
    public static class Constants
    {
        public const string FIREWALL_RULE_APP_NAME = "KON.Liwest.ChannelFactory";

        public const float EXCEL_PIXEL_TO_WIDTH_CONSTANT = (float)(8.43 / 64);
        public const float EXCEL_CHARACTER_TO_WIDTH_CONSTANT = 8 * EXCEL_PIXEL_TO_WIDTH_CONSTANT;

        public const string csCategory = "LIWEST";
        public const int ciTunerType = 0;                                   // 0 = dvb-c
        public const int ciModulationSystem = 1;                            // 1 = dvb-c
        public const int ciOrbitalPosition = 4000;                          // 4000 = default position for dvb-c
        public const int ciModulation = 5;                                  // 5 = qam256
        public const int ciSymbolRate = 6900;                               // 6900 = default symbolrate for dvb-c
        public const int ciFEC = 9;                                         // 9 = default fec for dvb-c
        public const int ciSpectralInversion = 1;                           // 1 = auto spectral inversion for dvb-c
        public const bool ciSatFeed = true;                                 // true = default satfeed for dvb-c
        public const string ciCountryCode = "AUT";                          // "AUT" = default country code for liwest
        public const string ciDefaultDVBCNamespace = "ffff0000";            // default namespace for dvb-c transponders

        public const string csDefaultDVBViewerChannelDatabaseID = "B2C2";   // default identifier for dvbviewer channel database
        public const int csDefaultDVBViewerChannelDatabaseVersionHi = 1;    // dvbviewer channel database major version
        public const int csDefaultDVBViewerChannelDatabaseVersionLo = 10;   // dvbviewer channel database minor version
    }
}