using System.Runtime.InteropServices;

namespace KON.Liwest.ChannelFactory
{
    public class DVBViewer
    {
        public enum AVFormat
        {
            AUDIO_MPEG = 0,
            AUDIO_AC3 = 1,
            VIDEO_MPEG2 = 0,
            VIDEO_H264 = 1,
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class ChannelDatabaseHeader
        {
            public byte IDLength;                                   // default: 4
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public char[]? ID;                                      // default: 'B2C2'
            public byte VersionHi;                                  // default: 1
            public byte VersionLo;                                  // default: 10
        }
        
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class ChannelDatabaseChannel
        {
            public ChannelDatabaseTuner? TunerData;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
            public char[]? Root;                                    // mostly the satellite position resp. DVB network name, can be user defined
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
            public char[]? ChannelName;                             // user defined
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
            public char[]? Category;                                // user defined
            public byte Encrypted;                                  // deprecated! Only set for compatibility. Same as TTuner.Flags.
            public byte Reserved;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public class ChannelDatabaseTuner
        {
            public byte TunerType;                                  // 0 = Cable, 1 = Satellite, 2 = Terrestrial, 3 = ATSC
            public byte ChannelGroup;                               // 0 = Group A, 1 = Group B, 2 = Group C
            public byte SatModulationSystem;                        // 0 = DVB-S, 1 = DVB-S2
            public byte ChannelFlags;
            public uint Frequency;                                  // (Required) - MHz for DVB-S, KHz for DVB-T/C and ATSC
            public uint Symbolrate;                                 // (Required) - DVB-S/C only
            public ushort LNB_LOF;                                  // DVB-S only, local oscillator frequency of the LNB
            public ushort PMT_PID;
            public ushort Reserved1;
            public byte SatModulation;                              // Bit 0..1: modulation. 0 = Auto, 1 = QPSK, 2 = 8PSK, 3 = 16QAM - Bit 2: modulation system. 0 = DVB-S, 1 = DVB-S2 - Bit 3..4: roll-off. 0 = 0.35, 1 = 0.25, 2 = 0.20, 3 = reserved - Bit 5..6: spectral inversion. 0 = undefined, 1 = auto, 2 = normal, 3 = inverted - Bit 7: pilot symbols. 0 = off, 1 = on - Note: Bit 3..7 only apply to DVB-S2 and Hauppauge HVR 4000 / Nova S2 Plus
            public byte ChannelAVFormat;
            public byte FEC;                                        // 0 = Auto, 1 = 1/2, 2 = 2/3, 3 = 3/4, 4 = 5/6, 5 = 7/8, 6 = 8/9, 7 = 3/5, 8 = 4/5, 9 = 9/10
            public byte Reserved2;                                  // must be 0
            public ushort ChannelNumber;
            public byte Polarity;                                   // (Required) - DVB-S polarity - 0 = horizontal, 1 = vertical, 2 = circular left, 3 = circular right - or DVB-C modulation, 0 = Auto, 1 = 16QAM, 2 = 32QAM, 3 = 64QAM, 4 = 128QAM, 5 = 256 QAM - or DVB-T bandwidth - 0 = 6 MHz, 1 = 7 MHz, 2 = 8 MHz
            public byte Reserved4;                                  // must be 0
            public ushort OrbitalPosition;
            public byte Tone;                                       // 0 = off, 1 = 22 khz
            public byte Reserved6;                                  // must be 0
            public ushort DiSEqCExt;                                // DiSEqC Extension: OrbitPos, or other value -> Positioner, GotoAngular, Command String (set to 0 if not required)
            public byte DiSEqC;                                     // 0 = None, 1 = Pos A (mostly translated to PosA/OptA), 2 = Pos B (mostly translated to PosB/OptA), 3 = PosA/OptA, 4 = PosB/OptA, 5 = PosA/OptB, 6 = PosB/OptB
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
            public char[]? Language;                                // Replaced Reserved7 (byte) = 0 + Reserved8 (ushort)
            public ushort Audio_PID;
            public ushort Reserved9;
            public ushort Video_PID;
            public ushort TransportStream_ID;
            public ushort Teletext_PID;
            public ushort OriginalNetwork_ID;
            public ushort Service_ID;                               // (Required)
            public ushort PCR_PID;
        }

        public static byte[] Binarize(object obj)
        {
            var size = Marshal.SizeOf(obj);
            var bytes = new byte[size];
            var ptr = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(obj, ptr, false);
            Marshal.Copy(ptr, bytes, 0, size);
            Marshal.FreeHGlobal(ptr);

            return bytes;
        }
        //public static unsafe dynamic? ToObject(byte[] bytes)
        //{
        //    var rf = __makeref(bytes);
        //    **(int**)&rf += 8;
        //    return GCHandle.Alloc(bytes).Target;
        //}
        //public static byte HexStringToByte(string hex)
        //{
        //    return Enumerable.Range(0, hex.Length)
        //        .Where(x => x % 2 == 0)
        //        .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
        //        .ToArray()[0];
        //}

        public static byte NibblesToByte(int iLocalLoNibble, int iLocalHiNibble)
        {
            return (byte)((iLocalLoNibble & 0x0F) | ((iLocalHiNibble & 0x0F) << 4));
        }

        public static byte GenerateChannelFlags(bool bLocalIsEncrypted, bool bLocalAutomaticRefresh, bool bLocalBroadcastRDS, bool bLocalIsVideoService, bool bLocalIsAudioService, bool bIsFollowingPID)
        {
            byte byReturnFlag = 0;

            // Bit 0: 1 = encrypted channel
            if (bLocalIsEncrypted) 
                byReturnFlag |= 1;
            // Bit 1: 1 = don't refresh channel automatically
            if (bLocalAutomaticRefresh)
                byReturnFlag |= 2;
            // Bit 2: 1 = channel broadcasts RDS data
            if (bLocalBroadcastRDS)
                byReturnFlag |= 4;
            // Bit 3: 1 = channels is a video service (even if the Video PID is temporarily = 0)
            if (bLocalIsVideoService)
                byReturnFlag |= 8;
            // Bit 4: 1 = channel is an audio service (even if the Audio PID is temporarily = 0)
            if (bLocalIsAudioService)
                byReturnFlag |= 16;
            // Bit 5..6 : reserved

            // Bit 7: 0 = channel is primary pid - 1 = channel is following pid
            if (bIsFollowingPID)
                byReturnFlag |= 128;

            return byReturnFlag;
        }

        public static byte GenerateChannelAVFormat(AVFormat cavfLocalChannelAudioFormat, AVFormat cavfLocalChannelVideoFormat)
        {
            return NibblesToByte((int)cavfLocalChannelAudioFormat, (int)cavfLocalChannelVideoFormat);
        }
    }
}