//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Runtime.InteropServices;
//using System.Text;
//using System.Threading.Channels;
//using System.Threading.Tasks;

//using static System.Net.Mime.MediaTypeNames;

//namespace KON.Liwest.ChannelFactory.global
//{
//    public class DelphiPorted
//    {
//        [StructLayout(LayoutKind.Sequential, Pack = 1)]
//        public struct TTuner
//        {
//            public byte TunerType; // 0 = cable, 1 = satellite, 2 = terrestrial, 3 = atsc
//            public byte ChannelGroup; // 0 = Group A, 1 = Group B, 2 = Group C
//            public byte SatModulationSystem; // 0 = DVB-S, 1 = DVB-S2
//            public byte Flags;
//            // Bit 0: 1 = encrypted channel
//            // Bit 1: 1 = don't display EPG on What's Now Tab & in Timeline
//            // Bit 2: 1 = channel broadcasts RDS data
//            // Bit 3: 1 = channels is a video service (even if the Video PID is temporarily = 0)
//            // Bit 4: 1 = channel is an audio service (even if the Audio PID is temporarily = 0)
//            // Bit 5..7: reserved
//            public uint Frequency; // MHz for DVB-S, KHz for DVB-T/C and ATSC
//            public uint Smbolrate; // DVB-S/C only
//            public ushort LNB_LOF; // DVB-S only, local oscillator frequency of the LNB
//            public ushort PMT_PID;
//            public ushort Reserved1;
//            public byte SatModulation;
//            // Bit 0..1: modulation. 0 = Auto, 1 = QPSK, 2 = 8PSK, 3 = 16QAM
//            // Bit 2: modulation system. 0 = DVB-S, 1 = DVB-S2
//            // Bit 3..4: roll-off. 0 = 0.35, 1 = 0.25, 2 = 0.20, 3 = reserved
//            // Bit 5..6: spectral inversion. 0 = undefined, 1 = auto, 2 = normal, 3 = inverted
//            // Bit 7: pilot symbols. 0 = off, 1 = on
//            // Note: Bit 3..7 only apply to DVB-S2 and Hauppauge HVR 4000 / Nova S2 Plus
//            public byte AVFormat;
//            // Low Nibble (Bit 0..3): audio format
//            //   0 = MPEG
//            //   1 = AC3
//            //   2..15 reserved
//            // High Nibble (Bit 4..7): video format
//            //   0 = MPEG2
//            //   1 = H.264
//            //   2..15 reserved
//            public byte FEC;
//            // 0 = Auto
//            // 1 = 1/2
//            // 2 = 2/3
//            // 3 = 3/4
//            // 4 = 5/6
//            // 5 = 7/8
//            // 6 = 8/9
//            // 7 = 3/5
//            // 8 = 4/5
//            // 9 = 9/10
//            public byte Reserved2; // must be 0
//            public ushort Reserved3;
//            public byte Polarity;
//            // DVB-S polarity
//            //   0 = horizontal
//            //   1 = vertical
//            //   2 = circular left
//            //   3 = circular right
//            // or DVB-C modulation
//            //   0 = Auto
//            //   1 = 16QAM
//            //   2 = 32QAM
//            //   3 = 64QAM
//            //   4 = 128QAM
//            //   5 = 256 QAM
//            // or DVB-T bandwidth
//            //   0 = 6 MHz
//            //   1 = 7 MHz
//            //   2 = 8 MHz
//            public byte Reserved4; // must be 0
//            public ushort Reserved5;
//            public byte Tone; // 0 = off, 1 = 22 khz
//            public byte Reserved6; // must be 0
//            public ushort DiSEqCExt;
//            // DiSEqC Extension: OrbitPos, or other value
//            // -> Positoner, GotoAngular, Command String (set to 0 if not required)
//            public byte DiSEqC;
// 0 = None
// 1 = Pos A (mostly translated to PosA/OptA)
// 2 = Pos B (mostly translated to PosB/OptA)
// 3 = PosA/OptA
// 4 = PosB/OptA
// 5 = PosA/OptB
// 6 = PosB/OptB
//            public byte Reserved7; // must be 0
//            public ushort Reserved8;
//            public ushort Audio_PID;
//            public ushort Reserved9;
//            public ushort Video_PID;
//            public ushort TransportStream_ID;
//            public ushort Teletext_PID;
//            public ushort OriginalNetwork_ID;
//            public ushort Service_ID;
//            public ushort PCR_PID;
//        }
//        public struct ShortString25
//        {
//            public byte Length;
//            public char[] Chars;

//            public ShortString25(string value)
//            {
//                Length = (byte)Math.Min(value.Length, 25);
//                Chars = new char[25];
//                value.CopyTo(0, Chars, 0, Length);
//                for (int i = Length; i < 25; i++)
//                    Chars[i] = '\0';
//            }

//            public override string ToString()
//            {
//                return new string(Chars, 0, Length);
//            }
//        }
//        public struct TChannel
//        {
//            public TTuner TunerData;
//            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 25)]
//            public string Root;
//            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 25)]
//            public string ChannelName;
//            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 25)]
//            public string Category;
//            public byte Encrypted;
//            public byte Reserved10;
//        }
//        public struct THeader
//        {
//            public byte IDLength; // = 4
//            [MarshalAs(UnmanagedType.FixedBuffer, SizeConst = 4)]
//            public char[] ID; //'B2C2' as ASCII chars
//            public byte VersionHi; //currently 1
//            public byte VersionLo; //currently 8 -> 1.8
//        }
//        [StructLayout(LayoutKind.Sequential, Pack = 1)]
//        public struct TEpgEvent
//        {
//            public ushort EventID; //see ETSI EN 300 468, 5.2.4 Event Information Table
//            public ushort Reserved; //Always 0 in DVBViewer GE. //May be used for other purpose in DVBViewer Pro
//            public double StartTime; //TDateTime
//            public double Duration; //TDateTime
//            [MarshalAs(UnmanagedType.LPStr)]
//            public string Event;
//            [MarshalAs(UnmanagedType.LPStr)]
//            public string Title;
//            [MarshalAs(UnmanagedType.LPStr)]
//            public string Description;
//            public byte CharSet;
//            public byte Content; //See ETSI EN 300 468, 6.2.9, Content Descriptor
//        }
//        public struct Ch_st
//        {
//            public ushort ID;
//            [MarshalAs(UnmanagedType.LPStr)]
//            public string St;
//        }

//        public string Bytes25ToString(short[] shortString25)
//        {
//            //string resultString;
//            //ushort stringLength;

//            //// ShortString25 = array[0..25] of byte; //fixed size, always 26 bytes
//            //// Byte 0: Length;
//            //// Byte 1..Length: Char;
//            //// Byte Length+1..25: Unused
//            //stringLength = (ushort)shortString25[0];
//            //resultString = new string(' ', stringLength);
//            //System.Runtime.InteropServices.Marshal.Copy(new IntPtr(shortString25), resultString.ToCharArray(), 1, stringLength);
//            //return resultString;

//            return string.Empty;
//        }
//        public string GetChannel(ushort channelID)
//        {
//            //string resultString = string.Empty;
//            //for (int index = 0; index < Channel_String.Length; index++)
//            //{
//            //    if (Channel_String[index].ID == channelID)
//            //    {
//            //        resultString = Channel_String[index].st;
//            //        break;
//            //    }
//            //}
//            //return resultString;

//            return string.Empty;
//        }
//        private void Button1Click(object sender, EventArgs e)
//        {
//            FileStream fileStream;
//            THeader header;
//            TChannel channel;
//            int index;
//            bool alreadyFound;

//            if (!File.Exists("channels.dat"))
//            {
//                MessageBox.Show("Copy channels.dat into this sample folder");
//                return;
//            }

//            for (index = 0; index < StringGrid1.RowCount; index++)
//            {
//                StringGrid1.Rows[index].Clear();
//            }
//            StringGrid1.ColCount = 0;

//            Application.DoEvents();

//            fileStream = new FileStream("channels.dat", FileMode.Open);
//            try
//            {
//                fileStream.Read(header, 0, Marshal.SizeOf(typeof(Header)));
//                Memo1.Lines.Add("Header information of Channels.dat");
//                Memo1.Lines.Add("-----------------------------------------");
//                Memo1.Lines.Add(header.IDLength.ToString());
//                Memo1.Lines.Add(header.ID);
//                Memo1.Lines.Add(header.VersionHi.ToString());
//                Memo1.Lines.Add(header.VersionLo.ToString());
//                Memo1.Lines.Add("-----------------------------------------");

//                while (fileStream.Read(channel, 0, Marshal.SizeOf(typeof(Channel))) > 0)
//                {
//                    Memo1.Lines.Add("Root         : " + Bytes25ToString(channel.Root));
//                    Memo1.Lines.Add("Channel Name : " + Bytes25ToString(channel.ChannelName));
//                    Memo1.Lines.Add("Service ID   : " + channel.TunerData.Service_ID.ToString());
//                    Memo1.Lines.Add("Channel Cat  : " + Bytes25ToString(channel.Category));
//                    Memo1.Lines.Add("-----------------------------------------");

//                    Array.Resize(ref Channel_String, Channel_String.Length + 1);

//                    alreadyFound = false;
//                    for (index = 0; index < StringGrid1.ColCount; index++)
//                    {
//                        if (StringGrid1.Cells[index, 0] == channel.TunerData.Service_ID.ToString())
//                        {
//                            alreadyFound = true;
//                            break;
//                        }
//                    }

//                    if (!alreadyFound)
//                    {
//                        Channel_String[Channel_String.Length - 1].ID = channel.TunerData.Service_ID;
//                        Channel_String[Channel_String.Length - 1].St = Bytes25ToString(channel.ChannelName);
//                        StringGrid1.Cells[StringGrid1.ColCount - 1, 0] = channel.TunerData.Service_ID.ToString();
//                        StringGrid1.Cells[StringGrid1.ColCount - 1, 1] = Bytes25ToString(channel.ChannelName);
//                        StringGrid1.ColCount++;
//                    }
//                }
//            }
//            finally
//            {
//                fileStream.Close();
//                StringGrid1.ColCount--;
//                Button1.Enabled = false;
//                Button2.Enabled = true;
//            }
//        }
//        //        void Button2Click(object sender, EventArgs e)
//        //        {
//        //            FileStream fileStream;
//        //            long filePosition;
//        //            long fileSize;
//        //            ushort formatVersion;
//        //            double timeZone;
//        //            int i, j;
//        //            uint numberOfEvents;
//        //            ushort eventID; // see ETSI EN 300 468, 5.2.4 Event Information Table
//        //            ushort reserved; // Always 0 in DVBViewer GE. // May be used for other purpose in DVBViewer Pro
//        //            double startTime;
//        //            double duration;
//        //            string eventDescription;
//        //            string title;
//        //            string description;
//        //            string temp;
//        //            byte charSet;
//        //            byte content; // See ETSI EN 300 468, 6.2.9, Content Descriptor
//        //            uint stringSize;
//        //            int epgColumn;

//        //            // new version 0x108
//        //            ushort serviceID;
//        //            ushort transportStreamID;
//        //            ushort originalNetworkID;
//        //            byte tunerType; // 0 = undefined, 1 = cable, 2 = satellite, 3 = terrestrial, 4 = ATSC
//        //            byte reserved1; // = 0
//        //            uint dataBlockLength;
//        //            uint pdc;
//        //            byte data;
//        //            byte minimumAge;

//        //            // new version 0x110
//        //            uint utf8Size;
//        //            string utf8Text; // SizeofUTF8Size

//        //            string ReadByteArray(FileStream fs)
//        //            {
//        //                fs.Read(BitConverter.GetBytes(stringSize), 0, sizeof(uint));
//        //                byte[] byteArray = new byte[stringSize];
//        //                fs.Read(byteArray, 0, (int)stringSize);
//        //                return System.Text.Encoding.Default.GetString(byteArray);
//        //            }

//        //            if (!File.Exists("epg.dat"))
//        //            {
//        //                MessageBox.Show("Copy epg.dat into this sample folder");
//        //                return;
//        //            }

//        //            fileStream = new FileStream("epg.dat", FileMode.Open);
//        //            try
//        //            {
//        //                fileSize = fileStream.Length;
//        //                fileStream.Read(BitConverter.GetBytes(formatVersion), 0, sizeof(ushort));
//        //                fileStream.Read(BitConverter.GetBytes(timeZone), 0, sizeof(double));
//        //                switch (formatVersion)
//        //                {
//        //                    case 0x0108:
//        //                        memo2.Lines.Add("FORMAT VERSION: 1.08");
//        //                        break;
//        //                    case 0x0109:
//        //                        memo2.Lines.Add("FORMAT VERSION: 1.09");
//        //                        break;
//        //                    case 0x010A:
//        //                        memo2.Lines.Add("FORMAT VERSION: 1.0A");
//        //                        break;
//        //                    default:
//        //                        MessageBox.Show("Unknown EPG.DAT Format ! -- STOP --");
//        //                        return;
//        //                }

//        //                if (showLog.Checked)
//        //                {
//        //                    memo2.Lines.Add("Header information of EPG.dat");
//        //                    memo2.Lines.Add("-----------------------------------------");
//        //                    memo2.Lines.Add(formatVersion.ToString());
//        //                    memo2.Lines.Add(timeZone.ToString());
//        //                    memo2.Lines.Add("-----------------------------------------");
//        //                }
//        //                progressBar1.Maximum = (int)fileSize;
//        //                progressBar1.Value = 0;

//        //                while (fileStream.Position < fileSize)
//        //                {
//        //                    progressBar1.Value = (int)fileStream.Position;

//        //                    fileStream.Read(BitConverter.GetBytes(serviceID), 0, sizeof(ushort));
//        //                    fileStream.Read(BitConverter.GetBytes(transportStreamID), 0, sizeof(ushort));

//        //                    if (showLog.Checked)
//        //                    {
//        //                        memo2.Lines.Add("ChannelID     : " + serviceID.ToString());
//        //                        memo2.Lines.Add("ChannelString : " + GetChannel(serviceID));
//        //                    }

//        //                    fileStream.Read(BitConverter.GetBytes(originalNetworkID), 0, sizeof(ushort));
//        //                    fileStream.Read(BitConverter.GetBytes(tunerType), 0, sizeof(byte));
//        //                    fileStream.Read(BitConverter.GetBytes(reserved1), 0, sizeof(byte));

//        //                    for (i = 0; i < stringGrid1.ColumnCount; i++)
//        //                        if (stringGrid1.Cells[i, 0] == serviceID.ToString())
//        //                            epgColumn = i;

//        //                    fileStream.Read(BitConverter.GetBytes(numberOfEvents), 0, sizeof(uint));
//        //                    if (showLog.Checked)
//        //            {
//        //                memo2.Lines.Add("Number of Events : " + numberOfEvents.ToString());
//        //                memo2.Lines.Add("-----------------------------------------");
//        //            }

//        //            if (stringGrid1.RowCount + 2 < numberOfEvents)
//        //                stringGrid1.RowCount = (int)(numberOfEvents + 2);

//        //            progressBar2.Maximum = (int)numberOfEvents;
//        //            progressBar2.Value = 0;
//        //            for (i = 0; i < numberOfEvents; i++)
//        //            {
//        //                progressBar2.PerformStep();
//        //                fileStream.Read(BitConverter.GetBytes(eventID), 0, sizeof(ushort));
//        //                fileStream.Read(BitConverter.GetBytes(reserved), 0, sizeof(ushort));
//        //                fileStream.Read(BitConverter.GetBytes(startTime), 0, sizeof(double));
//        //                fileStream.Read(BitConverter.GetBytes(duration), 0, sizeof(double));
//        //                if (showLog.Checked)
//        //                    memo2.Lines.Add(startTime.ToString("g") + " -(" + duration.ToString() + ")");

//        //                eventDescription = ReadByteArray(fileStream);
//        //                if (showLog.Checked)
//        //                    memo2.Lines.Add(eventDescription);
//        //                if (eventDescription[0] == '<')
//        //                    eventDescription = eventDescription.Remove(0, 5);

//        //                title = ReadByteArray(fileStream);
//        //                if (showLog.Checked)
//        //                    memo2.Lines.Add(title);
//        //                if (!string.IsNullOrEmpty(title) && title[0] == '<')
//        //                    title = title.Remove(0, 5);

//        //                if (!string.IsNullOrEmpty(title) && !string.IsNullOrEmpty(eventDescription))
//        //                {
//        //                    temp = eventDescription;
//        //                }
//        //                else
//        //                {
//        //                    temp = !string.IsNullOrEmpty(title) ? title : eventDescription;
//        //                }

//        //                stringGrid1.Cells[epgColumn, i + 2] = temp;

//        //                Application.DoEvents();
//        //                description = ReadByteArray(fileStream);
//        //                if (showLog.Checked)
//        //                    memo2.Lines.Add(description);

//        //                if (showLog.Checked)
//        //                    memo2.Lines.Add("-----------------------------------------");

//        //                // the following data block is reserved for various data. Skip it.
//        //                switch (formatVersion)
//        //                {
//        //                    case 0x0108:
//        //                        fileStream.Read(BitConverter.GetBytes(charSet), 0, sizeof(byte));
//        //                        fileStream.Read(BitConverter.GetBytes(content), 0, sizeof(byte));
//        //                        fileStream.Read(BitConverter.GetBytes(dataBlockLength), 0, sizeof(uint));
//        //                        for (j = 0; j < dataBlockLength; j++)
//        //                            fileStream.Read(BitConverter.GetBytes(data), 0, sizeof(byte));
//        //                        break;
//        //                    case 0x0109:
//        //                        fileStream.Read(BitConverter.GetBytes(content), 0, sizeof(byte));
//        //                        fileStream.Read(BitConverter.GetBytes(pdc), 0, sizeof(uint));
//        //                        fileStream.Read(BitConverter.GetBytes(minimumAge), 0, sizeof(byte));
//        //                        break;
//        //                    case 0x010A:
//        //                        fileStream.Read(BitConverter.GetBytes(content), 0, sizeof(byte));
//        //                        fileStream.Read(BitConverter.GetBytes(pdc), 0, sizeof(uint));
//        //                        fileStream.Read(BitConverter.GetBytes(minimumAge), 0, sizeof(byte));
//        //                        utf8Text = ReadByteArray(fileStream);
//        //                        break;
//        //                }
//        //            } // for i = 0 to numberOfEvents-1 do
//        //        } // while fileStream.Position < fileSize do
//        //    }
//        //    finally
//        //    {
//        //        fileStream.Close();
//        //        progressBar1.Value = 0;
//        //        Button2.Enabled = false;
//        //        Button1.Enabled = true;
//        //    }
//        //}
//    }
//}