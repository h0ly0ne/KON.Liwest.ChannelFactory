using System.Buffers.Binary;

using static KON.Liwest.ChannelFactory.Constants;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Panasonic channel list (hotel_setup export of the tv): encryption of svl.db and values of the sqlite tables SVL (services) / TSL (transponders)
    /// </summary>
    public static class Panasonic
    {
        /// <summary>
        /// Panasonic.ResolveTemplateDirectory() - template as directory (hotel_setup of the tv) or as model name (resources\Panasonic\&lt;model&gt;\hotel_setup next to the executable)
        /// </summary>
        /// <param name="strLocalTemplate"></param>
        /// <returns>Directory containing svl.db or null if not found</returns>
        public static string? ResolveTemplateDirectory(string? strLocalTemplate)
        {
            if (string.IsNullOrWhiteSpace(strLocalTemplate))
                return null;

            string[] saCurrentCandidates =
            [
                strLocalTemplate,
                Path.Combine(strLocalTemplate, csPanasonicHotelSetupDirectory),
                Path.Combine(AppContext.BaseDirectory, csPanasonicTemplateDirectory, strLocalTemplate, csPanasonicHotelSetupDirectory)
            ];

            return saCurrentCandidates.Select(Path.GetFullPath).FirstOrDefault(strCurrentCandidate => File.Exists(Path.Combine(strCurrentCandidate, csPanasonicChannelDatabaseFile)));
        }

        /// <summary>
        /// Panasonic.GetTemplateNames() - models of the templates next to the executable
        /// </summary>
        /// <returns>Model names</returns>
        public static string[] GetTemplateNames()
        {
            var strCurrentTemplateRoot = Path.Combine(AppContext.BaseDirectory, csPanasonicTemplateDirectory);

            return Directory.Exists(strCurrentTemplateRoot)
                ? Directory.GetDirectories(strCurrentTemplateRoot).Where(strCurrentDirectory => File.Exists(Path.Combine(strCurrentDirectory, csPanasonicHotelSetupDirectory, csPanasonicChannelDatabaseFile))).Select(Path.GetFileName).OfType<string>().Order(StringComparer.OrdinalIgnoreCase).ToArray()
                : [];
        }

        /// <summary>
        /// Panasonic.DecryptChannelDatabase() - checks the checksum and returns the sqlite database (without checksum)
        /// </summary>
        /// <param name="baLocalEncryptedData"></param>
        /// <returns>Decrypted sqlite database</returns>
        /// <exception cref="InvalidDataException">No encrypted Panasonic channel database or wrong checksum</exception>
        public static byte[] DecryptChannelDatabase(byte[] baLocalEncryptedData)
        {
            if (baLocalEncryptedData.Length < 8 || BinaryPrimitives.ReadUInt32LittleEndian(baLocalEncryptedData) != ciPanasonicChannelDatabaseMagic)
                throw new InvalidDataException("no encrypted Panasonic channel database (svl.db)");

            if (CalculateChecksum(baLocalEncryptedData.AsSpan(0, baLocalEncryptedData.Length - 4)) != BinaryPrimitives.ReadUInt32BigEndian(baLocalEncryptedData.AsSpan(baLocalEncryptedData.Length - 4)))
                throw new InvalidDataException("checksum of the Panasonic channel database (svl.db) is invalid");

            var baCurrentData = baLocalEncryptedData[..^4];
            ApplyCipher(baCurrentData, false);

            return baCurrentData;
        }

        /// <summary>
        /// Panasonic.EncryptChannelDatabase() - encrypts the sqlite database and appends the checksum
        /// </summary>
        /// <param name="baLocalPlainData"></param>
        /// <returns>Encrypted channel database</returns>
        public static byte[] EncryptChannelDatabase(byte[] baLocalPlainData)
        {
            var baCurrentData = new byte[baLocalPlainData.Length + 4];
            baLocalPlainData.CopyTo(baCurrentData, 0);

            ApplyCipher(baCurrentData.AsSpan(0, baLocalPlainData.Length), true);
            BinaryPrimitives.WriteUInt32BigEndian(baCurrentData.AsSpan(baLocalPlainData.Length), CalculateChecksum(baCurrentData.AsSpan(0, baLocalPlainData.Length)));

            return baCurrentData;
        }

        /// <summary>
        /// Panasonic.ApplyCipher() - xor cipher, the key depends on the previous encrypted bytes of the block (same as ChanSort)
        /// </summary>
        /// <param name="sbLocalData"></param>
        /// <param name="bLocalEncrypt"></param>
        private static void ApplyCipher(Span<byte> sbLocalData, bool bLocalEncrypt)
        {
            var iCurrentCipher = ciPanasonicCipherStart;
            var iCurrentStep = 0;

            for (var iCurrentPosition = 0; iCurrentPosition < sbLocalData.Length; iCurrentPosition++)
            {
                var byCurrentInput = sbLocalData[iCurrentPosition];
                var byCurrentOutput = (byte)(byCurrentInput ^ (iCurrentCipher >> 8));
                sbLocalData[iCurrentPosition] = byCurrentOutput;

                if (++iCurrentStep < ciPanasonicCipherBlockLength)
                {
                    iCurrentCipher += (bLocalEncrypt ? byCurrentOutput : byCurrentInput) + ciPanasonicCipherIncrement;
                }
                else
                {
                    iCurrentCipher = ciPanasonicCipherStart;
                    iCurrentStep = 0;
                }
            }
        }

        /// <summary>
        /// Panasonic.CalculateChecksum() - crc32 (mpeg-2: not reflected, start 0xFFFFFFFF, no final xor)
        /// </summary>
        /// <param name="sbLocalData"></param>
        /// <returns>uint</returns>
        private static uint CalculateChecksum(ReadOnlySpan<byte> sbLocalData)
        {
            var uiCurrentChecksum = 0xFFFFFFFF;

            foreach (var byCurrentByte in sbLocalData)
            {
                uiCurrentChecksum ^= (uint)byCurrentByte << 24;

                for (var iCurrentBit = 0; iCurrentBit < 8; iCurrentBit++)
                    uiCurrentChecksum = (uiCurrentChecksum & 0x80000000) != 0 ? (uiCurrentChecksum << 1) ^ ciPanasonicChecksumPolynomial : uiCurrentChecksum << 1;
            }

            return uiCurrentChecksum;
        }

        /// <summary>
        /// Panasonic.GetPhysicalChannel() - physical_ch of a frequency, links SVL and TSL
        /// </summary>
        /// <param name="iLocalFrequency">Frequency in kHz</param>
        /// <returns>int</returns>
        public static int GetPhysicalChannel(int iLocalFrequency)
        {
            return ciPanasonicPhysicalChannelBase + iLocalFrequency / 100;
        }

        /// <summary>
        /// Panasonic.GetCableChannelName() - cable channel of a frequency (hyperband S21-S41, uhf 21-69), empty for other frequencies
        /// </summary>
        /// <param name="iLocalFrequency">Frequency in kHz</param>
        /// <returns>string</returns>
        public static string GetCableChannelName(int iLocalFrequency)
        {
            if (iLocalFrequency >= ciCableChannelUHFFirstFrequency && (iLocalFrequency - ciCableChannelUHFFirstFrequency) % ciCableChannelStep == 0)
                return (21 + (iLocalFrequency - ciCableChannelUHFFirstFrequency) / ciCableChannelStep).ToString();

            if (iLocalFrequency >= ciCableChannelHyperbandFirstFrequency && iLocalFrequency < ciCableChannelUHFFirstFrequency && (iLocalFrequency - ciCableChannelHyperbandFirstFrequency) % ciCableChannelStep == 0)
                return $"S{21 + (iLocalFrequency - ciCableChannelHyperbandFirstFrequency) / ciCableChannelStep}";

            return string.Empty;
        }

        /// <summary>
        /// Panasonic.CreateDeliveryData() - delivery of TSL: frequency in 10 Hz (uint32 le), symbolrate in kSym/s (uint16 le), 2 bytes reserved, modulation, rest zero
        /// </summary>
        /// <param name="iLocalFrequency">Frequency in kHz</param>
        /// <returns>byte[]</returns>
        public static byte[] CreateDeliveryData(int iLocalFrequency)
        {
            var baCurrentDeliveryData = new byte[ciPanasonicDeliveryDataLength];

            BinaryPrimitives.WriteUInt32LittleEndian(baCurrentDeliveryData, (uint)iLocalFrequency * 100);
            BinaryPrimitives.WriteUInt16LittleEndian(baCurrentDeliveryData.AsSpan(4), ciSymbolRate);
            baCurrentDeliveryData[8] = ciPanasonicModulationQAM256;

            return baCurrentDeliveryData;
        }
    }
}
