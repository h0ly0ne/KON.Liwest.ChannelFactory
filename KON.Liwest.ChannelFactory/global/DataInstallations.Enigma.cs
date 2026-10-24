using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

using Renci.SshNet;
using Renci.SshNet.Common;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    public static partial class DataInstallations
    {
        [GeneratedRegex(@"^#SERVICE\s+(?:4097|5001|5002|8193):|^#SERVICE\s+\S*(?:https?|rtsp|rtmp)%3a", RegexOptions.IgnoreCase | RegexOptions.Multiline)]
        private static partial Regex rxEnigmaStreamingService();

        /// <summary>
        /// DataInstallations.InstallEnigmaData() - connects to the receiver via SSH / SFTP, backs up its settings and installs the files
        /// </summary>
        /// <param name="strLocalHost">host or host:port</param>
        /// <param name="strLocalUser"></param>
        /// <param name="strLocalPassword"></param>
        /// <param name="ltLocalFiles">Local file and absolute target path on the receiver</param>
        /// <param name="bLocalDryRun">Only back up and show what would be done</param>
        /// <returns>bool - false if the installation failed (error is logged)</returns>
        public static bool InstallEnigmaData(string strLocalHost, string strLocalUser, string strLocalPassword, IReadOnlyList<(string strSourceFile, string strTargetFile)> ltLocalFiles, bool bLocalDryRun)
        {
            LogTrace("DataInstallations.InstallEnigmaData()");
            LogInfo($"-» Install Enigma2 data into {strLocalHost}");

            var (strCurrentHost, iCurrentPort) = ParseHost(strLocalHost);
            var strCurrentHostKeyName = $"{strCurrentHost}:{iCurrentPort}";

            try
            {
                var ciCurrentConnectionInfo = new ConnectionInfo(strCurrentHost, iCurrentPort, strLocalUser, new PasswordAuthenticationMethod(strLocalUser, strLocalPassword)) { Timeout = TimeSpan.FromSeconds(ciEnigmaConnectTimeoutSeconds) };

                using var scCurrentSshClient = new SshClient(ciCurrentConnectionInfo);
                using var scCurrentSftpClient = new SftpClient(ciCurrentConnectionInfo);
                scCurrentSshClient.HostKeyReceived += (_, hkeaCurrentEventArgs) => hkeaCurrentEventArgs.CanTrust = IsTrustedHostKey(strCurrentHostKeyName, hkeaCurrentEventArgs);
                scCurrentSftpClient.HostKeyReceived += (_, hkeaCurrentEventArgs) => hkeaCurrentEventArgs.CanTrust = IsTrustedHostKey(strCurrentHostKeyName, hkeaCurrentEventArgs);

                scCurrentSshClient.Connect();
                scCurrentSftpClient.Connect();

                var strCurrentSystem = BuildEnigmaSystemInfo(RunCommand(scCurrentSshClient, $"cat {csEnigmaImageVersionFile} 2>/dev/null"), RunCommand(scCurrentSshClient, "head -n 1 /etc/issue.net 2>/dev/null || head -n 1 /etc/issue"), RunCommand(scCurrentSshClient, "hostname"));
                LogInfo($"-» Connected to {strCurrentHostKeyName}{(strCurrentSystem.Length > 0 ? $" ({strCurrentSystem})" : string.Empty)}");

                // Symbolic links (e.g. /etc/tuxbox -> /usr/share/tuxbox) are resolved, otherwise the link itself would be replaced
                var ltCurrentFiles = ltLocalFiles.Select(tCurrentFile => tCurrentFile with { strTargetFile = ResolveRemotePath(scCurrentSshClient, tCurrentFile.strTargetFile) }).ToList();

                var strCurrentBackupDirectory = BackupEnigmaSettings(scCurrentSftpClient, strCurrentHost, ltCurrentFiles.Select(tCurrentFile => tCurrentFile.strTargetFile));
                LogInfo($"-» Backup of the receiver settings in {strCurrentBackupDirectory}");
                ReportEnigmaStreamingServices(strCurrentBackupDirectory);

                // Install files are merged with the settings of the receiver - foreign entries, other bouquets (e.g. streaming) and other cables are kept
                var ltCurrentInstallFiles = PrepareEnigmaInstallFiles(ltCurrentFiles, strCurrentBackupDirectory, strCurrentBackupDirectory + csEnigmaInstallDirectorySuffix);

                if (bLocalDryRun)
                {
                    LogInfo("-» Dry run - would stop enigma2 (init 4), install the files and start enigma2 again (init 3)");

                    foreach (var (strCurrentSourceFile, strCurrentTargetFile) in ltCurrentInstallFiles)
                        LogInfo($"-» Dry run - would {(scCurrentSftpClient.Exists(strCurrentTargetFile) ? "replace" : "create")} {strCurrentTargetFile} with {strCurrentSourceFile}");

                    return true;
                }

                var bCurrentEnigmaStopped = false;

                try
                {
                    // enigma2 writes lamedb and bouquets on exit - it has to be stopped before the files are replaced
                    LogInfo("-» Stop enigma2");
                    RunCommand(scCurrentSshClient, "init 4", true);
                    bCurrentEnigmaStopped = true;

                    if (!WaitForEnigma(scCurrentSshClient, false))
                        throw new InvalidOperationException($"enigma2 has not stopped within {ciEnigmaServiceTimeoutSeconds} seconds");

                    foreach (var (strCurrentSourceFile, strCurrentTargetFile) in ltCurrentInstallFiles)
                    {
                        UploadEnigmaFile(scCurrentSshClient, scCurrentSftpClient, strCurrentSourceFile, strCurrentTargetFile);
                        LogInfo($"-» Installed {strCurrentTargetFile}");
                    }

                    return true;
                }
                catch (Exception excCurrentException) when (excCurrentException is SshException or IOException or InvalidOperationException)
                {
                    LogError($"-» Error on installing Enigma2 data into {strLocalHost}: {excCurrentException.Message} - the settings of the receiver can be restored from {strCurrentBackupDirectory}");
                    return false;
                }
                finally
                {
                    if (bCurrentEnigmaStopped)
                    {
                        LogInfo("-» Start enigma2");
                        RunCommand(scCurrentSshClient, "init 3", true);

                        if (!WaitForEnigma(scCurrentSshClient, true))
                            LogWarning($"-» enigma2 has not started within {ciEnigmaServiceTimeoutSeconds} seconds - please check the receiver");
                    }
                }
            }
            catch (Exception excCurrentException) when (excCurrentException is SshException or SocketException or IOException or InvalidOperationException or InvalidDataException or XmlException)
            {
                LogError($"-» Error on installing Enigma2 data into {strLocalHost}: {excCurrentException.Message}");
                return false;
            }
        }

        /// <summary>
        /// DataInstallations.BackupEnigmaSettings() - downloads all files of the settings directory and the target files outside of it
        /// </summary>
        /// <returns>Local backup directory</returns>
        private static string BackupEnigmaSettings(SftpClient scLocalSftpClient, string strLocalHost, IEnumerable<string> ieLocalTargetFiles)
        {
            var strCurrentBackupDirectory = Path.GetFullPath(Path.Combine(GetExportDirectory(), csEnigmaBackupDirectory, $"{string.Concat(strLocalHost.Select(cCurrentCharacter => Path.GetInvalidFileNameChars().Contains(cCurrentCharacter) ? '_' : cCurrentCharacter))}_{DateTime.Now.ToString(csBackupTimestampFormat)}"));
            Directory.CreateDirectory(strCurrentBackupDirectory);

            var lsCurrentRemoteFiles = scLocalSftpClient.ListDirectory(csEnigmaSettingsDirectory).Where(ssfCurrentFile => ssfCurrentFile.IsRegularFile).Select(ssfCurrentFile => ssfCurrentFile.FullName).ToList();
            lsCurrentRemoteFiles.AddRange(ieLocalTargetFiles.Where(strCurrentTargetFile => !strCurrentTargetFile.StartsWith(csEnigmaSettingsDirectory + "/", StringComparison.Ordinal) && scLocalSftpClient.Exists(strCurrentTargetFile)));

            foreach (var strCurrentRemoteFile in lsCurrentRemoteFiles.Distinct(StringComparer.Ordinal))
            {
                var strCurrentLocalFile = GetLocalEnigmaFile(strCurrentBackupDirectory, strCurrentRemoteFile);
                Directory.CreateDirectory(Path.GetDirectoryName(strCurrentLocalFile)!);

                using var fsCurrentFileStream = File.Create(strCurrentLocalFile);
                scLocalSftpClient.DownloadFile(strCurrentRemoteFile, fsCurrentFileStream);
            }

            return strCurrentBackupDirectory;
        }

        /// <summary>
        /// DataInstallations.PrepareEnigmaInstallFiles() - merges the exported files with the backed up settings of the receiver
        /// lamedb / lamedb5: entries of the DVB-C namespace are replaced, all other entries kept (both files written if the receiver uses lamedb5)
        /// bouquets.tv / .radio: kept, missing references to the exported bouquets added
        /// cables.xml: cables named like the export (LIWEST...) replaced, other cables kept
        /// blacklist / whitelist: kept (parental control of the receiver)
        /// </summary>
        /// <returns>Local install file and absolute target path on the receiver</returns>
        private static List<(string strSourceFile, string strTargetFile)> PrepareEnigmaInstallFiles(IReadOnlyList<(string strSourceFile, string strTargetFile)> ltLocalFiles, string strLocalBackupDirectory, string strLocalInstallDirectory)
        {
            Directory.CreateDirectory(strLocalInstallDirectory);
            List<(string strSourceFile, string strTargetFile)> ltCurrentInstallFiles = [];

            foreach (var (strCurrentSourceFile, strCurrentTargetFile) in ltLocalFiles)
            {
                var strCurrentFileName = strCurrentTargetFile[(strCurrentTargetFile.LastIndexOf('/') + 1)..];
                var strCurrentBackupFile = GetLocalEnigmaFile(strLocalBackupDirectory, strCurrentTargetFile);
                var strCurrentInstallFile = GetLocalEnigmaFile(strLocalInstallDirectory, strCurrentTargetFile);
                Directory.CreateDirectory(Path.GetDirectoryName(strCurrentInstallFile)!);

                switch (strCurrentFileName)
                {
                    case csEnigmaChannelDatabaseFile:
                    {
                        var ldCurrentExport = ParseLamedb4(strCurrentSourceFile);
                        var ldCurrentMerged = MergeLamedb(File.Exists(strCurrentBackupFile) ? ParseLamedb4(strCurrentBackupFile) : null, ldCurrentExport);
                        WriteLamedb4(ldCurrentMerged, strCurrentInstallFile);
                        ltCurrentInstallFiles.Add((strCurrentInstallFile, strCurrentTargetFile));
                        LogInfo($"-» Merged {strCurrentFileName}: {ldCurrentExport.lsaTransponders.Count} transponders / {ldCurrentExport.lsaServices.Count} services of the export, {ldCurrentMerged.lsaTransponders.Count - ldCurrentExport.lsaTransponders.Count} transponders / {ldCurrentMerged.lsaServices.Count - ldCurrentExport.lsaServices.Count} services of the receiver kept");

                        // lamedb5 if the receiver uses it or the export contains it - always consistent with lamedb
                        var strCurrentTarget5File = strCurrentTargetFile + "5";
                        var strCurrentBackup5File = strCurrentBackupFile + "5";
                        if (File.Exists(strCurrentBackup5File) || ltLocalFiles.Any(tCurrentFile => tCurrentFile.strTargetFile == strCurrentTarget5File))
                        {
                            WriteLamedb5(File.Exists(strCurrentBackup5File) ? File.ReadAllLines(strCurrentBackup5File) : [], ldCurrentExport, strCurrentInstallFile + "5");
                            ltCurrentInstallFiles.Add((strCurrentInstallFile + "5", strCurrentTarget5File));
                        }

                        break;
                    }
                    case csEnigmaChannelDatabaseFile + "5":
                        // Written together with lamedb
                        break;
                    case "blacklist":
                    case "whitelist":
                        LogInfo($"-» Kept {strCurrentFileName} of the receiver (parental control)");
                        break;
                    case "bouquets.tv":
                    case "bouquets.radio":
                        MergeBouquetsFile(File.Exists(strCurrentBackupFile) ? strCurrentBackupFile : null, strCurrentSourceFile, strCurrentInstallFile);
                        ltCurrentInstallFiles.Add((strCurrentInstallFile, strCurrentTargetFile));
                        break;
                    case "cables.xml":
                        MergeCablesFile(File.Exists(strCurrentBackupFile) ? strCurrentBackupFile : null, strCurrentSourceFile, strCurrentInstallFile);
                        ltCurrentInstallFiles.Add((strCurrentInstallFile, strCurrentTargetFile));
                        break;
                    default:
                        File.Copy(strCurrentSourceFile, strCurrentInstallFile, true);
                        ltCurrentInstallFiles.Add((strCurrentInstallFile, strCurrentTargetFile));
                        break;
                }
            }

            return ltCurrentInstallFiles;
        }

        /// <summary>
        /// Content of a lamedb (version 4): transponder blocks (key, data, "/") and service blocks (key, name, provider data)
        /// </summary>
        private sealed record LamedbData(List<string[]> lsaTransponders, List<string[]> lsaServices, List<string> lsTrailer);

        /// <summary>
        /// DataInstallations.ParseLamedb4() - parses a lamedb of version 4 (lamedb5 is converted to lamedb4) into transponder and service blocks, the trailer is kept
        /// </summary>
        /// <param name="strLocalFile"></param>
        /// <returns></returns>
        /// <exception cref="InvalidDataException"></exception>
        private static LamedbData ParseLamedb4(string strLocalFile)
        {
            var saCurrentLines = File.ReadAllLines(strLocalFile);
            var iCurrentLine = Array.IndexOf(saCurrentLines, "transponders") + 1;

            if (iCurrentLine == 0 || !saCurrentLines[0].StartsWith("eDVB services /4/", StringComparison.Ordinal))
                throw new InvalidDataException($"{strLocalFile} is no lamedb of version 4");

            List<string[]> lsaCurrentTransponders = [];
            while (iCurrentLine < saCurrentLines.Length && saCurrentLines[iCurrentLine] != "end")
            {
                var iCurrentBlockStart = iCurrentLine;
                while (iCurrentLine < saCurrentLines.Length && saCurrentLines[iCurrentLine] != "/")
                    iCurrentLine++;

                if (iCurrentLine >= saCurrentLines.Length)
                    throw new InvalidDataException($"{strLocalFile}: transponder section not terminated");

                lsaCurrentTransponders.Add(saCurrentLines[iCurrentBlockStart..(iCurrentLine + 1)]);
                iCurrentLine++;
            }

            iCurrentLine = Array.IndexOf(saCurrentLines, "services", iCurrentLine) + 1;
            if (iCurrentLine == 0)
                throw new InvalidDataException($"{strLocalFile}: services section missing");

            List<string[]> lsaCurrentServices = [];
            while (iCurrentLine < saCurrentLines.Length && saCurrentLines[iCurrentLine] != "end")
            {
                if (iCurrentLine + 2 >= saCurrentLines.Length)
                    throw new InvalidDataException($"{strLocalFile}: services section not terminated");

                lsaCurrentServices.Add(saCurrentLines[iCurrentLine..(iCurrentLine + 3)]);
                iCurrentLine += 3;
            }

            return new LamedbData(lsaCurrentTransponders, lsaCurrentServices, saCurrentLines.Skip(iCurrentLine + 1).Where(strCurrentLine => strCurrentLine.Length > 0).ToList());
        }

        /// <summary>
        /// DataInstallations.MergeLamedb() - entries of the receiver outside of the DVB-C namespace of the export are kept, the export is added
        /// </summary>
        private static LamedbData MergeLamedb(LamedbData? ldLocalReceiver, LamedbData ldLocalExport)
        {
            static bool IsExportNamespace(string strLocalNamespace) => strLocalNamespace.Equals(ciDefaultDVBCNamespace, StringComparison.OrdinalIgnoreCase);

            var lsaCurrentTransponders = (ldLocalReceiver?.lsaTransponders ?? []).Where(saCurrentBlock => !IsExportNamespace(saCurrentBlock[0].Split(':')[0])).Concat(ldLocalExport.lsaTransponders).ToList();
            var lsaCurrentServices = (ldLocalReceiver?.lsaServices ?? []).Where(saCurrentBlock => saCurrentBlock[0].Split(':') is not { Length: > 1 } saCurrentKey || !IsExportNamespace(saCurrentKey[1])).Concat(ldLocalExport.lsaServices).ToList();

            return new LamedbData(lsaCurrentTransponders, lsaCurrentServices, ldLocalExport.lsTrailer);
        }

        /// <summary>
        /// DataInstallations.WriteLamedb4() - writes a lamedb of version 4 from transponder and service blocks, the trailer is kept
        /// </summary>
        /// <param name="ldLocalLamedb"></param>
        /// <param name="strLocalFile"></param>
        private static void WriteLamedb4(LamedbData ldLocalLamedb, string strLocalFile)
        {
            List<string> lsCurrentLines =
            [
                "eDVB services /4/", "transponders",
                .. ldLocalLamedb.lsaTransponders.SelectMany(saCurrentBlock => saCurrentBlock),
                "end", "services",
                .. ldLocalLamedb.lsaServices.SelectMany(saCurrentBlock => saCurrentBlock),
                "end",
                .. ldLocalLamedb.lsTrailer
            ];

            WriteEnigmaTextFile(strLocalFile, lsCurrentLines, "\n");
        }

        /// <summary>
        /// DataInstallations.WriteLamedb5() - t: / s: lines of the receiver outside of the DVB-C namespace of the export are kept, the export is converted (t:key,data / s:key,"name",provider)
        /// </summary>
        private static void WriteLamedb5(string[] saLocalReceiverLines, LamedbData ldLocalExport, string strLocalFile)
        {
            static bool IsExportNamespace(string strLocalLine, int iLocalField) => strLocalLine[2..].Split(',')[0].Split(':') is { } saCurrentKey && saCurrentKey.Length > iLocalField && saCurrentKey[iLocalField].Equals(ciDefaultDVBCNamespace, StringComparison.OrdinalIgnoreCase);

            var lsCurrentTransponders = saLocalReceiverLines.Where(strCurrentLine => strCurrentLine.StartsWith("t:", StringComparison.Ordinal) && !IsExportNamespace(strCurrentLine, 0)).ToList();
            var lsCurrentServices = saLocalReceiverLines.Where(strCurrentLine => strCurrentLine.StartsWith("s:", StringComparison.Ordinal) && !IsExportNamespace(strCurrentLine, 1)).ToList();

            // "\tc 586000:6900000:2:5:9:0:0" -> "c:586000:6900000:2:5:9:0:0"
            lsCurrentTransponders.AddRange(ldLocalExport.lsaTransponders.Select(saCurrentBlock => $"t:{saCurrentBlock[0]},{ReplaceFirst(saCurrentBlock[1].Trim(), ' ', ':')}"));
            lsCurrentServices.AddRange(ldLocalExport.lsaServices.Select(saCurrentBlock => $"s:{saCurrentBlock[0]},\"{saCurrentBlock[1]}\",{saCurrentBlock[2]}"));

            List<string> lsCurrentLines =
            [
                "eDVB services /5/",
                .. lsCurrentTransponders,
                .. lsCurrentServices,
                $"# done. {lsCurrentTransponders.Count} channels and {lsCurrentServices.Count} services"
            ];

            WriteEnigmaTextFile(strLocalFile, lsCurrentLines, "\n");
        }

        /// <summary>
        /// DataInstallations.MergeBouquetsFile() - bouquets.tv / .radio of the receiver kept, missing references of the export added
        /// </summary>
        private static void MergeBouquetsFile(string? strLocalReceiverFile, string strLocalExportFile, string strLocalInstallFile)
        {
            if (strLocalReceiverFile == null)
            {
                File.Copy(strLocalExportFile, strLocalInstallFile, true);
                return;
            }

            static string? GetBouquetFile(string strLocalLine) => strLocalLine.Contains("FROM BOUQUET \"", StringComparison.OrdinalIgnoreCase) ? strLocalLine.Split('"') is { Length: > 1 } saCurrentParts ? saCurrentParts[1] : null : null;

            var strCurrentContent = File.ReadAllText(strLocalReceiverFile);
            var strCurrentNewLine = strCurrentContent.Contains("\r\n") ? "\r\n" : "\n";
            var lsCurrentLines = strCurrentContent.Split(["\r\n", "\n"], StringSplitOptions.None).ToList();
            while (lsCurrentLines.Count > 0 && lsCurrentLines[^1].Length == 0)
                lsCurrentLines.RemoveAt(lsCurrentLines.Count - 1);

            var hsCurrentBouquets = lsCurrentLines.Select(GetBouquetFile).OfType<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var strCurrentExportLine in File.ReadAllLines(strLocalExportFile))
            {
                if (GetBouquetFile(strCurrentExportLine) is { } strCurrentBouquet && hsCurrentBouquets.Add(strCurrentBouquet))
                {
                    lsCurrentLines.Add(strCurrentExportLine);
                    LogInfo($"-» Added bouquet {strCurrentBouquet} to {Path.GetFileName(strLocalInstallFile)}");
                }
            }

            WriteEnigmaTextFile(strLocalInstallFile, lsCurrentLines, strCurrentNewLine);
        }

        /// <summary>
        /// DataInstallations.MergeCablesFile() - cables of the receiver named like the exported cables (LIWEST...) replaced, other cables kept
        /// </summary>
        private static void MergeCablesFile(string? strLocalReceiverFile, string strLocalExportFile, string strLocalInstallFile)
        {
            if (strLocalReceiverFile == null)
            {
                File.Copy(strLocalExportFile, strLocalInstallFile, true);
                return;
            }

            var xdCurrentReceiver = XDocument.Load(strLocalReceiverFile);
            var lxeCurrentExportCables = XDocument.Load(strLocalExportFile).Root?.Elements("cable").ToList() ?? [];
            var lxeCurrentReplacedCables = xdCurrentReceiver.Root?.Elements("cable").Where(xeCurrentCable => ((string?)xeCurrentCable.Attribute("name"))?.StartsWith(csCategory, StringComparison.OrdinalIgnoreCase) == true).ToList() ?? [];

            if (xdCurrentReceiver.Root == null)
                throw new InvalidDataException($"{strLocalReceiverFile} contains no cables");

            // Exported cables at the position of the first replaced cable, otherwise at the end
            if (lxeCurrentReplacedCables.Count > 0)
                lxeCurrentReplacedCables[0].AddBeforeSelf(lxeCurrentExportCables);
            else
                xdCurrentReceiver.Root.Add(lxeCurrentExportCables);

            lxeCurrentReplacedCables.ForEach(xeCurrentCable => xeCurrentCable.Remove());
            LogInfo($"-» Merged cables.xml: {lxeCurrentReplacedCables.Count} cable(s) replaced, {xdCurrentReceiver.Root.Elements("cable").Count() - lxeCurrentExportCables.Count} cable(s) of the receiver kept");

            var eCurrentEncoding = xdCurrentReceiver.Declaration?.Encoding is { Length: > 0 } strCurrentEncodingName ? Encoding.GetEncoding(strCurrentEncodingName) : new UTF8Encoding(false);
            using var xwCurrentWriter = XmlWriter.Create(strLocalInstallFile, new XmlWriterSettings { Encoding = eCurrentEncoding, Indent = true });
            xdCurrentReceiver.Save(xwCurrentWriter);
        }

        /// <summary>
        /// DataInstallations.WriteEnigmaTextFile() - writes a text file with the given lines and new line, UTF-8 without BOM
        /// </summary>
        /// <param name="strLocalFile"></param>
        /// <param name="ieLocalLines"></param>
        /// <param name="strLocalNewLine"></param>
        private static void WriteEnigmaTextFile(string strLocalFile, IEnumerable<string> ieLocalLines, string strLocalNewLine)
        {
            File.WriteAllText(strLocalFile, string.Join(strLocalNewLine, ieLocalLines) + strLocalNewLine, new UTF8Encoding(false));
        }

        /// <summary>
        /// DataInstallations.ReplaceFirst() - replaces the first occurrence of a character in a string
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <param name="cLocalOld"></param>
        /// <param name="cLocalNew"></param>
        /// <returns></returns>
        private static string ReplaceFirst(string strLocalString, char cLocalOld, char cLocalNew)
        {
            var iCurrentIndex = strLocalString.IndexOf(cLocalOld);
            return iCurrentIndex < 0 ? strLocalString : strLocalString[..iCurrentIndex] + cLocalNew + strLocalString[(iCurrentIndex + 1)..];
        }

        /// <summary>
        /// DataInstallations.UploadEnigmaFile() - uploads to a temporary file and moves it to the target, the target is never partly written
        /// </summary>
        private static void UploadEnigmaFile(SshClient scLocalSshClient, SftpClient scLocalSftpClient, string strLocalSourceFile, string strLocalTargetFile)
        {
            var strCurrentTemporaryFile = strLocalTargetFile + csEnigmaUploadSuffix;

            using (var fsCurrentFileStream = File.OpenRead(strLocalSourceFile))
                scLocalSftpClient.UploadFile(fsCurrentFileStream, strCurrentTemporaryFile, true);

            RunCommand(scLocalSshClient, $"mv -f {QuoteShell(strCurrentTemporaryFile)} {QuoteShell(strLocalTargetFile)}", true);
        }

        /// <summary>
        /// DataInstallations.WaitForEnigma() - waits until enigma2 is running / stopped
        /// </summary>
        private static bool WaitForEnigma(SshClient scLocalSshClient, bool bLocalRunning)
        {
            var swCurrentStopwatch = System.Diagnostics.Stopwatch.StartNew();

            while (swCurrentStopwatch.Elapsed < TimeSpan.FromSeconds(ciEnigmaServiceTimeoutSeconds))
            {
                if (RunCommand(scLocalSshClient, "pidof enigma2 >/dev/null && echo running || echo stopped").Trim() == (bLocalRunning ? "running" : "stopped"))
                    return true;

                Thread.Sleep(TimeSpan.FromSeconds(1));
            }

            return false;
        }

        /// <summary>
        /// DataInstallations.GetLocalEnigmaFile() - files of the settings directory directly in the local directory, all others with their path (e.g. etc/tuxbox/cables.xml)
        /// </summary>
        private static string GetLocalEnigmaFile(string strLocalDirectory, string strLocalRemoteFile)
        {
            return Path.Combine(strLocalDirectory, strLocalRemoteFile.StartsWith(csEnigmaSettingsDirectory + "/", StringComparison.Ordinal) ? strLocalRemoteFile[(csEnigmaSettingsDirectory.Length + 1)..] : strLocalRemoteFile.TrimStart('/').Replace('/', Path.DirectorySeparatorChar));
        }

        /// <summary>
        /// DataInstallations.QuoteShell() - quotes a string for the shell, single quotes are escaped by closing and reopening the quotes
        /// </summary>
        /// <param name="strLocalString"></param>
        /// <returns></returns>
        private static string QuoteShell(string strLocalString)
        {
            return $"'{strLocalString.Replace("'", "'\\''")}'";
        }

        /// <summary>
        /// DataInstallations.ReportEnigmaStreamingServices() - streaming services (4097, 5001, 5002, 8193 or stream url) found in the backed up bouquets and lamedb
        /// </summary>
        private static void ReportEnigmaStreamingServices(string strLocalBackupDirectory)
        {
            foreach (var strCurrentFile in Directory.GetFiles(strLocalBackupDirectory).Where(strCurrentFile => Path.GetFileName(strCurrentFile).StartsWith("userbouquet.", StringComparison.OrdinalIgnoreCase) || Path.GetFileName(strCurrentFile).StartsWith("bouquets.", StringComparison.OrdinalIgnoreCase)))
            {
                var iCurrentCount = rxEnigmaStreamingService().Matches(File.ReadAllText(strCurrentFile)).Count;
                if (iCurrentCount > 0)
                    LogInfo($"-» Streaming services: {iCurrentCount} in {Path.GetFileName(strCurrentFile)}");
            }

            foreach (var strCurrentFile in Directory.GetFiles(strLocalBackupDirectory, csEnigmaChannelDatabaseFile + "*"))
            {
                var strCurrentContent = File.ReadAllText(strCurrentFile);
                if (strCurrentContent.Contains("http%3a", StringComparison.OrdinalIgnoreCase) || strCurrentContent.Contains("http://", StringComparison.OrdinalIgnoreCase) || strCurrentContent.Contains("https://", StringComparison.OrdinalIgnoreCase))
                    LogInfo($"-» Streaming services: stream urls found in {Path.GetFileName(strCurrentFile)}");
            }
        }

        /// <summary>
        /// DataInstallations.IsTrustedHostKey() - trust on first use: unknown hosts are stored, a changed host key is rejected
        /// </summary>
        private static bool IsTrustedHostKey(string strLocalHostKeyName, HostKeyEventArgs hkeaLocalEventArgs)
        {
            var strCurrentKnownHostsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), cstrCurrentProjectName, csEnigmaKnownHostsFile);
            var strCurrentEntry = $"{hkeaLocalEventArgs.HostKeyName} {hkeaLocalEventArgs.FingerPrintSHA256}";

            var dCurrentKnownHosts = File.Exists(strCurrentKnownHostsFile)
                ? File.ReadAllLines(strCurrentKnownHostsFile).Select(strCurrentLine => strCurrentLine.Split(' ', 2)).Where(saCurrentLine => saCurrentLine.Length == 2).GroupBy(saCurrentLine => saCurrentLine[0]).ToDictionary(igCurrentGroup => igCurrentGroup.Key, igCurrentGroup => igCurrentGroup.Select(saCurrentLine => saCurrentLine[1]).ToList())
                : [];

            if (dCurrentKnownHosts.TryGetValue(strLocalHostKeyName, out var lsCurrentKnownKeys))
            {
                if (lsCurrentKnownKeys.Contains(strCurrentEntry))
                    return true;

                LogError($"-» Host key of {strLocalHostKeyName} has changed ({strCurrentEntry}) - connection refused. If the receiver has been reinstalled, remove its entry from {strCurrentKnownHostsFile}");
                return false;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(strCurrentKnownHostsFile)!);
            File.AppendAllLines(strCurrentKnownHostsFile, [$"{strLocalHostKeyName} {strCurrentEntry}"]);
            LogWarning($"-» New host key of {strLocalHostKeyName} stored ({strCurrentEntry})");

            return true;
        }

        /// <summary>
        /// DataInstallations.BuildEnigmaSystemInfo() - image name and version of /etc/image-version (e.g. comment=openATV, version=7.5) and the host name,
        /// otherwise the login banner with its placeholders (%h, \n, \l, ...) replaced - they are only expanded by the login service
        /// </summary>
        /// <returns>e.g. "openATV 7.5 - vuduo4k"</returns>
        private static string BuildEnigmaSystemInfo(string strLocalImageVersion, string strLocalIssue, string strLocalHostName)
        {
            var strCurrentHostName = strLocalHostName.Trim();

            var dCurrentImageVersion = strLocalImageVersion.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(strCurrentLine => strCurrentLine.Split('=', 2, StringSplitOptions.TrimEntries))
                .Where(saCurrentKeyValue => saCurrentKeyValue is [_, { Length: > 0 }])
                .GroupBy(saCurrentKeyValue => saCurrentKeyValue[0], StringComparer.OrdinalIgnoreCase)
                .ToDictionary(igCurrentGroup => igCurrentGroup.Key, igCurrentGroup => igCurrentGroup.First()[1], StringComparer.OrdinalIgnoreCase);

            string? GetValue(params string[] saLocalKeys) => saLocalKeys.Select(dCurrentImageVersion.GetValueOrDefault).FirstOrDefault(strCurrentValue => strCurrentValue != null);

            var strCurrentInfo = $"{GetValue(csaEnigmaImageNameKeys)} {GetValue(csaEnigmaImageVersionKeys)}".Trim();

            if (strCurrentInfo.Length == 0)
                strCurrentInfo = rxLoginBannerEscape().Replace(strLocalIssue.Replace("%h", strCurrentHostName), string.Empty).Trim();

            return strCurrentHostName.Length > 0 && !strCurrentInfo.Contains(strCurrentHostName, StringComparison.OrdinalIgnoreCase) ? $"{strCurrentInfo} - {strCurrentHostName}".Trim(' ', '-') : strCurrentInfo;
        }

        /// <summary>
        /// DataInstallations.rxLoginBannerEscape() - regex for placeholders in the login banner (%h, \n, \l, ...) that are only expanded by the login service
        /// </summary>
        /// <returns></returns>
        [GeneratedRegex(@"\\[a-zA-Z]|%[a-zA-Z]")]
        private static partial Regex rxLoginBannerEscape();

        /// <summary>
        /// DataInstallations.ResolveRemotePath() - resolves symbolic links on the receiver, otherwise the link itself would be replaced
        /// </summary>
        /// <param name="scLocalSshClient"></param>
        /// <param name="strLocalPath"></param>
        /// <returns></returns>
        private static string ResolveRemotePath(SshClient scLocalSshClient, string strLocalPath)
        {
            var strCurrentResolvedPath = RunCommand(scLocalSshClient, $"readlink -f {QuoteShell(strLocalPath)}").Trim();
            return strCurrentResolvedPath.StartsWith('/') ? strCurrentResolvedPath : strLocalPath;
        }

        /// <summary>
        /// DataInstallations.RunCommand() - executes a command on the receiver and returns the output, optionally checks the exit status
        /// </summary>
        /// <param name="scLocalSshClient"></param>
        /// <param name="strLocalCommand"></param>
        /// <param name="bLocalCheckExitStatus"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        private static string RunCommand(SshClient scLocalSshClient, string strLocalCommand, bool bLocalCheckExitStatus = false)
        {
            LogTrace($"DataInstallations.RunCommand() - {strLocalCommand}");

            using var scCurrentCommand = scLocalSshClient.CreateCommand(strLocalCommand);
            scCurrentCommand.CommandTimeout = TimeSpan.FromSeconds(ciEnigmaCommandTimeoutSeconds);

            var strCurrentResult = scCurrentCommand.Execute();

            if (bLocalCheckExitStatus && scCurrentCommand.ExitStatus is { } iCurrentExitStatus && iCurrentExitStatus != 0)
                throw new InvalidOperationException($"'{strLocalCommand}' failed ({iCurrentExitStatus}): {scCurrentCommand.Error.Trim()}");

            return strCurrentResult;
        }

        /// <summary>
        /// DataInstallations.ParseHost() - parses a host string into host and port, IPv6 addresses only with port in brackets ([::1]:22)
        /// </summary>
        /// <param name="strLocalHost"></param>
        /// <returns></returns>
        private static (string strHost, int iPort) ParseHost(string strLocalHost)
        {
            var strCurrentHost = strLocalHost.Trim();
            var iCurrentSeparator = strCurrentHost.LastIndexOf(':');

            // host:port - IPv6 addresses only with port in brackets ([::1]:22)
            if (iCurrentSeparator > 0 && strCurrentHost.IndexOf(':') == iCurrentSeparator && int.TryParse(strCurrentHost[(iCurrentSeparator + 1)..], out var iCurrentPort))
                return (strCurrentHost[..iCurrentSeparator], iCurrentPort);

            if (strCurrentHost.StartsWith('[') && strCurrentHost.Contains("]:") && int.TryParse(strCurrentHost[(strCurrentHost.LastIndexOf("]:", StringComparison.Ordinal) + 2)..], out iCurrentPort))
                return (strCurrentHost[1..strCurrentHost.LastIndexOf(']')], iCurrentPort);

            return (strCurrentHost.Trim('[', ']'), ciEnigmaSshPort);
        }
    }
}