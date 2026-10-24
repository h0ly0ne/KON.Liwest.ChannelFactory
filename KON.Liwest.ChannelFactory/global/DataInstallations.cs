using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.ServiceProcess;

using Microsoft.Win32;

using static KON.Liwest.ChannelFactory.Constants;
using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Installed DVBViewer / DVBViewer Media Server instance
    /// </summary>
    /// <param name="strName">Product name of the installation (DisplayName of the uninstall entry)</param>
    /// <param name="strInstallDirectory">Program directory</param>
    /// <param name="strDataDirectory">Data directory (channels.dat, Transponders) - defined by usermode.ini</param>
    public record DVBViewerInstance(string strName, string strInstallDirectory, string strDataDirectory);

    /// <summary>
    /// Installation of exported data into installed applications
    /// </summary>
    public static partial class DataInstallations
    {
        /// <summary>
        /// DataInstallations.FindDVBViewerInstances() - DVBViewer / DVBViewer Media Server installations of the uninstall entries (64 bit, 32 bit and current user)
        /// </summary>
        /// <returns>List of found instances</returns>
        [SupportedOSPlatform("windows")]
        public static List<DVBViewerInstance> FindDVBViewerInstances()
        {
            LogTrace("DataInstallations.FindDVBViewerInstances()");

            List<DVBViewerInstance> ldviCurrentInstances = [];

            foreach (var (rhCurrentHive, rvCurrentView) in new[] { (RegistryHive.LocalMachine, RegistryView.Registry64), (RegistryHive.LocalMachine, RegistryView.Registry32), (RegistryHive.CurrentUser, RegistryView.Default) })
            {
                using var rkCurrentBaseKey = RegistryKey.OpenBaseKey(rhCurrentHive, rvCurrentView);
                using var rkCurrentUninstallKey = rkCurrentBaseKey.OpenSubKey(csUninstallRegistryKey);

                foreach (var strCurrentSubKeyName in rkCurrentUninstallKey?.GetSubKeyNames() ?? [])
                {
                    using var rkCurrentProductKey = rkCurrentUninstallKey!.OpenSubKey(strCurrentSubKeyName);

                    if (rkCurrentProductKey?.GetValue("DisplayName") is not string strCurrentDisplayName || !csaDVBViewerProductNames.Any(strCurrentProductName => strCurrentDisplayName.StartsWith(strCurrentProductName, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var strCurrentInstallDirectory = GetInstallDirectory(rkCurrentProductKey);
                    if (strCurrentInstallDirectory == null || !Directory.Exists(strCurrentInstallDirectory) || ldviCurrentInstances.Any(dviCurrentInstance => IsSamePath(dviCurrentInstance.strInstallDirectory, strCurrentInstallDirectory)))
                        continue;

                    ldviCurrentInstances.Add(new DVBViewerInstance(strCurrentDisplayName, strCurrentInstallDirectory, GetDataDirectory(strCurrentInstallDirectory)));
                }
            }

            return ldviCurrentInstances;
        }

        /// <summary>
        /// DataInstallations.InstallDVBViewerData() - stops running programs / services of the instance, backs up and replaces the files and restarts them
        /// </summary>
        /// <param name="dviLocalInstance"></param>
        /// <param name="ltLocalFiles">Source file and target file relative to the data directory</param>
        /// <param name="bLocalDryRun">Only show what would be done</param>
        /// <returns>bool - false if the installation failed (error is logged)</returns>
        [SupportedOSPlatform("windows")]
        public static bool InstallDVBViewerData(DVBViewerInstance dviLocalInstance, IReadOnlyList<(string strSourceFile, string strTargetFile)> ltLocalFiles, bool bLocalDryRun = false)
        {
            LogTrace("DataInstallations.InstallDVBViewerData()");
            LogInfo($"-» Install DVBViewer data into {dviLocalInstance.strName} ({dviLocalInstance.strDataDirectory})");

            // One timestamp for all backups of the installation
            var strCurrentBackupTimestamp = DateTime.Now.ToString(csBackupTimestampFormat);

            if (bLocalDryRun)
            {
                foreach (var scCurrentService in GetRunningServices(dviLocalInstance.strInstallDirectory))
                {
                    using (scCurrentService)
                        LogInfo($"-» Dry run - would stop and restart service {scCurrentService.ServiceName}");
                }

                foreach (var pCurrentProcess in GetRunningProcesses(dviLocalInstance.strInstallDirectory))
                {
                    using (pCurrentProcess)
                        LogInfo($"-» Dry run - would close and restart program {Path.GetFileName(pCurrentProcess.MainModule!.FileName)}");
                }

                foreach (var (strCurrentSourceFile, strCurrentTargetFile) in ltLocalFiles)
                {
                    var strCurrentTargetPath = Path.Combine(dviLocalInstance.strDataDirectory, strCurrentTargetFile);
                    LogInfo($"-» Dry run - would {(File.Exists(strCurrentTargetPath) ? "back up and replace" : "create")} {strCurrentTargetPath} with {strCurrentSourceFile}");
                }

                return true;
            }

            List<ServiceController> lscCurrentStoppedServices = [];
            List<string> lsCurrentStoppedPrograms = [];

            try
            {
                // Services first - their processes are part of the program directory as well
                foreach (var scCurrentService in GetRunningServices(dviLocalInstance.strInstallDirectory))
                {
                    LogInfo($"-» Stop service {scCurrentService.ServiceName}");
                    scCurrentService.Stop();
                    scCurrentService.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(ciDVBViewerServiceTimeoutSeconds));
                    lscCurrentStoppedServices.Add(scCurrentService);
                }

                foreach (var pCurrentProcess in GetRunningProcesses(dviLocalInstance.strInstallDirectory))
                {
                    using var _ = pCurrentProcess;
                    var strCurrentProgram = pCurrentProcess.MainModule!.FileName;
                    LogInfo($"-» Close program {Path.GetFileName(strCurrentProgram)}");

                    // Regular close first so the program saves its data, killed only if it does not end in time
                    if (!pCurrentProcess.CloseMainWindow() || !pCurrentProcess.WaitForExit(TimeSpan.FromSeconds(ciDVBViewerProcessExitTimeoutSeconds)))
                    {
                        LogWarning($"-» Program {Path.GetFileName(strCurrentProgram)} did not close - terminated");
                        pCurrentProcess.Kill();
                        pCurrentProcess.WaitForExit(TimeSpan.FromSeconds(ciDVBViewerProcessExitTimeoutSeconds));
                    }

                    lsCurrentStoppedPrograms.Add(strCurrentProgram);
                }

                foreach (var (strCurrentSourceFile, strCurrentTargetFile) in ltLocalFiles)
                {
                    var strCurrentTargetPath = Path.Combine(dviLocalInstance.strDataDirectory, strCurrentTargetFile);
                    Directory.CreateDirectory(Path.GetDirectoryName(strCurrentTargetPath)!);

                    if (File.Exists(strCurrentTargetPath))
                    {
                        var strCurrentBackupPath = $"{strCurrentTargetPath}.{strCurrentBackupTimestamp}.bak";
                        File.Copy(strCurrentTargetPath, strCurrentBackupPath, true);
                        LogInfo($"-» Backup {strCurrentBackupPath}");
                    }

                    File.Copy(strCurrentSourceFile, strCurrentTargetPath, true);
                    LogInfo($"-» Installed {strCurrentTargetPath}");
                }

                return true;
            }
            catch (Exception excCurrentException) when (excCurrentException is IOException or UnauthorizedAccessException or InvalidOperationException or Win32Exception or System.ServiceProcess.TimeoutException)
            {
                LogError($"-» Error on installing DVBViewer data into {dviLocalInstance.strName}: {excCurrentException.Message}");
                return false;
            }
            finally
            {
                RestartPrograms(lsCurrentStoppedPrograms);
                RestartServices(lscCurrentStoppedServices);
            }
        }

        /// <summary>
        /// DataInstallations.GetInstallDirectory() - InstallLocation, otherwise the directory of DisplayIcon / UninstallString
        /// </summary>
        [SupportedOSPlatform("windows")]
        private static string? GetInstallDirectory(RegistryKey rkLocalProductKey)
        {
            if (rkLocalProductKey.GetValue("InstallLocation") is string strCurrentInstallLocation && !string.IsNullOrWhiteSpace(strCurrentInstallLocation))
                return NormalizeDirectory(strCurrentInstallLocation);

            foreach (var strCurrentValueName in new[] { "DisplayIcon", "UninstallString" })
            {
                if (rkLocalProductKey.GetValue(strCurrentValueName) is not string strCurrentValue || string.IsNullOrWhiteSpace(strCurrentValue))
                    continue;

                // "C:\...\file.exe" /args or C:\...\file.ico,0
                var strCurrentFile = strCurrentValue.Trim().StartsWith('"') ? strCurrentValue.Trim().Split('"', StringSplitOptions.RemoveEmptyEntries)[0] : strCurrentValue.Split(',')[0].Trim();
                var strCurrentDirectory = Path.GetDirectoryName(strCurrentFile);

                if (!string.IsNullOrEmpty(strCurrentDirectory))
                    return NormalizeDirectory(strCurrentDirectory);
            }

            return null;
        }

        /// <summary>
        /// DataInstallations.GetDataDirectory() - usermode.ini of the program directory: UserMode 0 = program directory, 1 = application data of the user, 2 = common application data (each + Root)
        /// </summary>
        private static string GetDataDirectory(string strLocalInstallDirectory)
        {
            var strCurrentUserModeFile = Path.Combine(strLocalInstallDirectory, csDVBViewerUserModeFile);
            if (!File.Exists(strCurrentUserModeFile))
                return strLocalInstallDirectory;

            var iCurrentUserMode = 0;
            var strCurrentRoot = Path.Combine(csDVBViewerDefaultRootDirectory, Path.GetFileName(strLocalInstallDirectory));

            foreach (var strCurrentLine in File.ReadAllLines(strCurrentUserModeFile))
            {
                var saCurrentKeyValue = strCurrentLine.Split('=', 2, StringSplitOptions.TrimEntries);
                if (saCurrentKeyValue.Length != 2)
                    continue;

                if (saCurrentKeyValue[0].Equals("UserMode", StringComparison.OrdinalIgnoreCase))
                    iCurrentUserMode = saCurrentKeyValue[1].ToInt32();
                else if (saCurrentKeyValue[0].Equals("Root", StringComparison.OrdinalIgnoreCase) && saCurrentKeyValue[1].Length > 0)
                    strCurrentRoot = saCurrentKeyValue[1];
            }

            return iCurrentUserMode switch
            {
                1 => NormalizeDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), strCurrentRoot)),
                2 => NormalizeDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), strCurrentRoot)),
                _ => strLocalInstallDirectory
            };
        }

        /// <summary>
        /// DataInstallations.GetRunningServices() - running services whose executable is located in the program directory
        /// </summary>
        [SupportedOSPlatform("windows")]
        private static List<ServiceController> GetRunningServices(string strLocalInstallDirectory)
        {
            List<ServiceController> lscCurrentServices = [];

            foreach (var scCurrentService in ServiceController.GetServices())
            {
                // Services which are not part of the instance are released immediately
                if (scCurrentService.Status != ServiceControllerStatus.Running)
                {
                    scCurrentService.Dispose();
                    continue;
                }

                using var rkCurrentServiceKey = Registry.LocalMachine.OpenSubKey($@"{csServicesRegistryKey}\{scCurrentService.ServiceName}");
                var strCurrentImagePath = Environment.ExpandEnvironmentVariables(rkCurrentServiceKey?.GetValue("ImagePath") as string ?? string.Empty).Trim();
                var strCurrentExecutable = strCurrentImagePath.StartsWith('"') ? strCurrentImagePath.Split('"', StringSplitOptions.RemoveEmptyEntries)[0] : strCurrentImagePath.Split(' ')[0];

                if (strCurrentExecutable.Length > 0 && IsInDirectory(strCurrentExecutable, strLocalInstallDirectory))
                    lscCurrentServices.Add(scCurrentService);
                else
                    scCurrentService.Dispose();
            }

            return lscCurrentServices;
        }

        /// <summary>
        /// DataInstallations.GetRunningProcesses() - running processes whose executable is located in the program directory
        /// </summary>
        private static List<Process> GetRunningProcesses(string strLocalInstallDirectory)
        {
            List<Process> lpCurrentProcesses = [];

            foreach (var pCurrentProcess in Process.GetProcesses())
            {
                var bCurrentIsInstanceProcess = false;

                try
                {
                    bCurrentIsInstanceProcess = pCurrentProcess.Id != Environment.ProcessId && pCurrentProcess.MainModule?.FileName is { } strCurrentFileName && IsInDirectory(strCurrentFileName, strLocalInstallDirectory);
                }
                catch (Exception excCurrentException) when (excCurrentException is Win32Exception or InvalidOperationException)
                {
                    // No access to system / ended processes - not part of the instance
                }

                // Processes which are not part of the instance are released immediately
                if (bCurrentIsInstanceProcess)
                    lpCurrentProcesses.Add(pCurrentProcess);
                else
                    pCurrentProcess.Dispose();
            }

            return lpCurrentProcesses;
        }

        /// <summary>
        /// DataInstallations.RestartPrograms()
        /// </summary>
        /// <param name="lsLocalPrograms"></param>
        private static void RestartPrograms(List<string> lsLocalPrograms)
        {
            foreach (var strCurrentProgram in lsLocalPrograms.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    LogInfo($"-» Restart program {Path.GetFileName(strCurrentProgram)}");
                    Process.Start(new ProcessStartInfo(strCurrentProgram) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(strCurrentProgram) });
                }
                catch (Exception excCurrentException) when (excCurrentException is Win32Exception or InvalidOperationException)
                {
                    LogError($"-» Error on restarting program {strCurrentProgram}: {excCurrentException.Message}");
                }
            }
        }

        /// <summary>
        /// DataInstallations.RestartServices()
        /// </summary>
        /// <param name="lscLocalServices"></param>
        [SupportedOSPlatform("windows")]
        private static void RestartServices(List<ServiceController> lscLocalServices)
        {
            foreach (var scCurrentService in lscLocalServices)
            {
                try
                {
                    LogInfo($"-» Start service {scCurrentService.ServiceName}");
                    scCurrentService.Refresh();

                    if (scCurrentService.Status == ServiceControllerStatus.Stopped)
                        scCurrentService.Start();

                    scCurrentService.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(ciDVBViewerServiceTimeoutSeconds));
                }
                catch (Exception excCurrentException) when (excCurrentException is InvalidOperationException or Win32Exception or System.ServiceProcess.TimeoutException)
                {
                    LogError($"-» Error on starting service {scCurrentService.ServiceName}: {excCurrentException.Message}");
                }
                finally
                {
                    scCurrentService.Dispose();
                }
            }
        }

        /// <summary>
        /// DataInstallations.NormalizeDirectory() - removes quotes, trims and returns the full path without trailing directory separator
        /// </summary>
        /// <param name="strLocalDirectory"></param>
        /// <returns></returns>
        private static string NormalizeDirectory(string strLocalDirectory)
        {
            return Path.GetFullPath(strLocalDirectory.Trim().Trim('"')).TrimEnd(Path.DirectorySeparatorChar);
        }

        /// <summary>
        /// DataInstallations.IsSamePath() - compares two paths after normalization
        /// </summary>
        /// <param name="strLocalPath1"></param>
        /// <param name="strLocalPath2"></param>
        /// <returns></returns>
        private static bool IsSamePath(string strLocalPath1, string strLocalPath2)
        {
            return string.Equals(NormalizeDirectory(strLocalPath1), NormalizeDirectory(strLocalPath2), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// DataInstallations.IsInDirectory() - checks if a file is located in a directory (or its subdirectories) after normalization
        /// </summary>
        /// <param name="strLocalFile"></param>
        /// <param name="strLocalDirectory"></param>
        /// <returns></returns>
        private static bool IsInDirectory(string strLocalFile, string strLocalDirectory)
        {
            try
            {
                return Path.GetFullPath(strLocalFile).StartsWith(NormalizeDirectory(strLocalDirectory) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception excCurrentException) when (excCurrentException is ArgumentException or NotSupportedException or PathTooLongException)
            {
                return false;
            }
        }
    }
}