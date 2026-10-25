# KON.Liwest.ChannelFactory
Fetching, parsing and creating LIWEST DVB-C channel data for various backends (Excel, DVBViewer, Enigma2)

## Usage

![KON.Liwest.ChannelFactory --help](docs/help.svg)

<details>
<summary>Help as text</summary>

```text
KON.Liwest.ChannelFactory
Copyright (C) 2024-2026 Oswald Oliver

  --mode                    Required. Execution-Mode or Operation to use: GenerateData (collect
                            source data), ModifyData (channel removal / sorting of a source data
                            file), ExportData (export a source data file), DumpData (output an
                            existing file), AutomaticData (collect, modify and output / export in
                            one run).
  --submode                 Execution-Sub-Mode of AutomaticData: SubModeCustom1 (no output),
                            SubModeCustom2 (output to --target), SubModeCustom3 (exports of
                            --exportmode).
  --source                  Required. Source of the data: ChannelList (LIWEST channel list),
                            SmartcardList (LIWEST signal check), BlindscanList (SAT>IP blindscan),
                            CombinedList (all three), FileList (source data file), CustomFile (file
                            for DumpData).
  --target                  Output of the source data: None, ConsoleList (table on the console),
                            FileList (source data file --targetfile).
  --smartcardid             Smartcard number for the LIWEST signal check (11 digits, starting with
                            01359 or 02766). Not available between 23:00 and 1:00 UTC.
  --serverip                IP Address of SAT>IP server for the blindscan.
  --frequencies             Frequencies in MHz for the blindscan, comma separated. Ranges as
                            start-end:step.
  --sourcefile              Source data file (json) to load from - for DumpData the file to output.
  --targetfile              Source data file (json) to save to (--target FileList). Relative paths
                            are placed in Export\<start of the run> next to the executable.
  --channelvariationfile    Channel variation file (json) - unifies the channel names of the
                            different sources.
  --sourcecleanup           Cleanup of the source data: RemoveInvalidSIDs (remove channels without
                            SID), GroupAndInvalidate (merge the sources per channel, remove channels
                            without stream).
  --channelremovalfile      Channel removal file (json) - channels (name and frequency) to remove.
  --channelsortingfile      Channel sorting file (json) - sort number per channel (name and
                            frequency), unsorted channels are placed in front.
  --exportmode              Exports, space separated (one entry per export): ExcelChannelList,
                            DVBViewerTransponderList, DVBViewerChannelDatabase,
                            EnigmaTransponderList, EnigmaChannelDatabase, EnigmaBouquetsFiles.
  --exportfilename          Filename per export, same order as --exportmode ("" for
                            EnigmaBouquetsFiles). Relative paths are placed in Export\<start of the
                            run> next to the executable.
  --exportfilenameformat    Enigma format per export, same order as --exportmode: Enigma1Ver1,
                            Enigma1Ver2 (Enigma 1), Enigma2Ver3, Enigma2Ver4 (default of current
                            images), Enigma2Ver5 (Enigma 2 lamedb version) - None for excel and
                            DVBViewer exports.
  --installdvbviewerdata    Install the exported DVBViewer channel database and transponder list
                            into all installed DVBViewer / DVBViewer Media Server instances
                            (ExportData or AutomaticData SubModeCustom3). Running instances are
                            stopped and restarted, existing files are backed up.
  --installenigmadata       Install the exported Enigma2 data into the receivers of --enigmahost via
                            SSH (ExportData or AutomaticData SubModeCustom3). The settings of the
                            receiver are backed up locally and merged: other bouquets (e.g.
                            streaming), other cables and foreign lamedb entries are kept, enigma2 is
                            stopped and restarted.
  --enigmahost              Enigma2 receivers for --installenigmadata, space separated (host or
                            host:port of the SSH server).
  --enigmauser              (Default: root) SSH user of the Enigma2 receivers.
  --enigmapassword          SSH password of the Enigma2 receivers (required for
                            --installenigmadata).
  --dryrun                  Installation test without changes: finds the DVBViewer installations /
                            connects to the Enigma2 receivers and backs up their settings, nothing
                            is installed.
  --keepexports             Number of run directories in Export next to the executable to keep
                            (including the current run) - older ones are removed. Without this
                            option all run directories are kept.
  --dumpdatamode            Format of the file to output with DumpData.
  --verbose                 Set output to verbose level.
  --help                    Display this help screen.
  --version                 Display version information.

* e.g.: Channel list to the console
    --mode GenerateData --source ChannelList --target ConsoleList

* e.g.: Channel list to a source data file
    --mode GenerateData --source ChannelList --target FileList
    --targetfile SourceData_ChannelList.json

* e.g.: Smartcard (LIWEST signal check) to a source data file
    --mode GenerateData --source SmartcardList --target FileList --smartcardid 01359xxxxxx
    --targetfile SourceData_SmartcardList.json

* e.g.: Blindscan (SAT>IP) to a source data file
    --mode GenerateData --source BlindscanList --target FileList --serverip 192.168.1.10
    --frequencies 114,130-858:8 --targetfile SourceData_BlindscanList.json

* e.g.: All sources merged per channel to a source data file
    --mode GenerateData --source CombinedList --target FileList --smartcardid 01359xxxxxx
    --serverip 192.168.1.10 --frequencies 114,130-858:8 --sourcecleanup GroupAndInvalidate
    --channelvariationfile .\resources\ChannelVariation.json
    --targetfile SourceData_CombinedList.json

* e.g.: Check a source data file and output it to the console
    --mode GenerateData --source FileList --target ConsoleList
    --sourcefile SourceData_CombinedList.json

* e.g.: Remove channels without SID from a source data file
    --mode GenerateData --source FileList --target FileList
    --sourcefile SourceData_CombinedList.json --sourcecleanup RemoveInvalidSIDs
    --targetfile SourceData_Cleaned.json

* e.g.: Merge the sources of a source data file per channel
    --mode GenerateData --source FileList --target FileList
    --sourcefile SourceData_CombinedList.json --sourcecleanup GroupAndInvalidate
    --channelvariationfile .\resources\ChannelVariation.json
    --targetfile SourceData_Grouped.json

* e.g.: Remove and sort channels of a source data file
    --mode ModifyData --source FileList --target FileList --sourcefile SourceData_Grouped.json
    --channelremovalfile .\resources\ChannelRemoval.json
    --channelsortingfile .\resources\ChannelSorting.json
    --targetfile SourceData_Sorted.json

* e.g.: Export to an excel channel list
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode ExcelChannelList --exportfilename channels.xlsx --exportfilenameformat None

* e.g.: Export to DVBViewer transponder list and channel database
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase
    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None

* e.g.: Export to DVBViewer and install into DVBViewer / Media Server
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase
    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None
    --installdvbviewerdata

* e.g.: Test the installation into DVBViewer without changes
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode DVBViewerTransponderList DVBViewerChannelDatabase
    --exportfilename DVB-C_at_LIWEST.ini channels.dat --exportfilenameformat None None
    --installdvbviewerdata --dryrun

* e.g.: Test the installation into an Enigma2 receiver (SSH, backup only)
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles
    --exportfilename cables.xml lamedb ""
    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4
    --installenigmadata --enigmahost 192.168.1.20 --enigmauser root
    --enigmapassword xxxxxx --dryrun

* e.g.: Export to Enigma2 and install into a receiver (SSH)
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles
    --exportfilename cables.xml lamedb ""
    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4
    --installenigmadata --enigmahost 192.168.1.20 --enigmauser root
    --enigmapassword xxxxxx

* e.g.: Export to Enigma2 transponders, channel database and bouquets
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode EnigmaTransponderList EnigmaChannelDatabase EnigmaBouquetsFiles
    --exportfilename cables.xml lamedb ""
    --exportfilenameformat Enigma2Ver4 Enigma2Ver4 Enigma2Ver4

* e.g.: Export to all formats into the folder export
    --mode ExportData --source FileList --sourcefile SourceData_Sorted.json
    --exportmode ExcelChannelList DVBViewerTransponderList DVBViewerChannelDatabase
    EnigmaTransponderList EnigmaChannelDatabase
    --exportfilename export\channels.xlsx export\DVB-C_at_LIWEST.ini
    export\channels.dat export\cables.xml export\lamedb
    --exportfilenameformat None None None Enigma2Ver4 Enigma2Ver4

* e.g.: Output a DVBViewer channel database to the console
    --mode DumpData --source CustomFile --dumpdatamode DVBViewerChannelDatabase
    --sourcefile channels.dat

* e.g.: Automatic run without output (collect, cleanup, modify only)
    --mode AutomaticData --submode SubModeCustom1 --source CombinedList
    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8
    --sourcecleanup GroupAndInvalidate

* e.g.: Automatic run with output to a source data file
    --mode AutomaticData --submode SubModeCustom2 --source CombinedList
    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8
    --sourcecleanup GroupAndInvalidate
    --channelvariationfile .\resources\ChannelVariation.json
    --channelremovalfile .\resources\ChannelRemoval.json
    --channelsortingfile .\resources\ChannelSorting.json
    --target FileList --targetfile SourceData_Automatic.json

* e.g.: Automatic run with exports to excel, DVBViewer and Enigma2
    --mode AutomaticData --submode SubModeCustom3 --source CombinedList
    --smartcardid 01359xxxxxx --serverip 192.168.1.10 --frequencies 114,130-858:8
    --sourcecleanup GroupAndInvalidate
    --channelvariationfile .\resources\ChannelVariation.json
    --channelremovalfile .\resources\ChannelRemoval.json
    --channelsortingfile .\resources\ChannelSorting.json
    --exportmode ExcelChannelList DVBViewerChannelDatabase EnigmaChannelDatabase
    EnigmaBouquetsFiles --exportfilename channels.xlsx channels.dat lamedb ""
    --exportfilenameformat None None Enigma2Ver4 Enigma2Ver4

* e.g.: Channel list with verbose (trace) output
    --mode GenerateData --source ChannelList --target None --verbose
```

</details>

## Modes

| Mode | Description |
|---|---|
| `GenerateData` | Collects source data from the LIWEST channel list, the smartcard check, a blindscan (SAT>IP) or a source data file, optionally cleans it up and writes it to the console or a source data file |
| `ModifyData` | Loads a source data file, removes channels (`--channelremovalfile`) and sorts them (`--channelsortingfile`) |
| `ExportData` | Loads a source data file and exports it - several exports per run, `--exportmode`, `--exportfilename` and `--exportfilenameformat` are given as lists of the same length |
| `DumpData` | Prints the content of an existing file (e.g. DVBViewer channel database) |
| `AutomaticData` | Runs channel list, smartcard and blindscan, cleanup, removal and sorting in one go - `SubModeCustom2` writes the target, `SubModeCustom3` runs the exports |

## Sources and data

* Channel list: LIWEST channel list (tv and radio csv)
* Smartcard: LIWEST signal check api - not available between 23:00 and 1:00 UTC (maintenance window of the api, e.g. 1:00 to 3:00 in Austria during daylight saving time, 0:00 to 2:00 otherwise) - times in messages of the api are shown in the time zone of the system
* Blindscan: SAT>IP server via [KON.OctoScan.NET](https://github.com/h0ly0ne/KON.OctoScan.NET)
* Source data file: json array of channels with the property names of the source data (`ChannelName`, `ChannelFrequency`, `SID`, ...), not listed properties stay empty
* Frequencies are always stored in kHz - input values like `578`, `578000` or `578.00 MHz` are normalized
* `Sources` lists the sources of a channel (`ChannelList`, `SmartcardList`, `BlindscanList`) - after `--sourcecleanup GroupAndInvalidate` all merged sources
* `Package`: with `--sourcecleanup GroupAndInvalidate` the package of the smartcard wins, the channel list is only used if the smartcard delivers none
* `Package`: html of the smartcard api is converted to text - line breaks (`<br />`) become ` / `, other tags are removed, entities decoded

## Encryption

The blindscan delivers:

* `Encrypted` - decision of KON.OctoScan.NET: the measured scrambling bits win, without measurement the PMT CA_descriptor / SDT free_CA_mode is used
* `CASystemIDs` - CA_system_IDs of the PMT CA_descriptors (hex, comma separated)

Exports:

* DVBViewer: encrypted flag of the channel from `Encrypted`
* Enigma (lamedb): encrypted services get one `C:xxxx` entry per CA_system_ID, `C:0000` if no CA_system_ID is known - unencrypted services get none

Channels without `Encrypted` (e.g. no blindscan data or older source data files) are exported as not encrypted - regenerate older source data files with a current blindscan.
Channel list, smartcard and blindscan data of one channel are only combined with `--sourcecleanup GroupAndInvalidate`.

## Dolby

The blindscan delivers the codec per audio pid (`APIDSCODEC`: MPEG, AAC, AC3, EAC3, AC4, DTS, TRUEHD, UNKNOWN).
`Dolby` is true if at least one audio pid is AC3, EAC3 or AC4 - with `--sourcecleanup GroupAndInvalidate` the measured value wins over the channel list.
The DVBViewer export sets the audio format from the codec as well (stream type only if no codec is known).

## Codec

`Codec` contains the video codec (`MPEG2`, `H264`, `HEVC`, `VVC`, `AVS`, ...), for services without video the codec of the first audio pid (`MPEG`, `AAC`, `AC3`, ...).
The blindscan derives it from the video stream type (`VPIDST`) resp. the audio codecs (`APIDSCODEC`) - with `--sourcecleanup GroupAndInvalidate` this value wins over the smartcard.
Spellings of the smartcard are unified (`MPEG4` -> `H264`).

## Language

`Language` is always stored as 3 letter ISO 639-2/T code (e.g. `DE`, `ger` -> `deu`), several languages comma separated.
The blindscan sets it to the language of the first audio pid with a real language (placeholders like `und`, `mis` or `qaa` are skipped) - with `--sourcecleanup GroupAndInvalidate` this value wins over the channel list / smartcard.
Channels without any language get `und` (undetermined).

## Resource files

| File | Option | Content |
|---|---|---|
| `resources/ChannelVariation.json` | `--channelvariationfile` | Primary channel name and its variations (different spellings of the sources) |
| `resources/ChannelRemoval.json` | `--channelremovalfile` | Channels (name and frequency) to remove |
| `resources/ChannelSorting.json` | `--channelsortingfile` | Sort number per channel (name and frequency) - channels without sort number are placed in front |

Channel names are matched case-insensitive.

## Output files

Generated files (`--exportfilename`, `--targetfile`) with a relative path are placed in an export directory per run next to the executable:

```text
<directory of KON.Liwest.ChannelFactory.exe>\Export\<yyyyMMdd-HHmmss>\
```

* The timestamp is the start of the run - all files of a run are placed in the same directory
* Runs started in the same second get their own directory (`<yyyyMMdd-HHmmss>_2`, `_3`, ...) - marked by the hidden file `.channelfactory`
* `--keepexports <n>` keeps the newest n run directories (including the current run) and removes older ones - only directories named like a run are touched, without the option all run directories are kept
* Absolute paths are kept, missing directories are created
* Enigma bouquet files (`--exportmode EnigmaBouquetsFiles`) and the Enigma2 receiver backups (`EnigmaBackup`) are placed in the export directory of the run as well
* Input files (`--sourcefile`, resource files) are still resolved relative to the current working directory

## DVBViewer installation

`--installdvbviewerdata` (with `ExportData` or `AutomaticData` / `SubModeCustom3` and a DVBViewer export) installs the DVBViewer files exported in this run into all installed DVBViewer / DVBViewer Media Server instances:

* Installations are found by their uninstall entries (DVBViewer Pro, DVBViewer Media Server, DVBViewer Recording Service, DVBViewer GE)
* The data directory is defined by `usermode.ini` of the program directory: `UserMode` 0 = program directory, 1 = application data of the user, 2 = common application data (`C:\ProgramData`), each with `Root` (e.g. `CMUV\DVBViewer`)
* `channels.dat` is copied to the data directory, the transponder list to `<data directory>\Transponders`
* Existing files are backed up as `<file>.<yyyyMMdd-HHmmss>.bak`
* Running programs of the instance are closed (terminated if they do not close within 30 seconds) and restarted, running services (Media Server) are stopped and restarted
* Only files written successfully in this run are installed

## Enigma2 installation

`--installenigmadata` (with `ExportData` or `AutomaticData` / `SubModeCustom3` and an Enigma2 export) connects to the receivers of `--enigmahost` via SSH / SFTP:

* Credentials are passed with `--enigmauser` (default `root`) and `--enigmapassword` - no credentials are part of ChannelFactory
* The host key of a receiver is stored on first use in `%LOCALAPPDATA%\KON.Liwest.ChannelFactory\known_hosts`, a changed host key is refused
* All files of `/etc/enigma2` (except the program guide `epg.dat`) and the target files outside of it (e.g. `cables.xml`) are backed up to `EnigmaBackup\<host>_<yyyyMMdd-HHmmss>` (in the export directory of the run)
* Streaming services found in the backed up bouquets / lamedb are reported
* The export is merged with the backed up settings of the receiver:
   * `lamedb` / `lamedb5`: entries of the DVB-C namespace (`ffff0000`) are replaced, all other entries kept - `lamedb5` is written (consistent with `lamedb`) if the receiver uses it
   * `bouquets.tv` / `bouquets.radio`: kept (e.g. streaming bouquets), missing references to the exported bouquets added
   * `userbouquet.*`: replaced by the export
   * `cables.xml`: cables named `LIWEST...` replaced, other cables kept
   * `blacklist` / `whitelist`: kept (parental control of the receiver)
* The files are uploaded to temporary files (`<file>.channelfactory.tmp`) while enigma2 is still running - on an upload error they are removed and the receiver stays unchanged
* One SSH command chain stops enigma2 (`init 4`), moves the files to their targets, removes leftovers and always starts enigma2 again (`init 3`) - no script or other file is stored on the receiver, no round trip between stop and start
* A hangup of the SSH session is ignored by the chain (`trap '' HUP`), so it is completed and enigma2 started again also if the connection is lost
* If enigma2 does not stop within 60 seconds nothing is installed (temporary files removed, enigma2 kept running)
* Targets: `lamedb` / `lamedb5` and bouquet files to `/etc/enigma2`, `cables.xml` to `/etc/tuxbox/cables.xml` (symbolic links resolved)
* The merged files are kept locally in `EnigmaBackup\<host>_<yyyyMMdd-HHmmss>_install`
* The cable of the Enigma2 transponder export is always named `LIWEST` - the scan provider selection of the receiver is kept on every update

## Dry run

`--dryrun` together with `--installdvbviewerdata` / `--installenigmadata` only checks the installation: DVBViewer installations are found and the planned actions are shown, Enigma2 receivers are connected and backed up - nothing is changed.

## Enigma formats

`--exportfilenameformat` uses the names and versions of the EnigmaSettings library:

| Value | Format |
|---|---|
| `Enigma1Ver1` | Enigma 1, version 1 |
| `Enigma1Ver2` | Enigma 1, version 2 |
| `Enigma2Ver3` | Enigma 2, lamedb version 3 |
| `Enigma2Ver4` | Enigma 2, lamedb version 4 (default of current images) |
| `Enigma2Ver5` | Enigma 2, lamedb version 4 + lamedb5 |

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success (also for `--help` / `--version`) |
| `1` | Invalid or missing arguments |
| `2` | Operation failed - at least one error has been reported |
| `3` | Prerequisites not met (firewall rules could not be created - run as administrator on a Windows client) |

## Requirements

* Windows client, run as administrator (firewall rules are created like in KON.OctoScan.NET)
* .NET 10
