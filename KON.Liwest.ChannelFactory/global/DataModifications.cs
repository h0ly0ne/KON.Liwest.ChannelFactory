using static KON.Liwest.ChannelFactory.Global;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Cleanup, removal and sorting of source data
    /// </summary>
    public static class DataModifications
    {
        /// <summary>
        /// DataModifications.LoadChannelVariationDictionaryFromFile()
        /// </summary>
        /// <param name="strLocalChannelVariationDictionaryFilename"></param>
        /// <returns>bool - false if the file could not be loaded</returns>
        public static bool LoadChannelVariationDictionaryFromFile(string? strLocalChannelVariationDictionaryFilename)
        {
            if (string.IsNullOrEmpty(strLocalChannelVariationDictionaryFilename))
                return false;

            LogTrace("DataModifications.LoadChannelVariationDictionaryFromFile()");
            LogInfo("-» Load string variation dictionary from file");

            var cvdCurrentChannelVariationDictionary = LoadJsonFile<ChannelVariationDictionary>(strLocalChannelVariationDictionaryFilename, "channel variation file");
            if (cvdCurrentChannelVariationDictionary == null)
                return false;

            SetChannelVariationDictionary(cvdCurrentChannelVariationDictionary);
            return true;
        }

        /// <summary>
        /// DataModifications.CleanupSourceData()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="fscLocalFlagSourceCleanup"></param>
        public static void CleanupSourceData(List<SourceDataItem> lsdiLocalSourceData, FlagSourceCleanup fscLocalFlagSourceCleanup)
        {
            if (fscLocalFlagSourceCleanup == FlagSourceCleanup.None)
                return;

            LogTrace("DataModifications.CleanupSourceData()");
            LogInfo("-» Cleanup sourcedata by defined operation");

            switch (fscLocalFlagSourceCleanup)
            {
                case FlagSourceCleanup.RemoveInvalidSIDs:
                    CleanupSourceData_RemoveInvalidSIDs(lsdiLocalSourceData);
                    break;
                case FlagSourceCleanup.GroupAndInvalidate:
                    CleanupSourceData_GroupAndInvalidate(lsdiLocalSourceData);
                    CleanupSourceData_RemoveInvalidSIDs(lsdiLocalSourceData);
                    break;
            }
        }

        /// <summary>
        /// DataModifications.CleanupSourceData_RemoveInvalidSIDs()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        public static void CleanupSourceData_RemoveInvalidSIDs(List<SourceDataItem> lsdiLocalSourceData)
        {
            LogTrace("DataModifications.CleanupSourceData_RemoveInvalidSIDs()");
            LogInfo("-» Cleanup sourcedata by removing invalid SIDs");

            if (lsdiLocalSourceData.Count == 0)
                return;

            var iCurrentItemCount = lsdiLocalSourceData.Count;
            lsdiLocalSourceData.RemoveAll(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.SID == null);

            if (lsdiLocalSourceData.Count == 0)
                LogWarning($"-» All {iCurrentItemCount} channels have been removed as none of them contains a SID (blindscan data required)");
        }

        /// <summary>
        /// DataModifications.CleanupSourceData_GroupAndInvalidate()
        /// Channels with the same name (case-insensitive) and frequency are merged - the first non-empty value of every field wins.
        /// Channels without valid video or audio stream are removed.
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        public static void CleanupSourceData_GroupAndInvalidate(List<SourceDataItem> lsdiLocalSourceData)
        {
            LogTrace("DataModifications.CleanupSourceData_GroupAndInvalidate()");
            LogInfo("-» Cleanup sourcedata by grouping and invalidating");

            if (lsdiLocalSourceData.Count == 0)
                return;

            var lsdiCurrentMergedSourceData = new List<SourceDataItem>();

            foreach (var igCurrentGroup in lsdiLocalSourceData.GroupBy(sdiCurrentSourceDataItem => ChannelKey.Create(sdiCurrentSourceDataItem.ChannelName, sdiCurrentSourceDataItem.ChannelFrequency)))
            {
                var sdiCurrentMergedSourceDataItem = new SourceDataItem();

                foreach (var sdiCurrentSourceDataItem in igCurrentGroup)
                    sdiCurrentMergedSourceDataItem.FillEmptyValuesFrom(sdiCurrentSourceDataItem);

                // Dolby, codec and language measured by the blindscan win over the channel list / smartcard
                if (sdiCurrentMergedSourceDataItem.DolbyFromAudioCodecs is { } bCurrentDolby)
                    sdiCurrentMergedSourceDataItem.Dolby = bCurrentDolby;

                if (sdiCurrentMergedSourceDataItem.CodecFromBlindscan is { } strCurrentCodec)
                    sdiCurrentMergedSourceDataItem.Codec = strCurrentCodec;

                // Language: real audio language of the blindscan, otherwise the first real language of the group, otherwise undetermined (Language never is empty and thus not filled by FillEmptyValuesFrom)
                sdiCurrentMergedSourceDataItem.Language = sdiCurrentMergedSourceDataItem.LanguageFromAudioLanguages ?? igCurrentGroup.Select(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.Language).FirstOrDefault(strCurrentLanguage => strCurrentLanguage != null && !strCurrentLanguage.IsPlaceholderLanguageCode());

                // Package of the smartcard wins over the channel list (channel list only as fallback)
                if (igCurrentGroup.FirstOrDefault(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.IsFromSmartcard && !string.IsNullOrWhiteSpace(sdiCurrentSourceDataItem.Package))?.Package is { } strCurrentPackage)
                    sdiCurrentMergedSourceDataItem.Package = strCurrentPackage;

                // Sources of all merged channels
                var saCurrentSources = igCurrentGroup.SelectMany(sdiCurrentSourceDataItem => (sdiCurrentSourceDataItem.Sources ?? string.Empty).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                sdiCurrentMergedSourceDataItem.Sources = saCurrentSources.Length > 0 ? string.Join(",", saCurrentSources) : null;

                if (sdiCurrentMergedSourceDataItem.HasValidStream)
                    lsdiCurrentMergedSourceData.Add(sdiCurrentMergedSourceDataItem);
            }

            lsdiLocalSourceData.Clear();
            lsdiLocalSourceData.AddRange(lsdiCurrentMergedSourceData);
        }

        /// <summary>
        /// DataModifications.ModifyDataRemoveChannels()
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalChannelRemovalFile"></param>
        public static void ModifyDataRemoveChannels(List<SourceDataItem> lsdiLocalSourceData, string? strLocalChannelRemovalFile)
        {
            if (string.IsNullOrEmpty(strLocalChannelRemovalFile))
                return;

            LogTrace("DataModifications.ModifyDataRemoveChannels()");
            LogInfo("-» Remove channels by channelremoval file");

            var crdaCurrentChannelRemovalDictionary = LoadJsonFile<ChannelRemovalDictionary[]>(strLocalChannelRemovalFile, "channel removal file");
            if (crdaCurrentChannelRemovalDictionary is not { Length: > 0 } || lsdiLocalSourceData.Count == 0)
                return;

            var hsckCurrentRemovedChannels = crdaCurrentChannelRemovalDictionary.OfType<ChannelRemovalDictionary>().Select(crdCurrentEntry => ChannelKey.Create(crdCurrentEntry.ChannelName, crdCurrentEntry.ChannelFrequency.ToFrequencyKHz())).ToHashSet();
            lsdiLocalSourceData.RemoveAll(sdiCurrentSourceDataItem => hsckCurrentRemovedChannels.Contains(ChannelKey.Create(sdiCurrentSourceDataItem.ChannelName, sdiCurrentSourceDataItem.ChannelFrequency)));
        }

        /// <summary>
        /// DataModifications.ModifyDataSortChannels()
        /// Channel numbers are replaced by the sort numbers of the sorting file, unsorted channels (without sort number) are placed in front of the sorted ones.
        /// </summary>
        /// <param name="lsdiLocalSourceData"></param>
        /// <param name="strLocalChannelSortingFile"></param>
        public static void ModifyDataSortChannels(List<SourceDataItem> lsdiLocalSourceData, string? strLocalChannelSortingFile)
        {
            if (string.IsNullOrEmpty(strLocalChannelSortingFile))
                return;

            LogTrace("DataModifications.ModifyDataSortChannels()");
            LogInfo("-» Sorting channels by channelsorting file");

            var csdaCurrentChannelSortingDictionary = LoadJsonFile<ChannelSortingDictionary[]>(strLocalChannelSortingFile, "channel sorting file");
            if (csdaCurrentChannelSortingDictionary is not { Length: > 0 } || lsdiLocalSourceData.Count == 0)
                return;

            // First entry wins if a channel is listed more than once
            var dCurrentSortNumbers = new Dictionary<ChannelKey, int?>();
            foreach (var csdCurrentEntry in csdaCurrentChannelSortingDictionary.OfType<ChannelSortingDictionary>())
                dCurrentSortNumbers.TryAdd(ChannelKey.Create(csdCurrentEntry.ChannelName, csdCurrentEntry.ChannelFrequency.ToFrequencyKHz()), csdCurrentEntry.SortNumber.ToNullableInt32());

            foreach (var sdiCurrentSourceDataItem in lsdiLocalSourceData)
                sdiCurrentSourceDataItem.ChannelNumber = dCurrentSortNumbers.GetValueOrDefault(ChannelKey.Create(sdiCurrentSourceDataItem.ChannelName, sdiCurrentSourceDataItem.ChannelFrequency));

            var lsdiCurrentSortedSourceData = lsdiLocalSourceData.OrderBy(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelNumber.HasValue ? 1 : 0).ThenBy(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelNumber).ThenBy(sdiCurrentSourceDataItem => sdiCurrentSourceDataItem.ChannelName).ToList();

            lsdiLocalSourceData.Clear();
            lsdiLocalSourceData.AddRange(lsdiCurrentSortedSourceData);
        }
    }
}