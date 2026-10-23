using System.Collections.Concurrent;
using System.Reflection;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Reads an object property by property, using the target class as reference:
    /// - a [JsonProperty] name is resolved as (nested) token path, e.g. "frequency.value"
    /// - the property name itself is always used as fallback (case-insensitive)
    /// Thus API responses (nested names) and saved source data files (flat property names) can both be read without listing every field.
    /// </summary>
    public class NestedJsonPropertyConverter : JsonConverter
    {
        private static readonly ConcurrentDictionary<Type, (PropertyInfo piProperty, string[] saTokenPaths)[]> cdPropertyMappingCache = new();

        public override bool CanConvert(Type objectType) => true;
        public override bool CanWrite => false;

        public override object? ReadJson(JsonReader reader, Type objectType, object? existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null)
                return null;

            var joCurrentJsonObject = JObject.Load(reader);
            var oCurrentTargetObject = existingValue ?? Activator.CreateInstance(objectType)!;

            foreach (var (piCurrentProperty, saCurrentTokenPaths) in GetPropertyMappings(objectType))
            {
                var jtCurrentToken = saCurrentTokenPaths.Select(strCurrentTokenPath => GetToken(joCurrentJsonObject, strCurrentTokenPath)).FirstOrDefault(jtCurrentToken => jtCurrentToken is not null && jtCurrentToken.Type is not (JTokenType.Null or JTokenType.Undefined));
                if (jtCurrentToken == null)
                    continue;

                try
                {
                    piCurrentProperty.SetValue(oCurrentTargetObject, jtCurrentToken.ToObject(piCurrentProperty.PropertyType, serializer));
                }
                catch (Exception excCurrentException) when (excCurrentException is JsonException or FormatException or InvalidCastException or ArgumentException)
                {
                    Global.LogTrace($"NestedJsonPropertyConverter.ReadJson() - Skipped value '{jtCurrentToken}' for property '{piCurrentProperty.Name}': {excCurrentException.Message}");
                }
            }

            return oCurrentTargetObject;
        }

        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer) => throw new NotImplementedException();

        private static (PropertyInfo piProperty, string[] saTokenPaths)[] GetPropertyMappings(Type tLocalObjectType)
        {
            return cdPropertyMappingCache.GetOrAdd(tLocalObjectType, tCurrentObjectType => tCurrentObjectType
                .GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(piCurrentProperty => piCurrentProperty.CanWrite && piCurrentProperty.GetCustomAttribute<JsonIgnoreAttribute>() == null)
                .Select(piCurrentProperty =>
                {
                    var strCurrentJsonPropertyName = piCurrentProperty.GetCustomAttribute<JsonPropertyAttribute>()?.PropertyName;
                    string[] saCurrentTokenPaths = string.IsNullOrEmpty(strCurrentJsonPropertyName) || strCurrentJsonPropertyName == piCurrentProperty.Name ? [piCurrentProperty.Name] : [strCurrentJsonPropertyName, piCurrentProperty.Name];

                    return (piCurrentProperty, saCurrentTokenPaths);
                })
                .ToArray());
        }

        private static JToken? GetToken(JObject joLocalJsonObject, string strLocalTokenPath)
        {
            return strLocalTokenPath.Contains('.') ? joLocalJsonObject.SelectToken(strLocalTokenPath) : joLocalJsonObject.GetValue(strLocalTokenPath, StringComparison.OrdinalIgnoreCase);
        }
    }
}
