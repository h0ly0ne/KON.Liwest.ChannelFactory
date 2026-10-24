using System.Collections.Concurrent;
using System.Globalization;
using System.Reflection;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KON.Liwest.ChannelFactory
{
    /// <summary>
    /// Reads an object property by property, using the target class as reference:
    /// - a [JsonProperty] name is resolved as (nested) token path, e.g. "frequency.value"
    /// - the property name itself is always used as fallback (case-insensitive)
    /// - a [JsonConverter] of the property is used for its value
    /// Thus API responses (nested names) and saved source data files (flat property names) can both be read without listing every field.
    /// Writing creates a flat object of all public readable properties (property names, declaration order).
    /// </summary>
    public class NestedJsonPropertyConverter : JsonConverter
    {
        private static readonly ConcurrentDictionary<Type, (PropertyInfo piProperty, string[] saTokenPaths, JsonConverter? jcConverter)[]> cdReadMappingCache = new();
        private static readonly ConcurrentDictionary<Type, PropertyInfo[]> cdWriteMappingCache = new();

        public override bool CanConvert(Type objectType) => true;

        public override object? ReadJson(JsonReader reader, Type objectType, object? existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null)
                return null;

            var joCurrentJsonObject = JObject.Load(reader);
            var oCurrentTargetObject = existingValue ?? Activator.CreateInstance(objectType)!;

            foreach (var (piCurrentProperty, saCurrentTokenPaths, jcCurrentConverter) in GetReadMappings(objectType))
            {
                var jtCurrentToken = saCurrentTokenPaths.Select(strCurrentTokenPath => GetToken(joCurrentJsonObject, strCurrentTokenPath)).FirstOrDefault(jtCurrentToken => jtCurrentToken is not null && jtCurrentToken.Type is not (JTokenType.Null or JTokenType.Undefined));
                if (jtCurrentToken == null)
                    continue;

                try
                {
                    object? oCurrentValue;

                    if (jcCurrentConverter != null)
                    {
                        using var jrCurrentTokenReader = jtCurrentToken.CreateReader();
                        jrCurrentTokenReader.Read();
                        oCurrentValue = jcCurrentConverter.ReadJson(jrCurrentTokenReader, piCurrentProperty.PropertyType, null, serializer);
                    }
                    else
                        oCurrentValue = jtCurrentToken.ToObject(piCurrentProperty.PropertyType, serializer);

                    piCurrentProperty.SetValue(oCurrentTargetObject, oCurrentValue);
                }
                catch (Exception excCurrentException) when (excCurrentException is JsonException or FormatException or InvalidCastException or ArgumentException or OverflowException)
                {
                    Global.LogTrace($"NestedJsonPropertyConverter.ReadJson() - Skipped value '{jtCurrentToken}' for property '{piCurrentProperty.Name}': {excCurrentException.Message}");
                }
            }

            return oCurrentTargetObject;
        }

        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer)
        {
            if (value == null)
            {
                writer.WriteNull();
                return;
            }

            writer.WriteStartObject();

            foreach (var piCurrentProperty in GetWriteMappings(value.GetType()))
            {
                writer.WritePropertyName(piCurrentProperty.Name);
                serializer.Serialize(writer, piCurrentProperty.GetValue(value));
            }

            writer.WriteEndObject();
        }

        private static (PropertyInfo piProperty, string[] saTokenPaths, JsonConverter? jcConverter)[] GetReadMappings(Type tLocalObjectType)
        {
            return cdReadMappingCache.GetOrAdd(tLocalObjectType, tCurrentObjectType => tCurrentObjectType
                .GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(piCurrentProperty => piCurrentProperty.CanWrite && piCurrentProperty.GetCustomAttribute<JsonIgnoreAttribute>() == null)
                .Select(piCurrentProperty =>
                {
                    var strCurrentJsonPropertyName = piCurrentProperty.GetCustomAttribute<JsonPropertyAttribute>()?.PropertyName;
                    string[] saCurrentTokenPaths = string.IsNullOrEmpty(strCurrentJsonPropertyName) || strCurrentJsonPropertyName == piCurrentProperty.Name ? [piCurrentProperty.Name] : [strCurrentJsonPropertyName, piCurrentProperty.Name];
                    var tCurrentConverterType = piCurrentProperty.GetCustomAttribute<JsonConverterAttribute>()?.ConverterType;
                    var jcCurrentConverter = tCurrentConverterType != null ? (JsonConverter?)Activator.CreateInstance(tCurrentConverterType) : null;

                    return (piCurrentProperty, saCurrentTokenPaths, jcCurrentConverter);
                })
                .ToArray());
        }

        private static PropertyInfo[] GetWriteMappings(Type tLocalObjectType)
        {
            return cdWriteMappingCache.GetOrAdd(tLocalObjectType, tCurrentObjectType => tCurrentObjectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(piCurrentProperty => piCurrentProperty.CanRead && piCurrentProperty.GetCustomAttribute<JsonIgnoreAttribute>() == null)
                .ToArray());
        }

        private static JToken? GetToken(JObject joLocalJsonObject, string strLocalTokenPath)
        {
            return strLocalTokenPath.Contains('.') ? joLocalJsonObject.SelectToken(strLocalTokenPath) : joLocalJsonObject.GetValue(strLocalTokenPath, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// Reads a frequency given as number or text (e.g. 578000, 578, "578.00 MHz") and normalizes it to kHz.
    /// </summary>
    public class FrequencyJsonConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType) => objectType == typeof(int?) || objectType == typeof(int);

        public override object? ReadJson(JsonReader reader, Type objectType, object? existingValue, JsonSerializer serializer)
        {
            return reader.TokenType switch
            {
                JsonToken.Integer or JsonToken.Float => Convert.ToDecimal(reader.Value, CultureInfo.InvariantCulture).ToFrequencyKHz(),
                JsonToken.String => Convert.ToString(reader.Value, CultureInfo.InvariantCulture).ToFrequencyKHz(),
                _ => null
            };
        }

        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer) => writer.WriteValue(value);
    }
}
