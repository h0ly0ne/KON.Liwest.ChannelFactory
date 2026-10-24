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

        /// <summary>
        /// NestedJsonPropertyConverter.CanConvert() - Always returns true, as this converter is used for all types.
        /// </summary>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType) => true;

        /// <summary>
        /// NestedJsonPropertyConverter.ReadJson() - Reads a JSON object and maps its properties to the target object, using the target class as reference.
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="objectType"></param>
        /// <param name="existingValue"></param>
        /// <param name="serializer"></param>
        /// <returns></returns>
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

        /// <summary>
        /// NestedJsonPropertyConverter.WriteJson() - Writes a JSON object with all public readable properties of the target object, using the property names as keys.
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <param name="serializer"></param>
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

        /// <summary>
        /// NestedJsonPropertyConverter.GetReadMappings() - Gets the property mappings for reading, including the property info, token paths, and optional converter.
        /// </summary>
        /// <param name="tLocalObjectType"></param>
        /// <returns></returns>
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

        /// <summary>
        /// NestedJsonPropertyConverter.GetWriteMappings() - Gets the property mappings for writing, including only public readable properties that are not ignored.
        /// </summary>
        /// <param name="tLocalObjectType"></param>
        /// <returns></returns>
        private static PropertyInfo[] GetWriteMappings(Type tLocalObjectType)
        {
            return cdWriteMappingCache.GetOrAdd(tLocalObjectType, tCurrentObjectType => tCurrentObjectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(piCurrentProperty => piCurrentProperty.CanRead && piCurrentProperty.GetCustomAttribute<JsonIgnoreAttribute>() == null)
                .ToArray());
        }

        /// <summary>
        /// NestedJsonPropertyConverter.GetToken() - Gets a JToken from a JObject using a token path, which can be nested (e.g. "frequency.value") or flat (e.g. "frequency").
        /// </summary>
        /// <param name="joLocalJsonObject"></param>
        /// <param name="strLocalTokenPath"></param>
        /// <returns></returns>
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
        /// <summary>
        /// FrequencyJsonConverter.CanConvert() - Returns true for int and int? types, as this converter is used for frequency values.
        /// </summary>
        /// <param name="objectType"></param>
        /// <returns></returns>
        public override bool CanConvert(Type objectType) => objectType == typeof(int?) || objectType == typeof(int);

        /// <summary>
        /// FrequencyJsonConverter.ReadJson() - Reads a JSON token and converts it to a frequency in kHz, using the ToFrequencyKHz() extension method.
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="objectType"></param>
        /// <param name="existingValue"></param>
        /// <param name="serializer"></param>
        /// <returns></returns>
        public override object? ReadJson(JsonReader reader, Type objectType, object? existingValue, JsonSerializer serializer)
        {
            return reader.TokenType switch
            {
                JsonToken.Integer or JsonToken.Float => Convert.ToDecimal(reader.Value, CultureInfo.InvariantCulture).ToFrequencyKHz(),
                JsonToken.String => Convert.ToString(reader.Value, CultureInfo.InvariantCulture).ToFrequencyKHz(),
                _ => null
            };
        }

        /// <summary>
        /// FrequencyJsonConverter.WriteJson() - Writes a JSON token with the frequency value in kHz, using the writer.WriteValue() method.
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <param name="serializer"></param>
        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer) => writer.WriteValue(value);
    }
}