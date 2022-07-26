using System.Collections.Generic;

namespace MetadataExtractor.Common.Extensions
{
    public static class DictionaryExtensions 
    {
        public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue = default(TValue))
        {
            if (dictionary == null) return defaultValue;
            if (key == null) return defaultValue;

            TValue value;
            return dictionary.TryGetValue(key, out value) ? value : defaultValue;
        }  
    }
}