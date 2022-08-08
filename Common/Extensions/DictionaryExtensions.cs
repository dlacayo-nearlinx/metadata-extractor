using System.Collections.Generic;

namespace MetadataExtractor.Common.Extensions
{
    public static class DictionaryExtensions 
    {
        /// <summary>
        /// Returns the default value for the key or a default value if it cannot be read.
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        /// <typeparam name="TValue"></typeparam>
        /// <param name="dictionary"></param>
        /// <param name="key"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue = default(TValue))
        {
            if (dictionary == null) return defaultValue;
            if (key == null) return defaultValue;

            TValue value;
            return dictionary.TryGetValue(key, out value) ? value : defaultValue;
        }  
    }
}