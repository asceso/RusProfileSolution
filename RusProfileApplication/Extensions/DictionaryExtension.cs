using System.Collections.Generic;

namespace RusProfileApplication.Extensions
{
    public static class DictionaryExtension
    {
        public static void AppendDictionary<TKey, TValue>(this Dictionary<TKey, TValue> target, Dictionary<TKey, TValue> source)
        {
            foreach (TKey key in source.Keys)
            {
                target.AddWithKey(key, source[key]);
            }
        }
        public static void AddWithKey<TKey, TValue>(this Dictionary<TKey, TValue> dic, TKey key, TValue value)
        {
            if (dic.ContainsKey(key))
            {
                dic[key] = value;
            }
            else
            {
                dic.Add(key, value);
            }
        }
    }
}
