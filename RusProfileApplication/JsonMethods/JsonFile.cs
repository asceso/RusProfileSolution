using System;
using System.IO;
using System.Text;

namespace RusProfileApplication.JsonMethods
{
    public static class JsonFile
    {
        public static bool JsonCreate(string path, string buffer)
        {
            try
            {
                using StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8);
                writer.Write(buffer);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        public static string JsonRead(string path)
        {
            string result = string.Empty;
            try
            {
                using StreamReader reader = new StreamReader(path, Encoding.UTF8);
                result = reader.ReadToEnd();
            }
            catch (Exception)
            {
                result = string.Empty;
            }
            return result;
        }
    }
}
