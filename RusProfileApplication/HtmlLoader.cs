using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace RusProfileApplication
{
    public class HtmlLoader
    {
        #region constants
        private const string URLEmptyError = "URL string is Empty";
        private const string DefaultProxyAddress = "localhost";
        private const int DefaultProxyPort = 8866;
        #endregion

        public static async Task<string> LoadPage(string URL, WebProxy proxy = null)
        {
            if (string.IsNullOrEmpty(URL))
            {
                throw new ArgumentException(URLEmptyError);
            }
            HttpWebRequest request = WebRequest.CreateHttp(URL);
            request.Method = nameof(HttpMethod.Get);
            proxy ??= new WebProxy(DefaultProxyAddress, DefaultProxyPort);
            request.Proxy = proxy;

            string result = string.Empty;
            try
            {
                WebResponse response = await request.GetResponseAsync();
                using StreamReader reader = new StreamReader(response.GetResponseStream());
                result = reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return result;
        }
    }
}
