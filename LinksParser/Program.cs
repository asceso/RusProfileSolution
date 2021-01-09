using System;
using System.Collections.Generic;
using System.Configuration;
using RusProfileApplication.Executers;

namespace LinksParser
{
    internal class Program
    {
        private static readonly string OutDir = ConfigurationManager.AppSettings["OutputDirectory"];
        private static readonly string CaptchaApi = ConfigurationManager.AppSettings["CaptchaApiKey"];
        static void Main(string[] args)
        {
            //string BaseURL = string.Empty;
            string BaseURL = "https://www.rusprofile.ru/codes/411000/respublika-adygeya";
            foreach (string arg in args)
            {
                if (arg.Contains("BaseURL ="))
                {
                    BaseURL = arg.Replace("BaseURL = ", string.Empty);
                }
            }
            if (string.IsNullOrEmpty(BaseURL))
            {
                Console.WriteLine("Bad parameters");
                Console.ReadKey();
                Environment.Exit(0);
            }
            Parser parser = new Parser(CaptchaApi);
            Dictionary<string, string> links = parser.ParseUrls(BaseURL).Result;
        }
    }
}
