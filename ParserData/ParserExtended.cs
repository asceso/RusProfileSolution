using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace ParserData
{
    public class ParserExtended
    {
        private bool PagesFounded;
        private int BaseSleepLoading = 300;
        public async Task<Dictionary<string, string>> ParseUrls(string baseURL)
        {
            SoundPlayer player = new SoundPlayer(Environment.CurrentDirectory + "\\Speech On.wav");
            Dictionary<string, string> result = new Dictionary<string, string>();
            HtmlDocument document;
            HtmlWeb web = new HtmlWeb();
            PagesFounded = false;
            int page = 1;
            int max_page = 0;

            do
            {
            captcha_link:
                document = web.Load(baseURL + page);
                Thread.Sleep(BaseSleepLoading);

                if (document.Text.Contains("«Я не робот»."))
                {
                    player.Play();
                    Console.WriteLine("CAPTCHA");

                    string dsKey = document.DocumentNode.SelectSingleNode("//p[@class='g-recaptcha']").Attributes["data-sitekey"].Value;

                    CapMonsterCloud.CapMonsterClient antcap = new CapMonsterCloud.CapMonsterClient(CapMonsterKey);

                    NoCaptchaTask captchaTask = new NoCaptchaTask()
                    {
                        WebsiteUrl = baseURL + page,
                        WebsiteKey = dsKey
                    };

                    decimal balance = antcap.GetBalanceAsync().Result;
                    int taskId = antcap.CreateTaskAsync(captchaTask).Result;
                    NoCaptchaTaskResult taskResult = antcap.GetTaskResultAsync<NoCaptchaTaskResult>(taskId).Result;

                    HtmlNode input = document.DocumentNode.SelectSingleNode("//input[@id='recaptcha-token']");

                    goto captcha_link;
                }
                if (!PagesFounded)
                {
                    int size = Convert.ToInt32(document.DocumentNode.SelectSingleNode("//span[@class='page-navigation__num']")
                        .InnerText.Replace("из ", string.Empty));
                    int more_100 = (size - (size / 100)) > 0 ? 1 : 0;
                    max_page = (size / 100) + more_100;
                    PagesFounded = true;
                }
                HtmlNodeCollection titles = document.DocumentNode.SelectNodes("//div[@class='company-item__title']/a");
                foreach (HtmlNode title in titles)
                {
                    HtmlAttribute href = title.Attributes["href"];
                    result.AddWithKey(title.InnerText.Trim(), href.Value);
                }
                Console.WriteLine($"Page {page} parsed");
                page++;
            }
            while (page <= max_page);

            return result;
        }
        private string TryGetInnerText(HtmlNode node, string XPath)
        {
            HtmlNode parsed_value = node.SelectSingleNode(XPath);
            if (parsed_value != null)
            {
                return parsed_value.InnerText.Trim();
            }
            return string.Empty;
        }
        private string TryGetAttributeValue(HtmlNode node, string attrib_name, string XPath)
        {
            HtmlAttribute attribute = node.SelectSingleNode(XPath)
                                          .Attributes
                                          .FirstOrDefault(a => a.Name.Contains(attrib_name));

            if (attribute != null)
            {
                return attribute.Value;
            }
            return string.Empty;
        }
    }
}
