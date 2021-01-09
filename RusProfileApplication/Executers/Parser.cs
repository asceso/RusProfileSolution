using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Media;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using CapMonsterCloud.Models.CaptchaTasks;
using CapMonsterCloud.Models.CaptchaTasksResults;
using HtmlAgilityPack;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using RusProfileApplication.Extensions;
using RusProfileApplication.JsonMethods;
using RusProfileApplication.Models;

namespace RusProfileApplication.Executers
{
    public class Parser
    {
        #region maybe
        //HtmlDocument document;
        //if (useCustomLoad)
        //{
        //    string page = await HtmlLoader.LoadPage(URL);
        //    document = new HtmlDocument();
        //    document.LoadHtml(page);
        //}
        //else
        //{
        //    HtmlWeb web = new HtmlWeb();
        //    document = web.Load(URL);
        //}


        //var rootNode = document.DocumentNode.SelectNodes("//div[@class='hidden-parent']"); 


        //driver.SwitchTo().Frame(0);
        //driver.FindElement(By.XPath("//span[@id='recaptcha-anchor']")).Click();
        //Thread.Sleep(1000);

        //driver.FindElement(By.Id("solver-button")).Click();
        //Thread.Sleep(3000);
        //driver.SwitchTo().DefaultContent();
        #endregion
        private readonly string CapMonsterKey;

        private bool PagesFounded = false;
        private readonly int BaseSleepLoading = 300;

        public Parser(string capthca_key)
        {
            CapMonsterKey = capthca_key;
        }
        public IWebDriver CreateDriver(string extension_directory, AuthCredetials credentials, Proxy proxy = null)
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddExtension(extension_directory + "\\Buster.crx");
            options.AddExtension(extension_directory + "\\VPN.crx");
            options.AddExtension(extension_directory + "\\Add.crx");
            if (proxy != null)
            {
                options.AddArguments($"proxy-server={proxy.HttpProxy}");
            }
            IWebDriver driver = new ChromeDriver(options);
            Console.WriteLine("If you ready type anything");
            Console.ReadKey();
            driver.Url = "https://www.rusprofile.ru/";
            Console.Clear();
            if (proxy == null)
            {
                Console.WriteLine("Used proxy : No proxy");
            }
            else
            {
                Console.WriteLine("Used proxy : " + proxy.HttpProxy);
            }
            driver.FindElement(By.Id("tologin")).Click();
            Thread.Sleep(200);

            driver.FindElement(By.Id("mw-l_mail")).SendKeys(credentials.UserName);
            driver.FindElement(By.Id("mw-l_pass")).SendKeys(credentials.Password);

            driver.FindElement(By.Id("mw-l_entrance")).Click();

            return driver;
        }
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
        public async Task<Dictionary<string, string>> ParseUrls(IWebDriver driver, string baseURL)
        {
            SoundPlayer player = new SoundPlayer(Environment.CurrentDirectory + "\\Speech On.wav");
            Dictionary<string, string> result = new Dictionary<string, string>();
            PagesFounded = false;
            int JsonCounter = 1;
            int page = 1;
            int max_page = 0;
            do
            {
            captcha_link:
                driver.Url = baseURL + page;

                if (page % 50 == 0)
                {
                    string OutDir = ConfigurationManager.AppSettings["OutputDirectory"];
                    string json = JsonConvert.SerializeObject(result);
                    JsonFile.JsonCreate(OutDir + "\\" + baseURL.Replace("https://www.rusprofile.ru/codes/", string.Empty).Replace("/sankt-peterburg/", string.Empty) + $"_part{JsonCounter++}" + ".json", json);
                    result.Clear();
                }

                //Thread.Sleep(BaseSleepLoading);

                if (driver.PageSource.Contains("«Я не робот»."))
                {
                    SolveCaptcha(driver, baseURL + page);

                    goto captcha_link;
                }
                if (!PagesFounded)
                {
                    try
                    {
                        IWebElement counter = driver.FindElement(By.ClassName("page-navigation__num"));
                        if (counter != null)
                        {
                            int size = Convert.ToInt32(counter.Text.Replace("из ", string.Empty));
                            int more_100 = (size - (size / 100)) > 0 ? 1 : 0;
                            max_page = (size / 100) + more_100;
                            PagesFounded = true;
                        }
                    }
                    catch
                    {
                        max_page = 1;
                        PagesFounded = true;
                    }
                }
                System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> titles = driver.FindElements(By.XPath("//div[@class='company-item__title']//a"));
                foreach (IWebElement title in titles)
                {
                    string href = title.GetAttribute("href");
                    result.AddWithKey(title.Text, href);
                }
                page++;
            }
            while (page <= max_page);

            return result;
        }
        public async Task<CompanyCard> ParseCard(IWebDriver driver, string URL)
        {
            SoundPlayer player = new SoundPlayer(Environment.CurrentDirectory + "\\Speech On.wav");

        forward:
            driver.Url = URL;
            Thread.Sleep(BaseSleepLoading);

            if (driver.PageSource.Contains("«Я не робот»."))
            {
                SolveCaptcha(driver, URL);

                goto forward;
            }

            CompanyCard card = new CompanyCard();

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> showHidden = driver.FindElements(By.XPath("(//button[@class='all-text-link js-hidden-text-opener'])"));
            Regex checkMore = new Regex(@"[Еще ]{4}[\d.]");
            try
            {
                foreach (IWebElement elem in showHidden)
                {
                    if (checkMore.IsMatch(elem.Text))
                    {
                        elem.Click();
                        Thread.Sleep(100);
                    }
                }
            }
            catch (Exception)
            {

            }

            card.ShortName = driver.FindElement(By.ClassName("company-header__row")).Text.Trim();

            card.FullName = driver.FindElement(By.ClassName("company-name")).Text.Trim();

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> phones = driver.FindElements(By.XPath("//span[contains(@class,'company-info__contact phone')]"));
            foreach (IWebElement phone in phones)
            {
                string editedPhone = phone.Text.Replace("Телефон\r\n", string.Empty);
                card.Phones += editedPhone + Environment.NewLine;
            }
            card.Phones = phones.Count > 0 ? card.Phones.Trim() : "нет данных";

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> mails = driver.FindElements(By.XPath("//span[contains(@class,'company-info__contact mail')]"));
            foreach (IWebElement mail in mails)
            {
                string editedMail = mail.Text.Replace("Электронная почта\r\n", string.Empty);
                card.Mails += editedMail + Environment.NewLine;
            }
            card.Mails = mails.Count > 0 ? card.Mails.Trim() : "нет данных";

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> sites = driver.FindElements(By.XPath("//span[contains(@class,'company-info__contact site')]"));
            foreach (IWebElement site in sites)
            {
                string editedSite = site.Text.Replace("Сайт\r\n", string.Empty);
                card.Sites += editedSite + Environment.NewLine;
            }
            card.Sites = sites.Count > 0 ? card.Sites.Trim() : "нет данных";

            try
            {
                IWebElement inn = driver.FindElement(By.XPath("(//dd[@class='company-info__text has-copy'])[2]"));
                IWebElement kpp = driver.FindElement(By.XPath("(//dd[@class='company-info__text has-copy'])[3]"));
                card.INN = inn.Text + " / " + kpp.Text;
            }
            catch (NoSuchElementException)
            {
                card.INN = "нет данных";
            }

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> primary_occupation = driver.FindElements(By.ClassName("company-info__text"));
            Regex regex = new Regex(@"\D*[(]{1}[\d]*[.]{1}[\d.]*[)]{1}");
            foreach (IWebElement info in primary_occupation)
            {
                if (regex.IsMatch(info.Text.Trim()))
                {
                    card.PrimaryOccupation = info.Text.Trim();
                    break;
                }
            }
            if (card.PrimaryOccupation == null)
            {
                player.Play();
            }

            string organization_status = driver.FindElement(By.ClassName("company-status")).Text;
            card.OrganizationStatus = organization_status;

            return card;
        }
        public void SolveCaptcha(IWebDriver driver, string URL)
        {
            string dsKey = driver.FindElement(By.XPath("//p[@class='g-recaptcha']")).GetAttribute("data-sitekey");

            CapMonsterCloud.CapMonsterClient antcap = new CapMonsterCloud.CapMonsterClient(CapMonsterKey);

        solving:
            try
            {
                NoCaptchaTask captchaTask = new NoCaptchaTask()
                {
                    WebsiteUrl = URL,
                    WebsiteKey = dsKey
                };
                int taskId = antcap.CreateTaskAsync(captchaTask).Result;
                Console.Write($"TaskCreated...\t");
                NoCaptchaTaskResult taskResult = antcap.GetTaskResultAsync<NoCaptchaTaskResult>(taskId).Result;
                Console.WriteLine($"Response retrieved");

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("document.getElementById('g-recaptcha-response').style.removeProperty('display');");
                IWebElement input = driver.FindElement(By.Id("g-recaptcha-response"));
                input.SendKeys(taskResult.GRecaptchaResponse);
                input.Submit();
            }
            catch (Exception)
            {
                goto solving;
            }
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
