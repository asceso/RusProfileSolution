using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Media;
using OfficeOpenXml;
using OpenQA.Selenium;
using RusProfileApplication.Executers;
using RusProfileApplication.Models;

namespace RusProfileApplication
{
    internal class Program
    {
        private static readonly string OutDir = ConfigurationManager.AppSettings["OutputDirectory"];
        private static readonly string ExtDir = ConfigurationManager.AppSettings["ExtensionDirectory"];
        private static readonly string ExportedDir = ConfigurationManager.AppSettings["ExportedDirectory"];

        private static void Main(string[] args)
        {
            Proxy proxy = null;
            string LinksFileName = string.Empty;
            foreach (string arg in args)
            {
                if (arg.Contains("Links ="))
                {
                    LinksFileName = arg.Replace("Links = ", string.Empty).Trim();
                }
                if (arg.Contains("Proxy ="))
                {
                    proxy = new Proxy() { HttpProxy = arg.Replace("Proxy = ", string.Empty).Trim() };
                }
            }
            if (string.IsNullOrEmpty(LinksFileName))
            {
                LinksFileName = "Links2.xlsx";
            }
            SoundPlayer player = new SoundPlayer(Environment.CurrentDirectory + "\\Speech On.wav");
            Dictionary<string, string> urls = new Dictionary<string, string>()
            {
                //{"411000", "https://www.rusprofile.ru/codes/411000/sankt-peterburg/" } ,
                //{"412000", "https://www.rusprofile.ru/codes/412000/sankt-peterburg/" } ,
                //{"421100", "https://www.rusprofile.ru/codes/421100/sankt-peterburg/" } ,
                //{"421200", "https://www.rusprofile.ru/codes/421200/sankt-peterburg/" } ,
                //{"421300", "https://www.rusprofile.ru/codes/421300/sankt-peterburg/" } ,
                //{"422100", "https://www.rusprofile.ru/codes/422100/sankt-peterburg/" } ,
                //{"422200", "https://www.rusprofile.ru/codes/422200/sankt-peterburg/" } ,
                //{"422210", "https://www.rusprofile.ru/codes/422210/sankt-peterburg/" } ,
                //{"422220", "https://www.rusprofile.ru/codes/422220/sankt-peterburg/" } ,
                //{"422230", "https://www.rusprofile.ru/codes/422230/sankt-peterburg/" } ,
                //{"429100", "https://www.rusprofile.ru/codes/429100/sankt-peterburg/" } ,
                //{"429110", "https://www.rusprofile.ru/codes/429110/sankt-peterburg/" } ,
                //{"429120", "https://www.rusprofile.ru/codes/429120/sankt-peterburg/" } ,
                //{"429140", "https://www.rusprofile.ru/codes/429140/sankt-peterburg/" } ,
                //{"429150", "https://www.rusprofile.ru/codes/429150/sankt-peterburg/" } ,
                //{"429900", "https://www.rusprofile.ru/codes/429900/sankt-peterburg/" } ,
                //{"431100", "https://www.rusprofile.ru/codes/431100/sankt-peterburg/" } ,
                //{"431200", "https://www.rusprofile.ru/codes/431200/sankt-peterburg/" } ,
                //{"431210", "https://www.rusprofile.ru/codes/431210/sankt-peterburg/" } ,
                //{"431220", "https://www.rusprofile.ru/codes/431220/sankt-peterburg/" } ,
                //{"431230", "https://www.rusprofile.ru/codes/431230/sankt-peterburg/" } ,
                //{"431240", "https://www.rusprofile.ru/codes/431240/sankt-peterburg/" } ,
                //{"431300", "https://www.rusprofile.ru/codes/431300/sankt-peterburg/" } ,
                //{"432000", "https://www.rusprofile.ru/codes/432000/sankt-peterburg/" } ,
                //{"432100", "https://www.rusprofile.ru/codes/432100/sankt-peterburg/" } ,
                //{"432200", "https://www.rusprofile.ru/codes/432200/sankt-peterburg/" } ,
                //{"432900", "https://www.rusprofile.ru/codes/432900/sankt-peterburg/" } ,
                //{"433100", "https://www.rusprofile.ru/codes/433100/sankt-peterburg/" } ,
                //{"433200", "https://www.rusprofile.ru/codes/433200/sankt-peterburg/" } ,
                //{"433210", "https://www.rusprofile.ru/codes/433210/sankt-peterburg/" } ,
                //{"433220", "https://www.rusprofile.ru/codes/433220/sankt-peterburg/" } ,
                //{"433230", "https://www.rusprofile.ru/codes/433230/sankt-peterburg/" } ,
                //{"433300", "https://www.rusprofile.ru/codes/433300/sankt-peterburg/" } ,
                //{"433400", "https://www.rusprofile.ru/codes/433400/sankt-peterburg/" } ,
                //{"433410", "https://www.rusprofile.ru/codes/433410/sankt-peterburg/" } ,
                //{"433420", "https://www.rusprofile.ru/codes/433420/sankt-peterburg/" } ,
                //{"433900", "https://www.rusprofile.ru/codes/433900/sankt-peterburg/" } ,
                //{"439100", "https://www.rusprofile.ru/codes/439100/sankt-peterburg/" } ,
                //{"439900", "https://www.rusprofile.ru/codes/439900/sankt-peterburg/" } ,
                //{"439910", "https://www.rusprofile.ru/codes/439910/sankt-peterburg/" } ,
                //{"439920", "https://www.rusprofile.ru/codes/439920/sankt-peterburg/" } ,
                //{"439930", "https://www.rusprofile.ru/codes/439930/sankt-peterburg/" } ,
                //{"439940", "https://www.rusprofile.ru/codes/439940/sankt-peterburg/" } ,
                //{"439950", "https://www.rusprofile.ru/codes/439950/sankt-peterburg/" } ,
                //{"439960", "https://www.rusprofile.ru/codes/439960/sankt-peterburg/" } ,
                //{"439970", "https://www.rusprofile.ru/codes/439970/sankt-peterburg/" } ,
                //{"439990", "https://www.rusprofile.ru/codes/439990/sankt-peterburg/" } ,
            };

            Parser parser = new Parser();
            Exporter exporter = new Exporter(ExportedDir);

            AuthCredetials credentials = new AuthCredetials
            {
                UserName = "wokkiturk@gmail.com",
                Password = "wokkiturk1"
            };
            IWebDriver driver = null;
            if (proxy == null)
            {
                driver = parser.CreateDriver(ExtDir, credentials);
            }
            else
            {
                driver = parser.CreateDriver(ExtDir, credentials, proxy);
            }

            DirectoryInfo profiles = new DirectoryInfo(ExportedDir + "\\Profiles\\");
            FileInfo[] files = profiles.GetFiles();

            FileInfo links = new FileInfo(ExportedDir + "\\" + LinksFileName);
            ExcelPackage pack = new ExcelPackage(links);

            for (int sheetNum = 0; sheetNum < pack.Workbook.Worksheets.Count; sheetNum++)
            {
                foreach (ExcelWorksheet sheet in pack.Workbook.Worksheets)
                {
                second_trying:
                    ExcelWorksheet linkSheet = sheet;
                    int last_row = linkSheet.Cells.Where(c => c.Start.Column == 1 &&
                                                         !c.Value.ToString().Equals("")).Last().End.Row;
                    int last_profile_row = 0;
                    Console.WriteLine("START PARSING " + linkSheet.Name + "At time: " + DateTime.Now.ToString());
                    ExcelPackage profilePack = null;
                    try
                    {
                        profilePack = new ExcelPackage(files.FirstOrDefault(f => f.Name.Contains(linkSheet.Name)));
                        ExcelWorksheet profileSheet = profilePack.Workbook.Worksheets.FirstOrDefault();
                        last_profile_row = profileSheet.Cells.Where(c => c.Start.Column == 2 &&
                                                                !c.Value.ToString().Equals("")).Last().End.Row + 1;

                        IEnumerable<ExcelRangeBase> query = from row in linkSheet.Cells["C:XFD"] select row;
                        ExcelRangeBase trying_find_last = query.ToList().FirstOrDefault(c => c.Value.ToString() == "wait");
                        if (trying_find_last == null)
                        {
                            Console.WriteLine("SKIP  PARSING " + linkSheet.Name + "At time: " + DateTime.Now.ToString());
                            continue;
                        }
                        int last_not_parsed_row = trying_find_last.Start.Row;

                        for (int row_index = last_not_parsed_row; row_index <= last_row; row_index++)
                        {
                            string url = linkSheet.Cells[row_index, 2].Value.ToString();

                            if (linkSheet.Cells[row_index, 3].Value.ToString() == "ok")
                            {
                                continue;
                            }
                            CompanyCard card = parser.ParseCard(driver, url).Result;

                            #region map
                            profileSheet.Cells[last_profile_row, 2].Value = card.ShortName;
                            profileSheet.Cells[last_profile_row, 3].Value = card.FullName;
                            profileSheet.Cells[last_profile_row, 4].Value = card.Phones;
                            profileSheet.Cells[last_profile_row, 5].Value = card.Mails;
                            profileSheet.Cells[last_profile_row, 6].Value = card.Sites;
                            profileSheet.Cells[last_profile_row, 7].Value = card.INN;
                            profileSheet.Cells[last_profile_row, 8].Value = card.PrimaryOccupation;
                            profileSheet.Cells[last_profile_row, 9].Value = card.OrganizationStatus;
                            #endregion
                            last_profile_row++;
                            profilePack.Save();

                            linkSheet.Cells[row_index, 3].Value = "ok";
                        }
                        pack.Save();
                        Console.WriteLine("STOP  PARSING " + linkSheet.Name + "At time: " + DateTime.Now.ToString());
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("Error on links page number = " + sheetNum);
                        Console.WriteLine("      last profile row = " + last_profile_row);
                        Console.WriteLine("      executed again");
                    save_mark:
                        try
                        {
                            pack.Save();
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            goto save_mark;
                        }
                        goto second_trying;
                    }
                }
            }
            driver.Close();
            player.Play();
            Console.WriteLine("Done");

            Console.ReadKey();

            //exporter.CreateFiles("Links.xlsx");
            #region export links
            //DirectoryInfo directory = new DirectoryInfo(OutDir);
            //var files = directory.GetFiles();

            //Console.WriteLine("Export started");
            //int notpartmax = files.Where(f => !f.Name.Contains("part")).Count();
            //int partmax = files.Where(f => f.Name.Contains("part")).Count();
            //Console.WriteLine($"not parted count: {notpartmax}");
            //Console.WriteLine($"parted count: {partmax}");
            //int counter = 1;
            //Console.WriteLine("notpart started");
            //foreach (var notpart in files.Where(f => !f.Name.Contains("part")))
            //{
            //    var dic = JsonConvert.DeserializeObject<Dictionary<string, string>>(JsonFile.JsonRead(notpart.FullName));
            //    exporter.ExportLinks(notpart.Name.Remove(6), dic);
            //    Console.WriteLine($"{counter++} of {notpartmax} exported");
            //}
            //counter = 1;
            //bool appendState = false;
            //Console.WriteLine("part started");
            //foreach (var part in files.Where(f => f.Name.Contains("part")))
            //{
            //    var dic = JsonConvert.DeserializeObject<Dictionary<string, string>>(JsonFile.JsonRead(part.FullName));
            //    exporter.ExportLinks(part.Name.Remove(6), dic, appendrows: appendState);
            //    appendState = true;
            //    Console.WriteLine($"{counter++} of {partmax} exported");
            //}
            //Console.WriteLine("Done");
            #endregion
        }
    }
}
