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
            //Executer.ExecuteFromLinks(args);
            Executer.ExecuteFromSearch();

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
