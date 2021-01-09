using System;
using System.IO;
using System.Xml;

namespace SettingsAplifier
{
    internal class Program
    {
        private static void Main()
        {
            Console.WriteLine("Start settings change");
            XmlDocument xml = new XmlDocument();
            string path = Environment.CurrentDirectory + "\\RusProfileApplication.exe.config";
            StreamReader reader = new StreamReader(path);
            xml.LoadXml(reader.ReadToEnd());

            XmlElement root = xml.DocumentElement;
            foreach (XmlNode rootChild in root.ChildNodes)
            {
                if (rootChild.Name == "appSettings")
                {
                    foreach (XmlNode key in rootChild.ChildNodes)
                    {
                        if (key.Attributes[0].Value == "EPPlus:ExcelPackage.LicenseContext")
                        {
                            continue;
                        }
                        string insertString = string.Empty;
                        foreach (XmlAttribute attr in key.Attributes)
                        {
                            if (attr.Name == "key" && attr.Value == "OutputDirectory")
                            {
                                insertString = Environment.CurrentDirectory + "\\OutDir";
                            }
                            if (attr.Name == "key" && attr.Value == "ExtensionDirectory")
                            {
                                insertString = Environment.CurrentDirectory + "\\Extensions";
                            }
                            if (attr.Name == "key" && attr.Value == "ExportedDirectory")
                            {
                                insertString = Environment.CurrentDirectory + "\\Exported";
                            }
                            if (attr.Name == "value")
                            {
                                attr.Value = insertString;
                            }
                        }
                    }
                }
            }
            reader.Close();
            xml.Save(path);
            Console.WriteLine("Done");
        }
    }
}
