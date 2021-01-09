using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;

namespace DescriptionUpdater
{
    class Updater
    {
        internal void RunUpdater(string[] args)
        {
            string withDirPath = Environment.CurrentDirectory + "\\With description";
            string noDirPath = Environment.CurrentDirectory + "\\No description";

            DirectoryInfo withDir = new DirectoryInfo(withDirPath);
            int filesCount = withDir.GetFiles().Length;
            Console.WriteLine("Files founded: " + filesCount);

            Console.WriteLine("Start updating");
            int currentCounter = 1;
            foreach (FileInfo file in withDir.GetFiles())
            {
                Console.WriteLine($"Start updating {currentCounter++} of {filesCount}");
                ExcelPackage pack = new ExcelPackage(file);

                ExcelWorksheet sheet = pack.Workbook.Worksheets.FirstOrDefault();

                int start_number = Convert.ToInt32(args[0]);
                int collumn_number = Convert.ToInt32(args[1]);
                int last_row = 0;
                if (args.Length == 3)
                {
                    last_row = Convert.ToInt32(args[2]);
                }
                else
                {
                    last_row = sheet.Cells.Where(c => c.Start.Column == collumn_number && !c.Value.ToString().Equals("")).Last().End.Row;
                }
                int ppNum = 1;

                for (int i = start_number; i <= last_row; i++)
                {
                    object value = sheet.Cells[i, collumn_number].Value;
                    if (value == null)
                    {
                        continue;
                    }
                    sheet.Cells[i, collumn_number].Value = value.ToString().GetInteger();
                    sheet.Cells[i, 1].Value = ppNum++;
                }

                pack.SaveAs(new FileInfo(noDirPath + "\\" + file.Name));
            }
            Console.WriteLine("Done.");
            Thread.Sleep(3000);
        }
    }
}
