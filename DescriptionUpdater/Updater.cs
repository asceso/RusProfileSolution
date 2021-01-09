using System;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;

namespace DescriptionUpdater
{
    class Updater
    {
        internal void RunUpdater()
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

                int last_row = sheet.Cells.Where(c => c.Start.Column == 2 && !c.Value.ToString().Equals("")).Last().End.Row;
                int ppNum = 1;
                for (int i = 2; i < last_row; i++)
                {
                    string value = sheet.Cells[i, 8].Value.ToString();
                    sheet.Cells[i, 8].Value = value.GetInteger();
                    sheet.Cells[i, 1].Value = ppNum++;
                }

                pack.SaveAs(new FileInfo(noDirPath + "\\" + file.Name));
            }
            Console.WriteLine("Done.");
            Thread.Sleep(3000);
        }
    }
}
