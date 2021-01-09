using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using RusProfileApplication.Models;

namespace RusProfileApplication.Executers
{
    public class Exporter
    {
        private readonly string ExportDirectory;
        public Exporter(string ExportDirectory) => this.ExportDirectory = ExportDirectory;

        public void ExportLinks(string sheetname,
                                Dictionary<string, string> dic,
                                string filename = "Links.xlsx",
                                bool appendrows = false)
        {
            FileInfo filepath = new FileInfo(ExportDirectory + "\\" + filename);
            ExcelPackage pack = new ExcelPackage(filepath);

            ExcelWorksheet worksheet = pack.Workbook.Worksheets.FirstOrDefault(w => w.Name == sheetname);
            if (worksheet == null)
            {
                worksheet = pack.Workbook.Worksheets.Add(sheetname);
            }
            worksheet.Cells[1, 1].Value = "Наименование";
            worksheet.Cells[1, 2].Value = "Ссылка";
            worksheet.Cells[1, 3].Value = "Статус";

            int row_start = 2;
            if (appendrows)
            {
                int last_row = worksheet.Cells.Where(c => c.Start.Column == 1 &&
                                                          !c.Value.ToString().Equals("")).Last().End.Row;
                row_start = last_row;
            }
            foreach (string key in dic.Keys)
            {
                worksheet.Cells[row_start, 1].Value = key;
                worksheet.Cells[row_start, 2].Value = dic[key];
                worksheet.Cells[row_start, 3].Value = "wait";
                row_start++;
            }

            worksheet.Cells[1, 1].AutoFitColumns(300);
            worksheet.Cells[1, 2].AutoFitColumns(900);


            pack.Save();
        }
        public void UpdateWidth(string filename)
        {
            FileInfo filepath = new FileInfo(ExportDirectory + "\\" + filename);
            ExcelPackage pack = new ExcelPackage(filepath);

            foreach (ExcelWorksheet sheet in pack.Workbook.Worksheets)
            {
                sheet.Column(1).Width = 50;
                sheet.Column(2).Width = 75;
            }

            pack.Save();
        }
        public void CreateFiles(string filename)
        {
            CompanyCard fieldsModel = new CompanyCard();
            FileInfo filepath = new FileInfo(ExportDirectory + "\\" + filename);
            ExcelPackage pack = new ExcelPackage(filepath);

            foreach (ExcelWorksheet sheet in pack.Workbook.Worksheets)
            {
                FileInfo newBook = new FileInfo(ExportDirectory + "\\Profiles\\" + sheet.Name + ".xlsx");
                ExcelPackage elemPack = new ExcelPackage(newBook);

                ExcelWorksheet worksheet = elemPack.Workbook.Worksheets.FirstOrDefault(w => w.Name == "Информация");
                if (worksheet == null)
                {
                    worksheet = elemPack.Workbook.Worksheets.Add("Информация");
                }

                PropertyInfo[] props = fieldsModel.GetType().GetProperties();

                for (int p = 0; p < props.Length; p++)
                {
                    string displayName = props[p].CustomAttributes.FirstOrDefault()
                                                  .ConstructorArguments.FirstOrDefault()
                                                  .Value.ToString();
                    worksheet.Cells[1, p + 1].Value = displayName;
                }

                worksheet.Column(1).Width = 10;
                worksheet.Column(2).Width = 30;
                worksheet.Column(3).Width = 30;
                worksheet.Column(4).Width = 17;
                worksheet.Column(5).Width = 17;
                worksheet.Column(6).Width = 17;
                worksheet.Column(7).Width = 20;
                worksheet.Column(8).Width = 30;
                worksheet.Column(9).Width = 20;

                elemPack.Save();
            }
            System.Console.WriteLine("Done");
        }
    }
}
