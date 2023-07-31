using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace LineClown.Excel.Services;

public class ExcelReader
{
    public void ProcessLoot(ref IDictionary<int, IList<UserRowData>> dicData)
    {
        List<FileInfo> allXlsxFiles = Directory.EnumerateFiles(Settings.XlsxRoot, "*.xlsx").Select(_ => new FileInfo(_)).ToList();

        foreach (FileInfo fileInfo in allXlsxFiles.OrderByDescending(_ => _.LastWriteTime))
        {
            string file = fileInfo.FullName;
            if (file.StartsWith('~') || file.Contains("BankData"))
            {
                Console.WriteLine($"Skipping {file}");
                continue;
            }

            Console.WriteLine($"{file}");

            using (var package = new ExcelPackage(fileInfo))
            {
                for (int rowY = 1; rowY < 100; rowY++)
                {
                    UserRowData rowData = new UserRowData();

                    var cells = package.Workbook.Worksheets[0].Cells;
                    string userId = cells[rowY, 1].GetValue<string>();
                    if (string.IsNullOrWhiteSpace(userId))
                    {
                        break;
                    }

                    if (!int.TryParse(userId, out int uid) || uid == 0)
                    {
                        continue;
                    }

                    if (!dicData.ContainsKey(uid))
                    {
                        dicData.Add(uid, new List<UserRowData>());
                    }

                    dicData[uid].Add(rowData);
                    rowData.Name = cells[rowY, 2].GetValue<string>();
                    rowData.Hunt = cells[rowY, 4].GetValue<int>();
                    //
                    rowData.HuntL1 = cells[rowY, 7].GetValue<int>();
                    rowData.HuntL2 = cells[rowY, 8].GetValue<int>();
                    rowData.HuntL3 = cells[rowY, 9].GetValue<int>();
                    rowData.HuntL4 = cells[rowY, 10].GetValue<int>();
                    rowData.HuntL5 = cells[rowY, 11].GetValue<int>();
                    //
                    rowData.PurchaseL1 = cells[rowY, 13].GetValue<int>();
                    rowData.PurchaseL2 = cells[rowY, 14].GetValue<int>();
                    rowData.PurchaseL3 = cells[rowY, 15].GetValue<int>();
                    rowData.PurchaseL4 = cells[rowY, 16].GetValue<int>();
                    rowData.PurchaseL5 = cells[rowY, 17].GetValue<int>();
                    //
                    rowData.FirstHunt = cells[rowY, 25].GetValue<DateTime>();
                    rowData.LastHunt = cells[rowY, 26].GetValue<DateTime>();
                }
            }
        }
    }

    public void ProcessBank(ref IDictionary<int, IList<UserRowData>> dicData)
    {
        foreach (string file in Directory.EnumerateFiles(Settings.XlsxRoot, "*BankData.xlsx"))
        {
            if (file.StartsWith('~'))
            {
                Console.WriteLine($"Skipping {file}");
                continue;
            }

            Console.WriteLine($"{file}");

            using var package = new ExcelPackage(new FileInfo(file));
            for (int rowY = 1; rowY < 300; rowY++)
            {
                var cells = package.Workbook.Worksheets[0].Cells;

                string userId = cells[rowY, 2].GetValue<string>();
                if (string.IsNullOrWhiteSpace(userId))
                {
                    break;
                }

                if (!int.TryParse(userId, out int uid) || uid == 0)
                {
                    continue;
                }

                if (dicData.TryGetValue(uid, out IList<UserRowData> urdList))
                {
                    var aRow = urdList.First();
                    string food = cells[rowY, 3].GetValue<string>();
                    string stone = cells[rowY, 4].GetValue<string>();
                    string wood = cells[rowY, 5].GetValue<string>();
                    string ore = cells[rowY, 6].GetValue<string>();
                    string gold = cells[rowY, 7].GetValue<string>();

                    var dFood = ParsePrefix(food);
                    var dStone = ParsePrefix(stone);
                    var dWood = ParsePrefix(wood);
                    var dOre = ParsePrefix(ore);
                    var dGold = ParsePrefix(gold);

                    aRow.Rss = dFood + dStone + dWood + dOre + dGold;
                }
            }
        }
    }

    private double ParsePrefix(string value)
    {
        var numberFormat = CultureInfo.GetCultureInfo("en-US", true);
        string[] superSuffix = new [] { "K", "M", "B" };

        foreach (char c in value)
        {
            foreach (string s in superSuffix)
            {
                if (c.ToString().Equals(s.ToLower(), StringComparison.InvariantCultureIgnoreCase))
                {
                    char suffixChar = s[0];
                    string num = value.Substring(0, value.IndexOf(c));
                    double multiplier = Math.Pow(1000, superSuffix.ToList().IndexOf(suffixChar.ToString()) + 1);
                    return Convert.ToDouble(num, numberFormat) * multiplier;
                }
            }
        }

        return Convert.ToDouble(value, numberFormat);
    }

    internal void ProcessGf(ref IDictionary<int, IList<UserRowData>> dicData)
    {
        foreach (string file in Directory.EnumerateFiles(Settings.XlsxRoot, "GF *.xlsx"))
        {
            if (file.StartsWith('~'))
            {
                Console.WriteLine($"Skipping {file}");
                continue;
            }

            Console.WriteLine($"{file}");

            using var package = new ExcelPackage(new FileInfo(file));
            for (int rowY = 2; rowY < 200; rowY++)
            {
                var cells = package.Workbook.Worksheets[0].Cells;
                string name = cells[rowY, 1].GetValue<string>();

                if (string.IsNullOrWhiteSpace(name))
                {
                    break;
                }

                var aRows = dicData.FirstOrDefault(kvp => kvp.Value.Any(urd => urd.Name.Equals(name)));
                if (aRows.Value != null)
                {
                    var aRow = aRows.Value.First();

                    string gfScore = cells[rowY, 3].GetValue<string>();

                    var gf = ParsePrefix(gfScore);

                    aRow.GfScore = (int)gf;
                }
            }
        }
    }
}