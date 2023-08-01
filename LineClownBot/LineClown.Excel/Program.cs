using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using LineClown.Excel.Services;
using OfficeOpenXml;

namespace LineClown.Excel;

public class Program
{
    private static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Settings.XlsxRoot = ConfigurationManager.AppSettings.Get("XlsxRoot") ?? string.Empty;
        Settings.OutFilename = ConfigurationManager.AppSettings.Get("OutFilename") ?? string.Empty;

        if(!Directory.Exists(Settings.XlsxRoot))
        {
            throw new Exception($"Could not find source folder: {Settings.XlsxRoot}");
        }

        IDictionary<int, IList<UserRowData>> dicData = new Dictionary<int, IList<UserRowData>>();

        ExcelReader rd = new ExcelReader();
        ExcelWriter wr = new ExcelWriter();

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Precessing Loot, {Settings.XlsxRoot}");
        Console.ResetColor();

        rd.ProcessLoot(ref dicData);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Precessing Bank, *BankData.xlsx");
        Console.ResetColor();

        rd.ProcessBank(ref dicData);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Precessing GF, GF *.xlsx");
        Console.ResetColor();

        rd.ProcessGuildFest(ref dicData);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Precessing GuidStats, GuildStats*.xlsx");
        Console.ResetColor();

        rd.ProcessGuildStats(ref dicData);

        Console.WriteLine("-------------------------");
        Console.WriteLine("Name, Tot, Lvl1, Lvl2, Lvl3, Lvl4, Lvl5, Pts, p2p, First, Last, RSS, GF, Δ");


        PrintList(ref dicData);

        var fileData = wr.ExportList(ref dicData);

        string strPath = Path.Combine(Settings.XlsxRoot, string.Format(Settings.OutFilename, DateTime.Today.ToString("yyyy_MM_dd")));
            
        File.WriteAllBytes(strPath, fileData);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Wrote summary file {strPath}");
        Console.ResetColor();

        Console.ReadLine();
    }

    private static void PrintList(ref IDictionary<int, IList<UserRowData>> dicData)
    {
        foreach ((_, IList<UserRowData> value) in dicData.OrderBy(_ => _.Value.First().Name))
        {
            var name = value.First().Name;

            var totHunt = value.Sum(_ => _.Hunt);
            var totHuntL1 = value.Sum(_ => _.HuntL1);
            var totHuntL2 = value.Sum(_ => _.HuntL2);
            var totHuntL3 = value.Sum(_ => _.HuntL3);
            var totHuntL4 = value.Sum(_ => _.HuntL4);
            var totHuntL5 = value.Sum(_ => _.HuntL5);
            var totHuntPoints = value.Sum(_ => _.HuntBusPoints);
            int totRss = (int)(value.Sum(_ => _.Rss) / 1_000_000);
            int totGf = value.Sum(_ => _.GfScore);

            string totsRss = string.Empty;
            if (totRss > 0)
            {
                totsRss = $"{totRss}M";
            }
            string gfScore = totGf.ToString();


            if (!value.Any(_ => _.FirstHunt.Year > 2009))
            {
                Console.WriteLine($"{name}, {totHunt},,,,,,,,,,,, {totsRss}, {totGf}");

                continue;
            }

            bool hasPurchased = value.Any(_ => _.HasPurchased);
            DateTime startHunt = value.Where(_ => _.FirstHunt.Year > 2009).Min(_ => _.FirstHunt);
            DateTime endHunt = value.Where(_ => _.LastHunt.Year > 2009).Max(_ => _.LastHunt);

            Console.WriteLine($"{name}, {totHunt}, {totHuntL1}, {totHuntL2}, {totHuntL3}, {totHuntL4}, {totHuntL5}, {totHuntPoints}, {(hasPurchased ? "y" : "n")}, {startHunt.ToShortDateString()}, {endHunt.ToShortDateString()}, {totsRss}, {gfScore}, Δ");
        }
    }
}