using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace LineClown.Excel.Services;

public class ExcelWriter
{
    public byte[] ExportList(ref IDictionary<int, IList<UserRowData>> dicData)
    {
        using ExcelPackage excel = new ExcelPackage();

        var workSheet = excel.Workbook.Worksheets.Add("Stats");

        int recordIndex = 1;

        workSheet.Cells[recordIndex, 1].Value = "Name";
        workSheet.Cells[recordIndex, 2].Value = "Tot";
        workSheet.Cells[recordIndex, 3].Value = "lvl1";
        workSheet.Cells[recordIndex, 4].Value = "lvl2";
        workSheet.Cells[recordIndex, 5].Value = "lvl3";
        workSheet.Cells[recordIndex, 6].Value = "lvl4";
        workSheet.Cells[recordIndex, 7].Value = "lvl5";
        workSheet.Cells[recordIndex, 8].Value = "Pts";
        workSheet.Cells[recordIndex, 9].Value = "p2p";
        workSheet.Cells[recordIndex, 10].Value = "First";
        workSheet.Cells[recordIndex, 11].Value = "Last";
        workSheet.Cells[recordIndex, 12].Value = "RSS";
        workSheet.Cells[recordIndex, 13].Value = "GF";

        recordIndex++;

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
            double totRssRaw = value.Sum(_ => _.Rss);
            int totRss = (int)(value.Sum(_ => _.Rss) / 1000000);
            bool hasPurchased = value.Any(_ => _.HasPurchased);
            int totGf = value.Sum(_ => _.GfScore);

            string totFmtRss = string.Empty;

            if (totRss > 0)
            {
                totFmtRss = $"{totRss}M";
            }

            workSheet.Cells[recordIndex, 1].Value = name;
            workSheet.Cells[recordIndex, 2].Value = totHunt;
            workSheet.Cells[recordIndex, 3].Value = totHuntL1;
            workSheet.Cells[recordIndex, 4].Value = totHuntL2;
            workSheet.Cells[recordIndex, 5].Value = totHuntL3;
            workSheet.Cells[recordIndex, 6].Value = totHuntL4;
            workSheet.Cells[recordIndex, 7].Value = totHuntL5;
            workSheet.Cells[recordIndex, 8].Value = totHuntPoints;
            if (IsMaxPts(dicData, totHuntPoints))
            {
                workSheet.Cells[recordIndex, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[recordIndex, 8].Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent6);
            }
            else if (totHuntPoints < 100)
            {
                workSheet.Cells[recordIndex, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[recordIndex, 8].Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent3);
            }

            workSheet.Cells[recordIndex, 9].Value = $"{(hasPurchased ? "y" : "n")}";

            if (!value.Any(_ => _.FirstHunt.Year > 2009))
            {
                workSheet.Cells[recordIndex, 10].Value = "";
                workSheet.Cells[recordIndex, 11].Value = "";
            }
            else
            {
                // decimal huntGoalPct = value.Average(_ => _.HuntGoalPct);
                DateTime startHunt = value.Where(_ => _.FirstHunt.Year > 1900).Min(_ => _.FirstHunt);
                DateTime endHunt = value.Where(_ => _.LastHunt.Year > 1900).Max(_ => _.LastHunt);

                if ((DateTime.Now - endHunt).TotalDays > 2)
                {
                    workSheet.Cells[recordIndex, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[recordIndex, 11].Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent2);
                }

                workSheet.Cells[recordIndex, 9].Style.Numberformat.Format = "@";
                workSheet.Cells[recordIndex, 10].Value = startHunt.ToShortDateString();
                workSheet.Cells[recordIndex, 11].Value = endHunt.ToShortDateString();
            }

            workSheet.Cells[recordIndex, 12].Value = totFmtRss;
            if (IsMaxRss(dicData, totRssRaw))
            {
                workSheet.Cells[recordIndex, 12].Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Cells[recordIndex, 12].Style.Fill.BackgroundColor.SetColor(eThemeSchemeColor.Accent6);
            }
            workSheet.Cells[recordIndex, 13].Value = totGf;

            recordIndex++;
        }

        workSheet.Column(1).AutoFit();
        workSheet.Column(2).AutoFit();
        workSheet.Column(3).AutoFit();
        workSheet.Column(4).AutoFit();
        workSheet.Column(5).AutoFit();
        workSheet.Column(6).AutoFit();
        workSheet.Column(7).AutoFit();
        workSheet.Column(8).AutoFit();
        workSheet.Column(9).AutoFit();
        workSheet.Column(9).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Column(10).AutoFit();
        workSheet.Column(10).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Column(11).AutoFit();
        workSheet.Column(11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Column(12).AutoFit();
        workSheet.Column(12).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

        workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        workSheet.Cells[recordIndex, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 1].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 2].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 3].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 4].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 4].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 5].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 6].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 7].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 7].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 8].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 8].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 9].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 10].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 11].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 12].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 12].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
        workSheet.Cells[recordIndex, 13].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        workSheet.Cells[recordIndex, 13].Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);

        workSheet.Cells[recordIndex, 2].Formula = $"SUM($B$2:$B${recordIndex - 1})";
        workSheet.Cells[recordIndex, 3].Formula = $"SUM(C2:C{recordIndex - 1})";
        workSheet.Cells[recordIndex, 4].Formula = $"SUM(D2:D{recordIndex - 1})";
        workSheet.Cells[recordIndex, 5].Formula = $"SUM(E2:E{recordIndex - 1})";
        workSheet.Cells[recordIndex, 6].Formula = $"SUM(F2:F{recordIndex - 1})";
        workSheet.Cells[recordIndex, 7].Formula = $"SUM(G2:G{recordIndex - 1})";

        workSheet.Row(1).Style.Font.Bold = true;
        workSheet.Cells["A1:M1"].AutoFilter = true;

        workSheet.Cells["O3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["O3"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xE9, 0xE0, 0xE0));
        workSheet.Cells["O3"].Value = "Common";
        workSheet.Cells["P3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells["P3"].Value = $"{MonsterBusPoints.Common}p";

        workSheet.Cells["O4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["O4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x51, 0xE3, 0x69));
        workSheet.Cells["O4"].Value = "Uncommon";
        workSheet.Cells["P4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells["P4"].Value = $"{MonsterBusPoints.Uncommon}p";

        workSheet.Cells["O5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["O5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0x5C, 0xCF, 0xF5));
        workSheet.Cells["O5"].Value = "Rare";
        workSheet.Cells["P5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells["P5"].Value = $"{MonsterBusPoints.Rare}p";

        workSheet.Cells["O6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["O6"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xB5, 0x65, 0xB5));
        workSheet.Cells["O6"].Value = "Epic";
        workSheet.Cells["P6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells["P6"].Value = $"{MonsterBusPoints.Epic}p";

        workSheet.Cells["O7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["O7"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0xA5, 0x9A, 0x18));
        workSheet.Cells["O7"].Value = "Legendary";
        workSheet.Cells["P7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells["P7"].Value = $"{MonsterBusPoints.Legendary}p";

        workSheet.Column(14).AutoFit();
        workSheet.Column(15).AutoFit();
        workSheet.Column(16).AutoFit();

        workSheet.Cells["O11"].Style.Font.Bold = true;
        workSheet.Cells["O11"].Value = "Tools";

        workSheet.Cells["O12"].Hyperlink = new Uri("https://medio.pe/dev/LM_Calc/");
        workSheet.Cells["O12"].Value = "https://medio.pe/dev/LM_Calc/";

        workSheet.Cells["O13"].Hyperlink = new Uri("https://lordsmobile.igg.com/project/game_tool/index.php");

        workSheet.Cells["O14"].Hyperlink = new Uri("https://www.lordsmobilecalculator.com/tools/troop-training/");

        return excel.GetAsByteArray();
    }

    private bool IsMaxPts(IDictionary<int, IList<UserRowData>> dicData, int totHuntPoints)
    {
        int top5Min = dicData.Values.OrderByDescending(_ => _.Sum(r => r.HuntBusPoints)).Take(5).Min(m => m.Sum(i => i.HuntBusPoints));
        return totHuntPoints >= top5Min;
    }

    private bool IsMaxRss(IDictionary<int, IList<UserRowData>> dicData, double totRssRaw)
    {
        double top5Min = dicData.Values.OrderByDescending(_ => _.Sum(r => r.Rss)).Take(5).Min(m => m.Sum(i => i.Rss));
        return totRssRaw >= top5Min;
    }
}