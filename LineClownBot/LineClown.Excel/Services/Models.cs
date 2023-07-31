using System;

namespace LineClown.Excel.Services;

public static class Settings
{
    public static string XlsxRoot { get; set; } 
    public static string OutFilename { get; set; } 
}

public class UserRowData
{
    public string Name { get; set; }
    public int Hunt { get; set; }

    public int HuntL1 { get; set; }
    public int HuntL2 { get; set; }
    public int HuntL3 { get; set; }
    public int HuntL4 { get; set; }
    public int HuntL5 { get; set; }

    public int PurchaseL1 { get; set; }
    public int PurchaseL2 { get; set; }
    public int PurchaseL3 { get; set; }
    public int PurchaseL4 { get; set; }
    public int PurchaseL5 { get; set; }

    public bool HasPurchased => PurchaseL1 > 0 || PurchaseL2 > 0 || PurchaseL3 > 0 || PurchaseL4 > 0 || PurchaseL5 > 0;

    public int HuntBusPoints => 
        HuntL1 * MonsterBusPoints.Common + HuntL2 * MonsterBusPoints.Uncommon + HuntL3 * MonsterBusPoints.Epic + HuntL4 * MonsterBusPoints.Rare + HuntL5 * MonsterBusPoints.Legendary;

    public DateTime FirstHunt { get; set; }
    public DateTime LastHunt { get; set; }

    public double Rss { get; set; }
    public int GfScore { get; set; }
}

public static class MonsterBusPoints
{
    public const int Common = 2;
    public const int Uncommon = 5;
    public const int Rare = 20;
    public const int Epic = 50;
    public static int Legendary = 100;
}