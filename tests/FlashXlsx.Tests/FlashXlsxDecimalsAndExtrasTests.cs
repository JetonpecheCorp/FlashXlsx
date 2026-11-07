using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneComptableFlexible
{
     [XlsxDefault("FR")]
     public string Pays { get; set; } = string.Empty;

     // Auto-détection flexible : lit "1 234,56" ou "1,234.56"
     public decimal MontantTTC { get; set; }

     // Forcé sur le point même si le classeur est lu en culture française
     [XlsxDecimalSeparator('.')]
     public double Latitude { get; set; }

     [XlsxFormat("dd/MM/yyyy")]
     public DateOnly DateOperation { get; set; }

     [XlsxSplit(';')]
     public List<string> Tags { get; set; } = [];
}

[XlsxSerializable]
public class LigneSansEnTete
{
     [XlsxColumnLetter("A")]
     public string Code { get; set; } = string.Empty;

     [XlsxColumnLetter("D")] // Saute B et C
     public decimal Prix { get; set; }
}

public class FlashXlsxDecimalsAndExtrasTests
{
     [Fact]
     public async Task DecimalSeparator_Default_Format_And_Split_ShouldWorkEndToEnd()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["Pays"] = XlsxValue.Empty, // Doit prendre [XlsxDefault("FR")]
                ["MontantTTC"] = "1 450,75", // Virgule + espace milliers
                ["Latitude"] = "45.7640",    // Point explicite
                ["DateOperation"] = "07/11/2026",
                ["Tags"] = "Export; Urgent ; Comptabilité"
            },
            new()
            {
                ["Pays"] = "BE",
                ["MontantTTC"] = "2,500.25", // Format anglo-saxon dans la même colonne !
                ["Latitude"] = "50.8503",
                ["DateOperation"] = "08/11/2026",
                ["Tags"] = "Archive"
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<LigneComptableFlexible>(ms);

          Assert.Equal(2, imported.Count);

          Assert.Equal("FR", imported[0].Pays);
          Assert.Equal(1450.75m, imported[0].MontantTTC);
          Assert.Equal(45.7640, imported[0].Latitude, precision: 4);
          Assert.Equal(new DateOnly(2026, 11, 7), imported[0].DateOperation);
          Assert.Equal(["Export", "Urgent", "Comptabilité"], imported[0].Tags);

          Assert.Equal("BE", imported[1].Pays);
          Assert.Equal(2500.25m, imported[1].MontantTTC);
     }

     [Fact]
     public async Task ColumnLetter_WithoutHeaderRow_ShouldMapByExcelCoordinates()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new() { ["ColA"] = "REF-01", ["ColB"] = "Ignore1", ["ColC"] = "Ignore2", ["ColD"] = "99,90" }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows, new XlsxWriteOptions().WithHeader(false));

          ms.Position = 0;
          var options = new XlsxReadOptions<LigneSansEnTete>()
              .WithHeader(hasHeader: false);

          var rows = await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Single(rows);
          Assert.Equal("REF-01", rows[0].Code);
          Assert.Equal(99.90m, rows[0].Prix);
     }
}