using System.Globalization;
using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public record LigneFactureRecord(
    string CodeProduit,
    decimal PrixHT,
    string Pays,
    [property: XlsxConverter(typeof(CalculTtcParPaysConverter))] decimal PrixTTC
)
{
     public string? NoteInterne { get; init; }
}

public sealed class CalculTtcParPaysConverter : IXlsxRowConverter<LigneFactureRecord, decimal>
{
     public bool TryRead(ReadOnlySpan<char> rawValue, in LigneFactureRecord row, XlsxReadOptions options, out decimal value)
     {
          if (!rawValue.IsEmpty && decimal.TryParse(rawValue, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
               return true;

          decimal taux = row.Pays switch
          {
               "FR" => 1.20m,
               "CH" => 1.081m,
               _ => 1.00m
          };
          value = Math.Round(row.PrixHT * taux, 2);
          return true;
     }

     public string? Write(in decimal value, in LigneFactureRecord row, XlsxWriteOptions options)
         => value.ToString("0.00", CultureInfo.InvariantCulture);
}

public class FlashXlsxRecordsAndConvertersTests
{
     [Fact]
     public async Task PositionalRecord_InitProperties_And_RowConverter_ShouldRoundTrip()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["CodeProduit"] = "SRV-01",
                ["PrixHT"] = 100.00m,
                ["Pays"] = "FR",
                ["PrixTTC"] = XlsxValue.Empty, // Calculé automatiquement à partir de PrixHT et Pays !
                ["NoteInterne"] = "Client VIP"
            },
            new()
            {
                ["CodeProduit"] = "SRV-02",
                ["PrixHT"] = 200.00m,
                ["Pays"] = "CH",
                ["PrixTTC"] = XlsxValue.Empty,
                ["NoteInterne"] = "Suisse"
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<LigneFactureRecord>(ms);

          Assert.Equal(2, imported.Count);

          Assert.Equal("SRV-01", imported[0].CodeProduit);
          Assert.Equal(100.00m, imported[0].PrixHT);
          Assert.Equal("FR", imported[0].Pays);
          Assert.Equal(120.00m, imported[0].PrixTTC);
          Assert.Equal("Client VIP", imported[0].NoteInterne);

          Assert.Equal("SRV-02", imported[1].CodeProduit);
          Assert.Equal(216.20m, imported[1].PrixTTC);
          Assert.Equal("Suisse", imported[1].NoteInterne);
     }
}