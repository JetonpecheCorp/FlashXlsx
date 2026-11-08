using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneImportDevise
{
     public string Libelle { get; set; } = string.Empty;
     public decimal PrixDecimal { get; set; }
     public double PrixDouble { get; set; }
}

public class FlashXlsxEncryptedAndCurrencyTests
{
     [Fact]
     public async Task EncryptedOle2File_ShouldNotifyOnError_AndStopWithoutThrowing()
     {
          // Simulation d'un fichier Excel chiffré par mot de passe (en-tête OLE2/CFB D0 CF 11 E0 A1 B1 1A E1)
          byte[] ole2Header = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00, 0x00, 0x00];
          using var ms = new MemoryStream(ole2Header);

          var errors = new List<XlsxErrorContext>();
          var opts = new XlsxReadOptions<LigneImportDevise>()
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          var rows = await FlashXlsxSerializer.ReadAllAsync(ms, opts);

          Assert.Empty(rows);
          Assert.Single(errors);
          Assert.Equal(XlsxCellStatus.EncryptedFile, errors[0].Status);
     }

     [Fact]
     public async Task AutomaticCurrencyStripping_ShouldParseFormattedCurrencyStrings()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new() { ["Libelle"] = "France", ["PrixDecimal"] = "1\u202F250,50 €", ["PrixDouble"] = "3 400,75 EUR" },
            new() { ["Libelle"] = "USA", ["PrixDecimal"] = "$1,250.50", ["PrixDouble"] = "USD 999.25" },
            new() { ["Libelle"] = "Suisse", ["PrixDecimal"] = "CHF 450.00", ["PrixDouble"] = "£120.50" }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<LigneImportDevise>(ms);

          Assert.Equal(3, imported.Count);

          Assert.Equal(1250.50m, imported[0].PrixDecimal);
          Assert.Equal(3400.75, imported[0].PrixDouble, 2);

          Assert.Equal(1250.50m, imported[1].PrixDecimal);
          Assert.Equal(999.25, imported[1].PrixDouble, 2);

          Assert.Equal(450.00m, imported[2].PrixDecimal);
          Assert.Equal(120.50, imported[2].PrixDouble, 2);
     }
}