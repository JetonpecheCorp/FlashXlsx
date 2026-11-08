using System.Text;
using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneAvecConvertisseurFautif
{
     public string Nom { get; set; } = string.Empty;

     [XlsxConverter(typeof(ConvertisseurExplosif))]
     public int Valeur { get; set; }
}

public sealed class ConvertisseurExplosif : IXlsxConverter<int>
{
     public bool TryRead(ReadOnlySpan<char> rawValue, XlsxReadOptions options, out int value)
     {
          if (rawValue.Equals("BOOM", StringComparison.OrdinalIgnoreCase))
               throw new InvalidOperationException("Erreur critique de conversion");

          return int.TryParse(rawValue, out value);
     }

     public string? Write(in int value, XlsxWriteOptions options)
     {
          if (value < 0)
               throw new InvalidOperationException("Valeur négative interdite à l'export");

          return value.ToString();
     }
}

public class FlashXlsxErrorHandlingTests
{
     [Fact]
     public async Task Write_WithConverterException_And_ReportAndContinue_ShouldLogErrorAndContinue()
     {
          var items = new List<LigneAvecConvertisseurFautif>
        {
            new() { Nom = "OK1", Valeur = 10 },
            new() { Nom = "ERREUR", Valeur = -5 },
            new() { Nom = "OK2", Valeur = 20 }
        };

          var errors = new List<XlsxErrorContext>();
          var writeOpts = new XlsxWriteOptions<LigneAvecConvertisseurFautif>()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, items, writeOpts);

          Assert.Single(errors);
          Assert.Equal(3, errors[0].RowNumber); // Ligne 1 = Header, Ligne 2 = OK1, Ligne 3 = ERREUR
          Assert.Equal(XlsxCellStatus.ConverterError, errors[0].Status);

          ms.Position = 0;
          var readBack = await FlashXlsxSerializer.ReadAllAsync<LigneAvecConvertisseurFautif>(ms);
          Assert.Equal(3, readBack.Count);
          Assert.Equal(10, readBack[0].Valeur);
          Assert.Equal(0, readBack[1].Valeur); // Cellule laissée vide suite à l'erreur ignorée
          Assert.Equal(20, readBack[2].Valeur);
     }

     [Fact]
     public async Task Read_WithCorruptedZipStream_ShouldThrowDescriptiveInvalidDataException()
     {
          using var ms = new MemoryStream(Encoding.UTF8.GetBytes("Ceci n'est pas un fichier ZIP/XLSX"));

          await Assert.ThrowsAsync<InvalidDataException>(async () =>
          {
               await FlashXlsxSerializer.ReadAllAsync<LigneAvecConvertisseurFautif>(ms);
          });
     }
}