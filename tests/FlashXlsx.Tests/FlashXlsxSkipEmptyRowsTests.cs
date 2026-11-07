using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

[XlsxSerializable]
public class LigneAvecRequis
{
     [XlsxRequired]
     public string Code { get; set; } = string.Empty;

     public int Quantite { get; set; }
}

public class FlashXlsxSkipEmptyRowsTests
{
     [Fact]
     public async Task SkipEmptyRows_ShouldIgnoreBlankRowsByDefault_WithoutTriggeringRequiredErrors()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new() { ["Code"] = "A1", ["Quantite"] = 5 },
            new() { ["Code"] = "   ", ["Quantite"] = XlsxValue.Empty }, // Ligne vide (espaces uniquement)
            new() { ["Code"] = XlsxValue.Empty, ["Quantite"] = XlsxValue.Empty }, // Ligne totalement vide
            new() { ["Code"] = "B2", ["Quantite"] = 10 }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          // 1. Par défaut (SkipEmptyRows = true) : seules les 2 lignes valides sont lues, aucune erreur
          ms.Position = 0;
          var rowsDefault = await FlashXlsxSerializer.ReadAllAsync<LigneAvecRequis>(ms);
          Assert.Equal(2, rowsDefault.Count);
          Assert.Equal("A1", rowsDefault[0].Code);
          Assert.Equal("B2", rowsDefault[1].Code);

          // 2. Avec SkipEmptyRows = false : les lignes vides sont traitées et déclenchent RequiredMissing
          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var optsNoSkip = new XlsxReadOptions<LigneAvecRequis>()
              .WithSkipEmptyRows(false)
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in XlsxErrorContext err) => errors.Add(err));

          var rowsAll = await FlashXlsxSerializer.ReadAllAsync(ms, optsNoSkip);
          Assert.Equal(4, rowsAll.Count);
          Assert.Equal(2, errors.Count);
          Assert.All(errors, e => Assert.Equal(XlsxCellStatus.RequiredMissing, e.Status));
     }
}