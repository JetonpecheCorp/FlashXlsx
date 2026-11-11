using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using FlashXlsx;
using FlashXlsx.Abstractions;
using Xunit;

namespace FlashXlsx.Tests;

// --- Modèles et conditions avec contexte externe ---

public sealed record ImportContext(bool EstAdministrateur, decimal PlafondValidation);

public sealed class ApprobationRequiseSiPlafondDepasseCondition : IXlsxCondition<FactureContextuelle, ImportContext>
{
     public bool Evaluate(FactureContextuelle item, ImportContext context)
         => !context.EstAdministrateur && item.Montant > context.PlafondValidation;
}

public sealed class MasquerMargeSiNonAdminCondition : IXlsxCondition<FactureContextuelle, ImportContext>
{
     public bool Evaluate(FactureContextuelle item, ImportContext context)
         => !context.EstAdministrateur;
}

[XlsxSerializable]
public class FactureContextuelle
{
     [XlsxRequired]
     public string Reference { get; set; } = string.Empty;

     public decimal Montant { get; set; }

     [XlsxRequiredIf(typeof(ApprobationRequiseSiPlafondDepasseCondition))]
     public string? CodeApprobation { get; set; }

     [XlsxIgnoreIf(typeof(MasquerMargeSiNonAdminCondition))]
     public decimal MargeInterne { get; set; }
}

[XlsxSerializable]
public class LigneVolume
{
     public int Index { get; set; }
     public string Libelle { get; set; } = string.Empty;
     public double Valeur { get; set; }
}

// --- Suite de tests avancés ---

public class FlashXlsxAdvancedTests
{
     [Fact]
     public async Task ExternalContext_ShouldAdaptValidationAndFilteringPerUserRole()
     {
          var factures = new List<FactureContextuelle>
        {
            new() { Reference = "FAC-01", Montant = 500m, MargeInterne = 120m },
            new() { Reference = "FAC-02", Montant = 2500m, MargeInterne = 600m } // Dépasse 1000m sans CodeApprobation
        };

          // 1. Export en tant que non-administrateur : MargeInterne doit être masquée (0)
          var ctxUser = new ImportContext(EstAdministrateur: false, PlafondValidation: 1000m);
          var writeOpts = new XlsxWriteOptions<FactureContextuelle, ImportContext>(ctxUser);

        await using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, factures, writeOpts);

          // 2. Import en tant qu'utilisateur standard (plafond 1000m) => FAC-02 doit déclencher RequiredMissing sur CodeApprobation
          ms.Position = 0;
          var userErrors = new List<XlsxErrorContext>();
          var readUserOpts = new XlsxReadOptions<FactureContextuelle, ImportContext>(ctxUser)
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => userErrors.Add(err));

          var importedForUser = await FlashXlsxSerializer.ReadAllAsync(ms, readUserOpts);

          Assert.Equal(2, importedForUser.Count);
          Assert.Equal(0m, importedForUser[0].MargeInterne);
          Assert.Equal(0m, importedForUser[1].MargeInterne);
          Assert.Single(userErrors);
          Assert.Equal(XlsxCellStatus.RequiredMissing, userErrors[0].Status);
          Assert.Equal("CodeApprobation", userErrors[0].ColumnName);

          // 3. Import du même fichier en tant qu'administrateur => aucune erreur de plafond
          ms.Position = 0;
          var ctxAdmin = new ImportContext(EstAdministrateur: true, PlafondValidation: 1000m);
          var adminErrors = new List<XlsxErrorContext>();
          var readAdminOpts = new XlsxReadOptions<FactureContextuelle, ImportContext>(ctxAdmin)
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => adminErrors.Add(err));

          var importedForAdmin = await FlashXlsxSerializer.ReadAllAsync(ms, readAdminOpts);

          Assert.Equal(2, importedForAdmin.Count);
          Assert.Empty(adminErrors);
     }

     [Fact]
     public async Task LargeStream_100kRows_ShouldStreamWithoutBufferingAllInMemory()
     {
          const int rowCount = 100_000;

          static async IAsyncEnumerable<LigneVolume> GenerateRowsAsync(
              int count,
              [EnumeratorCancellation] CancellationToken ct = default)
          {
               for (int i = 1; i <= count; i++)
               {
                    ct.ThrowIfCancellationRequested();
                    yield return new LigneVolume
                    {
                         Index = i,
                         Libelle = $"Article_{i}",
                         Valeur = i * 1.5
                    };
               }
               await Task.CompletedTask;
          }

        await using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, GenerateRowsAsync(rowCount));

          ms.Position = 0;
          int readCount = 0;
          double checksum = 0;

          await foreach (var row in FlashXlsxSerializer.ReadAsync<LigneVolume>(ms))
          {
               readCount++;
               checksum += row.Index;
          }

          Assert.Equal(rowCount, readCount);
          Assert.Equal((double)rowCount * (rowCount + 1) / 2, checksum);
     }

     [Fact]
     public async Task CustomHeaderRow_And_SparseCells_ShouldMapCorrectColumns()
     {
        // Construction d'un fichier XLSX minimal avec un titre en ligne 1, une ligne vide en 2,
        // les en-têtes en ligne 3, et une cellule manquante (B4 omise dans le XML)
        await using var ms = new MemoryStream();
        await using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
          {
               AddEntry(archive, "[Content_Types].xml",
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                   "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                   "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                   "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                   "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                   "</Types>");

               AddEntry(archive, "_rels/.rels",
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                   "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                   "</Relationships>");

               AddEntry(archive, "xl/workbook.xml",
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                   "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");

               AddEntry(archive, "xl/_rels/workbook.xml.rels",
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                   "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                   "</Relationships>");

               // Ligne 1 = Titre décoratif, Ligne 3 = En-têtes, Ligne 4 = A4 et C4 présents (B4 sauté)
               AddEntry(archive, "xl/worksheets/sheet1.xml",
                   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                   "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                   "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Rapport Annuel 2026</t></is></c></row>" +
                   "<row r=\"3\">" +
                   "  <c r=\"A3\" t=\"inlineStr\"><is><t>Index</t></is></c>" +
                   "  <c r=\"B3\" t=\"inlineStr\"><is><t>Libelle</t></is></c>" +
                   "  <c r=\"C3\" t=\"inlineStr\"><is><t>Valeur</t></is></c>" +
                   "</row>" +
                   "<row r=\"4\">" +
                   "  <c r=\"A4\"><v>42</v></c>" +
                   "  <c r=\"C4\"><v>99.5</v></c>" +
                   "</row>" +
                   "</sheetData></worksheet>");
          }

          ms.Position = 0;
          var options = new XlsxReadOptions<LigneVolume>()
              .WithHeader(hasHeader: true, headerRowNumber: 3);

          var rows = await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Single(rows);
          Assert.Equal(42, rows[0].Index);
          Assert.Equal(string.Empty, rows[0].Libelle); // B4 était absent dans le XML
          Assert.Equal(99.5, rows[0].Valeur);
     }

     [Fact]
     public async Task ErrorHandling_StopOnFirstError_ShouldHaltImmediately()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new() { ["Index"] = 1, ["Libelle"] = "OK", ["Valeur"] = 10.0 },
            new() { ["Index"] = "INVALIDE", ["Libelle"] = "Erreur", ["Valeur"] = 20.0 },
            new() { ["Index"] = 3, ["Libelle"] = "JamaisLu", ["Valeur"] = 30.0 }
        };

        await using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var options = new XlsxReadOptions<LigneVolume>()
              .WithErrorMode(XlsxErrorMode.StopOnFirstError)
              .WithErrorHandler((in err) => errors.Add(err));

          var results = await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Single(results); // Seule la ligne 1 a été retournée
          Assert.Equal(1, results[0].Index);
          Assert.Single(errors); // Une seule erreur capturée avant l'arrêt
          Assert.Equal(3, errors[0].RowNumber); // Ligne 1 = header, Ligne 2 = OK, Ligne 3 = INVALIDE
     }

     private static void AddEntry(ZipArchive archive, string path, string xml)
     {
          var entry = archive.CreateEntry(path);
          using var s = entry.Open();
          s.Write(Encoding.UTF8.GetBytes(xml));
     }
}