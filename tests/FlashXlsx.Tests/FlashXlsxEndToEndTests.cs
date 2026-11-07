using FlashXlsx;
using FlashXlsx.Abstractions;
using System.Globalization;
using Xunit;

namespace FlashXlsx.Tests;

// --- Modèles de test décorés avec [XlsxSerializable] ---

public enum StatutCommande
{
     EnAttente,
     Validee,
     Annulee
}

public sealed class MotifAnnulationRequisCondition : IXlsxCondition<CommandeTest>
{
     public bool Evaluate(CommandeTest item) => item.Statut == StatutCommande.Annulee;
}

public sealed class IgnorerRemiseSiNonValideeCondition : IXlsxCondition<CommandeTest>
{
     public bool Evaluate(CommandeTest item) => item.Statut != StatutCommande.Validee;
}

[XlsxSerializable]
public class CommandeTest
{
     [XlsxName("Numéro")]
     [XlsxRequired]
     public string Numero { get; set; } = string.Empty;

     [XlsxName("Montant HT")]
     public decimal Montant { get; set; }

     [XlsxName("Date Commande")]
     public DateOnly DateCommande { get; set; }

     public TimeOnly HeureCreation { get; set; }

     public Guid ReferenceUnique { get; set; }

     public StatutCommande Statut { get; set; }

     [XlsxRequiredIf(typeof(MotifAnnulationRequisCondition))]
     public string? MotifAnnulation { get; set; }

     [XlsxIgnoreIf(typeof(IgnorerRemiseSiNonValideeCondition))]
     public decimal Remise { get; set; }

     [XlsxBoolean(trueValues: ["Oui", "X"], falseValues: ["Non"])]
     public bool EstPayee { get; set; }

     [XlsxBoolean(trueValues: ["VIP"])] // Tout le reste => false automatiquement
     public bool ClientVip { get; set; }

     [XlsxIgnore]
     public string ChampInterne { get; set; } = "Ignoré";
}

// --- Suite de tests ---

public class FlashXlsxEndToEndTests
{
     [Fact]
     public async Task RoundTrip_TypedModel_ShouldPreserveAllTypesAndRules()
     {
          var guid = Guid.NewGuid();
          var commandes = new List<CommandeTest>
        {
            new()
            {
                Numero = "CMD-001",
                Montant = 149.90m,
                DateCommande = new DateOnly(2026, 11, 7),
                HeureCreation = new TimeOnly(14, 30, 0),
                ReferenceUnique = guid,
                Statut = StatutCommande.Validee,
                Remise = 15.00m,
                EstPayee = true,
                ClientVip = true,
                ChampInterne = "Ne doit pas être exporté"
            },
            new()
            {
                Numero = "CMD-002",
                Montant = 50.00m,
                DateCommande = new DateOnly(2026, 11, 8),
                HeureCreation = new TimeOnly(9, 15, 0),
                ReferenceUnique = Guid.Empty,
                Statut = StatutCommande.Annulee,
                MotifAnnulation = "Rupture de stock",
                Remise = 10.00m, // Doit être ignoré à l'export car Statut != Validee
                EstPayee = false,
                ClientVip = false
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, commandes);

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var readOptions = new XlsxReadOptions<CommandeTest>()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => errors.Add(err));

          var imported = await FlashXlsxSerializer.ReadAllAsync(ms, readOptions);

          Assert.Empty(errors);
          Assert.Equal(2, imported.Count);

          // Vérification ligne 1
          Assert.Equal("CMD-001", imported[0].Numero);
          Assert.Equal(149.90m, imported[0].Montant);
          Assert.Equal(new DateOnly(2026, 11, 7), imported[0].DateCommande);
          Assert.Equal(new TimeOnly(14, 30, 0), imported[0].HeureCreation);
          Assert.Equal(guid, imported[0].ReferenceUnique);
          Assert.Equal(StatutCommande.Validee, imported[0].Statut);
          Assert.Equal(15.00m, imported[0].Remise);
          Assert.True(imported[0].EstPayee);
          Assert.True(imported[0].ClientVip);
          Assert.Equal("Ignoré", imported[0].ChampInterne);

          // Vérification ligne 2 (IgnoreIf a bien laissé Remise à 0)
          Assert.Equal("CMD-002", imported[1].Numero);
          Assert.Equal("Rupture de stock", imported[1].MotifAnnulation);
          Assert.Equal(0m, imported[1].Remise);
          Assert.False(imported[1].EstPayee);
          Assert.False(imported[1].ClientVip);
     }

     [Fact]
     public async Task Translations_FluentAndDictionary_ShouldMapColumnsInAnyLanguage()
     {
          var items = new List<CommandeTest>
        {
            new() { Numero = "EN-100", Montant = 250.50m, EstPayee = true }
        };

          // Export en Anglais via Fluent Translate
          var writeEn = new XlsxWriteOptions<CommandeTest>()
              .WithSheetName("Orders_EN")
              .Translate(x => x.Numero, "Order Number")
              .Translate(x => x.Montant, "Total Amount");

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, items, writeEn);

          // Import en simulant un mapping chargé depuis une base de données
          ms.Position = 0;
          var dbMapping = new Dictionary<string, string>
          {
               ["Numero"] = "Order Number",
               ["Montant"] = "Total Amount"
          };

          var readFromDb = new XlsxReadOptions<CommandeTest>()
              .WithSheet("Orders_EN")
              .WithTranslations(dbMapping);

          var imported = await FlashXlsxSerializer.ReadAllAsync(ms, readFromDb);

          Assert.Single(imported);
          Assert.Equal("EN-100", imported[0].Numero);
          Assert.Equal(250.50m, imported[0].Montant);
     }

     [Fact]
     public async Task ErrorHandling_ReportAndContinue_ShouldCaptureAllErrorsWithoutThrowing()
     {
          // Création d'un fichier avec données invalides via le mode dictionnaire
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["Numéro"] = "", // Erreur: [XlsxRequired]
                ["Montant HT"] = "PAS_UN_NOMBRE", // Erreur: InvalidFormat
                ["Statut"] = "Annulee",
                ["MotifAnnulation"] = "", // Erreur: [XlsxRequiredIf] car Statut == Annulee
                ["EstPayee"] = "Peut-être" // Erreur: ni dans trueValues ["Oui", "X"] ni dans falseValues ["Non"]
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var options = new XlsxReadOptions<CommandeTest>()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => errors.Add(err));

          var imported = await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Single(imported);
          Assert.Equal(4, errors.Count);
          Assert.Contains(errors, e => e.Status == XlsxCellStatus.InvalidFormat && e.RawValue == "PAS_UN_NOMBRE");
          Assert.Contains(errors, e => e.Status == XlsxCellStatus.InvalidFormat && e.RawValue == "Peut-être");
          Assert.Equal(2, errors.Count(e => e.Status == XlsxCellStatus.RequiredMissing));
     }

     [Fact]
     public async Task MultiSheet_WorkbookMap_ShouldImportAndExportMultipleSchemas()
     {
          var commandes = new List<CommandeTest>
        {
            new() { Numero = "CMD-A", Montant = 10m, EstPayee = true }
        };

          var configDynamique = new List<Dictionary<string, XlsxValue>>
        {
            new() { ["Cle"] = "Version", ["Valeur"] = 2 },
            new() { ["Cle"] = "Actif", ["Valeur"] = true }
        };

          // Export multi-onglets déclaratif
          var writeMap = new XlsxWriteWorkbookMap()
              .AddSheet("Commandes", commandes)
              .AddDictionarySheet("Configuration", configDynamique);

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteWorkbookAsync(ms, writeMap);

          // Import multi-onglets déclaratif
          ms.Position = 0;
          var readMap = new XlsxReadWorkbookMap()
              .MapSheet<CommandeTest>("Commandes", out var importedCommandes)
              .MapSheetAsDictionary("Configuration", out var importedConfig);

          await FlashXlsxSerializer.ReadWorkbookAsync(ms, readMap);

          Assert.Single(importedCommandes);
          Assert.Equal("CMD-A", importedCommandes[0].Numero);

          Assert.Equal(2, importedConfig.Count);
          Assert.Equal("Version", importedConfig[0]["Cle"].AsString());
          Assert.Equal(2, importedConfig[0]["Valeur"].AsInt32());
          Assert.True(importedConfig[1]["Valeur"].AsBoolean());
     }
}