# FlashXlsx

<p align="center">
  <img src="assets/logoFlashXlsx.png" alt="FlashXlsx Logo" width="220" />
</p>

**FlashXlsx** est le moteur d'import/export de fichiers Excel (`.xlsx`) le plus rapide, léger et sécurisé pour l'écosystème .NET.

Il a été pensé pour être **simple comme un jeu d'enfant** pour un développeur débutant, tout en offrant aux architectes et experts des **performances extrêmes** grâce à la compilation native (**Native AOT**), au streaming asynchrone et à la vectorisation matérielle (SIMD).

## 🧭 Sommaire

1. [Comprendre FlashXlsx en 2 minutes (Pour tous)](#comprendre-flashxlsx-en-2-minutes-pour-tous)
2. [Pourquoi choisir FlashXlsx ?](#pourquoi-choisir-flashxlsx-)
3. [Benchmarks de référence (100 000 lignes)](#benchmarks-de-référence-100-000-lignes)
4. [Matrice complète des fonctionnalités](#matrice-complète-des-fonctionnalités)
5. [Installation](#installation)
6. [Démarrage en 3 étapes simples](#démarrage-en-3-étapes-simples)
7. [Cas d'usage réel : API Web ASP.NET Core & Ingestion massive](#cas-dusage-réel--api-web-aspnet-core--ingestion-massive)
8. [Guide complet des fonctionnalités par l'exemple](#guide-complet-des-fonctionnalités-par-lexemple)
   * [1. Modéliser vos données simplement](#1-modéliser-vos-données-simplement)
   * [2. Tableau de référence des attributs](#2-tableau-de-référence-des-attributs)
   * [3. Vérifier les données à la volée (Validation déclarative)](#3-vérifier-les-données-à-la-volée-validation-déclarative)
   * [4. Règles dynamiques selon le contexte de l'utilisateur](#4-règles-dynamiques-selon-le-contexte-de-lutilisateur)
   * [5. Noms de colonnes variables et multilingues](#5-noms-de-colonnes-variables-et-multilingues)
   * [6. Importer sans planter (Gestion d'erreurs tolérante)](#6-importer-sans-planter-gestion-derreurs-tolérante)
   * [7. Détection des fichiers protégés par mot de passe](#7-détection-des-fichiers-protégés-par-mot-de-passe)
   * [8. Nettoyage magique des devises et espaces insécables](#8-nettoyage-magique-des-devises-et-espaces-insécables)
   * [9. Mode Dynamique sans schéma : XlsxValue & Dictionnaires (Zéro-Boxing)](#9-mode-dynamique-sans-schéma--xlsxvalue--dictionnaires-zéro-boxing)
   * [10. Gérer plusieurs onglets à la fois](#10-gérer-plusieurs-onglets-à-la-fois)
   * [11. Découvrir les onglets sans charger le fichier](#11-découvrir-les-onglets-sans-charger-le-fichier)
   * [12. Cellules fusionnées lues sans effort](#12-cellules-fusionnées-lues-sans-effort)
   * [13. Créer de vrais tableaux Excel avec filtres](#13-créer-de-vrais-tableaux-excel-avec-filtres)
   * [14. Aperçu rapide et pagination de gros fichiers](#14-aperçu-rapide-et-pagination-de-gros-fichiers)
   * [15. Bouclier automatique contre les formules pirates](#15-bouclier-automatique-contre-les-formules-pirates)
   * [16. Convertisseurs sur-mesure (IXlsxConverter & calculs entre colonnes)](#16-convertisseurs-sur-mesure-ixlsxconverter--calculs-entre-colonnes)
   * [17. Découpage de listes en cellule & valeurs par défaut](#17-découpage-de-listes-en-cellule--valeurs-par-défaut)
   * [18. Barre de progression en temps réel (IProgress)](#18-barre-de-progression-en-temps-réel-iprogress)
9. [Sous le capot : Pour les experts et passionnés d'optimisation](#sous-le-capot--pour-les-experts-et-passionnés-doptimisation)
   * [L'analogie du tapis roulant (Streaming vs DOM)](#lanalogie-du-tapis-roulant-streaming-vs-dom)
   * [Scanner vectorisé UTF-8 (`ReadOnlySpan<byte>` & AVX2)](#scanner-vectorisé-utf-8-readonlyspanbyte--avx2)
   * [Accélération SIMD ASCII -> UTF-16](#accélération-simd-ascii---utf-16)
   * [Zéro allocation sur le tas (`ArrayPool`, `XlsxValue` & structs sur pile)](#zéro-allocation-sur-le-tas-arraypool-xlsxvalue--structs-sur-pile)
   * [Source Generator Roslyn (100% Native AOT & Trimming-Safe)](#source-generator-roslyn-100-native-aot--trimming-safe)
10. [Licence](#licence)

## Comprendre FlashXlsx en 2 minutes (Pour tous)

### 💡 L'analogie du camion et du tapis roulant
Imaginez que vous receviez un conteneur rempli de 100 000 colis :

* **Les anciennes bibliothèques (ClosedXML, etc.) :** Elles vident tout le camion dans votre salon d'un seul coup. Résultat : votre salon est encombré, la mémoire de votre serveur explose (parfois 800 Mo de RAM), et votre machine se bloque pendant plusieurs secondes.
* **FlashXlsx :** Il installe un **tapis roulant ultra-rapide**. Chaque colis est déchargé, contrôlé, traité, puis rangé immédiatement. La mémoire de votre serveur reste vide (moins de 20 Mo de RAM), et le traitement file à toute allure sans jamais ralentir.

### 🎯 En résumé :
* **Débutant ?** Vous écrivez 3 lignes de C# simples et tout fonctionne sans configuration obscure.
* **Expert / Architecte ?** Vous bénéficiez du code machine généré à la compilation, d'instructions vectorielles SIMD/AVX2, d'une compatibilité Native AOT totale et d'une allocation mémoire quasi-nulle.

## Pourquoi choisir FlashXlsx ?

| Critère | Bibliothèques classiques (DOM) | Bibliothèques légères (Réflexion) | FlashXlsx ⚡ |
| :--- | :--- | :--- | :--- |
| **Prise en main** | Souvent verbeuse | Simple | **Ultra-simple et expressive** |
| **Mémoire (100 000 lignes)** | ~700 à 850 Mo ⚠️ | ~260 à 1 300 Mo ⚠️ | **13 à 19 Mo (Gen1 & Gen2 = 0) 🏆** |
| **Vitesse (100 000 lignes)** | ~2,5 secondes | ~1 seconde | **~0,13 seconde (750 000 lignes/s) 🚀** |
| **Compatibilité Native AOT** | ❌ Impossible | ⚠️ Avertissements / Instable | **✅ 100 % AOT sans avertissement** |
| **Données invalides dans le fichier** | Fait planter votre application | Lève des exceptions lentes | **Capture propre sans exception (`OnError`)** |
| **Sécurité (injections de formules)** | À coder soi-même | À coder soi-même | **Automatique et transparente** |
| **Nettoyage automatique devises** | Complexe (code manuel) | Complexe | **Automatique (`"1 250 €"` -> `1250m`)** |
| **Fichiers avec mot de passe** | Crash ou erreur bizarre | Erreur ZIP incompréhensible | **Détecté proprement sans crash** |
| **Mode Dictionnaire dynamique** | `Dictionary<string, object>` (Boxing lourd) | `ExpandoObject` / Boxing | **`XlsxValue` struct (100% Zéro-Boxing)** |

## Benchmarks de référence (100 000 lignes)

Mesures officielles avec **BenchmarkDotNet v0.14.0** sur **.NET 10 (X64 RyuJIT AVX2)**, processeur standard 4 cœurs (Intel Core i5-7600K 3.80GHz).
Scénario : Écriture et lecture d'un classeur de **100 000 lignes × 7 colonnes typées** (`int`, `string`, `string`, `decimal`, `double`, `DateTime`, `bool`).

```text
BenchmarkDotNet v0.14.0, Windows 10
Intel Core i5-7600K CPU 3.80GHz, 1 CPU, 4 logical and 4 physical cores
.NET SDK 10.0.403, X64 RyuJIT AVX2
```

| Opération testée | Temps moyen (`Mean`) | Vitesse relative | Mémoire allouée | Consommation RAM | Collectes Garbage Collector |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Écriture 100k - FlashXlsx (AOT)** | **129,7 ms** | **Référence (1.00x)** | **13,49 Mo** | **Référence (1.00x)** | **500 / 500 / 500** |
| Écriture 100k - MiniExcel | 283,0 ms | 2,18x plus lent | 265,79 Mo | 19,7x plus de RAM | 78000 / 2500 / 1500 |
| Écriture 100k - ClosedXML | 2 420,0 ms | 18,65x plus lent | 724,80 Mo | 53,7x plus de RAM | 126000 / 18000 / 5000 |
| **Lecture 100k - FlashXlsx (AOT)** | **132,5 ms** | **Référence (1.02x)** | **19,89 Mo** | **Référence (1.47x)** | **6500 / 0 / 0** |
| Lecture 100k - MiniExcel | 946,3 ms | 7,29x plus lent | 857,80 Mo | 43,1x plus de RAM | 286000 / 0 / 0 |
| Lecture 100k - ClosedXML | 1 968,5 ms | 15,17x plus lent | 835,98 Mo | 61,9x plus de RAM | 186000 / 22000 / 2000 |

> **Ce que révèlent ces chiffres :**
> - **À l'écriture (129,7 ms) :** FlashXlsx génère **~770 000 lignes par seconde**. Les 13 Mo alloués correspondent presque uniquement à la taille du fichier ZIP compressé final.
> - **À la lecture (132,5 ms) :** FlashXlsx décode **~755 000 lignes par seconde** et alloue seulement 19,89 Mo (qui correspondent aux objets C# finaux créés pour vous). Le moteur interne n'alloue quasiment aucun déchet, évitant toute pause de nettoyage mémoire (Gen1 et Gen2 restent à zéro).

---

## Matrice complète des fonctionnalités

| Catégorie | Fonctionnalité | Ce que ça fait (Explication simple) |
| :--- | :--- | :--- |
| **Moteur** | **Source Generator Roslyn** | Écrit le code de lecture/écriture pendant la compilation : zéro lenteur au démarrage. |
| | **Native AOT & Trimming** | Votre application peut être compilée en fichier exécutable autonome et léger sans dépendance. |
| | **Streaming Asynchrone** | Lit les lignes une par une (`await foreach`) sans jamais saturer la RAM. |
| **Types de données** | **Primitives complètes** | Lit et écrit `int`, `long`, `decimal`, `double`, `float`, `bool`, `Guid`, `string`. |
| | **Dates modernes** | Support natif de `DateOnly`, `TimeOnly`, `DateTime` et du vieux format calendrier Mac 1904. |
| | **Nettoyage automatique devises** | Transforme automatiquement `"1 250,50 €"` ou `"$99.00"` en vrais nombres. |
| | **Booléens souples** | Comprend automatiquement `"Oui"`, `"Non"`, `"1"`, `"0"`, `"True"`, `"False"`. |
| | **Listes délimitées** | Découpe automatiquement les cellules contenant des collections (`"A;B;C"`). |
| **Validation & Conversion** | **Règles déclaratives** | Vérifie les données avec des étiquettes simples (`[XlsxRequired]`, `[XlsxRange]`, etc.). |
| | **Règles conditionnelles** | Oblige ou masque un champ seulement sous certaines conditions (`[XlsxRequiredIf]`). |
| | **Convertisseurs typés** | Prise en charge des conversions avancées et calculs entre colonnes (`IXlsxConverter`). |
| | **Contexte dynamique** | Adapte la lecture selon qui est connecté (administrateur, client, etc.). |
| **Tolérance aux erreurs** | **Mode Report & Continue** | Ne plante pas au premier problème : continue de lire et vous dresse la liste des erreurs. |
| | **Fichiers avec mot de passe** | Détecte les classeurs chiffrés dès les 8 premiers octets sans lever d'exception. |
| | **Lignes vides & fantômes** | Ignore silencieusement les lignes vides ou remplies d'espaces inutiles. |
| **Structure & Métadonnées** | **Multi-onglets** | Lit ou écrit plusieurs feuilles différentes dans le même fichier en un seul appel. |
| | **Mode Dictionnaire Zéro-Boxing** | Lit et écrit n'importe quel fichier sans classe C# grâce à la struct `XlsxValue`. |
| | **Cellules fusionnées** | Recopie intelligemment la valeur du haut sur toutes les sous-lignes fusionnées. |
| | **Vrais tableaux Excel** | Génère de superbes tableaux avec auto-filtres et lignes colorées alternées. |
| | **Aperçu & Pagination** | Permet de lire seulement de la ligne 100 à 150 pour faire un affichage paginé. |
| | **Suivi de progression** | Émet des notifications régulières (`IProgress<int>`) sans bloquer le pipeline. |
| **Sécurité** | **Protection Anti-Injections** | Empêche les utilisateurs malveillants de pirater Excel avec des formules dangereuses. |

---

## Installation

Dans votre terminal ou console du gestionnaire de paquets NuGet :

```bash
dotnet add package FlashXlsx
```

Si votre projet cible **Native AOT** (exécutable autonome natif) :
```xml
<PropertyGroup>
  <TargetFramework>net10.0</TargetFramework>
  <PublishAot>true</PublishAot>
  <Nullable>enable</Nullable>
</PropertyGroup>
```
Tout fonctionne directement : FlashXlsx embarque son générateur de code, sans aucune configuration supplémentaire.

---

## Démarrage en 3 étapes simples

### Étape 1 : Créez votre classe C#
Ajoutez l'attribut `[XlsxSerializable]` sur votre classe, et `[XlsxName]` pour nommer vos colonnes comme vous le souhaitez :

```csharp
using FlashXlsx;
using FlashXlsx.Abstractions;

[XlsxSerializable]
public class Produit
{
    [XlsxName("Référence")]
    [XlsxRequired]
    public string Reference { get; set; } = string.Empty;

    [XlsxName("Désignation")]
    public string Nom { get; set; } = string.Empty;

    [XlsxName("Prix Unitaire HT")]
    public decimal Prix { get; set; }

    [XlsxName("En Stock")]
    [XlsxBoolean(trueValues: ["Oui", "1"], falseValues: ["Non", "0"])]
    public bool Disponible { get; set; }

    [XlsxName("Date Entrée")]
    public DateOnly DateEntree { get; set; }
}
```

### Étape 2 : Écrivez un fichier Excel en 2 lignes
```csharp
var liste = new List<Produit>
{
    new() { Reference = "PRD-01", Nom = "Clavier sans fil", Prix = 49.90m, Disponible = true, DateEntree = new DateOnly(2026, 11, 10) },
    new() { Reference = "PRD-02", Nom = "Souris optique", Prix = 19.50m, Disponible = false, DateEntree = new DateOnly(2026, 11, 09) }
};

using var flux = File.Create("produits.xlsx");
await FlashXlsxSerializer.WriteAsync(flux, liste);
```

### Étape 3 : Lisez le fichier ligne par ligne
```csharp
using var flux = File.OpenRead("produits.xlsx");

await foreach (Produit p in FlashXlsxSerializer.ReadAsync<Produit>(flux))
{
    Console.WriteLine($"Produit : {p.Nom} - {p.Prix} € (Disponible : {p.Disponible})");
}
```

---

## Cas d'usage réel : API Web ASP.NET Core & Ingestion massive

Dans une vraie application d'entreprise ou une API Cloud (sur Docker, Azure ou Kubernetes), un utilisateur peut envoyer un fichier fournisseur de 100 000 lignes. 
Si le code charge tout en mémoire, le serveur manque de RAM et plante pour tout le monde.

Voici le code complet d'une API ASP.NET Core moderne : elle reçoit le fichier, traite les lignes **par petits lots de 1 000**, les enregistre en base de données et consigne les anomalies de saisie sans jamais planter :

```csharp
using FlashXlsx;
using FlashXlsx.Abstractions;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateSlimBuilder(args);
var app = builder.Build();

app.MapPost("/api/import-commandes", async (IFormFile file, [FromServices] ICommandeRepository repo, CancellationToken ct) =>
{
    if (file is null || file.Length == 0)
        return Results.BadRequest("Le fichier envoyé est vide.");

    await using var stream = file.OpenReadStream();

    var journalAnomalies = new List<object>();

    // Configuration tolérante : on enregistre les erreurs au lieu de planter
    var options = new XlsxReadOptions<CommandeImportDto>()
        .WithHeader(hasHeader: true, headerRowNumber: 1)
        .WithSkipEmptyRows(true)
        .WithErrorMode(XlsxErrorMode.ReportAndContinue)
        .WithErrorHandler((in XlsxErrorContext err) =>
        {
            journalAnomalies.Add(new
            {
                Ligne = err.RowNumber,
                Colonne = err.ColumnName ?? err.ColumnLetter,
                ValeurInvalide = err.RawValue,
                TypeErreur = err.Status.ToString()
            });
        });

    var lot = new List<CommandeImportDto>(1000);
    int totalImporte = 0;

    // Lecture fluide au fur et à mesure du téléchargement
    await foreach (var commande in FlashXlsxSerializer.ReadAsync(stream, options, ct))
    {
        lot.Add(commande);

        // Sauvegarde par lot de 1 000 lignes
        if (lot.Count >= 1000)
        {
            await repo.InsererLotAsync(lot, ct);
            totalImporte += lot.Count;
            lot.Clear();
        }
    }

    // Sauvegarde du reliquat éventuel
    if (lot.Count > 0)
    {
        await repo.InsererLotAsync(lot, ct);
        totalImporte += lot.Count;
    }

    return Results.Ok(new
    {
        Message = "Importation terminée !",
        TotalLignesImportees = totalImporte,
        NombreAnomalies = journalAnomalies.Count,
        Anomalies = journalAnomalies
    });
});

app.Run();
```

---

## Guide complet des fonctionnalités par l'exemple

### 1. Modéliser vos données simplement
Vous décidez de l'ordre exact des colonnes dans Excel, du texte de chaque en-tête ou des données à ne pas exporter :

```csharp
[XlsxSerializable]
public class Employe
{
    [XlsxColumnOrder(1)]
    [XlsxName("Matricule")]
    public string Matricule { get; set; } = string.Empty;

    [XlsxColumnOrder(2)]
    [XlsxName("Nom Complet")]
    public string Nom { get; set; } = string.Empty;

    [XlsxColumnOrder(3)]
    [XlsxName("Date de Naissance")]
    public DateOnly DateNaissance { get; set; }

    [XlsxColumnOrder(4)]
    [XlsxName("Poste Cadre")]
    [XlsxBoolean(trueValues: ["OUI", "O", "VRAI"], falseValues: ["NON", "N", "FAUX"])]
    public bool EstCadre { get; set; }

    [XlsxIgnore] // Ce champ ne sera JAMAIS lu ni écrit dans Excel
    public string CleSecreteInterne { get; set; } = string.Empty;
}
```

### 2. Tableau de référence des attributs

| Attribut | Cible | Description |
| :--- | :--- | :--- |
| `[XlsxSerializable]` | Classe / Struct | Active la génération de code Native AOT par le Source Generator. |
| `[XlsxName("Titre")]` | Propriété | Définit le nom d'en-tête de la colonne dans le fichier Excel. |
| `[XlsxColumnLetter("C")]` | Propriété | Associe la propriété directement à une lettre de colonne (ex: fichiers sans en-tête). |
| `[XlsxColumnOrder(1)]` | Propriété | Détermine la position ordonnée de la colonne lors de l'exportation. |
| `[XlsxIgnore]` | Propriété | Exclut totalement la propriété de la lecture et de l'écriture. |
| `[XlsxRequired]` | Propriété | Déclenche une erreur si la cellule est vide ou absente à la lecture. |
| `[XlsxRequiredIf(typeof(T))]` | Propriété | Rend la cellule obligatoire selon le résultat d'une condition métier `IXlsxCondition`. |
| `[XlsxIgnoreIf(typeof(T))]` | Propriété | Ignore dynamiquement la propriété si la condition métier retourne `true`. |
| `[XlsxRange(min, max)]` | Propriété | Valide les bornes numériques (`[XlsxMin]` et `[XlsxMax]` sont également disponibles). |
| `[XlsxLength(min, max)]` | Propriété | Valide la longueur minimale / maximale d'une chaîne de texte. |
| `[XlsxRegex("pattern")]` | Propriété | Vérifie la conformité du texte avec une expression régulière. |
| `[XlsxAllowedValues(...)]` | Propriété | Restreint la saisie à une liste fermée de valeurs textuelles autorisées. |
| `[XlsxBoolean(...)]` | Propriété | Définit les textes admis pour représenter `true` et `false`. |
| `[XlsxSplit(';')]` | Propriété | Découpe une cellule en collection (`string[]` ou `List<string>`) selon un séparateur. |
| `[XlsxDefault(...)]` | Propriété | Assigne une valeur par défaut statique ou dynamique (`Now`, `Today`, `NewGuid`, etc.). |
| `[XlsxConverter(typeof(T))]` | Propriété | Branche un convertisseur sur-mesure (`IXlsxConverter`, `IXlsxRowConverter`). |
| `[XlsxNumberFormat("...")]` | Propriété | Applique un masque numérique OpenXML natif (`#,##0.00 €`, `0.00%`). |
| `[XlsxFormula("...")]` | Propriété | Écrit une formule Excel calculée dynamiquement. |
| `[XlsxCase(...)]` | Propriété | Force la transformation de la casse du texte (`Upper`, `Lower`, `Title`). |
| `[XlsxTrim(...)]` | Propriété | Configure la stratégie de rognage des espaces sur les colonnes textuelles. |
| `[XlsxDecimalSeparator(',')]`| Propriété | Impose un caractère séparateur décimal strict pour cette colonne. |

### 3. Vérifier les données à la volée (Validation déclarative)
*En clair : Vous posez des barrières simples pour refuser les mauvaises saisies.*

```csharp
public sealed class MotifRequisSiRefuse : IXlsxCondition<DemandeConge>
{
    public bool Evaluate(DemandeConge item) => item.Statut == "REFUSE";
}

[XlsxSerializable]
public class DemandeConge
{
    [XlsxRequired] // Ne doit pas être vide
    [XlsxLength(min: 3, max: 15)] // Entre 3 et 15 caractères
    public string NumeroDossier { get; set; } = string.Empty;

    [XlsxAllowedValues("ATTENTE", "ACCEPTE", "REFUSE")] // Valeurs précises imposées
    public string Statut { get; set; } = "ATTENTE";

    [XlsxRange(min: 0.5, max: 45.0)] // Entre 0.5 et 45 jours
    public double DureeJours { get; set; }

    [XlsxRegex(@"^[A-Z]{2}-\d{4}$")] // Doit respecter le format "RH-2026"
    public string CodeService { get; set; } = string.Empty;

    // Obligatoire UNIQUEMENT si le statut est "REFUSE" !
    [XlsxRequiredIf(typeof(MotifRequisSiRefuse))]
    public string? MotifRefus { get; set; }
}
```

### 4. Règles dynamiques selon le contexte de l'utilisateur
*En clair : La lecture ou l'écriture s'adapte selon qui est connecté (administrateur ou simple utilisateur).*

```csharp
// Contexte passé depuis votre application
public sealed record ContexteUtilisateur(bool EstAdmin, decimal PlafondMax);

public sealed class ApprobationObligatoireCondition : IXlsxCondition<Depense, ContexteUtilisateur>
{
    public bool Evaluate(Depense item, ContexteUtilisateur ctx)
        => !ctx.EstAdmin && item.Montant > ctx.PlafondMax;
}

[XlsxSerializable]
public class Depense
{
    public string Motif { get; set; } = string.Empty;
    public decimal Montant { get; set; }

    // Requis seulement si l'utilisateur n'est pas Admin ET que le montant dépasse son plafond
    [XlsxRequiredIf(typeof(ApprobationObligatoireCondition))]
    public string? CodeApprobation { get; set; }
}

// Utilisation :
var monContexte = new ContexteUtilisateur(EstAdmin: false, PlafondMax: 500m);
var options = new XlsxReadOptions<Depense, ContexteUtilisateur>(monContexte);

await foreach (var d in FlashXlsxSerializer.ReadAsync(flux, options))
{
    // Valide dynamiquement selon le profil de l'utilisateur !
}
```

### 5. Noms de colonnes variables et multilingues
*En clair : Le même code sait lire un fichier que les colonnes soient en français, anglais ou allemand.*

```csharp
var options = new XlsxReadOptions<Produit>()
    .WithHeader(hasHeader: true)
    .WithTranslations(new Dictionary<string, string[]>
    {
        ["Reference"] = ["Part Number", "Ref", "Item Code", "Artikelnummer"],
        ["Nom"] = ["Description", "Label", "Désignation", "Bezeichnung"],
        ["Prix"] = ["Price", "Unit Price", "Prix HT", "Preis"]
    });

var produits = await FlashXlsxSerializer.ReadAllAsync(flux, options);
```

### 6. Importer sans planter (Gestion d'erreurs tolérante)
*En clair : Si une ligne sur 10 000 a une faute de frappe, on ne jette pas tout le fichier à la poubelle.*

```csharp
var erreurs = new List<XlsxErrorContext>();

var options = new XlsxReadOptions<Produit>()
    .WithErrorMode(XlsxErrorMode.ReportAndContinue) // Continue la lecture
    .WithErrorHandler((in XlsxErrorContext err) =>
    {
        // Enregistre l'erreur sans ralentir le traitement
        erreurs.Add(err);
    });

await foreach (var p in FlashXlsxSerializer.ReadAsync(flux, options))
{
    // Vos données valides arrivent ici
}

// Affichez la liste des soucis rencontrés :
foreach (var e in erreurs)
{
    Console.WriteLine($"Ligne {e.RowNumber}, Colonne '{e.ColumnName}' : valeur invalide '{e.RawValue}' ({e.Status})");
}
```

### 7. Détection des fichiers protégés par mot de passe
*En clair : Si un utilisateur envoie un fichier verrouillé, vous le savez tout de suite avec un message clair.*

Les fichiers Excel chiffrés sont des conteneurs binaires OLE2. FlashXlsx inspecte les 8 premiers octets du fichier en une fraction de milliseconde :
```csharp
using var reader = new XlsxWorkbookReader(flux);

if (reader.IsEncrypted)
{
    Console.WriteLine("Désolé, ce classeur est verrouillé par un mot de passe.");
    return;
}
```

### 8. Nettoyage magique des devises et espaces insécables
*En clair : Fini les erreurs de parsing parce qu'Excel a écrit "1 250,50 €" sous forme de texte.*

FlashXlsx nettoie automatiquement les valeurs numériques en mémoire :
* Supprime les symboles de monnaie : `€`, `$`, `£`, `CHF`, `USD`, `EUR`, `CAD`.
* Nettoie les espaces insécables invisibles (`\u00A0`, `\u202F`) souvent insérés par les logiciels comptables.
* Détecte si la virgule ou le point sert de séparateur décimal.

Une cellule contenant `" 1 250,50 € "` sera convertie directement dans votre propriété C# `decimal` valant `1250.50m`.

### 9. Mode Dynamique sans schéma : XlsxValue & Dictionnaires (Zéro-Boxing)
*En clair : Lire ou générer n'importe quel fichier Excel à la volée sans créer de classe C#, tout en conservant les performances d'un type valeur.*

Contrairement aux solutions traditionnelles qui utilisent `Dictionary<string, object>` (ce qui alloue chaque entier, booléen ou date sur le tas par boxing), FlashXlsx utilise **`XlsxValue`** : une union discriminée allouée sur la pile (`readonly struct`), sans aucun boxing.

#### A. Lecture dynamique
```csharp
await foreach (Dictionary<string, XlsxValue> ligne in FlashXlsxSerializer.ReadDictionaryAsync(flux))
{
    // Accesseurs zéro-boxing avec conversion automatique sûre :
    string? client    = ligne["Client"].AsString();
    int quantite      = ligne["Quantite"].AsInt32();
    long id           = ligne["Identifiant"].AsInt64();
    double total      = ligne["Total"].AsDouble();
    decimal prix      = ligne["Prix"].AsDecimal();
    bool valide       = ligne["Valide"].AsBoolean();
    DateOnly dateLiv  = ligne["Livraison"].AsDateOnly();
    TimeOnly heure    = ligne["Heure"].AsTimeOnly();
    DateTime creeLe   = ligne["CreeLe"].AsDateTime();
    Guid refUnique    = ligne["Guid"].AsGuid();

    // Vérification de présence :
    if (!ligne["Option"].IsEmpty) { /* ... */ }
}
```

#### B. Écriture dynamique sans allocation
Grâce aux opérateurs de conversion implicite de `XlsxValue`, la création de lignes dynamiques est totalement naturelle :
```csharp
var rows = new List<Dictionary<string, XlsxValue>>
{
    new()
    {
        ["Reference"] = "PRD-99",                      // implicit string
        ["Prix"]      = 129.99m,                       // implicit decimal (zéro boxing)
        ["Stock"]     = 42,                            // implicit int (zéro boxing)
        ["Actif"]     = true,                          // implicit bool
        ["Date"]      = new DateOnly(2026, 11, 10),    // implicit DateOnly
        ["Id"]        = Guid.NewGuid()                 // implicit Guid
    }
};

await FlashXlsxSerializer.WriteDictionaryAsync(flux, rows);
```

### 10. Gérer plusieurs onglets à la fois
*En clair : Écrire ou lire un classeur complet avec ses feuilles "Clients", "Factures" et "Paramètres" en une seule fois.*

```csharp
// --- ÉCRITURE MULTI-FEUILLES ---
var exportMap = new XlsxWriteWorkbookMap()
    .AddSheet("Clients", listeClients)
    .AddSheet("Commandes", listeCommandes);

using (var outStream = File.Create("gestion.xlsx"))
{
    await FlashXlsxSerializer.WriteWorkbookAsync(outStream, exportMap);
}

// --- LECTURE MULTI-FEUILLES ---
using (var inStream = File.OpenRead("gestion.xlsx"))
{
    var importMap = new XlsxReadWorkbookMap()
        .MapSheet<Client>("Clients", out var mesClients)
        .MapSheet<Commande>("Commandes", out var mesCommandes);

    await FlashXlsxSerializer.ReadWorkbookAsync(inStream, importMap);

    // Les variables 'mesClients' et 'mesCommandes' sont maintenant remplies !
}
```

### 11. Découvrir les onglets sans charger le fichier
*En clair : Afficher la liste des feuilles à un utilisateur avant qu'il ne choisisse laquelle importer.*

```csharp
using var reader = new XlsxWorkbookReader(flux);

foreach (XlsxSheetInfo feuille in reader.Sheets)
{
    Console.WriteLine($"Onglet #{feuille.Index} : '{feuille.Name}' (Visible : {feuille.IsVisible})");
}

// Lire ensuite uniquement la feuille choisie :
await foreach (var ligne in reader.ReadSheetAsync<Produit>("Produits 2026"))
{
    // ...
}
```

### 12. Cellules fusionnées lues sans effort
*En clair : Dans les tableaux où une catégorie est fusionnée sur 10 lignes, la valeur est recopiée sur chaque ligne.*

```csharp
var options = new XlsxReadOptions<RapportVente>()
    .WithFillMergedCells(true); // Active la propagation automatique

await foreach (var ligne in FlashXlsxSerializer.ReadAsync(flux, options))
{
    // 'ligne.Categorie' est automatiquement renseigné, même sur les lignes fusionnées !
}
```

### 13. Créer de vrais tableaux Excel avec filtres
*En clair : Votre fichier Excel s'ouvre avec un tableau officiel stylisé, des couleurs alternées et les petits boutons de filtre sur les en-têtes.*

```csharp
var options = new XlsxWriteOptions<Produit>()
    .WithExcelTable(tableName: "TableauProduits", styleName: "TableStyleMedium2", showRowStripes: true, enableAutoFilter: true);

await FlashXlsxSerializer.WriteAsync(flux, liste, options);
```

### 14. Aperçu rapide et pagination de gros fichiers
*En clair : Charger rapidement les 20 premières lignes pour une prévisualisation à l'écran, sans lire tout le fichier.*

```csharp
var options = new XlsxReadOptions<Produit>()
    .WithStartRow(2)  // Commence à la première ligne de données
    .WithMaxRows(20); // Lit 20 lignes et s'arrête immédiatement

var apercu = await FlashXlsxSerializer.ReadAllAsync(flux, options);
```

### 15. Bouclier automatique contre les formules pirates
*En clair : Vous protège contre les pirates qui glissent des formules malveillantes dans les formulaires.*

Si un utilisateur saisit un texte commençant par `=`, `+`, `-`, `@`, `\t` ou `\r` (par exemple `=cmd|' /C calc'!A0`), Excel tente normalement de l'exécuter à l'ouverture.
FlashXlsx repère ces caractères et applique sa politique de sécurité (neutralisation automatique par apostrophe ou levée de violation) sans configuration superflue.

### 16. Convertisseurs sur-mesure (IXlsxConverter & calculs entre colonnes)
*En clair : Vous gérez des formats de données exotiques ou calculez une colonne à partir des autres sans aucune allocation mémoire.*

FlashXlsx propose trois niveaux de convertisseurs zéro-allocation :
* **`IXlsxConverter<TProp>`** : conversion directe cellule ↔ propriété via `ReadOnlySpan<char>`.
* **`IXlsxRowConverter<TRow, TProp>`** : accès aux autres propriétés déjà lues de la même ligne.
* **`IXlsxRowConverter<TRow, TContext, TProp>`** : accès combiné à la ligne et à votre contexte applicatif.

```csharp
// Exemple : Calculer le montant TTC à partir du HT et de la TVA de la ligne
public sealed class CalculTtcConverter : IXlsxRowConverter<LigneFacture, decimal>
{
    public bool TryRead(ReadOnlySpan<char> rawValue, in LigneFacture row, XlsxReadOptions options, out decimal value)
    {
        // Si la cellule est vide, on calcule dynamiquement : HT * (1 + TVA/100)
        value = Math.Round(row.MontantHt * (1m + (row.TauxTva / 100m)), 2);
        return true;
    }

    public string? Write(in decimal value, in LigneFacture row, XlsxWriteOptions options)
        => value.ToString("F2", options.Culture);
}

[XlsxSerializable]
public class LigneFacture
{
    public decimal MontantHt { get; set; }
    public decimal TauxTva { get; set; }

    [XlsxConverter(typeof(CalculTtcConverter))]
    public decimal MontantTtc { get; set; }
}
```

### 17. Découpage de listes en cellule & valeurs par défaut
*En clair : Lire une liste de tags séparés par des points-virgules et injecter des valeurs automatiques si la cellule est vide.*

```csharp
[XlsxSerializable]
public class Article
{
    public string Titre { get; set; } = string.Empty;

    // Découpe automatiquement "Tech;Dev;DotNet" en tableau de chaînes
    [XlsxSplit(separator: ';', TrimEntries = true, RemoveEmptyEntries = true)]
    public string[] Tags { get; set; } = [];

    // Génère un nouveau GUID si la cellule est vide à l'import
    [XlsxDefault(XlsxDefaultKind.NewGuid)]
    public Guid Identifiant { get; set; }

    // Attribue la date du jour si non renseignée
    [XlsxDefault(XlsxDefaultKind.Today)]
    public DateOnly DateCreation { get; set; }

    // Format numérique comptable natif Excel (la cellule reste calculable)
    [XlsxNumberFormat("#,##0.00 €")]
    public decimal PrixTTC { get; set; }
}
```

### 18. Barre de progression en temps réel (IProgress)
*En clair : Suivre l'avancement précis du traitement de gros fichiers pour l'afficher dans votre console ou votre interface.*

```csharp
var progress = new Progress<int>(lignesLues =>
{
    Console.Write($"\rLignes traitées : {lignesLues:N0}");
});

var options = new XlsxReadOptions<Produit>()
    .WithProgress(progress, intervalRows: 5000); // Notification toutes les 5 000 lignes

await foreach (var p in FlashXlsxSerializer.ReadAsync(flux, options))
{
    // Traitement fluide
}
```

---

## Sous le capot : Pour les experts et passionnés d'optimisation

Cette section détaille les choix d'ingénierie qui permettent à FlashXlsx d'atteindre **750 000 lignes/seconde** avec moins de 20 Mo de RAM.

```text
Flux .xlsx (ZipArchive)
       │
       ▼
 [ Décompression Deflate (zlib) ]
       │
       ▼  Buffer 128 Ko loué sur ArrayPool<byte>
 ┌────────────────────────────────────────────────────────┐
 │            Utf8SheetRowReader                          │
 │  • Vectorisation SIMD (SearchValues<byte> - AVX2)      │
 │  • Détection directe des balises <row>, <c>, <v>       │
 └─────────────────────────┬──────────────────────────────┘
                           │
      ┌────────────────────┴─────────────────────┐
      ▼                                          ▼
[ Nombres, Dates & Booléens ]           [ Textes & Chaînes ]
  Utf8Parser / Ascii.ToUtf16              Encoding.UTF8 (Tampon partagé)
  ZÉRO allocation sur le tas               HtmlDecode sans allocations inutiles
      │                                          │
      └────────────────────┬─────────────────────┘
                           ▼
          [ IXlsxTypeMap<T> Compilée AOT ]
          • Affectation directe par index (switch jump-table)
          • Zéro réflexion, zéro boxing, zéro Activator
          • Validation compilée inline
                           │
                           ▼
                  yield return instance;
```

### L'analogie du tapis roulant (Streaming vs DOM)
Alors que les bibliothèques DOM construisent un arbre géant de nœuds XML en mémoire (des millions d'objets `XmlNode` ou `XElement`), FlashXlsx est un **consommateur linéaire en streaming**. Une ligne lue est immédiatement convertie en votre objet métier et transmise à votre boucle. Dès que la ligne suivante arrive, les tampons internes sont réutilisés.

### Scanner vectorisé UTF-8 (`ReadOnlySpan<byte>` & AVX2)
FlashXlsx court-circuite `System.Xml.XmlReader` lors de la lecture des lignes :
* Un fichier `.xlsx` est nativement compressé en UTF-8.
* Plutôt que de décoder d'abord les octets UTF-8 en caractères UTF-16 (`string`), **`Utf8SheetRowReader`** travaille directement sur les tranches d'octets `ReadOnlySpan<byte>` issues du décompresseur ZIP.
* **Vectorisation matérielle :** L'instruction vectorielle AVX2 `SearchValues<byte>` scanne les délimiteurs de balises `<` et `>` par blocs de **32 octets par cycle CPU**.
* Les coordonnées de cellules Excel (`r="C14"`) sont décodées directement depuis les octets ASCII en arithmétique entière, sans créer d'objet `string`.

### Accélération SIMD ASCII -> UTF-16
Dans un classeur d'entreprise, plus de 80 % des données sont des chiffres (`12345.67`), des dates ISO, des booléens ou des codes alphanumériques.
FlashXlsx utilise `System.Text.Ascii.ToUtf16` : sur les processeurs modernes, la conversion s'effectue en une seule passe vectorielle SIMD, contournant les tables de correspondances complexes d'UTF-8.

### Zéro allocation sur le tas (`ArrayPool`, `XlsxValue` & structs sur pile)
* Les mémoires tampons de lecture (128 Ko) et les tableaux de cellules temporaires sont loués via `ArrayPool<byte>.Shared` et `ArrayPool<char>.Shared`.
* Les rapports d'erreur `XlsxErrorContext` sont des `readonly record struct` transmis par référence (`in`). Lors d'une erreur de saisie, **aucun objet `Exception` n'est instancié**, préservant ainsi le ramasse-miettes.
* **Le mode dynamique sans boxing (`XlsxValue`) :** En mode dictionnaire, `XlsxValue` stocke les entiers, flottants, booléens et dates sous forme de bits bruts (`long _bits`), de `decimal` ou de `Guid` continus dans une même structure de pile avec disposition mémoire compacte (`[StructLayout(LayoutKind.Auto)]`). Aucune allocation sur le tas (Gen0) n'a lieu lors de la lecture ou de l'écriture dynamique.

### Source Generator Roslyn (100% Native AOT & Trimming-Safe)
Au moment où vous compilez votre projet (`dotnet build`), le générateur de code FlashXlsx inspecte vos classes annotées avec `[XlsxSerializable]` et génère du code C# pur :
1. Une méthode `CreateInstance()` instanciant directement l'objet (`new MonObjet()`).
2. Un aiguillage `switch(propertyIndex)` direct qui assigne les valeurs champ par champ.
3. Des vérifications de validation compilées sous forme de conditions `if` directes.

**Résultat :**
* Zéro appel à `System.Reflection`.
* Zéro appel à `Type.GetProperties()` ou `PropertyInfo.SetValue()`.
* Zéro compilation dynamique d'expressions (`Expression.Compile()`).
* **Zéro avertissement** de compatibilité AOT ou Trimming (`IL2026`, `IL3050`).

---

## Licence

Ce projet est distribué sous la **licence MIT**. 

Vous pouvez l'utiliser, le modifier et l'intégrer librement dans tous vos projets personnels, d'entreprise ou commerciaux sans aucune redevance.