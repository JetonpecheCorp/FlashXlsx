using System;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Convertisseur personnalisé sans état conçu pour transformer une valeur textuelle brute de cellule en type C# <typeparamref name="TProperty"/> (zéro-allocation) et inversement.
/// </summary>
/// <typeparam name="TProperty">Le type C# de la propriété ciblée par la conversion.</typeparam>
public interface IXlsxConverter<TProperty>
{
     /// <summary>
     /// Tente d'analyser et de convertir le contenu textuel brut d'une cellule Excel vers le type cible.
     /// </summary>
     /// <param name="rawValue">Le segment de mémoire contenant le texte brut extrait de la cellule (évite les allocations de chaînes intermédiaires).</param>
     /// <param name="options">Les options de lecture actives (culture, séparateurs, tolérances).</param>
     /// <param name="value">Lorsque la méthode retourne <see langword="true"/>, contient la valeur convertie ; sinon, la valeur par défaut de <typeparamref name="TProperty"/>.</param>
     /// <returns><see langword="true"/> si la conversion a réussi ; <see langword="false"/> si le format est invalide ou incompatible.</returns>
     bool TryRead(ReadOnlySpan<char> rawValue, XlsxReadOptions options, out TProperty value);

     /// <summary>
     /// Convertit l'instance de la propriété en sa représentation textuelle formatée pour l'écriture dans la cellule OpenXML.
     /// </summary>
     /// <param name="value">La valeur de la propriété à sérialiser.</param>
     /// <param name="options">Les options d'écriture actives (culture, encodage, format par défaut).</param>
     /// <returns>La chaîne de caractères prête à être injectée dans la cellule, ou <see langword="null"/> pour écrire une cellule vide.</returns>
     string? Write(in TProperty value, XlsxWriteOptions options);
}

/// <summary>
/// Convertisseur avancé dépendant de l'état de la ligne.
/// Permet de valider ou de calculer la valeur d'une colonne en accédant aux autres propriétés déjà lues sur l'objet <typeparamref name="TRow"/>.
/// Exécuté après le parsing des colonnes standard et indépendantes.
/// </summary>
/// <typeparam name="TRow">Le type de l'objet représentant l'intégralité de la ligne.</typeparam>
/// <typeparam name="TProperty">Le type C# de la propriété ciblée par la conversion.</typeparam>
public interface IXlsxRowConverter<TRow, TProperty>
{
     /// <summary>
     /// Tente d'analyser la valeur de la cellule en s'appuyant sur les données déjà désérialisées de la même ligne.
     /// </summary>
     /// <param name="rawValue">Le segment de mémoire contenant le texte brut extrait de la cellule.</param>
     /// <param name="row">L'instance de la ligne en cours de construction, contenant les valeurs préalablement parsées.</param>
     /// <param name="options">Les options de lecture actives.</param>
     /// <param name="value">Lorsque la méthode retourne <see langword="true"/>, contient la valeur convertie ; sinon, la valeur par défaut de <typeparamref name="TProperty"/>.</param>
     /// <returns><see langword="true"/> si la conversion a réussi ; sinon, <see langword="false"/>.</returns>
     bool TryRead(ReadOnlySpan<char> rawValue, in TRow row, XlsxReadOptions options, out TProperty value);

     /// <summary>
     /// Calcule ou formate la valeur à inscrire dans la cellule en exploitant l'ensemble des données de la ligne parente.
     /// </summary>
     /// <param name="value">La valeur courante de la propriété.</param>
     /// <param name="row">L'instance complète de la ligne exportée.</param>
     /// <param name="options">Les options d'écriture actives.</param>
     /// <returns>La chaîne formatée à insérer dans la cellule, ou <see langword="null"/> si la cellule doit rester vide.</returns>
     string? Write(in TProperty value, in TRow row, XlsxWriteOptions options);
}

/// <summary>
/// Convertisseur contextuel complet ayant accès à la fois aux données de la ligne <typeparamref name="TRow"/>
/// et à un contexte d'exécution externe <typeparamref name="TContext"/> (ex: services, dictionnaires de mapping, identifiants de session).
/// </summary>
/// <typeparam name="TRow">Le type de l'objet représentant l'intégralité de la ligne.</typeparam>
/// <typeparam name="TContext">Le type de l'objet de contexte partagé passé aux options de sérialisation.</typeparam>
/// <typeparam name="TProperty">Le type C# de la propriété ciblée par la conversion.</typeparam>
public interface IXlsxRowConverter<TRow, TContext, TProperty>
{
     /// <summary>
     /// Tente d'analyser la valeur de la cellule en s'appuyant sur l'état de la ligne et le contexte externe.
     /// </summary>
     /// <param name="rawValue">Le segment de mémoire contenant le texte brut extrait de la cellule.</param>
     /// <param name="row">L'instance de la ligne en cours de construction.</param>
     /// <param name="context">Le contexte d'exécution injecté lors de l'appel de lecture.</param>
     /// <param name="options">Les options typées de lecture actives.</param>
     /// <param name="value">Lorsque la méthode retourne <see langword="true"/>, contient la valeur convertie ; sinon, la valeur par défaut de <typeparamref name="TProperty"/>.</param>
     /// <returns><see langword="true"/> si la conversion contextuelle a réussi ; sinon, <see langword="false"/>.</returns>
     bool TryRead(ReadOnlySpan<char> rawValue, in TRow row, in TContext context, XlsxReadOptions<TRow, TContext> options, out TProperty value);

     /// <summary>
     /// Calcule ou formate la valeur à exporter en exploitant les données de la ligne et le contexte externe.
     /// </summary>
     /// <param name="value">La valeur courante de la propriété.</param>
     /// <param name="row">L'instance complète de la ligne exportée.</param>
     /// <param name="context">Le contexte d'exécution injecté lors de l'appel d'écriture.</param>
     /// <param name="options">Les options typées d'écriture actives.</param>
     /// <returns>La chaîne formatée à insérer dans la cellule, ou <see langword="null"/> si la cellule doit rester vide.</returns>
     string? Write(in TProperty value, in TRow row, in TContext context, XlsxWriteOptions<TRow, TContext> options);
}