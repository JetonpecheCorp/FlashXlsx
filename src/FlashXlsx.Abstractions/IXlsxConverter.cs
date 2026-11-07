using System;

namespace FlashXlsx.Abstractions;

/// <summary>
/// Convertisseur personnalisé sans état pour transformer une cellule Excel en type C# TProperty et inversement.
/// </summary>
public interface IXlsxConverter<TProperty>
{
     bool TryRead(ReadOnlySpan<char> rawValue, XlsxReadOptions options, out TProperty value);
     string? Write(in TProperty value, XlsxWriteOptions options);
}

/// <summary>
/// Convertisseur avancé ayant accès à l'ensemble des autres propriétés de la ligne TRow.
/// Évalué après le parsing des colonnes standard de la ligne.
/// </summary>
public interface IXlsxRowConverter<TRow, TProperty>
{
     bool TryRead(ReadOnlySpan<char> rawValue, in TRow row, XlsxReadOptions options, out TProperty value);
     string? Write(in TProperty value, in TRow row, XlsxWriteOptions options);
}

/// <summary>
/// Convertisseur avancé ayant accès à la ligne TRow ainsi qu'au contexte externe TContext.
/// </summary>
public interface IXlsxRowConverter<TRow, TContext, TProperty>
{
     bool TryRead(ReadOnlySpan<char> rawValue, in TRow row, in TContext context, XlsxReadOptions<TRow, TContext> options, out TProperty value);
     string? Write(in TProperty value, in TRow row, in TContext context, XlsxWriteOptions<TRow, TContext> options);
}
