namespace FlashXlsx.Abstractions;

/// <summary>
/// Définit une condition métier évaluée sur l'élément en cours.
/// </summary>
public interface IXlsxCondition<in T>
{
     bool Evaluate(T item);
}

/// <summary>
/// Définit une condition métier évaluée sur l'élément en cours avec un contexte externe.
/// </summary>
public interface IXlsxCondition<in T, in TContext>
{
     bool Evaluate(T item, TContext context);
}
