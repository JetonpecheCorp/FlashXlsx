namespace FlashXlsx.Abstractions;

/// <summary>
/// Définit une condition métier évaluée dynamiquement sur l'élément en cours de traitement.
/// Utilisée notamment par <c>[XlsxRequiredIf]</c> pour rendre une cellule obligatoire sous condition,
/// ou par <c>[XlsxIgnoreIf]</c> pour exclure dynamiquement une colonne à la lecture ou à l'écriture.
/// </summary>
/// <typeparam name="T">Le type du modèle de ligne manipulé (classe, record ou struct).</typeparam>
public interface IXlsxCondition<in T>
{
     /// <summary>
     /// Évalue si la règle métier est satisfaite pour l'instance donnée.
     /// </summary>
     /// <param name="item">L'instance de l'élément en cours de lecture ou d'écriture.</param>
     /// <returns>
     /// <see langword="true"/> si la condition s'applique (ex: champ requis ou ignoré selon l'attribut associé) ;
     /// sinon, <see langword="false"/>.
     /// </returns>
     bool Evaluate(T item);
}

/// <summary>
/// Définit une condition métier contextuelle évaluée sur l'élément en cours combiné à un contexte d'exécution externe.
/// Permet d'adapter dynamiquement les règles (validation, exclusion) en fonction de paramètres applicatifs
/// (ex: droits utilisateur, environnement d'exécution, configuration globale).
/// </summary>
/// <typeparam name="T">Le type du modèle de ligne manipulé.</typeparam>
/// <typeparam name="TContext">Le type de l'objet de contexte partagé fourni aux options de sérialisation.</typeparam>
public interface IXlsxCondition<in T, in TContext>
{
     /// <summary>
     /// Évalue si la règle métier contextuelle est satisfaite.
     /// </summary>
     /// <param name="item">L'instance de l'élément en cours de lecture ou d'écriture.</param>
     /// <param name="context">Le contexte d'exécution externe injecté dans les options.</param>
     /// <returns>
     /// <see langword="true"/> si la condition s'applique compte tenu de l'état de l'élément et du contexte ;
     /// sinon, <see langword="false"/>.
     /// </returns>
     bool Evaluate(T item, TContext context);
}