using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

namespace FlashXlsx.Abstractions.Internal;

/// <summary>
/// Registre statique alimenté à la compilation via [ModuleInitializer].
/// Garantit un lookup O(1) 100% Native AOT sans réflexion.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class XlsxTypeRegistry
{
     private static readonly ConcurrentDictionary<RuntimeTypeHandle, object> Maps = new();

     public static void Register<T>(IXlsxTypeMap<T> map)
     {
          Maps[typeof(T).TypeHandle] = map;
     }

     public static bool TryGetMap<T>([NotNullWhen(true)] out IXlsxTypeMap<T>? map)
     {
          if (Maps.TryGetValue(typeof(T).TypeHandle, out var obj) && obj is IXlsxTypeMap<T> typedMap)
          {
               map = typedMap;
               return true;
          }

          map = null;
          return false;
     }

     public static IXlsxTypeMap<T, TContext> GetRequiredContextMap<T, TContext>()
     {
          var map = GetRequiredMap<T>();
          if (map is IXlsxTypeMap<T, TContext> contextMap)
               return contextMap;

          throw new InvalidOperationException(
              $"Le mapper généré pour '{typeof(T).FullName}' ne supporte pas le contexte '{typeof(TContext).FullName}'. " +
              $"Vérifiez que les attributs [XlsxRequiredIf] ou [XlsxIgnoreIf] utilisent bien une condition implémentant IXlsxCondition<{typeof(T).Name}, {typeof(TContext).Name}>.");
     }

     public static IXlsxTypeMap<T> GetRequiredMap<T>()
     {
          if (TryGetMap<T>(out var map))
               return map;

          throw new InvalidOperationException(
              $"Aucun mapper FlashXlsx n'a été généré pour le type '{typeof(T).FullName}'. " +
              "Assurez-vous que le type est décoré avec [XlsxSerializable] ou utilisé directement avec FlashXlsx.");
     }
}
