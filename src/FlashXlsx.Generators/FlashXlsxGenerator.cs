using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FlashXlsx.Generators;

[Generator(LanguageNames.CSharp)]
public sealed class FlashXlsxGenerator : IIncrementalGenerator
{
     private const string SerializableAttrName = "FlashXlsx.Abstractions.XlsxSerializableAttribute";
     private const string NameAttrName = "FlashXlsx.Abstractions.XlsxNameAttribute";
     private const string IgnoreAttrName = "FlashXlsx.Abstractions.XlsxIgnoreAttribute";
     private const string RequiredAttrName = "FlashXlsx.Abstractions.XlsxRequiredAttribute";
     private const string RequiredIfAttrName = "FlashXlsx.Abstractions.XlsxRequiredIfAttribute";
     private const string IgnoreIfAttrName = "FlashXlsx.Abstractions.XlsxIgnoreIfAttribute";
     private const string BooleanAttrName = "FlashXlsx.Abstractions.XlsxBooleanAttribute";

     private const string CaseAttrName = "FlashXlsx.Abstractions.XlsxCaseAttribute";
     private const string TrimAttrName = "FlashXlsx.Abstractions.XlsxTrimAttribute";
     private const string RangeAttrName = "FlashXlsx.Abstractions.XlsxRangeAttribute";
     private const string MinAttrName = "FlashXlsx.Abstractions.XlsxMinAttribute";
     private const string MaxAttrName = "FlashXlsx.Abstractions.XlsxMaxAttribute";
     private const string LengthAttrName = "FlashXlsx.Abstractions.XlsxLengthAttribute";
     private const string RegexAttrName = "FlashXlsx.Abstractions.XlsxRegexAttribute";
     private const string AllowedValuesAttrName = "FlashXlsx.Abstractions.XlsxAllowedValuesAttribute";
     private const string ColumnOrderAttrName = "FlashXlsx.Abstractions.XlsxColumnOrderAttribute";
     private const string FormulaAttrName = "FlashXlsx.Abstractions.XlsxFormulaAttribute";
     private const string EnumAttrName = "FlashXlsx.Abstractions.XlsxEnumAttribute";
     private const string EnumMemberAttrName = "FlashXlsx.Abstractions.XlsxEnumMemberAttribute";

     private const string DecimalSeparatorAttrName = "FlashXlsx.Abstractions.XlsxDecimalSeparatorAttribute";
     private const string DefaultAttrName = "FlashXlsx.Abstractions.XlsxDefaultAttribute";
     private const string FormatAttrName = "FlashXlsx.Abstractions.XlsxFormatAttribute";
     private const string ColumnLetterAttrName = "FlashXlsx.Abstractions.XlsxColumnLetterAttribute";
     private const string SplitAttrName = "FlashXlsx.Abstractions.XlsxSplitAttribute";
     private const string ConverterAttrName = "FlashXlsx.Abstractions.XlsxConverterAttribute";

     private static readonly DiagnosticDescriptor EmptyBooleanArraysError = new(
         id: "FX001",
         title: "Attribut [XlsxBoolean] invalide",
         messageFormat: "La propriété '{0}' déclare [XlsxBoolean] sans aucune valeur dans trueValues ni falseValues",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor InvalidConditionTypeError = new(
         id: "FX002",
         title: "Type de condition invalide",
         messageFormat: "Le type de condition '{0}' sur la propriété '{1}' doit avoir un constructeur sans paramètre et implémenter IXlsxCondition<{2}> ou IXlsxCondition<{2}, TContext>",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor UnsupportedPropertyTypeError = new(
         id: "FX003",
         title: "Type de propriété non supporté",
         messageFormat: "Le type '{0}' de la propriété '{1}' n'est pas supporté par FlashXlsx. Ajoutez [XlsxIgnore] pour l'exclure.",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor InvalidRegexPatternError = new(
         id: "FX004",
         title: "Expression régulière invalide",
         messageFormat: "L'expression régulière '{0}' sur la propriété '{1}' est invalide : {2}",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor InvalidRangeBoundsError = new(
         id: "FX005",
         title: "Bornes de validation invalides",
         messageFormat: "La borne minimale ({0}) ne peut pas être supérieure à la borne maximale ({1}) sur la propriété '{2}'",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor InvalidColumnLetterError = new(
         id: "FX006",
         title: "Lettre de colonne Excel invalide",
         messageFormat: "La lettre de colonne '{0}' sur la propriété '{1}' est invalide (attendu : A-Z, AA-ZZ, etc.)",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor MissingConstructorMappingError = new(
         id: "FX007",
         title: "Constructeur non mappable",
         messageFormat: "Le paramètre '{0}' du constructeur de '{1}' ne correspond à aucune propriété mappée et n'a pas de valeur par défaut",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     private static readonly DiagnosticDescriptor InvalidConverterTypeError = new(
         id: "FX008",
         title: "Convertisseur personnalisé invalide",
         messageFormat: "Le type '{0}' sur la propriété '{1}' doit avoir un constructeur sans paramètre et implémenter IXlsxConverter<{2}>, IXlsxRowConverter<{3}, {2}> ou IXlsxRowConverter<{3}, TContext, {2}>",
         category: "FlashXlsx",
         DiagnosticSeverity.Error,
         isEnabledByDefault: true);

     public void Initialize(IncrementalGeneratorInitializationContext context)
     {
          var typeDeclarations = context.SyntaxProvider.ForAttributeWithMetadataName(
              SerializableAttrName,
              predicate: static (node, _) => node is TypeDeclarationSyntax,
              transform: static (ctx, _) => ExtractTypeModel(ctx));

          var allModels = typeDeclarations.Collect();

          context.RegisterSourceOutput(allModels, static (spc, results) =>
          {
               var validModels = ImmutableArray.CreateBuilder<TypeModel>();

               foreach (var res in results)
               {
                    foreach (var diag in res.Diagnostics)
                         spc.ReportDiagnostic(diag);

                    if (res.Model is not null && res.Diagnostics.Length == 0)
                    {
                         validModels.Add(res.Model);
                         string code = XlsxCodeEmitter.EmitTypeMap(res.Model);
                         string fileName = $"{res.Model.FullTypeName.Replace("global::", "").Replace(".", "_")}.XlsxMap.g.cs";
                         spc.AddSource(fileName, SourceText.From(code, Encoding.UTF8));
                    }
               }

               if (validModels.Count > 0)
               {
                    string initCode = XlsxCodeEmitter.EmitModuleInitializer(validModels.ToImmutable());
                    spc.AddSource("FlashXlsx.ModuleInitializer.g.cs", SourceText.From(initCode, Encoding.UTF8));
               }
          });
     }

     private static GenerationResult ExtractTypeModel(GeneratorAttributeSyntaxContext ctx)
     {
          var diagnostics = ImmutableArray.CreateBuilder<Diagnostic>();

          if (ctx.TargetSymbol is not INamedTypeSymbol typeSymbol)
               return new GenerationResult(null, diagnostics.ToImmutable());

          string fullTypeName = typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
          string ns = typeSymbol.ContainingNamespace.IsGlobalNamespace ? string.Empty : typeSymbol.ContainingNamespace.ToDisplayString();

          var publicCtors = typeSymbol.InstanceConstructors
              .Where(c => c.DeclaredAccessibility == Accessibility.Public && !IsCopyConstructor(c, typeSymbol))
              .ToList();

          bool hasParameterlessCtor = typeSymbol.IsValueType || publicCtors.Any(c => c.Parameters.Length == 0);
          IMethodSymbol? chosenParameterizedCtor = null;

          if (!hasParameterlessCtor && publicCtors.Count > 0)
          {
               chosenParameterizedCtor = publicCtors
                   .OrderByDescending(c => c.Parameters.Length)
                   .First();
          }

          var ctorParamNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
          if (chosenParameterizedCtor is not null)
          {
               for (int pIdx = 0; pIdx < chosenParameterizedCtor.Parameters.Length; pIdx++)
                    ctorParamNames[chosenParameterizedCtor.Parameters[pIdx].Name] = pIdx;
          }

          var rawProps = ImmutableArray.CreateBuilder<(PropertyModel Prop, int DeclarationIndex)>();
          string? detectedContextType = null;
          int declarationIndex = 0;

          foreach (var member in typeSymbol.GetMembers().OfType<IPropertySymbol>())
          {
               if (member.IsStatic || member.IsIndexer || member.GetMethod is null)
                    continue;

               var attrs = member.GetAttributes();
               if (attrs.Any(a => a.AttributeClass?.ToDisplayString() == IgnoreAttrName))
                    continue;

               bool isMappedToCtor = ctorParamNames.TryGetValue(member.Name, out int ctorParamIdx);
               if (!isMappedToCtor)
                    ctorParamIdx = -1;

               bool isInitOnly = member.SetMethod is null || member.SetMethod.IsInitOnly;

               // Si ni setter ni paramètre de constructeur, on ignore (ex: propriété calculée { get => ... })
               if (member.SetMethod is null && !isMappedToCtor)
                    continue;

               string headerName = member.Name;
               bool isRequired = member.IsRequired;
               string? reqIfType = null;
               bool reqIfCtx = false;
               string? ignIfType = null;
               bool ignIfCtx = false;
               ImmutableArray<string> boolTrue = ImmutableArray<string>.Empty;
               ImmutableArray<string> boolFalse = ImmutableArray<string>.Empty;

               int? columnOrder = null;
               string? formula = null;
               int textCase = 0;
               int trimMode = 0;
               double? minValue = null;
               double? maxValue = null;
               int? minLength = null;
               int? maxLength = null;
               string? regexPattern = null;
               bool regexIgnoreCase = false;
               ImmutableArray<string> allowedValues = ImmutableArray<string>.Empty;
               bool allowedValuesIgnoreCase = true;
               int enumMode = 0;

               char? decimalSeparator = null;
               object? defaultValue = null;
               int defaultKind = 0;
               bool hasDefaultValue = false;
               bool defaultApplyOnExport = false;
               string? format = null;
               int fixedColumnIndex = -1;
               char? splitSeparator = null;
               bool splitTrimEntries = true;
               bool splitRemoveEmptyEntries = true;

               string? converterTypeFullName = null;
               ConverterKind converterKind = ConverterKind.None;

               foreach (var attr in attrs)
               {
                    string? attrFullName = attr.AttributeClass?.ToDisplayString();
                    if (attrFullName == NameAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         headerName = attr.ConstructorArguments[0].Value?.ToString() ?? member.Name;
                    }
                    else if (attrFullName == RequiredAttrName)
                    {
                         isRequired = true;
                    }
                    else if (attrFullName == RequiredIfAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is INamedTypeSymbol condSym1)
                    {
                         if (ValidateConditionSymbol(condSym1, typeSymbol, member, diagnostics, out string condFullName, out string? ctxType))
                         {
                              reqIfType = condFullName;
                              reqIfCtx = ctxType is not null;
                              detectedContextType ??= ctxType;
                         }
                    }
                    else if (attrFullName == IgnoreIfAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is INamedTypeSymbol condSym2)
                    {
                         if (ValidateConditionSymbol(condSym2, typeSymbol, member, diagnostics, out string condFullName, out string? ctxType))
                         {
                              ignIfType = condFullName;
                              ignIfCtx = ctxType is not null;
                              detectedContextType ??= ctxType;
                         }
                    }
                    else if (attrFullName == ConverterAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is INamedTypeSymbol convSym)
                    {
                         if (ValidateConverterSymbol(convSym, typeSymbol, member, diagnostics, out string convFull, out ConverterKind cKind, out string? ctxType))
                         {
                              converterTypeFullName = convFull;
                              converterKind = cKind;
                              detectedContextType ??= ctxType;
                         }
                    }
                    else if (attrFullName == BooleanAttrName)
                    {
                         boolTrue = ExtractStringArrayArg(attr, 0);
                         boolFalse = ExtractStringArrayArg(attr, 1);
                         if (boolTrue.IsEmpty && boolFalse.IsEmpty)
                              diagnostics.Add(Diagnostic.Create(EmptyBooleanArraysError, member.Locations.FirstOrDefault(), member.Name));
                    }
                    else if (attrFullName == ColumnOrderAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is int ord)
                    {
                         columnOrder = ord;
                    }
                    else if (attrFullName == FormulaAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         formula = attr.ConstructorArguments[0].Value?.ToString();
                    }
                    else if (attrFullName == CaseAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is int cVal)
                    {
                         textCase = cVal;
                    }
                    else if (attrFullName == TrimAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is int tVal)
                    {
                         trimMode = tVal;
                    }
                    else if (attrFullName == RangeAttrName && attr.ConstructorArguments.Length >= 2)
                    {
                         minValue = Convert.ToDouble(attr.ConstructorArguments[0].Value);
                         maxValue = Convert.ToDouble(attr.ConstructorArguments[1].Value);
                    }
                    else if (attrFullName == MinAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         minValue = Convert.ToDouble(attr.ConstructorArguments[0].Value);
                    }
                    else if (attrFullName == MaxAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         maxValue = Convert.ToDouble(attr.ConstructorArguments[0].Value);
                    }
                    else if (attrFullName == LengthAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         if (attr.ConstructorArguments.Length == 2 && attr.ConstructorArguments[1].Value is bool isExact)
                         {
                              int len = Convert.ToInt32(attr.ConstructorArguments[0].Value);
                              minLength = isExact ? len : 0;
                              maxLength = len;
                         }
                         else if (attr.ConstructorArguments.Length == 2)
                         {
                              minLength = Convert.ToInt32(attr.ConstructorArguments[0].Value);
                              maxLength = Convert.ToInt32(attr.ConstructorArguments[1].Value);
                         }
                    }
                    else if (attrFullName == RegexAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         regexPattern = attr.ConstructorArguments[0].Value?.ToString();
                         if (attr.ConstructorArguments.Length > 1 && attr.ConstructorArguments[1].Value is bool ic)
                              regexIgnoreCase = ic;

                         if (regexPattern is not null)
                         {
                              try { _ = new Regex(regexPattern); }
                              catch (ArgumentException ex)
                              {
                                   diagnostics.Add(Diagnostic.Create(InvalidRegexPatternError, member.Locations.FirstOrDefault(), regexPattern, member.Name, ex.Message));
                              }
                         }
                    }
                    else if (attrFullName == AllowedValuesAttrName)
                    {
                         allowedValues = ExtractStringArrayArg(attr, 0);
                         foreach (var named in attr.NamedArguments)
                         {
                              if (named.Key == "IgnoreCase" && named.Value.Value is bool ic)
                                   allowedValuesIgnoreCase = ic;
                         }
                    }
                    else if (attrFullName == EnumAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is int eMode)
                    {
                         enumMode = eMode;
                    }
                    else if (attrFullName == DecimalSeparatorAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is char sep)
                    {
                         decimalSeparator = sep;
                    }
                    else if (attrFullName == DefaultAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         hasDefaultValue = true;
                         var arg0 = attr.ConstructorArguments[0];
                         if (arg0.Type?.ToDisplayString() == "FlashXlsx.Abstractions.XlsxDefaultKind" && arg0.Value is int kindVal)
                         {
                              defaultKind = kindVal;
                         }
                         else
                         {
                              defaultKind = 0;
                              defaultValue = arg0.Value;
                         }

                         foreach (var named in attr.NamedArguments)
                         {
                              if (named.Key == "ApplyOnExport" && named.Value.Value is bool aoe)
                                   defaultApplyOnExport = aoe;
                         }
                    }
                    else if (attrFullName == FormatAttrName && attr.ConstructorArguments.Length > 0)
                    {
                         format = attr.ConstructorArguments[0].Value?.ToString();
                    }
                    else if (attrFullName == ColumnLetterAttrName && attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is string colLetters)
                    {
                         fixedColumnIndex = ParseExcelColumnLetters(colLetters);
                         if (fixedColumnIndex < 0)
                         {
                              diagnostics.Add(Diagnostic.Create(InvalidColumnLetterError, member.Locations.FirstOrDefault(), colLetters, member.Name));
                         }
                    }
                    else if (attrFullName == SplitAttrName)
                    {
                         splitSeparator = attr.ConstructorArguments.Length > 0 && attr.ConstructorArguments[0].Value is char s ? s : ';';
                         foreach (var named in attr.NamedArguments)
                         {
                              if (named.Key == "TrimEntries" && named.Value.Value is bool te) splitTrimEntries = te;
                              if (named.Key == "RemoveEmptyEntries" && named.Value.Value is bool re) splitRemoveEmptyEntries = re;
                         }
                    }
               }

               if (minValue.HasValue && maxValue.HasValue && minValue.Value > maxValue.Value)
                    diagnostics.Add(Diagnostic.Create(InvalidRangeBoundsError, member.Locations.FirstOrDefault(), minValue.Value, maxValue.Value, member.Name));

               if (minLength.HasValue && maxLength.HasValue && minLength.Value > maxLength.Value)
                    diagnostics.Add(Diagnostic.Create(InvalidRangeBoundsError, member.Locations.FirstOrDefault(), minLength.Value, maxLength.Value, member.Name));

               var (kind, isNullable, underlyingType, enumMembers) = ClassifyPropertyType(member.Type);
               if (kind == SupportedPropertyKind.Unsupported && converterKind == ConverterKind.None)
               {
                    diagnostics.Add(Diagnostic.Create(UnsupportedPropertyTypeError, member.Locations.FirstOrDefault(), member.Type.ToDisplayString(), member.Name));
                    continue;
               }

               if (kind is SupportedPropertyKind.StringArray or SupportedPropertyKind.StringList && splitSeparator is null)
               {
                    splitSeparator = ';';
               }

               var propModel = new PropertyModel(
                   Name: member.Name,
                   HeaderName: headerName,
                   TypeFullName: member.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                   UnderlyingTypeFullName: underlyingType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                   Kind: kind,
                   IsNullable: isNullable,
                   IsValueType: member.Type.IsValueType,
                   IsRequired: isRequired,
                   RequiredIfConditionType: reqIfType,
                   RequiredIfHasContext: reqIfCtx,
                   IgnoreIfConditionType: ignIfType,
                   IgnoreIfHasContext: ignIfCtx,
                   BoolTrueValues: boolTrue,
                   BoolFalseValues: boolFalse,
                   ColumnOrder: columnOrder,
                   Formula: formula,
                   TextCase: textCase,
                   TrimMode: trimMode,
                   MinValue: minValue,
                   MaxValue: maxValue,
                   MinLength: minLength,
                   MaxLength: maxLength,
                   RegexPattern: regexPattern,
                   RegexIgnoreCase: regexIgnoreCase,
                   AllowedValues: allowedValues,
                   AllowedValuesIgnoreCase: allowedValuesIgnoreCase,
                   EnumMode: enumMode,
                   EnumMembers: enumMembers,
                   DecimalSeparator: decimalSeparator,
                   DefaultValue: defaultValue,
                   DefaultKind: defaultKind,
                   HasDefaultValue: hasDefaultValue,
                   DefaultApplyOnExport: defaultApplyOnExport,
                   Format: format,
                   FixedColumnIndex: fixedColumnIndex,
                   SplitSeparator: splitSeparator,
                   SplitTrimEntries: splitTrimEntries,
                   SplitRemoveEmptyEntries: splitRemoveEmptyEntries,
                   IsInitOnly: isInitOnly,
                   ConstructorParamIndex: ctorParamIdx,
                   ConverterTypeFullName: converterTypeFullName,
                   ConverterKind: converterKind);

               rawProps.Add((propModel, declarationIndex++));
          }

          var orderedProps = rawProps
              .OrderBy(x => x.Prop.ColumnOrder ?? (x.Prop.FixedColumnIndex >= 0 ? x.Prop.FixedColumnIndex : int.MaxValue))
              .ThenBy(x => x.DeclarationIndex)
              .Select(x => x.Prop)
              .ToImmutableArray();

          // Construction des paramètres du constructeur si utilisé
          var ctorParamsBuilder = ImmutableArray.CreateBuilder<ConstructorParamModel>();
          if (chosenParameterizedCtor is not null)
          {
               foreach (var param in chosenParameterizedCtor.Parameters)
               {
                    int mappedPropIdx = -1;
                    for (int i = 0; i < orderedProps.Length; i++)
                    {
                         if (string.Equals(orderedProps[i].Name, param.Name, StringComparison.OrdinalIgnoreCase))
                         {
                              mappedPropIdx = i;
                              break;
                         }
                    }

                    if (mappedPropIdx < 0 && !param.HasExplicitDefaultValue)
                    {
                         diagnostics.Add(Diagnostic.Create(MissingConstructorMappingError, typeSymbol.Locations.FirstOrDefault(), param.Name, typeSymbol.Name));
                    }

                    string? defLiteral = param.HasExplicitDefaultValue
                        ? FormatCSharpConstant(param.ExplicitDefaultValue, param.Type)
                        : null;

                    ctorParamsBuilder.Add(new ConstructorParamModel(
                        ParameterName: param.Name,
                        TypeFullName: param.Type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                        MappedPropertyIndex: mappedPropIdx,
                        HasDefaultValue: param.HasExplicitDefaultValue,
                        DefaultValueLiteral: defLiteral));
               }
          }

          bool usesParamCtor = chosenParameterizedCtor is not null;
          bool hasInitOrRowConverters = usesParamCtor
              || orderedProps.Any(p => p.IsInitOnly || p.ConverterKind is ConverterKind.RowAware or ConverterKind.RowAndContext);

          var model = new TypeModel(
              Namespace: ns,
              TypeName: typeSymbol.Name,
              FullTypeName: fullTypeName,
              IsValueType: typeSymbol.IsValueType,
              HasParameterlessConstructor: hasParameterlessCtor,
              ContextTypeFullName: detectedContextType,
              Properties: orderedProps,
              UsesParameterizedConstructor: usesParamCtor,
              ConstructorParameters: ctorParamsBuilder.ToImmutable(),
              HasInitOrRowConverters: hasInitOrRowConverters);

          return new GenerationResult(model, diagnostics.ToImmutable());
     }

     private static bool IsCopyConstructor(IMethodSymbol ctor, INamedTypeSymbol typeSymbol)
     {
          return typeSymbol.IsRecord
              && ctor.Parameters.Length == 1
              && SymbolEqualityComparer.Default.Equals(ctor.Parameters[0].Type, typeSymbol);
     }

     private static string FormatCSharpConstant(object? val, ITypeSymbol type)
     {
          if (val is null) return "default!";
          if (val is string s) return $"\"{s.Replace("\\", "\\\\").Replace("\"", "\\\"")}\"";
          if (val is bool b) return b ? "true" : "false";
          if (val is char c) return $"'{c}'";
          if (type.TypeKind == TypeKind.Enum) return $"({type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)}){Convert.ToInt64(val)}";
          if (val is decimal m) return $"{m.ToString(CultureInfo.InvariantCulture)}m";
          if (val is double d) return $"{d.ToString("R", CultureInfo.InvariantCulture)}d";
          if (val is float f) return $"{f.ToString("R", CultureInfo.InvariantCulture)}f";
          if (val is long l) return $"{l.ToString(CultureInfo.InvariantCulture)}L";
          return Convert.ToString(val, CultureInfo.InvariantCulture) ?? "default!";
     }

     private static int ParseExcelColumnLetters(string letters)
     {
          if (string.IsNullOrWhiteSpace(letters))
               return -1;

          int index = 0;
          foreach (char c in letters.Trim())
          {
               char upper = char.ToUpperInvariant(c);
               if (upper is < 'A' or > 'Z')
                    return -1;

               index = (index * 26) + (upper - 'A' + 1);
          }
          return index - 1;
     }

     private static bool ValidateConditionSymbol(
         INamedTypeSymbol condSymbol,
         INamedTypeSymbol targetType,
         IPropertySymbol prop,
         ImmutableArray<Diagnostic>.Builder diagnostics,
         out string condFullName,
         out string? contextTypeFullName)
     {
          condFullName = condSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
          contextTypeFullName = null;

          bool hasDefaultCtor = condSymbol.InstanceConstructors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);
          if (!hasDefaultCtor)
          {
               diagnostics.Add(Diagnostic.Create(InvalidConditionTypeError, prop.Locations.FirstOrDefault(), condSymbol.Name, prop.Name, targetType.Name));
               return false;
          }

          foreach (var iface in condSymbol.AllInterfaces)
          {
               string ifaceName = iface.OriginalDefinition.ToDisplayString();
               if (ifaceName == "FlashXlsx.Abstractions.IXlsxCondition<T>" && SymbolEqualityComparer.Default.Equals(iface.TypeArguments[0], targetType))
                    return true;

               if (ifaceName == "FlashXlsx.Abstractions.IXlsxCondition<T, TContext>" && SymbolEqualityComparer.Default.Equals(iface.TypeArguments[0], targetType))
               {
                    contextTypeFullName = iface.TypeArguments[1].ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                    return true;
               }
          }

          diagnostics.Add(Diagnostic.Create(InvalidConditionTypeError, prop.Locations.FirstOrDefault(), condSymbol.Name, prop.Name, targetType.Name));
          return false;
     }

     private static bool ValidateConverterSymbol(
         INamedTypeSymbol convSymbol,
         INamedTypeSymbol targetType,
         IPropertySymbol prop,
         ImmutableArray<Diagnostic>.Builder diagnostics,
         out string convFullName,
         out ConverterKind kind,
         out string? contextTypeFullName)
     {
          convFullName = convSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
          kind = ConverterKind.None;
          contextTypeFullName = null;

          bool hasDefaultCtor = convSymbol.InstanceConstructors.Any(c => c.Parameters.Length == 0 && c.DeclaredAccessibility == Accessibility.Public);
          if (!hasDefaultCtor)
          {
               diagnostics.Add(Diagnostic.Create(InvalidConverterTypeError, prop.Locations.FirstOrDefault(), convSymbol.Name, prop.Name, prop.Type.Name, targetType.Name));
               return false;
          }

          foreach (var iface in convSymbol.AllInterfaces)
          {
               string ifaceName = iface.OriginalDefinition.ToDisplayString();
               if (ifaceName == "FlashXlsx.Abstractions.IXlsxConverter<TProperty>" &&
                   SymbolEqualityComparer.Default.Equals(iface.TypeArguments[0], prop.Type))
               {
                    kind = ConverterKind.Simple;
                    return true;
               }

               if (ifaceName == "FlashXlsx.Abstractions.IXlsxRowConverter<TRow, TProperty>" &&
                   SymbolEqualityComparer.Default.Equals(iface.TypeArguments[0], targetType) &&
                   SymbolEqualityComparer.Default.Equals(iface.TypeArguments[1], prop.Type))
               {
                    kind = ConverterKind.RowAware;
                    return true;
               }

               if (ifaceName == "FlashXlsx.Abstractions.IXlsxRowConverter<TRow, TContext, TProperty>" &&
                   SymbolEqualityComparer.Default.Equals(iface.TypeArguments[0], targetType) &&
                   SymbolEqualityComparer.Default.Equals(iface.TypeArguments[2], prop.Type))
               {
                    kind = ConverterKind.RowAndContext;
                    contextTypeFullName = iface.TypeArguments[1].ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
                    return true;
               }
          }

          diagnostics.Add(Diagnostic.Create(InvalidConverterTypeError, prop.Locations.FirstOrDefault(), convSymbol.Name, prop.Name, prop.Type.Name, targetType.Name));
          return false;
     }

     private static ImmutableArray<string> ExtractStringArrayArg(AttributeData attr, int index)
     {
          if (attr.ConstructorArguments.Length <= index)
               return ImmutableArray<string>.Empty;

          var arg = attr.ConstructorArguments[index];
          if (arg.IsNull || arg.Kind != TypedConstantKind.Array)
               return ImmutableArray<string>.Empty;

          var builder = ImmutableArray.CreateBuilder<string>();
          foreach (var item in arg.Values)
          {
               if (item.Value is string s)
                    builder.Add(s);
          }
          return builder.ToImmutable();
     }

     private static (SupportedPropertyKind Kind, bool IsNullable, ITypeSymbol Underlying, ImmutableArray<EnumMemberInfo> EnumMembers) ClassifyPropertyType(ITypeSymbol type)
     {
          bool isNullable = false;
          ITypeSymbol underlying = type;

          if (type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T } nullableType)
          {
               isNullable = true;
               underlying = nullableType.TypeArguments[0];
          }
          else if (type.NullableAnnotation == NullableAnnotation.Annotated && !type.IsValueType)
          {
               isNullable = true;
          }

          if (underlying is IArrayTypeSymbol { ElementType.SpecialType: SpecialType.System_String })
          {
               return (SupportedPropertyKind.StringArray, isNullable, underlying, ImmutableArray<EnumMemberInfo>.Empty);
          }

          if (underlying is INamedTypeSymbol namedList &&
              namedList.OriginalDefinition.ToDisplayString() == "System.Collections.Generic.List<T>" &&
              namedList.TypeArguments[0].SpecialType == SpecialType.System_String)
          {
               return (SupportedPropertyKind.StringList, isNullable, underlying, ImmutableArray<EnumMemberInfo>.Empty);
          }

          if (underlying.TypeKind == TypeKind.Enum)
          {
               var members = underlying.GetMembers()
                   .OfType<IFieldSymbol>()
                   .Where(f => f.HasConstantValue)
                   .Select(f =>
                   {
                        long numVal = f.ConstantValue is not null ? Convert.ToInt64(f.ConstantValue) : 0L;
                        string label = f.Name;
                        foreach (var attr in f.GetAttributes())
                        {
                             if (attr.AttributeClass?.ToDisplayString() == EnumMemberAttrName &&
                              attr.ConstructorArguments.Length > 0 &&
                              attr.ConstructorArguments[0].Value is string customLabel)
                             {
                                  label = customLabel;
                             }
                        }
                        return new EnumMemberInfo(f.Name, numVal, label);
                   })
                   .ToImmutableArray();

               return (SupportedPropertyKind.Enum, isNullable, underlying, members);
          }

          string fullName = underlying.WithNullableAnnotation(NullableAnnotation.NotAnnotated).ToDisplayString();
          var kind = fullName switch
          {
               "string" => SupportedPropertyKind.String,
               "int" => SupportedPropertyKind.Int32,
               "long" => SupportedPropertyKind.Int64,
               "double" => SupportedPropertyKind.Double,
               "decimal" => SupportedPropertyKind.Decimal,
               "bool" => SupportedPropertyKind.Boolean,
               "System.DateTime" => SupportedPropertyKind.DateTime,
               "System.DateOnly" => SupportedPropertyKind.DateOnly,
               "System.TimeOnly" => SupportedPropertyKind.TimeOnly,
               "System.Guid" => SupportedPropertyKind.Guid,
               _ => SupportedPropertyKind.Unsupported
          };

          return (kind, isNullable, underlying, ImmutableArray<EnumMemberInfo>.Empty);
     }
}