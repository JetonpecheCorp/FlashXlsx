using Microsoft.CodeAnalysis;
using System.Collections.Immutable;

namespace FlashXlsx.Generators;

internal enum SupportedPropertyKind
{
     Unsupported = 0,
     String,
     Int32,
     Int64,
     Double,
     Decimal,
     Boolean,
     DateTime,
     DateOnly,
     TimeOnly,
     Guid,
     Enum,
     StringArray,
     StringList
}

internal enum ConverterKind
{
     None = 0,
     Simple = 1,       // IXlsxConverter<TProp>
     RowAware = 2,     // IXlsxRowConverter<TRow, TProp>
     RowAndContext = 3 // IXlsxRowConverter<TRow, TContext, TProp>
}

internal sealed record EnumMemberInfo(
    string MemberName,
    long NumericValue,
    string Label
);

internal sealed record PropertyModel(
    string Name,
    string HeaderName,
    string TypeFullName,
    string UnderlyingTypeFullName,
    SupportedPropertyKind Kind,
    bool IsNullable,
    bool IsValueType,
    bool IsRequired,
    string? RequiredIfConditionType,
    bool RequiredIfHasContext,
    string? IgnoreIfConditionType,
    bool IgnoreIfHasContext,
    ImmutableArray<string> BoolTrueValues,
    ImmutableArray<string> BoolFalseValues,
    // Nouvelles fonctionnalités
    int? ColumnOrder,
    string? Formula,
    int TextCase, // 0=None, 1=Upper, 2=Lower, 3=Title
    int TrimMode, // 0=None, 1=Both, 2=Start, 3=End, 4=AllWhitespace
    double? MinValue,
    double? MaxValue,
    int? MinLength,
    int? MaxLength,
    string? RegexPattern,
    bool RegexIgnoreCase,
    ImmutableArray<string> AllowedValues,
    bool AllowedValuesIgnoreCase,
    int EnumMode, // 0=Name, 1=Value, 2=Description
    ImmutableArray<EnumMemberInfo> EnumMembers,
    char? DecimalSeparator,
    object? DefaultValue,
    int DefaultKind, // 0=Literal, 1=Now, 2=UtcNow, 3=Today, 4=NewGuid, 5=EmptyString
    bool HasDefaultValue,
    bool DefaultApplyOnExport,
    string? Format,
    int FixedColumnIndex, // -1 si non spécifié
    char? SplitSeparator,
    bool SplitTrimEntries,
    bool SplitRemoveEmptyEntries,
    bool IsInitOnly,
    int ConstructorParamIndex,
    string? ConverterTypeFullName,
    ConverterKind ConverterKind
);

internal sealed record TypeModel(
    string Namespace,
    string TypeName,
    string FullTypeName,
    bool IsValueType,
    bool HasParameterlessConstructor,
    string? ContextTypeFullName,
    ImmutableArray<PropertyModel> Properties,
    bool UsesParameterizedConstructor,
    ImmutableArray<ConstructorParamModel> ConstructorParameters,
    bool HasInitOrRowConverters
);

internal sealed record GenerationResult(
    TypeModel? Model,
    ImmutableArray<Diagnostic> Diagnostics
);

internal sealed record ConstructorParamModel(
    string ParameterName,
    string TypeFullName,
    int MappedPropertyIndex, // -1 si paramètre optionnel non mappé à une propriété
    bool HasDefaultValue,
    string? DefaultValueLiteral
);
