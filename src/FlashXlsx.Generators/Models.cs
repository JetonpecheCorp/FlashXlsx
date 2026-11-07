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
    bool SplitRemoveEmptyEntries
);

internal sealed record TypeModel(
    string Namespace,
    string TypeName,
    string FullTypeName,
    bool IsValueType,
    bool HasParameterlessConstructor,
    string? ContextTypeFullName,
    ImmutableArray<PropertyModel> Properties
);

internal sealed record GenerationResult(
    TypeModel? Model,
    ImmutableArray<Diagnostic> Diagnostics
);