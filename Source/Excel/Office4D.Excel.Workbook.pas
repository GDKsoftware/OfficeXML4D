unit Office4D.Excel.Workbook;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections,
  System.Generics.Defaults,
  System.RegularExpressions,
  System.Zip,
  Office4D.Excel,
  Office4D.Metadata,
  Office4D.Package;

type
  {$SCOPEDENUMS ON}
  TCellType = (Empty, StringValue, Number, Boolean, DateTime);
  {$SCOPEDENUMS OFF}

  TExcelWorkbook = class;

  TExcelCell = class(TInterfacedObject, IExcelCell)
  private
    FStringValue: string;
    FFloatValue: Double;
    FBooleanValue: Boolean;
    FDateTimeValue: TDateTime;
    FCellType: TCellType;
    FFormula: string;
    FBold: Boolean;
    FItalic: Boolean;
    FUnderline: Boolean;
    FFontName: string;
    FFontSize: Double;
    FBackgroundColor: Cardinal;
    FNumberFormat: string;
    FBorderStyle: array[TExcelBorderSide] of TExcelBorderStyle;
    FBorderColor: array[TExcelBorderSide] of Cardinal;
    FHAlign: TExcelHAlign;
    FVAlign: TExcelVAlign;
    FWrapText: Boolean;
    FFontColor: Cardinal;
    FStrikeout: Boolean;
  public
    function GetAsString: string;
    procedure SetAsString(const Value: string);
    function GetAsFloat: Double;
    procedure SetAsFloat(const Value: Double);
    function GetAsBoolean: Boolean;
    procedure SetAsBoolean(const Value: Boolean);
    function GetAsDateTime: TDateTime;
    procedure SetAsDateTime(const Value: TDateTime);
    function GetFormula: string;
    function GetHasFormula: Boolean;
    procedure SetFormula(const Formula: string; const CalculatedValue: Double);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetStrikeout: Boolean;
    procedure SetStrikeout(const Value: Boolean);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Double;
    procedure SetFontSize(const Value: Double);
    function GetBackgroundColor: Cardinal;
    procedure SetBackgroundColor(const Value: Cardinal);
    function GetNumberFormat: string;
    procedure SetNumberFormat(const Value: string);
    function GetBorderStyle(ASides: TExcelBorderSides): TExcelBorderStyle;
    procedure SetBorderStyle(ASides: TExcelBorderSides; const Value: TExcelBorderStyle);
    function GetBorderColor(ASides: TExcelBorderSides): Cardinal;
    procedure SetBorderColor(ASides: TExcelBorderSides; const Value: Cardinal);
    function GetHAlign: TExcelHAlign;
    procedure SetHAlign(const Value: TExcelHAlign);
    function GetVAlign: TExcelVAlign;
    procedure SetVAlign(const Value: TExcelVAlign);
    function GetWrapText: Boolean;
    procedure SetWrapText(const Value: Boolean);
    function GetFontColor: Cardinal;
    procedure SetFontColor(const Value: Cardinal);
    function GetFontStyle: TExcelFontStyles;
    procedure SetFontStyle(const Value: TExcelFontStyles);

    function GetIsString: Boolean;
    function HasStyle: Boolean;

    property CellType: TCellType read FCellType;
    property IsString: Boolean read GetIsString;
  end;

  TExcelSheet = class(TInterfacedObject, IExcelSheet)
  private
    FName: string;
    FCells: TDictionary<string, IExcelCell>;
    FColumnWidths: TDictionary<string, Double>;
    FRowHeights: TDictionary<Integer, Double>;
    FMergedRanges: TList<string>;
    FVisibility: TExcelSheetVisibility;
    FFrozenRows: Integer;
    FFrozenColumns: Integer;
    FNotes: TDictionary<string, string>;
    FOwner: TExcelWorkbook;

    class function ColumnLetterToNumber(const Column: string): Integer; static;
    class function ColumnNumberToLetters(const Column: Integer): string; static;
    class procedure ParseCellAddress(const Address: string; out Col, Row: Integer); static;

    procedure ClearLine(const AIsColumn: Boolean; const AIndex: Integer);
    procedure DeleteLine(const AIsColumn: Boolean; const AIndex: Integer);
    procedure RewriteOwnFormulas(const AIsColumn: Boolean; const AIndex: Integer);
  public
    constructor Create(const Name: string);
    destructor Destroy; override;

    function GetName: string;
    function GetCell(const Address: string): IExcelCell;

    procedure SetColumnWidth(const Column: string; const Width: Double);
    function GetColumnWidth(const Column: string): Double;

    procedure SetRowHeight(const Row: Integer; const Height: Double);
    function GetRowHeight(const Row: Integer): Double;

    procedure MergeCells(const Range: string);
    function GetMergedRanges: TArray<string>;

    function GetVisibility: TExcelSheetVisibility;
    procedure SetVisibility(const Value: TExcelSheetVisibility);

    procedure FreezePanes(const TopLeftCell: string);
    procedure UnfreezePanes;
    function GetFrozenRows: Integer;
    function GetFrozenColumns: Integer;

    function GetNote(const Address: string): string;
    procedure SetNote(const Address: string; const Value: string);
    function HasNotes: Boolean;

    procedure ClearColumn(const Column: string);
    procedure ClearRow(const Row: Integer);
    procedure DeleteColumn(const Column: string);
    procedure DeleteRow(const Row: Integer);

    procedure SetCellValue(const Address, Value: string; IsString: Boolean);
    procedure SetBooleanValue(const Address: string; Value: Boolean);
    procedure SetCellFormula(const Address, Formula, Value: string);
    procedure SetDateTimeValue(const Address: string; const Value: Double);

    function GetCells: TDictionary<string, IExcelCell>;

    property Cells: TDictionary<string, IExcelCell> read FCells;
    property ColumnWidths: TDictionary<string, Double> read FColumnWidths;
    property RowHeights: TDictionary<Integer, Double> read FRowHeights;
    property MergedRangesList: TList<string> read FMergedRanges;
    property Visibility: TExcelSheetVisibility read FVisibility write FVisibility;
    property FrozenRows: Integer read GetFrozenRows;
    property FrozenColumns: Integer read GetFrozenColumns;
    property Notes: TDictionary<string, string> read FNotes;
  end;

  TExcelWorkbook = class(TInterfacedObject, IExcelWorkbook)
  private
    FSheets: TList<IExcelSheet>;
    FMetadata: TDocumentMetadata;
    FSharedStrings: TList<string>;
    FStyleBold: TList<Boolean>;
    FStyleItalic: TList<Boolean>;
    FStyleUnderline: TList<Boolean>;
    FStyleStrikeout: TList<Boolean>;
    FStyleFontName: TList<string>;
    FStyleFontSize: TList<Double>;
    FStyleColors: TList<Cardinal>;
    FStyleFontColor: TList<Cardinal>;
    FStyleBorderTopStyle: TList<TExcelBorderStyle>;
    FStyleBorderTopColor: TList<Cardinal>;
    FStyleBorderRightStyle: TList<TExcelBorderStyle>;
    FStyleBorderRightColor: TList<Cardinal>;
    FStyleBorderBottomStyle: TList<TExcelBorderStyle>;
    FStyleBorderBottomColor: TList<Cardinal>;
    FStyleBorderLeftStyle: TList<TExcelBorderStyle>;
    FStyleBorderLeftColor: TList<Cardinal>;
    FStyleHAlign: TList<TExcelHAlign>;
    FStyleVAlign: TList<TExcelVAlign>;
    FStyleWrapText: TList<Boolean>;
    FStyleIsDate: TList<Boolean>;
    FStyleNumberFormat: TList<string>;

    procedure ParseWorkbook(const Xml: string);
    procedure ParseSharedStrings(const Xml: string);
    procedure ParseStyles(const Xml: string);
    procedure ParseSheet(const Sheet: TExcelSheet; const Xml: string);
    procedure ParseComments(const Sheet: TExcelSheet; const Xml: string);

    function BuildSharedStrings: TList<string>;
    function GetSharedStringIndex(const Strings: TList<string>; const Value: string): Integer;
    function AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;
    function GenerateContentTypes: string;
    function GenerateRels: string;
    function GenerateWorkbook: string;
    function GenerateWorkbookRels: string;
    function GenerateSheet(const Sheet: TExcelSheet; const SharedStrings: TList<string>; const StyleMap: TDictionary<string, Integer>): string;
    function GenerateComments(const Sheet: TExcelSheet): string;
    function GenerateVmlDrawing(const Sheet: TExcelSheet): string;
    function GenerateSheetRels(const CommentsFileIndex: Integer): string;
    function GenerateSharedStrings(const Strings: TList<string>): string;
    function GenerateStyles(const StyleMap: TDictionary<string, Integer>): string;
    function BuildStyleMap: TDictionary<string, Integer>;
    function GetStyleKey(const Cell: TExcelCell): string;
    function GetCellStyleIndex(const Cell: TExcelCell; const StyleMap: TDictionary<string, Integer>): Integer;

    procedure ApplyFormulaDelete(const ATargetSheet: TExcelSheet; const AIsColumn: Boolean; const ALineIndex: Integer);
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetSheetCount: Integer;
    function GetSheet(Index: Integer): IExcelSheet;
    function GetSheetByName(const Name: string): IExcelSheet;
    function GetMetadata: TDocumentMetadata;

    function AddSheet(const Name: string): IExcelSheet;
    function SheetByName(const Name: string): IExcelSheet;

    procedure RemoveSheet(Index: Integer);
    procedure RemoveSheetByName(const Name: string);
  end;

implementation

uses
  System.Math,
  System.StrUtils,
  Office4D.Errors,
  Office4D.Relationships,
  Office4D.Types,
  Office4D.Xml;

const
  // Default OOXML indexed colour palette (indices 0-63). Indices 0-7 are redundant
  // copies of 8-15. Only non-zero colours are listed; everything else maps to 0.
  OoxmlIndexedColors: array[0..63] of Cardinal = (
    $000000, $FFFFFF, $FF0000, $00FF00, $0000FF, $FFFF00, $FF00FF, $00FFFF,  // 0-7
    $000000, $FFFFFF, $FF0000, $00FF00, $0000FF, $FFFF00, $FF00FF, $00FFFF,  // 8-15
    $800000, $008000, $000080, $808000, $800080, $008080, $C0C0C0, $808080,  // 16-23
    $9999FF, $993366, $FFFFCC, $CCFFFF, $660066, $FF8080, $0066CC, $CCCCFF,  // 24-31
    $000080, $FF00FF, $FFFF00, $00FFFF, $800080, $800000, $008080, $0000FF,  // 32-39
    $00CCFF, $CCFFFF, $CCFFCC, $FFFF99, $99CCFF, $FF99CC, $CC99FF, $FFCC99,  // 40-47
    $3366FF, $33CCCC, $99CC00, $FFCC00, $FF9900, $FF6600, $666699, $969696,  // 48-55
    $003366, $339966, $003300, $333300, $993300, $993366, $333399, $333333   // 56-63
  );

  XmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  RelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  ContentTypesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
  SpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  OfficeDocRelsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  PartCoreProps = 'docProps/core.xml';
  PartSharedStrings = 'xl/sharedStrings.xml';
  PartStyles = 'xl/styles.xml';
  PartWorkbook = 'xl/workbook.xml';
  PartWorkbookRels = 'xl/_rels/workbook.xml.rels';
  PartSheetPrefix = 'xl/worksheets/sheet';
  PartSheetSuffix = '.xml';

{ Formula reference rewriting for DeleteColumn / DeleteRow }

// Left edge of a reference range on the deleted axis. A coordinate before the
// deleted line stays put; one after it shifts down by one; one exactly on the
// deleted line clamps to the deleted position (the next surviving line moves there).
function ClampLo(const AValue, ADeletedIndex: Integer): Integer;
begin
  if AValue < ADeletedIndex then
    Result := AValue
  else if AValue > ADeletedIndex then
    Result := AValue - 1
  else
    Result := ADeletedIndex;
end;

// Right edge of a reference range on the deleted axis. Identical to ClampLo except
// a coordinate exactly on the deleted line clamps to the line just before it, so a
// range whose far edge sat on the deleted line shrinks (e.g. A1:C1 -> A1:B1).
function ClampHi(const AValue, ADeletedIndex: Integer): Integer;
begin
  if AValue < ADeletedIndex then
    Result := AValue
  else if AValue > ADeletedIndex then
    Result := AValue - 1
  else
    Result := ADeletedIndex - 1;
end;

// Splits an optional leading '$' from a column/row part, e.g. '$D' -> ('D', True).
function SplitDollar(const APart: string; out AIsAbsolute: Boolean): string;
begin
  AIsAbsolute := APart.StartsWith('$');
  if AIsAbsolute then
    Result := APart.Substring(1)
  else
    Result := APart;
end;

function DollarIf(const AIsAbsolute: Boolean): string;
begin
  if AIsAbsolute then
    Result := '$'
  else
    Result := '';
end;

// Turns a matched sheet prefix (including the trailing '!', possibly quoted) into
// the bare sheet name, e.g. '''My Sheet''!' -> 'My Sheet'.
function UnquoteSheetPrefix(const APrefix: string): string;
begin
  Result := APrefix;
  if Result.EndsWith('!') then
    Result := Result.Substring(0, Result.Length - 1);
  if Result.StartsWith('''') and Result.EndsWith('''') and (Result.Length >= 2) then
  begin
    Result := Result.Substring(1, Result.Length - 2);
    Result := StringReplace(Result, '''''', '''', [rfReplaceAll]);
  end;
end;

// Rewrites a single matched cell reference or range for a column/row deletion.
// References that do not resolve to the deleted sheet are returned verbatim.
// Group layout: 1=sheet prefix, 2=col1, 3=row1, 4=col2, 5=row2 (4/5 only on a range).
function TransformRefMatch(const AMatch: TMatch; const ATargetSheet, AFormulaSheet: string;
  const AIsColumn: Boolean; const ALineIndex: Integer): string;

  // Delphi's TMatch only reports groups up to the highest one that participated, so an
  // optional trailing group is absent (not merely unsuccessful) on a non-range match.
  function GroupValue(const AIndex: Integer): string;
  begin
    if AIndex < AMatch.Groups.Count then
      Result := AMatch.Groups[AIndex].Value
    else
      Result := '';
  end;

  function GroupSuccess(const AIndex: Integer): Boolean;
  begin
    Result := (AIndex < AMatch.Groups.Count) and AMatch.Groups[AIndex].Success;
  end;

begin
  const Prefix1 = GroupValue(1);
  const IsRange = GroupSuccess(5);

  var RefSheet := AFormulaSheet;
  if Prefix1 <> '' then
    RefSheet := UnquoteSheetPrefix(Prefix1);

  // Only references pointing at the sheet whose line was deleted are affected.
  if not SameText(RefSheet, ATargetSheet) then
    Exit(AMatch.Value);

  var Col1Abs, Row1Abs: Boolean;
  const Col1Letters = SplitDollar(GroupValue(2), Col1Abs);
  const Row1Digits = SplitDollar(GroupValue(3), Row1Abs);
  var Col1 := TExcelSheet.ColumnLetterToNumber(Col1Letters);
  var Row1 := StrToIntDef(Row1Digits, 0);

  var Col2Abs := Col1Abs;
  var Row2Abs := Row1Abs;
  var Col2 := Col1;
  var Row2 := Row1;
  if IsRange then
  begin
    Col2 := TExcelSheet.ColumnLetterToNumber(SplitDollar(GroupValue(4), Col2Abs));
    Row2 := StrToIntDef(SplitDollar(GroupValue(5), Row2Abs), 0);
  end;

  if AIsColumn then
  begin
    const Lo = Min(Col1, Col2);
    const Hi = Max(Col1, Col2);
    if (Lo = ALineIndex) and (Hi = ALineIndex) then
      Exit('#REF!');
    const NewLo = ClampLo(Lo, ALineIndex);
    const NewHi = ClampHi(Hi, ALineIndex);
    if Col1 <= Col2 then
    begin
      Col1 := NewLo;
      Col2 := NewHi;
    end
    else
    begin
      Col1 := NewHi;
      Col2 := NewLo;
    end;
  end
  else
  begin
    const Lo = Min(Row1, Row2);
    const Hi = Max(Row1, Row2);
    if (Lo = ALineIndex) and (Hi = ALineIndex) then
      Exit('#REF!');
    const NewLo = ClampLo(Lo, ALineIndex);
    const NewHi = ClampHi(Hi, ALineIndex);
    if Row1 <= Row2 then
    begin
      Row1 := NewLo;
      Row2 := NewHi;
    end
    else
    begin
      Row1 := NewHi;
      Row2 := NewLo;
    end;
  end;

  Result := Prefix1 + DollarIf(Col1Abs) + TExcelSheet.ColumnNumberToLetters(Col1) +
    DollarIf(Row1Abs) + IntToStr(Row1);
  if IsRange then
    Result := Result + ':' + DollarIf(Col2Abs) + TExcelSheet.ColumnNumberToLetters(Col2) +
      DollarIf(Row2Abs) + IntToStr(Row2);
end;

// Applies the reference transform to a stretch of formula text that contains no
// string literals.
function TransformFormulaSegment(const ASegment, ATargetSheet, AFormulaSheet: string;
  const AIsColumn: Boolean; const ALineIndex: Integer): string;
const
  // [not a name char] [optional sheet prefix] col row  [ ':' [prefix] col row ]  [not '(' or name char].
  // The trailing '(' guard keeps function names like LOG10 from being read as a reference.
  // Groups: 1=sheet prefix, 2=col1, 3=row1, 4=col2, 5=row2. The second endpoint's sheet
  // prefix (rare) is matched but not captured.
  RefPattern =
    '(?<![A-Za-z0-9_])' +
    '((?:''[^'']*''|[A-Za-z_][A-Za-z0-9_.]*)!)?' +
    '(\$?[A-Za-z]{1,3})(\$?\d+)' +
    '(?::(?:(?:''[^'']*''|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3})(\$?\d+))?' +
    '(?![A-Za-z0-9_(])';
begin
  Result := ASegment;
  const Matches = TRegEx.Matches(ASegment, RefPattern, [roIgnoreCase]);
  var OffsetAdjust := 0;
  for var M in Matches do
  begin
    const OldRef = M.Value;
    const NewRef = TransformRefMatch(M, ATargetSheet, AFormulaSheet, AIsColumn, ALineIndex);
    if NewRef <> OldRef then
    begin
      const RefPos = M.Index + OffsetAdjust; // TMatch.Index is 1-based (Delphi string convention)
      Delete(Result, RefPos, Length(OldRef));
      Insert(NewRef, Result, RefPos);
      Inc(OffsetAdjust, Length(NewRef) - Length(OldRef));
    end;
  end;
end;

// Rewrites all cell references in a formula for a column/row deletion on ATargetSheet.
// String literals inside the formula are copied verbatim so cell-like text is never touched.
function TransformFormulaForLineDelete(const AFormula, ATargetSheet, AFormulaSheet: string;
  const AIsColumn: Boolean; const ALineIndex: Integer): string;
begin
  var Output := TStringBuilder.Create;
  var Segment := TStringBuilder.Create;
  try
    var CharIndex := 1;
    while CharIndex <= Length(AFormula) do
    begin
      if AFormula[CharIndex] = '"' then
      begin
        if Segment.Length > 0 then
        begin
          Output.Append(TransformFormulaSegment(Segment.ToString, ATargetSheet, AFormulaSheet, AIsColumn, ALineIndex));
          Segment.Clear;
        end;
        // Copy the literal verbatim, treating a doubled "" as an escaped quote.
        Output.Append(AFormula[CharIndex]);
        Inc(CharIndex);
        while CharIndex <= Length(AFormula) do
        begin
          Output.Append(AFormula[CharIndex]);
          if AFormula[CharIndex] = '"' then
          begin
            if (CharIndex < Length(AFormula)) and (AFormula[CharIndex + 1] = '"') then
            begin
              Output.Append(AFormula[CharIndex + 1]);
              Inc(CharIndex, 2);
              Continue;
            end;
            Inc(CharIndex);
            Break;
          end;
          Inc(CharIndex);
        end;
      end
      else
      begin
        Segment.Append(AFormula[CharIndex]);
        Inc(CharIndex);
      end;
    end;
    if Segment.Length > 0 then
      Output.Append(TransformFormulaSegment(Segment.ToString, ATargetSheet, AFormulaSheet, AIsColumn, ALineIndex));
    Result := Output.ToString;
  finally
    Segment.Free;
    Output.Free;
  end;
end;

{ TExcelCell }

function TExcelCell.GetAsString: string;
begin
  Result := FStringValue;
end;

procedure TExcelCell.SetAsString(const Value: string);
begin
  FStringValue := Value;
  FCellType := TCellType.StringValue;
end;

function TExcelCell.GetAsFloat: Double;
begin
  Result := FFloatValue;
end;

procedure TExcelCell.SetAsFloat(const Value: Double);
begin
  FFloatValue := Value;
  FStringValue := '';
  FCellType := TCellType.Number;
end;

function TExcelCell.GetAsBoolean: Boolean;
begin
  Result := FBooleanValue;
end;

procedure TExcelCell.SetAsBoolean(const Value: Boolean);
begin
  FBooleanValue := Value;
  FCellType := TCellType.Boolean;
end;

function TExcelCell.GetAsDateTime: TDateTime;
begin
  if FCellType = TCellType.DateTime then
    Result := FDateTimeValue
  else if FCellType = TCellType.Number then
    Result := FFloatValue
  else
    Result := 0;
end;

procedure TExcelCell.SetAsDateTime(const Value: TDateTime);
begin
  FDateTimeValue := Value;
  FFloatValue := Value;
  FCellType := TCellType.DateTime;
end;

function TExcelCell.GetIsString: Boolean;
begin
  Result := FCellType = TCellType.StringValue;
end;

function TExcelCell.GetFormula: string;
begin
  Result := FFormula;
end;

function TExcelCell.GetHasFormula: Boolean;
begin
  Result := FFormula <> '';
end;

procedure TExcelCell.SetFormula(const Formula: string; const CalculatedValue: Double);
begin
  FFormula := Formula;
  FFloatValue := CalculatedValue;
  FCellType := TCellType.Number;
end;

function TExcelCell.GetBold: Boolean;
begin
  Result := FBold;
end;

procedure TExcelCell.SetBold(const Value: Boolean);
begin
  FBold := Value;
end;

function TExcelCell.GetBackgroundColor: Cardinal;
begin
  Result := FBackgroundColor;
end;

procedure TExcelCell.SetBackgroundColor(const Value: Cardinal);
begin
  FBackgroundColor := Value;
end;

function TExcelCell.GetNumberFormat: string;
begin
  Result := FNumberFormat;
end;

procedure TExcelCell.SetNumberFormat(const Value: string);
begin
  FNumberFormat := Value;
end;

function TExcelCell.GetStrikeout: Boolean;
begin
  Result := FStrikeout;
end;

procedure TExcelCell.SetStrikeout(const Value: Boolean);
begin
  FStrikeout := Value;
end;

function TExcelCell.GetItalic: Boolean;
begin
  Result := FItalic;
end;

procedure TExcelCell.SetItalic(const Value: Boolean);
begin
  FItalic := Value;
end;

function TExcelCell.GetUnderline: Boolean;
begin
  Result := FUnderline;
end;

procedure TExcelCell.SetUnderline(const Value: Boolean);
begin
  FUnderline := Value;
end;

function TExcelCell.GetFontColor: Cardinal;
begin
  Result := FFontColor;
end;

function TExcelCell.GetFontName: string;
begin
  Result := FFontName;
end;

procedure TExcelCell.SetFontColor(const Value: Cardinal);
begin
  FFontColor := Value;
end;

procedure TExcelCell.SetFontName(const Value: string);
begin
  FFontName := Value;
end;

function TExcelCell.GetFontSize: Double;
begin
  Result := FFontSize;
end;

procedure TExcelCell.SetFontSize(const Value: Double);
begin
  FFontSize := Value;
end;

function TExcelCell.GetFontStyle: TExcelFontStyles;
begin
  Result := [];
  if FBold then Include(Result, TExcelFontStyle.Bold);
  if FItalic then Include(Result, TExcelFontStyle.Italic);
  if FUnderline then Include(Result, TExcelFontStyle.Underline);
  if FStrikeout then Include(Result, TExcelFontStyle.Strikeout);
end;

procedure TExcelCell.SetFontStyle(const Value: TExcelFontStyles);
begin
  FBold := TExcelFontStyle.Bold in Value;
  FItalic := TExcelFontStyle.Italic in Value;
  FUnderline := TExcelFontStyle.Underline in Value;
  FStrikeout := TExcelFontStyle.Strikeout in Value;
end;

function TExcelCell.GetBorderStyle(ASides: TExcelBorderSides): TExcelBorderStyle;
begin
  Result := TExcelBorderStyle.None;
  for var Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
    if Side in ASides then
      Exit(FBorderStyle[Side]); // returns the first matching side in Top/Right/Bottom/Left order
end;

procedure TExcelCell.SetBorderStyle(ASides: TExcelBorderSides; const Value: TExcelBorderStyle);
begin
  for var Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
    if Side in ASides then
      FBorderStyle[Side] := Value;
end;

function TExcelCell.GetBorderColor(ASides: TExcelBorderSides): Cardinal;
begin
  Result := 0;
  for var Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
    if Side in ASides then
      Exit(FBorderColor[Side]); // returns the first matching side in Top/Right/Bottom/Left order
end;

procedure TExcelCell.SetBorderColor(ASides: TExcelBorderSides; const Value: Cardinal);
begin
  for var Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
    if Side in ASides then
      FBorderColor[Side] := Value;
end;

function TExcelCell.GetHAlign: TExcelHAlign;
begin
  Result := FHAlign;
end;

procedure TExcelCell.SetHAlign(const Value: TExcelHAlign);
begin
  FHAlign := Value;
end;

function TExcelCell.GetVAlign: TExcelVAlign;
begin
  Result := FVAlign;
end;

procedure TExcelCell.SetVAlign(const Value: TExcelVAlign);
begin
  FVAlign := Value;
end;

function TExcelCell.GetWrapText: Boolean;
begin
  Result := FWrapText;
end;

procedure TExcelCell.SetWrapText(const Value: Boolean);
begin
  FWrapText := Value;
end;

function TExcelCell.HasStyle: Boolean;
begin
  const HasFont = (FBold) or (FItalic) or (FUnderline) or (FFontName <> '') or (FFontSize <> 0) or (FFontColor <> 0) or (FStrikeout);
  const HasFill = (FBackgroundColor <> 0);
  const HasFormat = (FCellType = TCellType.DateTime);
  var HasBorder := False;
  for var Side := Low(TExcelBorderSide) to High(TExcelBorderSide) do
    if FBorderStyle[Side] <> TExcelBorderStyle.None then
    begin
      HasBorder := True;
      Break;
    end;
  const HasAlign = (FHAlign <> TExcelHAlign.None) or (FVAlign <> TExcelVAlign.None) or (FWrapText);
  Result := (HasFont) or (HasFill) or (HasFormat) or (HasBorder) or (HasAlign);
end;

{ TExcelSheet }

constructor TExcelSheet.Create(const Name: string);
begin
  inherited Create;
  FName := Name;
  FCells := TDictionary<string, IExcelCell>.Create;
  FColumnWidths := TDictionary<string, Double>.Create;
  FRowHeights := TDictionary<Integer, Double>.Create;
  FMergedRanges := TList<string>.Create;
  FNotes := TDictionary<string, string>.Create;
end;

destructor TExcelSheet.Destroy;
begin
  FNotes.Free;
  FMergedRanges.Free;
  FRowHeights.Free;
  FColumnWidths.Free;
  FCells.Free;
  inherited;
end;

function TExcelSheet.GetName: string;
begin
  Result := FName;
end;

function TExcelSheet.GetCell(const Address: string): IExcelCell;
begin
  if not FCells.TryGetValue(UpperCase(Address), Result) then
  begin
    Result := TExcelCell.Create;
    FCells.Add(UpperCase(Address), Result);
  end;
end;

function TExcelSheet.GetCells: TDictionary<string, IExcelCell>;
begin
  Result := FCells;
end;

procedure TExcelSheet.SetColumnWidth(const Column: string; const Width: Double);
begin
  FColumnWidths.AddOrSetValue(UpperCase(Column), Width);
end;

function TExcelSheet.GetColumnWidth(const Column: string): Double;
begin
  if not FColumnWidths.TryGetValue(UpperCase(Column), Result) then
    Result := 0;
end;

procedure TExcelSheet.SetRowHeight(const Row: Integer; const Height: Double);
begin
  FRowHeights.AddOrSetValue(Row, Height);
end;

function TExcelSheet.GetRowHeight(const Row: Integer): Double;
begin
  if not FRowHeights.TryGetValue(Row, Result) then
    Result := 0;
end;

procedure TExcelSheet.MergeCells(const Range: string);
begin
  if not FMergedRanges.Contains(UpperCase(Range)) then
    FMergedRanges.Add(UpperCase(Range));
end;

function TExcelSheet.GetMergedRanges: TArray<string>;
begin
  Result := FMergedRanges.ToArray;
end;

function TExcelSheet.GetVisibility: TExcelSheetVisibility;
begin
  Result := FVisibility;
end;

procedure TExcelSheet.SetVisibility(const Value: TExcelSheetVisibility);
begin
  FVisibility := Value;
end;

procedure TExcelSheet.FreezePanes(const TopLeftCell: string);
var
  Col, Row: Integer;
begin
  ParseCellAddress(TopLeftCell, Col, Row);
  // TopLeftCell is the first *unfrozen* (scrollable) cell -- e.g. 'C2' means columns
  // A-B and row 1 are frozen, so the frozen counts are one less than the parsed position.
  FFrozenColumns := Max(0, Col - 1);
  FFrozenRows := Max(0, Row - 1);
end;

procedure TExcelSheet.UnfreezePanes;
begin
  FFrozenRows := 0;
  FFrozenColumns := 0;
end;

function TExcelSheet.GetFrozenRows: Integer;
begin
  Result := FFrozenRows;
end;

function TExcelSheet.GetFrozenColumns: Integer;
begin
  Result := FFrozenColumns;
end;

function TExcelSheet.GetNote(const Address: string): string;
begin
  if not FNotes.TryGetValue(UpperCase(Address), Result) then
    Result := '';
end;

procedure TExcelSheet.SetNote(const Address: string; const Value: string);
begin
  if Value = '' then
    FNotes.Remove(UpperCase(Address))
  else
    FNotes.AddOrSetValue(UpperCase(Address), Value);
end;

function TExcelSheet.HasNotes: Boolean;
begin
  Result := FNotes.Count > 0;
end;

class function TExcelSheet.ColumnLetterToNumber(const Column: string): Integer;
begin
  Result := 0;
  const UpperColumn = UpperCase(Column);
  for var CharIndex := 1 to Length(UpperColumn) do
  begin
    const CharVal = Ord(UpperColumn[CharIndex]) - Ord('A') + 1;
    Result := Result * 26 + CharVal;
  end;
end;

class function TExcelSheet.ColumnNumberToLetters(const Column: Integer): string;
begin
  Result := '';
  var ColNum := Column;
  while ColNum > 0 do
  begin
    const Remainder = (ColNum - 1) mod 26;
    Result := Chr(Ord('A') + Remainder) + Result;
    ColNum := (ColNum - 1) div 26;
  end;
end;

class procedure TExcelSheet.ParseCellAddress(const Address: string; out Col, Row: Integer);
begin
  const UpperAddress = UpperCase(Trim(Address));
  var ColPart := '';
  var RowPart := '';
  for var CharIndex := 1 to Length(UpperAddress) do
    if CharInSet(UpperAddress[CharIndex], ['A'..'Z']) then
      ColPart := ColPart + UpperAddress[CharIndex]
    else
      RowPart := RowPart + UpperAddress[CharIndex];

  if (ColPart = '') or (RowPart = '') or not TryStrToInt(RowPart, Row) then
    raise EExcelWorkbookException.CreateFmt('Invalid cell address: "%s"', [Address]);

  Col := ColumnLetterToNumber(ColPart);
end;

procedure TExcelSheet.SetCellValue(const Address, Value: string; IsString: Boolean);
begin
  var Cell := GetCell(Address) as TExcelCell;
  if IsString then
    Cell.SetAsString(Value)
  else
  begin
    Cell.SetAsFloat(StrToFloatDef(Value, 0, TFormatSettings.Invariant));
    Cell.FStringValue := Value;
  end;
end;

procedure TExcelSheet.SetBooleanValue(const Address: string; Value: Boolean);
begin
  var Cell := GetCell(Address) as TExcelCell;
  Cell.SetAsBoolean(Value);
end;

procedure TExcelSheet.SetDateTimeValue(const Address: string; const Value: Double);
begin
  // Value is the raw Excel serial number as read from <v>. A Delphi
  // TDateTime and an Excel/OOXML date serial number are the same value
  // numerically (see TExcelCell.SetAsDateTime), so it can be passed
  // straight through without any epoch conversion.
  var Cell := GetCell(Address) as TExcelCell;
  Cell.SetAsDateTime(Value);
end;

procedure TExcelSheet.SetCellFormula(const Address, Formula, Value: string);
begin
  var Cell := GetCell(Address) as TExcelCell;
  Cell.FFormula := Formula;
  Cell.SetAsFloat(StrToFloatDef(Value, 0, TFormatSettings.Invariant));
  Cell.FStringValue := Value;
end;

procedure TExcelSheet.ClearColumn(const Column: string);
begin
  ClearLine(True, ColumnLetterToNumber(Column));
end;

procedure TExcelSheet.ClearRow(const Row: Integer);
begin
  ClearLine(False, Row);
end;

procedure TExcelSheet.ClearLine(const AIsColumn: Boolean; const AIndex: Integer);

  function IsOnLine(const AAddress: string): Boolean;
  begin
    var Col, Row: Integer;
    ParseCellAddress(AAddress, Col, Row);
    Result := (AIsColumn and (Col = AIndex)) or ((not AIsColumn) and (Row = AIndex));
  end;

begin
  // Cells and notes are stored sparse and keyed by address, so clearing a line is
  // simply dropping the matching keys. Column widths, row heights, merges and freeze
  // panes are left intact -- this clears contents, not structure.
  var KeysToRemove := TList<string>.Create;
  try
    for var Pair in FCells do
      if IsOnLine(Pair.Key) then
        KeysToRemove.Add(Pair.Key);
    for var Key in KeysToRemove do
      FCells.Remove(Key);

    KeysToRemove.Clear;
    for var Pair in FNotes do
      if IsOnLine(Pair.Key) then
        KeysToRemove.Add(Pair.Key);
    for var Key in KeysToRemove do
      FNotes.Remove(Key);
  finally
    KeysToRemove.Free;
  end;
end;

procedure TExcelSheet.DeleteColumn(const Column: string);
begin
  DeleteLine(True, ColumnLetterToNumber(Column));
end;

procedure TExcelSheet.DeleteRow(const Row: Integer);
begin
  DeleteLine(False, Row);
end;

procedure TExcelSheet.DeleteLine(const AIsColumn: Boolean; const AIndex: Integer);
begin
  // 1. Cells: rebuild the address-keyed dictionary, dropping the deleted line and
  //    shifting everything past it back by one. The cell (and its styling) moves as one.
  var NewCells := TDictionary<string, IExcelCell>.Create;
  for var Pair in FCells do
  begin
    var Col, Row: Integer;
    ParseCellAddress(Pair.Key, Col, Row);
    var Axis := Col;
    if not AIsColumn then
      Axis := Row;
    if Axis = AIndex then
      Continue;
    if Axis > AIndex then
      Dec(Axis);
    if AIsColumn then
      NewCells.Add(ColumnNumberToLetters(Axis) + IntToStr(Row), Pair.Value)
    else
      NewCells.Add(ColumnNumberToLetters(Col) + IntToStr(Axis), Pair.Value);
  end;
  FCells.Free;
  FCells := NewCells;

  // 1b. Notes are keyed by address just like cells, so they shift (and drop on the
  //     deleted line) the same way, otherwise a note would detach from its cell.
  var NewNotes := TDictionary<string, string>.Create;
  for var Pair in FNotes do
  begin
    var Col, Row: Integer;
    ParseCellAddress(Pair.Key, Col, Row);
    var Axis := Col;
    if not AIsColumn then
      Axis := Row;
    if Axis = AIndex then
      Continue;
    if Axis > AIndex then
      Dec(Axis);
    if AIsColumn then
      NewNotes.Add(ColumnNumberToLetters(Axis) + IntToStr(Row), Pair.Value)
    else
      NewNotes.Add(ColumnNumberToLetters(Col) + IntToStr(Axis), Pair.Value);
  end;
  FNotes.Free;
  FNotes := NewNotes;

  // 2. Column widths (column delete) or row heights (row delete) shift the same way.
  if AIsColumn then
  begin
    var NewWidths := TDictionary<string, Double>.Create;
    for var Pair in FColumnWidths do
    begin
      var ColIdx := ColumnLetterToNumber(Pair.Key);
      if ColIdx = AIndex then
        Continue;
      if ColIdx > AIndex then
        Dec(ColIdx);
      NewWidths.Add(ColumnNumberToLetters(ColIdx), Pair.Value);
    end;
    FColumnWidths.Free;
    FColumnWidths := NewWidths;
  end
  else
  begin
    var NewHeights := TDictionary<Integer, Double>.Create;
    for var Pair in FRowHeights do
    begin
      var RowIdx := Pair.Key;
      if RowIdx = AIndex then
        Continue;
      if RowIdx > AIndex then
        Dec(RowIdx);
      NewHeights.Add(RowIdx, Pair.Value);
    end;
    FRowHeights.Free;
    FRowHeights := NewHeights;
  end;

  // 3. Merged ranges: shrink, shift or drop each range on the deleted axis.
  var NewMerges := TList<string>.Create;
  for var Range in FMergedRanges do
  begin
    const ColonPos = Pos(':', Range);
    if ColonPos = 0 then
      Continue;
    var C1, R1, C2, R2: Integer;
    ParseCellAddress(Copy(Range, 1, ColonPos - 1), C1, R1);
    ParseCellAddress(Copy(Range, ColonPos + 1, Length(Range)), C2, R2);

    if AIsColumn then
    begin
      const Lo = Min(C1, C2);
      const Hi = Max(C1, C2);
      if (Lo = AIndex) and (Hi = AIndex) then
        Continue;
      C1 := ClampLo(Lo, AIndex);
      C2 := ClampHi(Hi, AIndex);
    end
    else
    begin
      const Lo = Min(R1, R2);
      const Hi = Max(R1, R2);
      if (Lo = AIndex) and (Hi = AIndex) then
        Continue;
      R1 := ClampLo(Lo, AIndex);
      R2 := ClampHi(Hi, AIndex);
    end;

    // A range that has collapsed to a single cell is no longer a merge.
    if (C1 = C2) and (R1 = R2) then
      Continue;
    NewMerges.Add(ColumnNumberToLetters(C1) + IntToStr(R1) + ':' + ColumnNumberToLetters(C2) + IntToStr(R2));
  end;
  FMergedRanges.Free;
  FMergedRanges := NewMerges;

  // 4. Frozen pane counts: a deleted line at or before the frozen boundary shrinks it.
  if AIsColumn then
  begin
    if AIndex <= FFrozenColumns then
      Dec(FFrozenColumns);
  end
  else if AIndex <= FFrozenRows then
    Dec(FFrozenRows);

  // 5. Formulas: rewrite references across the whole workbook so cross-sheet refs to
  //    this sheet are adjusted too. Falls back to self-scope if there is no owner.
  if FOwner <> nil then
    FOwner.ApplyFormulaDelete(Self, AIsColumn, AIndex)
  else
    RewriteOwnFormulas(AIsColumn, AIndex);
end;

procedure TExcelSheet.RewriteOwnFormulas(const AIsColumn: Boolean; const AIndex: Integer);
begin
  for var Pair in FCells do
  begin
    var Cell := Pair.Value as TExcelCell;
    if Cell.GetHasFormula then
    begin
      const NewFormula = TransformFormulaForLineDelete(Cell.GetFormula, FName, FName, AIsColumn, AIndex);
      if NewFormula <> Cell.GetFormula then
        Cell.FFormula := NewFormula;
    end;
  end;
end;

{ TExcelWorkbook }

constructor TExcelWorkbook.Create;
begin
  inherited;
  FSheets := TList<IExcelSheet>.Create;
  FSharedStrings := TList<string>.Create;
  FStyleBold := TList<Boolean>.Create;
  FStyleItalic := TList<Boolean>.Create;
  FStyleUnderline := TList<Boolean>.Create;
  FStyleFontName := TList<string>.Create;
  FStyleFontSize := TList<Double>.Create;
  FStyleColors := TList<Cardinal>.Create;
  FStyleFontColor := TList<Cardinal>.Create;
  FStyleStrikeout := TList<Boolean>.Create;
  FStyleBorderTopStyle := TList<TExcelBorderStyle>.Create;
  FStyleBorderTopColor := TList<Cardinal>.Create;
  FStyleBorderRightStyle := TList<TExcelBorderStyle>.Create;
  FStyleBorderRightColor := TList<Cardinal>.Create;
  FStyleBorderBottomStyle := TList<TExcelBorderStyle>.Create;
  FStyleBorderBottomColor := TList<Cardinal>.Create;
  FStyleBorderLeftStyle := TList<TExcelBorderStyle>.Create;
  FStyleBorderLeftColor := TList<Cardinal>.Create;
  FStyleHAlign := TList<TExcelHAlign>.Create;
  FStyleVAlign := TList<TExcelVAlign>.Create;
  FStyleWrapText := TList<Boolean>.Create;
  FStyleIsDate := TList<Boolean>.Create;
  FStyleNumberFormat := TList<string>.Create;
end;

destructor TExcelWorkbook.Destroy;
begin
  FStyleNumberFormat.Free;
  FStyleIsDate.Free;
  FStyleWrapText.Free;
  FStyleVAlign.Free;
  FStyleHAlign.Free;
  FStyleBorderLeftColor.Free;
  FStyleBorderLeftStyle.Free;
  FStyleBorderBottomColor.Free;
  FStyleBorderBottomStyle.Free;
  FStyleBorderRightColor.Free;
  FStyleBorderRightStyle.Free;
  FStyleBorderTopColor.Free;
  FStyleBorderTopStyle.Free;
  FStyleFontColor.Free;
  FStyleStrikeout.Free;
  FStyleColors.Free;
  FStyleFontSize.Free;
  FStyleFontName.Free;
  FStyleUnderline.Free;
  FStyleItalic.Free;
  FStyleBold.Free;
  FSharedStrings.Free;
  FSheets.Free;
  inherited;
end;

procedure TExcelWorkbook.LoadFromFile(const FileName: string);
begin
  if not FileExists(FileName) then
    raise EPackageNotFound.Create('File not found: ' + FileName);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FileName);

    if Package.PartExists(PartCoreProps) then
    begin
      const CoreXml = Package.GetPartContent(PartCoreProps);
      var Parser := TMetadataParser.Create;
      try
        FMetadata := Parser.Parse(CoreXml);
      finally
        Parser.Free;
      end;
    end;

    if Package.PartExists(PartSharedStrings) then
      ParseSharedStrings(Package.GetPartContent(PartSharedStrings));

    if Package.PartExists(PartStyles) then
      ParseStyles(Package.GetPartContent(PartStyles));

    if Package.PartExists(PartWorkbook) then
      ParseWorkbook(Package.GetPartContent(PartWorkbook));

    var Rels := TRelationships.Create;
    try
      if Package.PartExists(PartWorkbookRels) then
        Rels.LoadFromXml(Package.GetPartContent(PartWorkbookRels));

      for var I := 0 to FSheets.Count - 1 do
      begin
        const SheetPath = PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix;
        if Package.PartExists(SheetPath) then
          ParseSheet(FSheets[I] as TExcelSheet, Package.GetPartContent(SheetPath));

        const SheetRelsPath = 'xl/worksheets/_rels/sheet' + IntToStr(I + 1) + '.xml.rels';
        if Package.PartExists(SheetRelsPath) then
        begin
          var SheetRels := TRelationships.Create;
          try
            SheetRels.LoadFromXml(Package.GetPartContent(SheetRelsPath));
            const CommentsTarget = SheetRels.GetTargetByType(OfficeDocRelsNs + '/comments');
            if CommentsTarget <> '' then
            begin
              // Targets are written relative to xl/worksheets/ (e.g. '../comments1.xml').
              // This resolves that specific convention rather than implementing full OPC
              // relative-URI resolution -- consistent with this library's other pragmatic,
              // regex-based parsing rather than a general-purpose XML/URI toolchain.
              var ResolvedPath := CommentsTarget;
              if ResolvedPath.StartsWith('../') then
                ResolvedPath := 'xl/' + Copy(ResolvedPath, 4, MaxInt);
              if Package.PartExists(ResolvedPath) then
                ParseComments(FSheets[I] as TExcelSheet, Package.GetPartContent(ResolvedPath));
            end;
          finally
            SheetRels.Free;
          end;
        end;
      end;
    finally
      Rels.Free;
    end;

    Package.Close;
  finally
    Package.Free;
  end;
end;

procedure TExcelWorkbook.LoadFromStream(const Stream: TStream);
begin
  var Package := TOXMLPackage.Create;
  try
    Package.Open(Stream);

    if Package.PartExists(PartCoreProps) then
    begin
      const CoreXml = Package.GetPartContent(PartCoreProps);
      var Parser := TMetadataParser.Create;
      try
        FMetadata := Parser.Parse(CoreXml);
      finally
        Parser.Free;
      end;
    end;

    if Package.PartExists(PartSharedStrings) then
      ParseSharedStrings(Package.GetPartContent(PartSharedStrings));

    if Package.PartExists(PartStyles) then
      ParseStyles(Package.GetPartContent(PartStyles));

    if Package.PartExists(PartWorkbook) then
      ParseWorkbook(Package.GetPartContent(PartWorkbook));

    var Rels := TRelationships.Create;
    try
      if Package.PartExists(PartWorkbookRels) then
        Rels.LoadFromXml(Package.GetPartContent(PartWorkbookRels));

      for var I := 0 to FSheets.Count - 1 do
      begin
        const SheetPath = PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix;
        if Package.PartExists(SheetPath) then
          ParseSheet(FSheets[I] as TExcelSheet, Package.GetPartContent(SheetPath));

        const SheetRelsPath = 'xl/worksheets/_rels/sheet' + IntToStr(I + 1) + '.xml.rels';
        if Package.PartExists(SheetRelsPath) then
        begin
          var SheetRels := TRelationships.Create;
          try
            SheetRels.LoadFromXml(Package.GetPartContent(SheetRelsPath));
            const CommentsTarget = SheetRels.GetTargetByType(OfficeDocRelsNs + '/comments');
            if CommentsTarget <> '' then
            begin
              var ResolvedPath := CommentsTarget;
              if ResolvedPath.StartsWith('../') then
                ResolvedPath := 'xl/' + Copy(ResolvedPath, 4, MaxInt);
              if Package.PartExists(ResolvedPath) then
                ParseComments(FSheets[I] as TExcelSheet, Package.GetPartContent(ResolvedPath));
            end;
          finally
            SheetRels.Free;
          end;
        end;
      end;
    finally
      Rels.Free;
    end;

    Package.Close;
  finally
    Package.Free;
  end;
end;

procedure TExcelWorkbook.SaveToFile(const FileName: string);
begin
  var SharedStrings := BuildSharedStrings;
  var StyleMap := BuildStyleMap;
  try
    var Zip := TZipFile.Create;
    try
      Zip.Open(FileName, zmWrite);

      const ContentTypes = TEncoding.UTF8.GetBytes(GenerateContentTypes);
      Zip.Add(ContentTypes, '[Content_Types].xml');

      const Rels = TEncoding.UTF8.GetBytes(GenerateRels);
      Zip.Add(Rels, '_rels/.rels');

      const Workbook = TEncoding.UTF8.GetBytes(GenerateWorkbook);
      Zip.Add(Workbook, PartWorkbook);

      const WorkbookRels = TEncoding.UTF8.GetBytes(GenerateWorkbookRels);
      Zip.Add(WorkbookRels, PartWorkbookRels);

      var CommentsFileIndex := 0;
      for var I := 0 to FSheets.Count - 1 do
      begin
        var ExcelSheet := FSheets[I] as TExcelSheet;
        const SheetXml = TEncoding.UTF8.GetBytes(GenerateSheet(ExcelSheet, SharedStrings, StyleMap));
        Zip.Add(SheetXml, PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix);

        if ExcelSheet.HasNotes then
        begin
          Inc(CommentsFileIndex);

          const CommentsXml = TEncoding.UTF8.GetBytes(GenerateComments(ExcelSheet));
          Zip.Add(CommentsXml, 'xl/comments' + IntToStr(CommentsFileIndex) + '.xml');

          const VmlXml = TEncoding.UTF8.GetBytes(GenerateVmlDrawing(ExcelSheet));
          Zip.Add(VmlXml, 'xl/drawings/vmlDrawing' + IntToStr(CommentsFileIndex) + '.vml');

          const SheetRelsXml = TEncoding.UTF8.GetBytes(GenerateSheetRels(CommentsFileIndex));
          Zip.Add(SheetRelsXml, 'xl/worksheets/_rels/sheet' + IntToStr(I + 1) + '.xml.rels');
        end;
      end;

      if SharedStrings.Count > 0 then
      begin
        const SharedStringsXml = TEncoding.UTF8.GetBytes(GenerateSharedStrings(SharedStrings));
        Zip.Add(SharedStringsXml, PartSharedStrings);
      end;

      const StylesXml = TEncoding.UTF8.GetBytes(GenerateStyles(StyleMap));
      Zip.Add(StylesXml, PartStyles);

      Zip.Close;
    finally
      Zip.Free;
    end;
  finally
    StyleMap.Free;
    SharedStrings.Free;
  end;
end;

procedure TExcelWorkbook.SaveToStream(const Stream: TStream);
begin
  var SharedStrings := BuildSharedStrings;
  var StyleMap := BuildStyleMap;
  try
    var Zip := TZipFile.Create;
    try
      Zip.Open(Stream, zmWrite);

      const ContentTypes = TEncoding.UTF8.GetBytes(GenerateContentTypes);
      Zip.Add(ContentTypes, '[Content_Types].xml');

      const Rels = TEncoding.UTF8.GetBytes(GenerateRels);
      Zip.Add(Rels, '_rels/.rels');

      const Workbook = TEncoding.UTF8.GetBytes(GenerateWorkbook);
      Zip.Add(Workbook, PartWorkbook);

      const WorkbookRels = TEncoding.UTF8.GetBytes(GenerateWorkbookRels);
      Zip.Add(WorkbookRels, PartWorkbookRels);

      var CommentsFileIndex := 0;
      for var I := 0 to FSheets.Count - 1 do
      begin
        var ExcelSheet := FSheets[I] as TExcelSheet;
        const SheetXml = TEncoding.UTF8.GetBytes(GenerateSheet(ExcelSheet, SharedStrings, StyleMap));
        Zip.Add(SheetXml, PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix);

        if ExcelSheet.HasNotes then
        begin
          Inc(CommentsFileIndex);

          const CommentsXml = TEncoding.UTF8.GetBytes(GenerateComments(ExcelSheet));
          Zip.Add(CommentsXml, 'xl/comments' + IntToStr(CommentsFileIndex) + '.xml');

          const VmlXml = TEncoding.UTF8.GetBytes(GenerateVmlDrawing(ExcelSheet));
          Zip.Add(VmlXml, 'xl/drawings/vmlDrawing' + IntToStr(CommentsFileIndex) + '.vml');

          const SheetRelsXml = TEncoding.UTF8.GetBytes(GenerateSheetRels(CommentsFileIndex));
          Zip.Add(SheetRelsXml, 'xl/worksheets/_rels/sheet' + IntToStr(I + 1) + '.xml.rels');
        end;
      end;

      if SharedStrings.Count > 0 then
      begin
        const SharedStringsXml = TEncoding.UTF8.GetBytes(GenerateSharedStrings(SharedStrings));
        Zip.Add(SharedStringsXml, PartSharedStrings);
      end;

      const StylesXml = TEncoding.UTF8.GetBytes(GenerateStyles(StyleMap));
      Zip.Add(StylesXml, PartStyles);

      Zip.Close;
    finally
      Zip.Free;
    end;
  finally
    StyleMap.Free;
    SharedStrings.Free;
  end;
end;

function TExcelWorkbook.BuildSharedStrings: TList<string>;
begin
  Result := TList<string>.Create;
  for var Sheet in FSheets do
  begin
    var ExcelSheet := Sheet as TExcelSheet;
    for var Pair in ExcelSheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      if (Cell.IsString) and (Cell.GetAsString <> '') then
        if not Result.Contains(Cell.GetAsString) then
          Result.Add(Cell.GetAsString);
    end;
  end;
end;

function TExcelWorkbook.GetSharedStringIndex(const Strings: TList<string>; const Value: string): Integer;
begin
  Result := Strings.IndexOf(Value);
end;

function TExcelWorkbook.GenerateContentTypes: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Types xmlns="' + ContentTypesNs + '">');
    SB.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    SB.Append('<Default Extension="xml" ContentType="application/xml"/>');

    var HasAnyNotes := False;
    for var Sheet in FSheets do
      if (Sheet as TExcelSheet).HasNotes then
        HasAnyNotes := True;
    if HasAnyNotes then
      SB.Append('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');

    SB.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    for var I := 0 to FSheets.Count - 1 do
      SB.Append('<Override PartName="/xl/worksheets/sheet' + IntToStr(I + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');

    if HasAnyNotes then
    begin
      var CommentsFileIndex := 0;
      for var Sheet in FSheets do
        if (Sheet as TExcelSheet).HasNotes then
        begin
          Inc(CommentsFileIndex);
          SB.Append('<Override PartName="/xl/comments' + IntToStr(CommentsFileIndex) +
            '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>');
        end;
    end;

    SB.Append('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
    SB.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
    SB.Append('</Types>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateRels: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + RelationshipsNs + '">');
    SB.Append('<Relationship Id="rId1" Type="' + OfficeDocRelsNs + '/officeDocument" Target="xl/workbook.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateWorkbook: string;
begin
  var HasVisibleSheet := False;
  for var Sheet in FSheets do
    if Sheet.Visibility = TExcelSheetVisibility.Visible then
    begin
      HasVisibleSheet := True;
      Break;
    end;
  // Excel refuses to open a workbook in which every sheet is hidden.
  if (FSheets.Count > 0) and not HasVisibleSheet then
    raise EExcelWorkbookException.Create('A workbook must contain at least one visible sheet');

  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<workbook xmlns="' + SpreadsheetNs + '" xmlns:r="' + OfficeDocRelsNs + '">');
    SB.Append('<sheets>');
    for var I := 0 to FSheets.Count - 1 do
    begin
      var StateAttr := '';
      if FSheets[I].Visibility = TExcelSheetVisibility.Hidden then
        StateAttr := ' state="hidden"'
      else if FSheets[I].Visibility = TExcelSheetVisibility.VeryHidden then
        StateAttr := ' state="veryHidden"';
      SB.Append('<sheet name="' + FSheets[I].Name + '" sheetId="' + IntToStr(I + 1) + '"' + StateAttr + ' r:id="rId' + IntToStr(I + 1) + '"/>');
    end;
    SB.Append('</sheets>');
    SB.Append('</workbook>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateWorkbookRels: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + RelationshipsNs + '">');
    for var I := 0 to FSheets.Count - 1 do
      SB.Append('<Relationship Id="rId' + IntToStr(I + 1) + '" Type="' + OfficeDocRelsNs + '/worksheet" Target="worksheets/sheet' + IntToStr(I + 1) + '.xml"/>');
    SB.Append('<Relationship Id="rId' + IntToStr(FSheets.Count + 1) + '" Type="' + OfficeDocRelsNs + '/sharedStrings" Target="sharedStrings.xml"/>');
    SB.Append('<Relationship Id="rId' + IntToStr(FSheets.Count + 2) + '" Type="' + OfficeDocRelsNs + '/styles" Target="styles.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateSheet(const Sheet: TExcelSheet; const SharedStrings: TList<string>; const StyleMap: TDictionary<string, Integer>): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<worksheet xmlns="' + SpreadsheetNs + '" xmlns:r="' + OfficeDocRelsNs + '">');

    const HasFrozenPanes = (Sheet.FrozenRows > 0) or (Sheet.FrozenColumns > 0);
    if HasFrozenPanes then
    begin
      const TopLeftCell = TExcelSheet.ColumnNumberToLetters(Sheet.FrozenColumns + 1) + IntToStr(Sheet.FrozenRows + 1);
      var ActivePane := 'topLeft';
      if (Sheet.FrozenRows > 0) and (Sheet.FrozenColumns > 0) then
        ActivePane := 'bottomRight'
      else if Sheet.FrozenRows > 0 then
        ActivePane := 'bottomLeft'
      else if Sheet.FrozenColumns > 0 then
        ActivePane := 'topRight';

      SB.Append('<sheetViews><sheetView workbookViewId="0">');
      SB.Append('<pane xSplit="' + IntToStr(Sheet.FrozenColumns) + '" ySplit="' + IntToStr(Sheet.FrozenRows) +
        '" topLeftCell="' + TopLeftCell + '" activePane="' + ActivePane + '" state="frozen"/>');

      // Excel emits one <selection> per pane that owns a selection: three for a both-axes
      // freeze (topRight, bottomLeft, bottomRight), one for a single-axis freeze. Each
      // pane's active cell is the top-left cell of that pane's own region.
      if (Sheet.FrozenRows > 0) and (Sheet.FrozenColumns > 0) then
      begin
        const TopRightCell = TExcelSheet.ColumnNumberToLetters(Sheet.FrozenColumns + 1) + '1';
        const BottomLeftCell = 'A' + IntToStr(Sheet.FrozenRows + 1);
        SB.Append('<selection pane="topRight" activeCell="' + TopRightCell + '" sqref="' + TopRightCell + '"/>');
        SB.Append('<selection pane="bottomLeft" activeCell="' + BottomLeftCell + '" sqref="' + BottomLeftCell + '"/>');
        SB.Append('<selection pane="bottomRight" activeCell="' + TopLeftCell + '" sqref="' + TopLeftCell + '"/>');
      end
      else
        SB.Append('<selection pane="' + ActivePane + '" activeCell="' + TopLeftCell + '" sqref="' + TopLeftCell + '"/>');

      SB.Append('</sheetView></sheetViews>');
    end;

    const HasColumnWidths = (Sheet.ColumnWidths.Count > 0);
    if HasColumnWidths then
    begin
      SB.Append('<cols>');
      for var ColPair in Sheet.ColumnWidths do
      begin
        const ColNum = TExcelSheet.ColumnLetterToNumber(ColPair.Key);
        const WidthStr = FormatFloat('0.##', ColPair.Value, TFormatSettings.Invariant);
        SB.Append('<col min="' + IntToStr(ColNum) + '" max="' + IntToStr(ColNum) + '" width="' + WidthStr + '" customWidth="1"/>');
      end;
      SB.Append('</cols>');
    end;

    SB.Append('<sheetData>');

    var RowCells := TDictionary<Integer, TList<TPair<string, IExcelCell>>>.Create;
    try
      for var Pair in Sheet.Cells do
      begin
        const Address = Pair.Key;
        var RowNum := 0;
        for var CharPos := 1 to Length(Address) do
          if CharInSet(Address[CharPos], ['0'..'9']) then
          begin
            RowNum := StrToIntDef(Copy(Address, CharPos, Length(Address)), 0);
            Break;
          end;

        if RowNum > 0 then
        begin
          if not RowCells.ContainsKey(RowNum) then
            RowCells.Add(RowNum, TList<TPair<string, IExcelCell>>.Create);
          RowCells[RowNum].Add(TPair<string, IExcelCell>.Create(Address, Pair.Value));
        end;
      end;

      var SortedRows := TList<Integer>.Create;
      try
        for var Row in RowCells.Keys do
          SortedRows.Add(Row);
        for var HeightRow in Sheet.RowHeights.Keys do
          if not SortedRows.Contains(HeightRow) then
            SortedRows.Add(HeightRow);
        SortedRows.Sort;

        for var Row in SortedRows do
        begin
          var RowHeight: Double := 0;
          if Sheet.RowHeights.TryGetValue(Row, RowHeight) then
            SB.Append('<row r="' + IntToStr(Row) + '" ht="' +
              FormatFloat('0.##', RowHeight, TFormatSettings.Invariant) + '" customHeight="1">')
          else
            SB.Append('<row r="' + IntToStr(Row) + '">');

          if not RowCells.ContainsKey(Row) then
          begin
            SB.Append('</row>');
            Continue;
          end;

          RowCells[Row].Sort(TComparer<TPair<string, IExcelCell>>.Construct(
            function(const Left, Right: TPair<string, IExcelCell>): Integer
            var
              LeftCol, RightCol: string;
              CharIndex: Integer;
            begin
              LeftCol := '';
              for CharIndex := 1 to Length(Left.Key) do
                if CharInSet(Left.Key[CharIndex], ['A'..'Z']) then
                  LeftCol := LeftCol + Left.Key[CharIndex]
                else
                  Break;

              RightCol := '';
              for CharIndex := 1 to Length(Right.Key) do
                if CharInSet(Right.Key[CharIndex], ['A'..'Z']) then
                  RightCol := RightCol + Right.Key[CharIndex]
                else
                  Break;

              if Length(LeftCol) <> Length(RightCol) then
                Result := Length(LeftCol) - Length(RightCol)
              else
                Result := CompareStr(LeftCol, RightCol);
            end));

          for var CellPair in RowCells[Row] do
          begin
            var Cell := CellPair.Value as TExcelCell;
            const StyleIdx = GetCellStyleIndex(Cell, StyleMap);
            var StyleAttr := '';
            if StyleIdx > 0 then
              StyleAttr := ' s="' + IntToStr(StyleIdx) + '"';

            if Cell.GetHasFormula then
            begin
              const FloatVal = FormatFloat('0.##############', Cell.GetAsFloat, TFormatSettings.Invariant);
              SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '><f>' + Cell.GetFormula + '</f><v>' + FloatVal + '</v></c>');
            end
            else
            case Cell.CellType of
              TCellType.Empty:
                begin
                  if StyleIdx > 0 then
                    SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '/>');
                end;
              TCellType.StringValue:
                begin
                  // Empty strings are excluded from sharedStrings (see BuildSharedStrings), so a
                  // lookup would yield the invalid index -1. Write them as value-less cells instead.
                  if Cell.GetAsString = '' then
                  begin
                    if StyleIdx > 0 then
                      SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '/>');
                  end
                  else
                  begin
                    const StrIdx = GetSharedStringIndex(SharedStrings, Cell.GetAsString);
                    SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + ' t="s"><v>' + IntToStr(StrIdx) + '</v></c>');
                  end;
                end;
              TCellType.Boolean:
                begin
                  const BoolVal = IfThen(Cell.GetAsBoolean, '1', '0');
                  SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + ' t="b"><v>' + BoolVal + '</v></c>');
                end;
              TCellType.Number:
                begin
                  const FloatVal = FormatFloat('0.##############', Cell.GetAsFloat, TFormatSettings.Invariant);
                  SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '><v>' + FloatVal + '</v></c>');
                end;
              TCellType.DateTime:
                begin
                  const FloatVal = FormatFloat('0.##############', Cell.GetAsFloat, TFormatSettings.Invariant);
                  SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '><v>' + FloatVal + '</v></c>');
                end;
            end;
          end;
          SB.Append('</row>');
        end;
      finally
        SortedRows.Free;
      end;

      for var CellList in RowCells.Values do
        CellList.Free;
    finally
      RowCells.Free;
    end;

    SB.Append('</sheetData>');

    const HasMergedRanges = (Sheet.MergedRangesList.Count > 0);
    if HasMergedRanges then
    begin
      SB.Append('<mergeCells count="' + IntToStr(Sheet.MergedRangesList.Count) + '">');
      for var MergeRange in Sheet.MergedRangesList do
        SB.Append('<mergeCell ref="' + MergeRange + '"/>');
      SB.Append('</mergeCells>');
    end;

    // legacyDrawing must come very late in the worksheet element per the CT_Worksheet
    // schema order (after sheetData/mergeCells, before the closing tag). Its r:id always
    // resolves to "rId1" -- see GenerateSheetRels, where the vmlDrawing relationship is
    // written first and the comments relationship (not referenced here at all; Excel
    // discovers it purely via the .rels Type) second.
    if Sheet.HasNotes then
      SB.Append('<legacyDrawing r:id="rId1"/>');

    SB.Append('</worksheet>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateComments(const Sheet: TExcelSheet): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<comments xmlns="' + SpreadsheetNs + '">');
    // A single generic author is used for every note -- per-note author names are out of
    // scope for this plain-text-only feature.
    SB.Append('<authors><author>Author</author></authors>');
    SB.Append('<commentList>');
    for var NotePair in Sheet.Notes do
      SB.Append('<comment ref="' + NotePair.Key + '" authorId="0"><text><r><t>' +
        TXml.Escape(NotePair.Value) + '</t></r></text></comment>');
    SB.Append('</commentList>');
    SB.Append('</comments>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateVmlDrawing(const Sheet: TExcelSheet): string;
begin
  var SB := TStringBuilder.Create;
  try
    // VML is a legacy, non-SpreadsheetML vocabulary Excel still requires for the visual
    // note box (position/size/default-hidden state). The shapetype is defined once and
    // shared by every note's <v:shape> via type="#_x0000_t202". Shape ids just need to be
    // unique within this file; 1000+N avoids colliding with the shapetype's own id.
    SB.Append('<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:x="urn:schemas-microsoft-com:office:excel">');
    SB.Append('<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>');
    SB.Append('<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
      'path="m,l,21600r21600,l21600,xe">' +
      '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>' +
      '</v:shapetype>');

    var ShapeId := 1000;
    for var NotePair in Sheet.Notes do
    begin
      Inc(ShapeId);
      var Col, Row: Integer;
      TExcelSheet.ParseCellAddress(NotePair.Key, Col, Row);
      // x:Row/x:Column are 0-based, unlike the rest of this library's 1-based row/column
      // convention -- ParseCellAddress returns 1-based values, so both are adjusted here.
      const ZeroRow = Row - 1;
      const ZeroCol = Col - 1;

      SB.Append('<v:shape id="_x0000_s' + IntToStr(ShapeId) + '" type="#_x0000_t202" ' +
        'style=''position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;' +
        'height:59.25pt;z-index:' + IntToStr(ShapeId) + ';visibility:hidden'' ' +
        'fillcolor="#ffffe1" o:insetmode="auto">');
      SB.Append('<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>');
      SB.Append('<v:path o:connecttype="none"/>');
      SB.Append('<v:textbox style=''mso-direction-alt:auto''><div style=''text-align:left''></div></v:textbox>');
      // x:Anchor positions the note box relative to its own cell (from/to column and row
      // with in-cell offsets), so each note sits next to its cell instead of every box
      // landing on the same fixed margin. This mirrors Excel's default: the box starts one
      // column to the right, spans two columns and about four rows.
      const Anchor = Format('%d, 15, %d, 2, %d, 15, %d, 4',
        [ZeroCol + 1, ZeroRow, ZeroCol + 3, ZeroRow + 4]);
      SB.Append('<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        '<x:Anchor>' + Anchor + '</x:Anchor>' +
        '<x:AutoFill>False</x:AutoFill>' +
        '<x:Row>' + IntToStr(ZeroRow) + '</x:Row>' +
        '<x:Column>' + IntToStr(ZeroCol) + '</x:Column>' +
        '</x:ClientData>');
      SB.Append('</v:shape>');
    end;

    SB.Append('</xml>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateSheetRels(const CommentsFileIndex: Integer): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + RelationshipsNs + '">');
    SB.Append('<Relationship Id="rId1" Type="' + OfficeDocRelsNs + '/vmlDrawing" Target="../drawings/vmlDrawing' +
      IntToStr(CommentsFileIndex) + '.vml"/>');
    SB.Append('<Relationship Id="rId2" Type="' + OfficeDocRelsNs + '/comments" Target="../comments' +
      IntToStr(CommentsFileIndex) + '.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GenerateSharedStrings(const Strings: TList<string>): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<sst xmlns="' + SpreadsheetNs + '" count="' + IntToStr(Strings.Count) + '" uniqueCount="' + IntToStr(Strings.Count) + '">');
    for var StringItem in Strings do
      SB.Append('<si><t>' + TXml.Escape(StringItem) + '</t></si>');
    SB.Append('</sst>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbook.GetStyleKey(const Cell: TExcelCell): string;
begin
  var FontSizeStr := '';
  if Cell.FFontSize <> 0 then
    FontSizeStr := FormatFloat('0.##', Cell.FFontSize, TFormatSettings.Invariant);
  // 0 = no date format, 1 = date only (built-in numFmtId 14), 2 = date with time (built-in numFmtId 22)
  var DateFlag := 0;
  if Cell.CellType = TCellType.DateTime then
    if Frac(Cell.FDateTimeValue) <> 0 then
      DateFlag := 2
    else
      DateFlag := 1;
  Result := Format('%d|%d|%d|%s|%d|%d|%s|%s|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d', [
    Ord(Cell.FBold),
    Cell.FBackgroundColor,
    DateFlag,
    Cell.FNumberFormat,
    Ord(Cell.FItalic),
    Ord(Cell.FUnderline),
    Cell.FFontName,
    FontSizeStr,
    Ord(Cell.FBorderStyle[TExcelBorderSide.Top]),             // 8
    Cell.FBorderColor[TExcelBorderSide.Top],                  // 9
    Ord(Cell.FBorderStyle[TExcelBorderSide.Right]),           // 10
    Cell.FBorderColor[TExcelBorderSide.Right],                // 11
    Ord(Cell.FBorderStyle[TExcelBorderSide.Bottom]),          // 12
    Cell.FBorderColor[TExcelBorderSide.Bottom],               // 13
    Ord(Cell.FBorderStyle[TExcelBorderSide.Left]),            // 14
    Cell.FBorderColor[TExcelBorderSide.Left],                 // 15
    Ord(Cell.FHAlign),
    Ord(Cell.FVAlign),
    Ord(Cell.FWrapText),
    Cell.FFontColor,
    Ord(Cell.FStrikeout)
  ]);
end;

function TExcelWorkbook.BuildStyleMap: TDictionary<string, Integer>;
begin
  Result := TDictionary<string, Integer>.Create;

  // Built from real TExcelCell instances via GetStyleKey rather than hand-typed literal
  // strings, so these can never drift out of sync with GetStyleKey's field list/order.
  var PlainCell := TExcelCell.Create;
  try
    Result.Add(GetStyleKey(PlainCell), 0);
  finally
    PlainCell.Free;
  end;

  var DateCell := TExcelCell.Create;
  try
    DateCell.SetAsDateTime(Trunc(Now)); // Trunc -> Frac = 0 -> DateFlag = 1 (date-only, numFmtId 14)
    Result.Add(GetStyleKey(DateCell), 1);
  finally
    DateCell.Free;
  end;

  var NextIndex := 2;
  for var Sheet in FSheets do
  begin
    var ExcelSheet := Sheet as TExcelSheet;
    for var Pair in ExcelSheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      const Key = GetStyleKey(Cell);
      if not Result.ContainsKey(Key) then
      begin
        Result.Add(Key, NextIndex);
        Inc(NextIndex);
      end;
    end;
  end;
end;

function TExcelWorkbook.GetCellStyleIndex(const Cell: TExcelCell; const StyleMap: TDictionary<string, Integer>): Integer;
begin
  const Key = GetStyleKey(Cell);
  if StyleMap.ContainsKey(Key) then
    Result := StyleMap[Key]
  else
    Result := 0;
end;

function TExcelWorkbook.GenerateStyles(const StyleMap: TDictionary<string, Integer>): string;

  function BorderStyleToString(const Style: TExcelBorderStyle): string;
  begin
    case Style of
      TExcelBorderStyle.Thin:   Result := 'thin';
      TExcelBorderStyle.Medium: Result := 'medium';
      TExcelBorderStyle.Thick:  Result := 'thick';
      TExcelBorderStyle.Dashed: Result := 'dashed';
      TExcelBorderStyle.Dotted: Result := 'dotted';
      TExcelBorderStyle.Double: Result := 'double';
    else
      Result := '';
    end;
  end;

  function HAlignToString(const Align: TExcelHAlign): string;
  begin
    case Align of
      TExcelHAlign.Left:    Result := 'left';
      TExcelHAlign.Center:  Result := 'center';
      TExcelHAlign.Right:   Result := 'right';
      TExcelHAlign.Justify: Result := 'justify';
    else
      Result := '';
    end;
  end;

  function VAlignToString(const Align: TExcelVAlign): string;
  begin
    case Align of
      TExcelVAlign.Top:    Result := 'top';
      TExcelVAlign.Center: Result := 'center';
      TExcelVAlign.Bottom: Result := 'bottom';
    else
      Result := '';
    end;
  end;

  function SideElement(const Tag: string; AStyle: TExcelBorderStyle; AColor: Cardinal): string;
  begin
    if AStyle = TExcelBorderStyle.None then
      Exit('<' + Tag + '/>');
    var ColorAttr := '';
    if AColor <> 0 then
      ColorAttr := '<color rgb="FF' + IntToHex(AColor, 6) + '"/>'
    else
      ColorAttr := '<color auto="1"/>';
    Result := '<' + Tag + ' style="' + BorderStyleToString(AStyle) + '">' + ColorAttr + '</' + Tag + '>';
  end;
begin
  var Colors := TList<Cardinal>.Create;
  var NumFormats := TList<string>.Create;
  var FontKeys := TList<string>.Create;
  var BorderKeys := TList<string>.Create;

  for var Sheet in FSheets do
  begin
    var ExcelSheet := Sheet as TExcelSheet;
    for var Pair in ExcelSheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      if (Cell.FBackgroundColor <> 0) and (not Colors.Contains(Cell.FBackgroundColor)) then
        Colors.Add(Cell.FBackgroundColor);
      if (Cell.FNumberFormat <> '') and (not NumFormats.Contains(Cell.FNumberFormat)) then
        NumFormats.Add(Cell.FNumberFormat);

      var FontSizeStr := '';
      if Cell.FFontSize <> 0 then
        FontSizeStr := FormatFloat('0.##', Cell.FFontSize, TFormatSettings.Invariant);
      const FontKey = Format('%d|%d|%d|%s|%s|%d|%d', [
        Ord(Cell.FBold), Ord(Cell.FItalic), Ord(Cell.FUnderline), Cell.FFontName, FontSizeStr, Cell.FFontColor, Ord(Cell.FStrikeout)]);
      if (FontKey <> '0|0|0|||0|0') and (not FontKeys.Contains(FontKey)) then
        FontKeys.Add(FontKey);

      const BorderKey = Format('%d|%d|%d|%d|%d|%d|%d|%d', [
        Ord(Cell.FBorderStyle[TExcelBorderSide.Top]), Cell.FBorderColor[TExcelBorderSide.Top],
        Ord(Cell.FBorderStyle[TExcelBorderSide.Right]), Cell.FBorderColor[TExcelBorderSide.Right],
        Ord(Cell.FBorderStyle[TExcelBorderSide.Bottom]), Cell.FBorderColor[TExcelBorderSide.Bottom],
        Ord(Cell.FBorderStyle[TExcelBorderSide.Left]), Cell.FBorderColor[TExcelBorderSide.Left]]);
      if (BorderKey <> '0|0|0|0|0|0|0|0') and (not BorderKeys.Contains(BorderKey)) then
        BorderKeys.Add(BorderKey);
    end;
  end;

  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<styleSheet xmlns="' + SpreadsheetNs + '">');

    // Date cells use the built-in locale-aware formats (14 = short date, 22 = date + time),
    // so only user-supplied custom formats need numFmt entries. Custom ids start at 165.
    if NumFormats.Count > 0 then
    begin
      SB.Append('<numFmts count="' + IntToStr(NumFormats.Count) + '">');
      for var I := 0 to NumFormats.Count - 1 do
        SB.Append('<numFmt numFmtId="' + IntToStr(165 + I) + '" formatCode="' + TXml.Escape(NumFormats[I]) + '"/>');
      SB.Append('</numFmts>');
    end;

    const FontCount = 1 + FontKeys.Count;
    SB.Append('<fonts count="' + IntToStr(FontCount) + '">');
    SB.Append('<font><sz val="11"/><name val="Calibri"/></font>');
    for var FontKey in FontKeys do
    begin
      const FontParts = FontKey.Split(['|']);
      const IsBold = FontParts[0] = '1';
      const IsItalic = FontParts[1] = '1';
      const IsUnderline = FontParts[2] = '1';
      const Name = FontParts[3];
      const Size = FontParts[4];
      const FontColor = StrToIntDef(FontParts[5], 0);
      const IsStrikeout = FontParts[6] = '1';
      SB.Append('<font>');
      if IsBold then SB.Append('<b/>');
      if IsItalic then SB.Append('<i/>');
      if IsUnderline then SB.Append('<u/>');
      if IsStrikeout then SB.Append('<strike/>');
      if Size <> '' then
        SB.Append('<sz val="' + Size + '"/>')
      else
        SB.Append('<sz val="11"/>');
      if FontColor <> 0 then
        SB.Append('<color rgb="FF' + IntToHex(FontColor, 6) + '"/>');
      if Name <> '' then
        SB.Append('<name val="' + TXml.Escape(Name) + '"/>')
      else
        SB.Append('<name val="Calibri"/>');
      SB.Append('</font>');
    end;
    SB.Append('</fonts>');

    const FillCount = 2 + Colors.Count;
    SB.Append('<fills count="' + IntToStr(FillCount) + '">');
    SB.Append('<fill><patternFill patternType="none"/></fill>');
    SB.Append('<fill><patternFill patternType="gray125"/></fill>');
    for var Color in Colors do
      SB.Append('<fill><patternFill patternType="solid"><fgColor rgb="FF' + IntToHex(Color, 6) + '"/></patternFill></fill>');
    SB.Append('</fills>');

    const BorderCount = 1 + BorderKeys.Count;
    SB.Append('<borders count="' + IntToStr(BorderCount) + '">');
    SB.Append('<border/>');
    for var BorderKey in BorderKeys do
    begin
      const BorderParts = BorderKey.Split(['|']);
      const TopStyle    = TExcelBorderStyle(StrToIntDef(BorderParts[0], 0));
      const TopColor    = StrToIntDef(BorderParts[1], 0);
      const RightStyle  = TExcelBorderStyle(StrToIntDef(BorderParts[2], 0));
      const RightColor  = StrToIntDef(BorderParts[3], 0);
      const BottomStyle = TExcelBorderStyle(StrToIntDef(BorderParts[4], 0));
      const BottomColor = StrToIntDef(BorderParts[5], 0);
      const LeftStyle   = TExcelBorderStyle(StrToIntDef(BorderParts[6], 0));
      const LeftColor   = StrToIntDef(BorderParts[7], 0);

      SB.Append('<border>');
      SB.Append(SideElement('left', LeftStyle, LeftColor));
      SB.Append(SideElement('right', RightStyle, RightColor));
      SB.Append(SideElement('top', TopStyle, TopColor));
      SB.Append(SideElement('bottom', BottomStyle, BottomColor));
      SB.Append('</border>');
    end;
    SB.Append('</borders>');

    SB.Append('<cellStyleXfs count="1">');
    SB.Append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    SB.Append('</cellStyleXfs>');

    SB.Append('<cellXfs count="' + IntToStr(StyleMap.Count) + '">');
    var SortedStyles := TList<TPair<string, Integer>>.Create;
    try
      for var Pair in StyleMap do
        SortedStyles.Add(Pair);
      SortedStyles.Sort(TComparer<TPair<string, Integer>>.Construct(
        function(const Left, Right: TPair<string, Integer>): Integer
        begin
          Result := Left.Value - Right.Value;
        end));

      for var StylePair in SortedStyles do
      begin
        const Parts = StylePair.Key.Split(['|']);
        const IsBold = Parts[0] = '1';
        const BgColor = StrToIntDef(Parts[1], 0);
        const DateFlag = StrToIntDef(Parts[2], 0);
        const CustomFormat = Parts[3];
        const IsItalic = Parts[4] = '1';
        const IsUnderline = Parts[5] = '1';
        const CellFontName = Parts[6];
        const CellFontSize = Parts[7];
        const CellBorderTopStyle    = TExcelBorderStyle(StrToIntDef(Parts[8], 0));
        const CellBorderTopColor    = StrToIntDef(Parts[9], 0);
        const CellBorderRightStyle  = TExcelBorderStyle(StrToIntDef(Parts[10], 0));
        const CellBorderRightColor  = StrToIntDef(Parts[11], 0);
        const CellBorderBottomStyle = TExcelBorderStyle(StrToIntDef(Parts[12], 0));
        const CellBorderBottomColor = StrToIntDef(Parts[13], 0);
        const CellBorderLeftStyle   = TExcelBorderStyle(StrToIntDef(Parts[14], 0));
        const CellBorderLeftColor   = StrToIntDef(Parts[15], 0);
        const CellHAlign = TExcelHAlign(StrToIntDef(Parts[16], 0));
        const CellVAlign = TExcelVAlign(StrToIntDef(Parts[17], 0));
        const CellWrapText = Parts[18] = '1';
        const CellFontColor = StrToIntDef(Parts[19], 0);
        const CellStrikeout = Parts[20] = '1';

        // Must match the FontKeys population format in the collection loop above,
        // field-for-field — this 5-vs-6-field mismatch was the original FontColor bug.
        const FontKey = Format('%d|%d|%d|%s|%s|%d|%d', [
          Ord(IsBold), Ord(IsItalic), Ord(IsUnderline), CellFontName, CellFontSize, CellFontColor, Ord(CellStrikeout)]);
        var FontId := 0;
        if FontKey <> '0|0|0|||0|0' then
          FontId := 1 + FontKeys.IndexOf(FontKey);

        var FillId := 0;
        if BgColor <> 0 then
          FillId := 2 + Colors.IndexOf(BgColor);

        const BorderKey = Format('%d|%d|%d|%d|%d|%d|%d|%d', [
          Ord(CellBorderTopStyle), CellBorderTopColor,
          Ord(CellBorderRightStyle), CellBorderRightColor,
          Ord(CellBorderBottomStyle), CellBorderBottomColor,
          Ord(CellBorderLeftStyle), CellBorderLeftColor]);
        var BorderId := 0;
        if BorderKey <> '0|0|0|0|0|0|0|0' then
          BorderId := 1 + BorderKeys.IndexOf(BorderKey);

        // A user-supplied NumberFormat overrides the default date format, so custom
        // formats also work on date cells.
        var NumFmtId := 0;
        if CustomFormat <> '' then
          NumFmtId := 165 + NumFormats.IndexOf(CustomFormat)
        else if DateFlag = 1 then
          NumFmtId := 14
        else if DateFlag = 2 then
          NumFmtId := 22;

        var ApplyAttrs := '';
        if NumFmtId <> 0 then ApplyAttrs := ApplyAttrs + ' applyNumberFormat="1"';
        if FontId <> 0 then ApplyAttrs := ApplyAttrs + ' applyFont="1"';
        if FillId <> 0 then ApplyAttrs := ApplyAttrs + ' applyFill="1"';
        if BorderId <> 0 then ApplyAttrs := ApplyAttrs + ' applyBorder="1"';

        const HasAlignment = (CellHAlign <> TExcelHAlign.None) or (CellVAlign <> TExcelVAlign.None) or (CellWrapText);
        if HasAlignment then ApplyAttrs := ApplyAttrs + ' applyAlignment="1"';

        if HasAlignment then
        begin
          SB.Append('<xf numFmtId="' + IntToStr(NumFmtId) + '" fontId="' + IntToStr(FontId) +
            '" fillId="' + IntToStr(FillId) + '" borderId="' + IntToStr(BorderId) + '" xfId="0"' + ApplyAttrs + '>');
          var AlignAttrs := '';
          if CellHAlign <> TExcelHAlign.None then
            AlignAttrs := AlignAttrs + ' horizontal="' + HAlignToString(CellHAlign) + '"';
          if CellVAlign <> TExcelVAlign.None then
            AlignAttrs := AlignAttrs + ' vertical="' + VAlignToString(CellVAlign) + '"';
          if CellWrapText then
            AlignAttrs := AlignAttrs + ' wrapText="1"';
          SB.Append('<alignment' + AlignAttrs + '/>');
          SB.Append('</xf>');
        end
        else
          SB.Append('<xf numFmtId="' + IntToStr(NumFmtId) + '" fontId="' + IntToStr(FontId) +
            '" fillId="' + IntToStr(FillId) + '" borderId="' + IntToStr(BorderId) + '" xfId="0"' + ApplyAttrs + '/>');
      end;
    finally
      SortedStyles.Free;
    end;
    SB.Append('</cellXfs>');

    SB.Append('</styleSheet>');
    Result := SB.ToString;
  finally
    SB.Free;
    BorderKeys.Free;
    FontKeys.Free;
    NumFormats.Free;
    Colors.Free;
  end;
end;

procedure TExcelWorkbook.ParseWorkbook(const Xml: string);
begin
  const Matches = TRegEx.Matches(Xml, '<sheet\s+name="([^"]+)"([^>]*?)(?:/>|>)', [roIgnoreCase]);
  for var Match in Matches do
    if Match.Groups.Count > 1 then
    begin
      const Sheet = TExcelSheet.Create(Match.Groups[1].Value);
      Sheet.FOwner := Self;
      const StateMatch = TRegEx.Match(Match.Groups[2].Value, 'state="([^"]+)"', [roIgnoreCase]);
      if StateMatch.Success then
      begin
        if SameText(StateMatch.Groups[1].Value, 'hidden') then
          Sheet.Visibility := TExcelSheetVisibility.Hidden
        else if SameText(StateMatch.Groups[1].Value, 'veryHidden') then
          Sheet.Visibility := TExcelSheetVisibility.VeryHidden;
      end;
      FSheets.Add(Sheet);
    end;
end;

procedure TExcelWorkbook.ParseSharedStrings(const Xml: string);
begin
  FSharedStrings.Clear;
  // Every <si> occupies exactly one index slot, including empty entries (<si><t/></si>)
  // and rich-text entries (<si><r><t>..</t></r><r><t>..</t></r></si>). Matching across
  // si boundaries shifts all subsequent indices, mapping cells to the wrong strings.
  const SiMatches = TRegEx.Matches(Xml, '<si>(.*?)</si>', [roIgnoreCase, roSingleLine]);
  for var SiMatch in SiMatches do
    if SiMatch.Groups.Count > 1 then
    begin
      // Phonetic guide blocks carry their own <t> elements that are not part of the text.
      const SiXml = TRegEx.Replace(SiMatch.Groups[1].Value, '<rPh\s[^>]*>.*?</rPh>', '', [roIgnoreCase, roSingleLine]);
      var Text := '';
      const TextMatches = TRegEx.Matches(SiXml, '<t(?:\s[^>]*)?>([^<]*)</t>', [roIgnoreCase]);
      for var TextMatch in TextMatches do
        if TextMatch.Groups.Count > 1 then
          Text := Text + TXml.Unescape(TextMatch.Groups[1].Value);
      FSharedStrings.Add(Text);
    end;
end;

procedure TExcelWorkbook.ParseStyles(const Xml: string);

  function StringToBorderStyle(const S: string): TExcelBorderStyle;
  begin
    if S = 'thin' then Result := TExcelBorderStyle.Thin
    else if S = 'medium' then Result := TExcelBorderStyle.Medium
    else if S = 'thick' then Result := TExcelBorderStyle.Thick
    else if S = 'dashed' then Result := TExcelBorderStyle.Dashed
    else if S = 'dotted' then Result := TExcelBorderStyle.Dotted
    else if S = 'double' then Result := TExcelBorderStyle.Double
    else Result := TExcelBorderStyle.None;
  end;

  function StringToHAlign(const S: string): TExcelHAlign;
  begin
    if S = 'left' then Result := TExcelHAlign.Left
    else if S = 'center' then Result := TExcelHAlign.Center
    else if S = 'right' then Result := TExcelHAlign.Right
    else if S = 'justify' then Result := TExcelHAlign.Justify
    else Result := TExcelHAlign.None;
  end;

  function StringToVAlign(const S: string): TExcelVAlign;
  begin
    if S = 'top' then Result := TExcelVAlign.Top
    else if S = 'center' then Result := TExcelVAlign.Center
    else if S = 'bottom' then Result := TExcelVAlign.Bottom
    else Result := TExcelVAlign.None;
  end;

  procedure ParseSide(const BorderXml, Tag: string; out AStyle: TExcelBorderStyle; out AColor: Cardinal);
  begin
    AStyle := TExcelBorderStyle.None;
    AColor := 0;

    // Isolate just this side's element first. Matching the whole <border> and then
    // scanning to the first <color> lets a self-closing sibling (e.g. <top/>) bleed the
    // colour of a later side (e.g. <bottom>) into this one, so restrict style/colour to
    // this element. The alternation handles both <top/> and <top ...>...</top>.
    const SideMatch = TRegEx.Match(BorderXml,
      '<' + Tag + '\b[^>]*(?:/>|>.*?</' + Tag + '>)', [roIgnoreCase, roSingleLine]);
    if not SideMatch.Success then
      Exit;

    const SideXml = SideMatch.Value;

    const StyleMatch = TRegEx.Match(SideXml, 'style="([^"]*)"', [roIgnoreCase]);
    if StyleMatch.Success then
      AStyle := StringToBorderStyle(StyleMatch.Groups[1].Value);

    const ColorMatch = TRegEx.Match(SideXml, '<color\s+rgb="FF([0-9A-Fa-f]{6})"', [roIgnoreCase]);
    if ColorMatch.Success then
      AColor := StrToInt64Def('$' + ColorMatch.Groups[1].Value, 0);
  end;

  // Built-in numFmtId values 14-22 are the standard date/time formats
  // (e.g. 14 = m/d/yyyy, 22 = m/d/yyyy h:mm); 45-47 are the built-in
  // duration/time formats (mm:ss, [h]:mm:ss, mmss.0). IDs 0-163 are
  // reserved for built-ins; anything else is either "General"/numeric
  // or unused, so it isn't a date.
  function IsBuiltInDateNumFmtId(const NumFmtId: Integer): Boolean;
  begin
    Result := ((NumFmtId >= 14) and (NumFmtId <= 22)) or
              ((NumFmtId >= 45) and (NumFmtId <= 47));
  end;

  // Format codes for the non-date built-in numFmtIds (ECMA-376 part 1,
  // section 18.8.30). Files authored by Excel reference these by id only,
  // without a <numFmt> element, so the codes have to be supplied here to
  // survive a round-trip (they are re-emitted as custom formats on save).
  function BuiltInFormatCode(const NumFmtId: Integer): string;
  begin
    case NumFmtId of
      1: Result := '0';
      2: Result := '0.00';
      3: Result := '#,##0';
      4: Result := '#,##0.00';
      9: Result := '0%';
      10: Result := '0.00%';
      11: Result := '0.00E+00';
      12: Result := '# ?/?';
      13: Result := '# ??/??';
      37: Result := '#,##0 ;(#,##0)';
      38: Result := '#,##0 ;[Red](#,##0)';
      39: Result := '#,##0.00;(#,##0.00)';
      40: Result := '#,##0.00;[Red](#,##0.00)';
      48: Result := '##0.0E+0';
      49: Result := '@';
    else
      Result := '';
    end;
  end;

  // Heuristic used to classify a custom (non-built-in) format code, e.g.
  // "yyyy-mm-dd" or "dd/mm/yyyy hh:mm". Quoted literal text and bracketed
  // sections (colour tags like [Red], locale tags like [$-409], or
  // conditional tags like [>=100]) are stripped first so that literal
  // text or non-date bracket tags can't be mistaken for date components.
  // What remains is checked for the date/time placeholder letters
  // (y, d, h, s, and AM/PM); if any are present the format is a date.
  // "m" is deliberately excluded on its own, since standalone "m" is
  // ambiguous with numeric formats, and a month-only format with no
  // year/day/time component is vanishingly rare in practice.
  function IsDateFormatCode(const FormatCode: string): Boolean;
  var
    Cleaned: string;
  begin
    if (FormatCode = '') or SameText(FormatCode, 'General') then
      Exit(False);

    // Only the first section (before an unescaped ';') applies to
    // positive values, which is what a date serial number is.
    Cleaned := TRegEx.Match(FormatCode, '^((?:[^;"\[]|"[^"]*"|\[[^\]]*\])*)').Groups[1].Value;
    Cleaned := TRegEx.Replace(Cleaned, '"[^"]*"', '', [roIgnoreCase]);
    Cleaned := TRegEx.Replace(Cleaned, '\[[^\]]*\]', '', [roIgnoreCase]);

    Result := TRegEx.IsMatch(Cleaned, '[ydhsYDHS]') or (Pos('AM/PM', UpperCase(Cleaned)) > 0);
  end;

begin
  FStyleBold.Clear;
  FStyleItalic.Clear;
  FStyleUnderline.Clear;
  FStyleFontName.Clear;
  FStyleFontSize.Clear;
  FStyleColors.Clear;
  FStyleFontColor.Clear;
  FStyleStrikeout.Clear;
  FStyleBorderTopStyle.Clear;
  FStyleBorderTopColor.Clear;
  FStyleBorderRightStyle.Clear;
  FStyleBorderRightColor.Clear;
  FStyleBorderBottomStyle.Clear;
  FStyleBorderBottomColor.Clear;
  FStyleBorderLeftStyle.Clear;
  FStyleBorderLeftColor.Clear;
  FStyleHAlign.Clear;
  FStyleVAlign.Clear;
  FStyleWrapText.Clear;
  FStyleIsDate.Clear;
  FStyleNumberFormat.Clear;

  var FontsBold := TList<Boolean>.Create;
  var FontsItalic := TList<Boolean>.Create;
  var FontsUnderline := TList<Boolean>.Create;
  var FontsStrikeout := TList<Boolean>.Create;
  var FontsName := TList<string>.Create;
  var FontsSize := TList<Double>.Create;
  var Fills := TList<Cardinal>.Create;
  var TopStyles := TList<TExcelBorderStyle>.Create;
  var TopColors := TList<Cardinal>.Create;
  var RightStyles := TList<TExcelBorderStyle>.Create;
  var RightColors := TList<Cardinal>.Create;
  var BottomStyles := TList<TExcelBorderStyle>.Create;
  var BottomColors := TList<Cardinal>.Create;
  var LeftStyles := TList<TExcelBorderStyle>.Create;
  var LeftColors := TList<Cardinal>.Create;
  var FontsColor := TList<Cardinal>.Create;
  var CustomNumFmts := TDictionary<Integer, string>.Create;
  try
    const NumFmtMatches = TRegEx.Matches(Xml,
      '<numFmt\s+numFmtId="(\d+)"\s+formatCode="([^"]*)"', [roIgnoreCase]);
    for var NumFmtMatch in NumFmtMatches do
      CustomNumFmts.AddOrSetValue(
        StrToIntDef(NumFmtMatch.Groups[1].Value, -1),
        TXml.Unescape(NumFmtMatch.Groups[2].Value));

    const FontMatches = TRegEx.Matches(Xml, '<font>(.*?)</font>', [roIgnoreCase, roSingleLine]);
    for var Match in FontMatches do
    begin
      const FontXml = Match.Groups[1].Value;
      FontsBold.Add(Pos('<b/>', FontXml) > 0);
      FontsItalic.Add(Pos('<i/>', FontXml) > 0);
      FontsUnderline.Add(Pos('<u/>', FontXml) > 0);
      FontsStrikeout.Add(Pos('<strike/>', FontXml) > 0);

      const ColorMatch = TRegEx.Match(FontXml, '<color\s+rgb="FF([0-9A-Fa-f]{6})"', [roIgnoreCase]);
      if ColorMatch.Success then
        FontsColor.Add(StrToInt64Def('$' + ColorMatch.Groups[1].Value, 0))
      else
        FontsColor.Add(0);

      const NameMatch = TRegEx.Match(FontXml, '<name\s+val="([^"]*)"', [roIgnoreCase]);
      if NameMatch.Success then
        FontsName.Add(NameMatch.Groups[1].Value)
      else
        FontsName.Add('');

      const SizeMatch = TRegEx.Match(FontXml, '<sz\s+val="([^"]*)"', [roIgnoreCase]);
      if SizeMatch.Success then
        FontsSize.Add(StrToFloatDef(SizeMatch.Groups[1].Value, 0, TFormatSettings.Invariant))
      else
        FontsSize.Add(0);
    end;

    // Build the indexed colour palette. Start with the OOXML default palette.
    // If the styles XML contains a custom <indexedColors> block it replaces
    // the defaults entirely, so we rebuild the list from those entries instead.
    var IndexedPalette := TList<Cardinal>.Create;
    try
      for var I := 0 to High(OoxmlIndexedColors) do
        IndexedPalette.Add(OoxmlIndexedColors[I]);

      const CustomPaletteMatch = TRegEx.Match(Xml,
        '<indexedColors>(.*?)</indexedColors>', [roIgnoreCase, roSingleLine]);
      if CustomPaletteMatch.Success then
      begin
        const RgbMatches = TRegEx.Matches(CustomPaletteMatch.Groups[1].Value,
          '<rgbColor\s+rgb="00([0-9A-Fa-f]{6})"', [roIgnoreCase]);
        if RgbMatches.Count > 0 then
        begin
          IndexedPalette.Clear;
          for var RgbMatch in RgbMatches do
            IndexedPalette.Add(StrToInt64Def('$' + RgbMatch.Groups[1].Value, 0));
        end;
      end;

    const FillMatches = TRegEx.Matches(Xml, '<fill>(.*?)</fill>', [roIgnoreCase, roSingleLine]);
    for var Match in FillMatches do
    begin
      const FillXml = Match.Groups[1].Value;
      const RgbMatch = TRegEx.Match(FillXml, 'fgColor\s+rgb="FF([0-9A-Fa-f]{6})"', [roIgnoreCase]);
      if RgbMatch.Success then
        Fills.Add(StrToInt64Def('$' + RgbMatch.Groups[1].Value, 0))
      else
      begin
        const IdxMatch = TRegEx.Match(FillXml, 'fgColor\s+indexed="(\d+)"', [roIgnoreCase]);
        if IdxMatch.Success then
        begin
          const Idx = StrToIntDef(IdxMatch.Groups[1].Value, -1);
          if (Idx >= 0) and (Idx < IndexedPalette.Count) then
            Fills.Add(IndexedPalette[Idx])
          else
            Fills.Add(0);
        end
        else
          Fills.Add(0);
      end;
    end;
    finally
      IndexedPalette.Free;
    end;

    const BorderMatches = TRegEx.Matches(Xml, '<border\s*/>', [roIgnoreCase]);
    for var Match in BorderMatches do
    begin
      TopStyles.Add(TExcelBorderStyle.None);
      TopColors.Add(0);
      RightStyles.Add(TExcelBorderStyle.None);
      RightColors.Add(0);
      BottomStyles.Add(TExcelBorderStyle.None);
      BottomColors.Add(0);
      LeftStyles.Add(TExcelBorderStyle.None);
      LeftColors.Add(0);
    end;
    const BorderFullMatches = TRegEx.Matches(Xml, '<border>(.*?)</border>', [roIgnoreCase, roSingleLine]);
    for var Match in BorderFullMatches do
    begin
      const BorderXml = Match.Groups[1].Value;
      var AStyle: TExcelBorderStyle;
      var AColor: Cardinal;

      ParseSide(BorderXml, 'top', AStyle, AColor);
      TopStyles.Add(AStyle);
      TopColors.Add(AColor);
      ParseSide(BorderXml, 'right', AStyle, AColor);
      RightStyles.Add(AStyle);
      RightColors.Add(AColor);
      ParseSide(BorderXml, 'bottom', AStyle, AColor);
      BottomStyles.Add(AStyle);
      BottomColors.Add(AColor);
      ParseSide(BorderXml, 'left', AStyle, AColor);
      LeftStyles.Add(AStyle);
      LeftColors.Add(AColor);
    end;

    const CellXfsMatch = TRegEx.Match(Xml, '<cellXfs[^>]*>(.*?)</cellXfs>', [roIgnoreCase, roSingleLine]);
    if CellXfsMatch.Success then
    begin
      const CellXfsXml = CellXfsMatch.Groups[1].Value;
      const XfMatches = TRegEx.Matches(CellXfsXml, '<xf\s[^/>]*(?:/>|(?:[^>]*>.*?</xf>))', [roIgnoreCase, roSingleLine]);
      for var Match in XfMatches do
      begin
        const XfXml = Match.Value;

        const NumFmtIdMatch = TRegEx.Match(XfXml, 'numFmtId="(\d+)"', [roIgnoreCase]);
        var NumFmtId := 0;
        if NumFmtIdMatch.Success then
          NumFmtId := StrToIntDef(NumFmtIdMatch.Groups[1].Value, 0);

        var CustomFormatCode: string;
        if IsBuiltInDateNumFmtId(NumFmtId) then
        begin
          // Built-in date formats are re-emitted as numFmtId 14/22 on save,
          // so no format code needs to be carried on the cell.
          FStyleIsDate.Add(True);
          FStyleNumberFormat.Add('');
        end
        else if CustomNumFmts.TryGetValue(NumFmtId, CustomFormatCode) then
        begin
          FStyleIsDate.Add(IsDateFormatCode(CustomFormatCode));
          FStyleNumberFormat.Add(CustomFormatCode);
        end
        else
        begin
          FStyleIsDate.Add(False);
          FStyleNumberFormat.Add(BuiltInFormatCode(NumFmtId));
        end;

        const FontIdMatch = TRegEx.Match(XfXml, 'fontId="(\d+)"', [roIgnoreCase]);
        var FontId := 0;
        if FontIdMatch.Success then
          FontId := StrToIntDef(FontIdMatch.Groups[1].Value, 0);

        const FillIdMatch = TRegEx.Match(XfXml, 'fillId="(\d+)"', [roIgnoreCase]);
        var FillId := 0;
        if FillIdMatch.Success then
          FillId := StrToIntDef(FillIdMatch.Groups[1].Value, 0);

        const BorderIdMatch = TRegEx.Match(XfXml, 'borderId="(\d+)"', [roIgnoreCase]);
        var BorderId := 0;
        if BorderIdMatch.Success then
          BorderId := StrToIntDef(BorderIdMatch.Groups[1].Value, 0);

        FStyleBold.Add((FontId < FontsBold.Count) and (FontsBold[FontId]));
        FStyleItalic.Add((FontId < FontsItalic.Count) and (FontsItalic[FontId]));
        FStyleUnderline.Add((FontId < FontsUnderline.Count) and (FontsUnderline[FontId]));

        var FName := '';
        if FontId < FontsName.Count then
          FName := FontsName[FontId];
        if SameText(FName, 'Calibri') then FName := '';
        FStyleFontName.Add(FName);

        var FSize: Double := 0;
        if FontId < FontsSize.Count then
          FSize := FontsSize[FontId];
        if SameValue(FSize, 11, 0.001) then FSize := 0;
        FStyleFontSize.Add(FSize);

        var BgColor: Cardinal := 0;
        if FillId < Fills.Count then
          BgColor := Fills[FillId];
        FStyleColors.Add(BgColor);

        var FColor: Cardinal := 0;
        if FontId < FontsColor.Count then
          FColor := FontsColor[FontId];
        FStyleFontColor.Add(FColor);

        FStyleStrikeout.Add((FontId < FontsStrikeout.Count) and (FontsStrikeout[FontId]));

        if BorderId < TopStyles.Count then
          FStyleBorderTopStyle.Add(TopStyles[BorderId])
        else
          FStyleBorderTopStyle.Add(TExcelBorderStyle.None);
        if BorderId < TopColors.Count then
          FStyleBorderTopColor.Add(TopColors[BorderId])
        else
          FStyleBorderTopColor.Add(0);
        if BorderId < RightStyles.Count then
          FStyleBorderRightStyle.Add(RightStyles[BorderId])
        else
          FStyleBorderRightStyle.Add(TExcelBorderStyle.None);
        if BorderId < RightColors.Count then
          FStyleBorderRightColor.Add(RightColors[BorderId])
        else
          FStyleBorderRightColor.Add(0);
        if BorderId < BottomStyles.Count then
          FStyleBorderBottomStyle.Add(BottomStyles[BorderId])
        else
          FStyleBorderBottomStyle.Add(TExcelBorderStyle.None);
        if BorderId < BottomColors.Count then
          FStyleBorderBottomColor.Add(BottomColors[BorderId])
        else
          FStyleBorderBottomColor.Add(0);
        if BorderId < LeftStyles.Count then
          FStyleBorderLeftStyle.Add(LeftStyles[BorderId])
        else
          FStyleBorderLeftStyle.Add(TExcelBorderStyle.None);
        if BorderId < LeftColors.Count then
          FStyleBorderLeftColor.Add(LeftColors[BorderId])
        else
          FStyleBorderLeftColor.Add(0);

        const AlignMatch = TRegEx.Match(XfXml, '<alignment([^/]*)/>', [roIgnoreCase]);
        if AlignMatch.Success then
        begin
          const AlignXml = AlignMatch.Groups[1].Value;
          const HMatch = TRegEx.Match(AlignXml, 'horizontal="([^"]*)"', [roIgnoreCase]);
          if HMatch.Success then
            FStyleHAlign.Add(StringToHAlign(HMatch.Groups[1].Value))
          else
            FStyleHAlign.Add(TExcelHAlign.None);
          const VMatch = TRegEx.Match(AlignXml, 'vertical="([^"]*)"', [roIgnoreCase]);
          if VMatch.Success then
            FStyleVAlign.Add(StringToVAlign(VMatch.Groups[1].Value))
          else
            FStyleVAlign.Add(TExcelVAlign.None);
          FStyleWrapText.Add(Pos('wrapText="1"', AlignXml) > 0);
        end
        else
        begin
          FStyleHAlign.Add(TExcelHAlign.None);
          FStyleVAlign.Add(TExcelVAlign.None);
          FStyleWrapText.Add(False);
        end;
      end;
    end;
  finally
    CustomNumFmts.Free;
    LeftColors.Free;
    LeftStyles.Free;
    BottomColors.Free;
    BottomStyles.Free;
    RightColors.Free;
    RightStyles.Free;
    TopColors.Free;
    TopStyles.Free;
    Fills.Free;
    FontsColor.Free;
    FontsSize.Free;
    FontsName.Free;
    FontsUnderline.Free;
    FontsStrikeout.Free;
    FontsItalic.Free;
    FontsBold.Free;
  end;
end;

function TExcelWorkbook.AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;
// Adjusts relative cell references in a formula by the given row and column
// deltas. Absolute references ($A$1) are left unchanged. Used when expanding shared formulas.
var
  Adjusted: string;
begin
  if (RowDelta = 0) and (ColDelta = 0) then
    Exit(Formula);

  Adjusted := Formula;
  // Match cell references of the form [SheetName!][$]ColLetters[$]RowNumber.
  // We handle relative row (no $ before row number) and relative column (no $ before col letters).
  // Pattern: optionally anchored sheet prefix, then optional $, column letters, optional $, row digits.
  const RefPattern = '(?<![A-Z$])(\$?[A-Z]+)(\$?\d+)';
  var Matches := TRegEx.Matches(Formula, RefPattern, [roIgnoreCase]);
  var OffsetAdjust := 0;
  for var M in Matches do
  begin
    const ColPart = M.Groups[1].Value;  // e.g. 'A' or '$A'
    const RowPart = M.Groups[2].Value;  // e.g. '1' or '$1'
    const ColAbs = ColPart.StartsWith('$');
    const RowAbs = RowPart.StartsWith('$');

    var NewColPart := ColPart;
    var NewRowPart := RowPart;

    if (not RowAbs) and (RowDelta <> 0) then
    begin
      const OldRow = StrToIntDef(RowPart, 0);
      if OldRow > 0 then
        NewRowPart := IntToStr(OldRow + RowDelta);
    end;

    if (not ColAbs) and (ColDelta <> 0) then
    begin
      var ColNum := TExcelSheet.ColumnLetterToNumber(ColPart);
      Inc(ColNum, ColDelta);
      NewColPart := TExcelSheet.ColumnNumberToLetters(ColNum);
    end;

    if (NewColPart <> ColPart) or (NewRowPart <> RowPart) then
    begin
      const OldRef = ColPart + RowPart;
      const NewRef = NewColPart + NewRowPart;
      const Pos = M.Index + OffsetAdjust;  // TMatch.Index is 1-based (Delphi string convention)
      Delete(Adjusted, Pos, Length(OldRef));
      Insert(NewRef, Adjusted, Pos);
      Inc(OffsetAdjust, Length(NewRef) - Length(OldRef));
    end;
  end;
  Result := Adjusted;
end;

procedure TExcelWorkbook.ParseComments(const Sheet: TExcelSheet; const Xml: string);
begin
  const CommentMatches = TRegEx.Matches(Xml, '<comment\s+ref="([^"]+)"[^>]*>(.*?)</comment>', [roIgnoreCase, roSingleLine]);
  for var Match in CommentMatches do
  begin
    const Address = Match.Groups[1].Value;
    const CommentBody = Match.Groups[2].Value;
    var Text := '';
    // Mirrors ParseSharedStrings' <t> extraction: a note's text can be split across
    // multiple <r><t>...</t></r> runs (e.g. if a run has different formatting), which
    // are concatenated. Rich per-run formatting itself is not preserved -- out of scope
    // for plain-text notes.
    const TextMatches = TRegEx.Matches(CommentBody, '<t(?:\s[^>]*)?>([^<]*)</t>', [roIgnoreCase]);
    for var TextMatch in TextMatches do
      if TextMatch.Groups.Count > 1 then
        Text := Text + TXml.Unescape(TextMatch.Groups[1].Value);
    Sheet.SetNote(Address, Text);
  end;
end;

procedure TExcelWorkbook.ParseSheet(const Sheet: TExcelSheet; const Xml: string);
begin
  const PaneMatch = TRegEx.Match(Xml, '<pane\s+[^/]*/>', [roIgnoreCase]);
  if PaneMatch.Success then
  begin
    const PaneXml = PaneMatch.Value;
    // Only a frozen pane encodes xSplit/ySplit as whole row/column counts. An unfrozen
    // "split" pane stores the split-bar position in twentieths of a point (e.g. 2160), so
    // reading that as a frozen count would be meaningless -- skip anything whose state is
    // not frozen or frozenSplit.
    const StateMatch = TRegEx.Match(PaneXml, 'state="([^"]*)"', [roIgnoreCase]);
    const IsFrozen = StateMatch.Success and
      (SameText(StateMatch.Groups[1].Value, 'frozen') or SameText(StateMatch.Groups[1].Value, 'frozenSplit'));
    if IsFrozen then
    begin
      const XSplitMatch = TRegEx.Match(PaneXml, 'xSplit="([^"]*)"', [roIgnoreCase]);
      const YSplitMatch = TRegEx.Match(PaneXml, 'ySplit="([^"]*)"', [roIgnoreCase]);
      if XSplitMatch.Success then
        Sheet.FFrozenColumns := Trunc(StrToFloatDef(XSplitMatch.Groups[1].Value, 0, TFormatSettings.Invariant));
      if YSplitMatch.Success then
        Sheet.FFrozenRows := Trunc(StrToFloatDef(YSplitMatch.Groups[1].Value, 0, TFormatSettings.Invariant));
    end;
  end;

  const ColMatches = TRegEx.Matches(Xml, '<col\s[^/]*/>', [roIgnoreCase]);
  for var ColMatch in ColMatches do
  begin
    const ColXml = ColMatch.Value;
    const WidthMatch = TRegEx.Match(ColXml, 'width="([^"]*)"', [roIgnoreCase]);
    const CustomMatch = TRegEx.Match(ColXml, 'customWidth="1"', [roIgnoreCase]);
    if WidthMatch.Success and CustomMatch.Success then
    begin
      const Width = StrToFloatDef(WidthMatch.Groups[1].Value, 0, TFormatSettings.Invariant);
      if Width > 0 then
      begin
        const MinMatch = TRegEx.Match(ColXml, 'min="(\d+)"', [roIgnoreCase]);
        const MaxMatch = TRegEx.Match(ColXml, 'max="(\d+)"', [roIgnoreCase]);
        const ColMin = StrToIntDef(MinMatch.Groups[1].Value, 0);
        const ColMax = StrToIntDef(MaxMatch.Groups[1].Value, 0);
        for var ColNum := ColMin to ColMax do
        begin
          const ColLetters = TExcelSheet.ColumnNumberToLetters(ColNum);
          Sheet.SetColumnWidth(ColLetters, Width);
        end;
      end;
    end;
  end;

  const RowMatches = TRegEx.Matches(Xml, '<row\s[^>]*>', [roIgnoreCase]);
  for var RowMatch in RowMatches do
  begin
    const RowXml = RowMatch.Value;
    const HtMatch = TRegEx.Match(RowXml, 'ht="([^"]*)"', [roIgnoreCase]);
    const CustomMatch = TRegEx.Match(RowXml, 'customHeight="1"', [roIgnoreCase]);
    if (HtMatch.Success) and (CustomMatch.Success) then
    begin
      const RowNumMatch = TRegEx.Match(RowXml, 'r="(\d+)"', [roIgnoreCase]);
      if RowNumMatch.Success then
      begin
        const RowNum = StrToIntDef(RowNumMatch.Groups[1].Value, 0);
        const RowHt = StrToFloatDef(HtMatch.Groups[1].Value, 0, TFormatSettings.Invariant);
        if (RowNum > 0) and (RowHt > 0) then
          Sheet.SetRowHeight(RowNum, RowHt);
      end;
    end;
  end;

  // Build a map of shared formula index -> (masterAddress, formulaString).
  // Excel writes shared formulas as <f t="shared" si="N" ref="...">formula</f> on the
  // master cell and <f t="shared" si="N"/> (self-closing, no text) on the dependent cells.
  // We need the master address to compute row/column offsets for relative-reference adjustment.
  var SharedFormulas := TDictionary<Integer, TPair<string, string>>.Create;
  try
    // Scan backwards in the XML: for each master cell (<c r="ADDR">...<f ... si="N">formula</f>...)
    // capture ADDR, N, and formula. We match the whole <c>...</c> block.
    const SharedMasterMatches = TRegEx.Matches(Xml,
      '<c\s+r="([A-Z]+\d+)"[^>]*>(?:(?!</c>).)*<f\s[^>]*\bsi="(\d+)"[^>]*>([^<]+)</f>',
      [roIgnoreCase, roSingleLine]);
    for var SfMatch in SharedMasterMatches do
      if SfMatch.Groups.Count > 3 then
      begin
        const SiIdx = StrToIntDef(SfMatch.Groups[2].Value, -1);
        if SiIdx >= 0 then
          SharedFormulas.AddOrSetValue(SiIdx,
            TPair<string, string>.Create(SfMatch.Groups[1].Value, SfMatch.Groups[3].Value));
      end;

    // Match each cell element. Self-closing <c r=".."/> empty cells are excluded via
    // (?<!/)> so they cannot steal <v> values from adjacent cells. (?:(?!</c>).)*
    // prevents crossing cell boundaries before reaching <v>.
    // The <f> element may have attributes (e.g. t="shared" si="0"), so we use
    // <f[^>]*> for the opening tag. Self-closing <f.../> cells have an empty formula
    // text; the si index is extracted separately for shared-formula dependent-cell resolution.
    // Group layout: 1=address, 2=style, 3=cell-type, 4=formula-text, 5=si-index, 6=value.
    const CellMatches = TRegEx.Matches(Xml,
      '<c\s+r="([A-Z]+\d+)"(?:\s+s="(\d+)")?(?:\s+t="([^"]*)")?[^>]*(?<!/)>' +
      '(?:<f(?:\s+[^>]*)?>([^<]*)</f>|<f(?:\s+[^>]*)?\bsi="(\d+)"[^/]*/>)?' +
      '(?:(?!</c>).)*<v>([^<]*)</v>.*?</c>',[roIgnoreCase, roSingleLine]);
    for var Match in CellMatches do
    begin
      if Match.Groups.Count > 6 then
      begin
        const Address = Match.Groups[1].Value;
        const StyleIdx = StrToIntDef(Match.Groups[2].Value, 0);
        const CellType = Match.Groups[3].Value;
        var   Formula: string := Match.Groups[4].Value;
        const SiIndex  = Match.Groups[5].Value;
        const Value    = Match.Groups[6].Value;

        // Dependent shared-formula cells have an empty formula text but carry a si index.
        // Resolve and adjust the formula string from the shared formula map.
        if (Formula = '') and (SiIndex <> '') then
        begin
          const SiIdx = StrToIntDef(SiIndex, -1);
          var MasterInfo: TPair<string, string>;
          if (SiIdx >= 0) and SharedFormulas.TryGetValue(SiIdx, MasterInfo) then
          begin
            // Compute row and column offsets from master cell to this dependent cell.
            const MasterAddr = MasterInfo.Key;
            const MasterFormula = MasterInfo.Value;
            var MasterCol := '';
            var MasterRowStr := '';
            for var Ch in MasterAddr do
              if CharInSet(Ch, ['A'..'Z', 'a'..'z']) then MasterCol := MasterCol + Ch
              else MasterRowStr := MasterRowStr + Ch;
            var DependentCol := '';
            var DependentRowStr := '';
            for var Ch in Address do
              if CharInSet(Ch, ['A'..'Z', 'a'..'z']) then DependentCol := DependentCol + Ch
              else DependentRowStr := DependentRowStr + Ch;
            const RowDelta = StrToIntDef(DependentRowStr, 0) - StrToIntDef(MasterRowStr, 0);
            const ColDelta = TExcelSheet.ColumnLetterToNumber(DependentCol) -
                             TExcelSheet.ColumnLetterToNumber(MasterCol);
            Formula := AdjustFormulaRefs(MasterFormula, RowDelta, ColDelta);
          end;
        end;

        var Cell: TExcelCell := nil;

        if Formula <> '' then
        begin
          Sheet.SetCellFormula(Address, Formula, Value);
          Cell := Sheet.GetCell(Address) as TExcelCell;
        end
        else if CellType = 's' then
        begin
          const StrIndex = StrToIntDef(Value, -1);
          if (StrIndex >= 0) and (StrIndex < FSharedStrings.Count) then
          begin
            Sheet.SetCellValue(Address, FSharedStrings[StrIndex], True);
            Cell := Sheet.GetCell(Address) as TExcelCell;
          end;
        end
        else if CellType = 'b' then
        begin
          Sheet.SetBooleanValue(Address, Value = '1');
          Cell := Sheet.GetCell(Address) as TExcelCell;
        end
        else
        begin
          if (StyleIdx > 0) and (StyleIdx < FStyleIsDate.Count) and FStyleIsDate[StyleIdx] then
            Sheet.SetDateTimeValue(Address, StrToFloatDef(Value, 0, TFormatSettings.Invariant))
          else
            Sheet.SetCellValue(Address, Value, False);
          Cell := Sheet.GetCell(Address) as TExcelCell;
        end;

        if (Cell <> nil) and (StyleIdx > 0) then
        begin
          if (StyleIdx < FStyleBold.Count) and (FStyleBold[StyleIdx]) then
            Cell.FBold := True;
          if (StyleIdx < FStyleItalic.Count) and (FStyleItalic[StyleIdx]) then
            Cell.FItalic := True;
          if (StyleIdx < FStyleUnderline.Count) and (FStyleUnderline[StyleIdx]) then
            Cell.FUnderline := True;
          if (StyleIdx < FStyleFontName.Count) and (FStyleFontName[StyleIdx] <> '') then
            Cell.FFontName := FStyleFontName[StyleIdx];
          if (StyleIdx < FStyleFontSize.Count) and (FStyleFontSize[StyleIdx] <> 0) then
            Cell.FFontSize := FStyleFontSize[StyleIdx];
          if (StyleIdx < FStyleColors.Count) and (FStyleColors[StyleIdx] <> 0) then
            Cell.FBackgroundColor := FStyleColors[StyleIdx];
          if (StyleIdx < FStyleFontColor.Count) and (FStyleFontColor[StyleIdx] <> 0) then
            Cell.FFontColor := FStyleFontColor[StyleIdx];
          if (StyleIdx < FStyleStrikeout.Count) and (FStyleStrikeout[StyleIdx]) then
            Cell.FStrikeout := True;
          if (StyleIdx < FStyleBorderTopStyle.Count) and (FStyleBorderTopStyle[StyleIdx] <> TExcelBorderStyle.None) then
            Cell.FBorderStyle[TExcelBorderSide.Top] := FStyleBorderTopStyle[StyleIdx];
          if (StyleIdx < FStyleBorderTopColor.Count) and (FStyleBorderTopColor[StyleIdx] <> 0) then
            Cell.FBorderColor[TExcelBorderSide.Top] := FStyleBorderTopColor[StyleIdx];
          if (StyleIdx < FStyleBorderRightStyle.Count) and (FStyleBorderRightStyle[StyleIdx] <> TExcelBorderStyle.None) then
            Cell.FBorderStyle[TExcelBorderSide.Right] := FStyleBorderRightStyle[StyleIdx];
          if (StyleIdx < FStyleBorderRightColor.Count) and (FStyleBorderRightColor[StyleIdx] <> 0) then
            Cell.FBorderColor[TExcelBorderSide.Right] := FStyleBorderRightColor[StyleIdx];
          if (StyleIdx < FStyleBorderBottomStyle.Count) and (FStyleBorderBottomStyle[StyleIdx] <> TExcelBorderStyle.None) then
            Cell.FBorderStyle[TExcelBorderSide.Bottom] := FStyleBorderBottomStyle[StyleIdx];
          if (StyleIdx < FStyleBorderBottomColor.Count) and (FStyleBorderBottomColor[StyleIdx] <> 0) then
            Cell.FBorderColor[TExcelBorderSide.Bottom] := FStyleBorderBottomColor[StyleIdx];
          if (StyleIdx < FStyleBorderLeftStyle.Count) and (FStyleBorderLeftStyle[StyleIdx] <> TExcelBorderStyle.None) then
            Cell.FBorderStyle[TExcelBorderSide.Left] := FStyleBorderLeftStyle[StyleIdx];
          if (StyleIdx < FStyleBorderLeftColor.Count) and (FStyleBorderLeftColor[StyleIdx] <> 0) then
            Cell.FBorderColor[TExcelBorderSide.Left] := FStyleBorderLeftColor[StyleIdx];
          if (StyleIdx < FStyleHAlign.Count) and (FStyleHAlign[StyleIdx] <> TExcelHAlign.None) then
            Cell.FHAlign := FStyleHAlign[StyleIdx];
          if (StyleIdx < FStyleVAlign.Count) and (FStyleVAlign[StyleIdx] <> TExcelVAlign.None) then
            Cell.FVAlign := FStyleVAlign[StyleIdx];
          if (StyleIdx < FStyleWrapText.Count) and (FStyleWrapText[StyleIdx]) then
            Cell.FWrapText := True;
          if (StyleIdx < FStyleNumberFormat.Count) and (FStyleNumberFormat[StyleIdx] <> '') then
            Cell.FNumberFormat := FStyleNumberFormat[StyleIdx];
        end;
      end;
    end;
  finally
    SharedFormulas.Free;
  end;

  // Apply styles to self-closing empty cells (<c r="X" s="N"/>).
  // These have no <v> element so the main cell loop skips them, but they may
  // carry a meaningful style (e.g. background colour) that must be preserved.
  const EmptyCellMatches = TRegEx.Matches(Xml,
    '<c\s+r="([A-Z]+\d+)"\s+s="(\d+)"[^>]*/>', [roIgnoreCase]);
  for var Match in EmptyCellMatches do
    if Match.Groups.Count > 2 then
    begin
      const Address  = Match.Groups[1].Value;
      const StyleIdx = StrToIntDef(Match.Groups[2].Value, 0);
      if StyleIdx > 0 then
      begin
        var Cell := Sheet.GetCell(Address) as TExcelCell;
        if (StyleIdx < FStyleBold.Count) and (FStyleBold[StyleIdx]) then
          Cell.FBold := True;
        if (StyleIdx < FStyleItalic.Count) and (FStyleItalic[StyleIdx]) then
          Cell.FItalic := True;
        if (StyleIdx < FStyleUnderline.Count) and (FStyleUnderline[StyleIdx]) then
          Cell.FUnderline := True;
        if (StyleIdx < FStyleFontName.Count) and (FStyleFontName[StyleIdx] <> '') then
          Cell.FFontName := FStyleFontName[StyleIdx];
        if (StyleIdx < FStyleFontSize.Count) and (FStyleFontSize[StyleIdx] <> 0) then
          Cell.FFontSize := FStyleFontSize[StyleIdx];
        if (StyleIdx < FStyleColors.Count) and (FStyleColors[StyleIdx] <> 0) then
          Cell.FBackgroundColor := FStyleColors[StyleIdx];
        if (StyleIdx < FStyleFontColor.Count) and (FStyleFontColor[StyleIdx] <> 0) then
          Cell.FFontColor := FStyleFontColor[StyleIdx];
        if (StyleIdx < FStyleStrikeout.Count) and (FStyleStrikeout[StyleIdx]) then
          Cell.FStrikeout := True;
        if (StyleIdx < FStyleBorderTopStyle.Count) and (FStyleBorderTopStyle[StyleIdx] <> TExcelBorderStyle.None) then
          Cell.FBorderStyle[TExcelBorderSide.Top] := FStyleBorderTopStyle[StyleIdx];
        if (StyleIdx < FStyleBorderTopColor.Count) and (FStyleBorderTopColor[StyleIdx] <> 0) then
          Cell.FBorderColor[TExcelBorderSide.Top] := FStyleBorderTopColor[StyleIdx];
        if (StyleIdx < FStyleBorderRightStyle.Count) and (FStyleBorderRightStyle[StyleIdx] <> TExcelBorderStyle.None) then
          Cell.FBorderStyle[TExcelBorderSide.Right] := FStyleBorderRightStyle[StyleIdx];
        if (StyleIdx < FStyleBorderRightColor.Count) and (FStyleBorderRightColor[StyleIdx] <> 0) then
          Cell.FBorderColor[TExcelBorderSide.Right] := FStyleBorderRightColor[StyleIdx];
        if (StyleIdx < FStyleBorderBottomStyle.Count) and (FStyleBorderBottomStyle[StyleIdx] <> TExcelBorderStyle.None) then
          Cell.FBorderStyle[TExcelBorderSide.Bottom] := FStyleBorderBottomStyle[StyleIdx];
        if (StyleIdx < FStyleBorderBottomColor.Count) and (FStyleBorderBottomColor[StyleIdx] <> 0) then
          Cell.FBorderColor[TExcelBorderSide.Bottom] := FStyleBorderBottomColor[StyleIdx];
        if (StyleIdx < FStyleBorderLeftStyle.Count) and (FStyleBorderLeftStyle[StyleIdx] <> TExcelBorderStyle.None) then
          Cell.FBorderStyle[TExcelBorderSide.Left] := FStyleBorderLeftStyle[StyleIdx];
        if (StyleIdx < FStyleBorderLeftColor.Count) and (FStyleBorderLeftColor[StyleIdx] <> 0) then
          Cell.FBorderColor[TExcelBorderSide.Left] := FStyleBorderLeftColor[StyleIdx];
        if (StyleIdx < FStyleHAlign.Count) and (FStyleHAlign[StyleIdx] <> TExcelHAlign.None) then
          Cell.FHAlign := FStyleHAlign[StyleIdx];
        if (StyleIdx < FStyleVAlign.Count) and (FStyleVAlign[StyleIdx] <> TExcelVAlign.None) then
          Cell.FVAlign := FStyleVAlign[StyleIdx];
        if (StyleIdx < FStyleWrapText.Count) and (FStyleWrapText[StyleIdx]) then
          Cell.FWrapText := True;
        if (StyleIdx < FStyleNumberFormat.Count) and (FStyleNumberFormat[StyleIdx] <> '') then
          Cell.FNumberFormat := FStyleNumberFormat[StyleIdx];
      end;
    end;

  const MergeMatches = TRegEx.Matches(Xml, '<mergeCell\s+ref="([^"]+)"', [roIgnoreCase]);
  for var MergeMatch in MergeMatches do
    if MergeMatch.Groups.Count > 1 then
      Sheet.MergeCells(MergeMatch.Groups[1].Value);
end;

function TExcelWorkbook.GetSheetCount: Integer;
begin
  Result := FSheets.Count;
end;

function TExcelWorkbook.GetSheet(Index: Integer): IExcelSheet;
begin
  Result := FSheets[Index];
end;

function TExcelWorkbook.GetSheetByName(const Name: string): IExcelSheet;
begin
  Result := SheetByName(Name);
end;

function TExcelWorkbook.GetMetadata: TDocumentMetadata;
begin
  Result := FMetadata;
end;

function TExcelWorkbook.AddSheet(const Name: string): IExcelSheet;
begin
  var Sheet := TExcelSheet.Create(Name);
  Sheet.FOwner := Self;
  Result := Sheet;
  FSheets.Add(Result);
end;

function TExcelWorkbook.SheetByName(const Name: string): IExcelSheet;
begin
  for var Sheet in FSheets do
    if SameText(Sheet.Name, Name) then
      Exit(Sheet);
  Result := nil;
end;

procedure TExcelWorkbook.RemoveSheet(Index: Integer);
begin
  if (Index < 0) or (Index >= FSheets.Count) then
    raise EExcelWorkbookException.CreateFmt('Sheet index out of range: %d', [Index]);
  // Excel refuses to open a workbook with no sheets, so keep at least one.
  if FSheets.Count <= 1 then
    raise EExcelWorkbookException.Create('A workbook must retain at least one sheet');
  // sheetId/r:id and the sheetN.xml part names are derived positionally from the
  // sheet order on save, so removing an entry is all that is needed.
  FSheets.Delete(Index);
end;

procedure TExcelWorkbook.RemoveSheetByName(const Name: string);
begin
  for var I := 0 to FSheets.Count - 1 do
    if SameText(FSheets[I].Name, Name) then
    begin
      RemoveSheet(I);
      Exit;
    end;
  raise EExcelWorkbookException.CreateFmt('Sheet not found: "%s"', [Name]);
end;

procedure TExcelWorkbook.ApplyFormulaDelete(const ATargetSheet: TExcelSheet; const AIsColumn: Boolean; const ALineIndex: Integer);
begin
  for var SheetIntf in FSheets do
  begin
    var Sheet := SheetIntf as TExcelSheet;
    for var Pair in Sheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      if Cell.GetHasFormula then
      begin
        const NewFormula = TransformFormulaForLineDelete(Cell.GetFormula, ATargetSheet.GetName, Sheet.GetName, AIsColumn, ALineIndex);
        if NewFormula <> Cell.GetFormula then
          Cell.FFormula := NewFormula;
      end;
    end;
  end;
end;

end.
