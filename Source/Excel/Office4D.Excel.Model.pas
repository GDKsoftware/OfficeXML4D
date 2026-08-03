unit Office4D.Excel.Model;

interface

uses
  System.SysUtils,
  System.Generics.Collections,
  Office4D.Excel,
  Office4D.Metadata;

type
  {$SCOPEDENUMS ON}
  TCellType = (Empty, StringValue, Number, Boolean, DateTime);
  {$SCOPEDENUMS OFF}

  TExcelWorkbookContent = class;

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
    FOwner: TExcelWorkbookContent;

    procedure ClearLine(const AIsColumn: Boolean; const AIndex: Integer);
    procedure DeleteLine(const AIsColumn: Boolean; const AIndex: Integer);
    procedure RewriteOwnFormulas(const AIsColumn: Boolean; const AIndex: Integer);
  public
    /// Address arithmetic, shared with whoever reads or writes cell references.
    class function ColumnLetterToNumber(const Column: string): Integer; static;
    class function ColumnNumberToLetters(const Column: Integer): string; static;
    class procedure ParseCellAddress(const Address: string; out Col, Row: Integer); static;

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

    /// Not part of IExcelSheet: only a reader restores a frozen state directly.
    procedure SetFrozenRows(const Value: Integer);
    procedure SetFrozenColumns(const Value: Integer);

    property Cells: TDictionary<string, IExcelCell> read FCells;
    property ColumnWidths: TDictionary<string, Double> read FColumnWidths;
    property RowHeights: TDictionary<Integer, Double> read FRowHeights;
    property MergedRangesList: TList<string> read FMergedRanges;
    property Visibility: TExcelSheetVisibility read FVisibility write FVisibility;
    property FrozenRows: Integer read GetFrozenRows;
    property FrozenColumns: Integer read GetFrozenColumns;
    property Notes: TDictionary<string, string> read FNotes;
  end;

  /// <summary>
  /// Everything a workbook holds: what a reader fills and a writer emits. It
  /// also owns the sheets, which is why a sheet deleting a row can ask it to
  /// rewrite formulas on every other sheet.
  /// </summary>
  TExcelWorkbookContent = class
  private
    FSheets: TList<IExcelSheet>;
    FMetadata: TDocumentMetadata;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;

    /// Creates a sheet owned by this content and appends it.
    function CreateSheet(const Name: string): TExcelSheet;
    function AddSheet(const Name: string): IExcelSheet;
    function SheetByName(const Name: string): IExcelSheet;
    procedure RemoveSheet(Index: Integer);
    procedure RemoveSheetByName(const Name: string);

    procedure ApplyFormulaDelete(const ATargetSheet: TExcelSheet; const AIsColumn: Boolean;
      const ALineIndex: Integer);

    property Sheets: TList<IExcelSheet> read FSheets;
    property Metadata: TDocumentMetadata read FMetadata write FMetadata;
  end;

implementation

uses
  System.Math,
  System.StrUtils,
  System.RegularExpressions,
  Office4D.Errors;

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

procedure TExcelSheet.SetFrozenRows(const Value: Integer);
begin
  FFrozenRows := Value;
end;

procedure TExcelSheet.SetFrozenColumns(const Value: Integer);
begin
  FFrozenColumns := Value;
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


{ TExcelWorkbookContent }

constructor TExcelWorkbookContent.Create;
begin
  inherited Create;
  FSheets := TList<IExcelSheet>.Create;
  FMetadata.Clear;
end;

destructor TExcelWorkbookContent.Destroy;
begin
  FSheets.Free;
  inherited;
end;

procedure TExcelWorkbookContent.Clear;
begin
  FSheets.Clear;
  FMetadata.Clear;
end;

function TExcelWorkbookContent.CreateSheet(const Name: string): TExcelSheet;
begin
  Result := TExcelSheet.Create(Name);
  Result.FOwner := Self;
  FSheets.Add(Result);
end;

function TExcelWorkbookContent.AddSheet(const Name: string): IExcelSheet;
begin
  Result := CreateSheet(Name);
end;

function TExcelWorkbookContent.SheetByName(const Name: string): IExcelSheet;
begin
  for var Sheet in FSheets do
    if SameText(Sheet.Name, Name) then
      Exit(Sheet);
  Result := nil;
end;

procedure TExcelWorkbookContent.RemoveSheet(Index: Integer);
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

procedure TExcelWorkbookContent.RemoveSheetByName(const Name: string);
begin
  for var I := 0 to FSheets.Count - 1 do
    if SameText(FSheets[I].Name, Name) then
    begin
      RemoveSheet(I);
      Exit;
    end;
  raise EExcelWorkbookException.CreateFmt('Sheet not found: "%s"', [Name]);
end;

procedure TExcelWorkbookContent.ApplyFormulaDelete(const ATargetSheet: TExcelSheet; const AIsColumn: Boolean;
  const ALineIndex: Integer);
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
