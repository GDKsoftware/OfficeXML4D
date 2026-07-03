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
    FBorderStyle: TExcelBorderStyle;
    FBorderColor: Cardinal;
    FHAlign: TExcelHAlign;
    FVAlign: TExcelVAlign;
    FWrapText: Boolean;
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
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Double;
    procedure SetFontSize(const Value: Double);
    function GetBackgroundColor: Cardinal;
    procedure SetBackgroundColor(const Value: Cardinal);
    function GetNumberFormat: string;
    procedure SetNumberFormat(const Value: string);
    function GetBorderStyle: TExcelBorderStyle;
    procedure SetBorderStyle(const Value: TExcelBorderStyle);
    function GetBorderColor: Cardinal;
    procedure SetBorderColor(const Value: Cardinal);
    function GetHAlign: TExcelHAlign;
    procedure SetHAlign(const Value: TExcelHAlign);
    function GetVAlign: TExcelVAlign;
    procedure SetVAlign(const Value: TExcelVAlign);
    function GetWrapText: Boolean;
    procedure SetWrapText(const Value: Boolean);

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

    class function ColumnLetterToNumber(const Column: string): Integer; static;
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

    procedure SetCellValue(const Address, Value: string; IsString: Boolean);
    procedure SetBooleanValue(const Address: string; Value: Boolean);
    procedure SetCellFormula(const Address, Formula, Value: string);

    function GetCells: TDictionary<string, IExcelCell>;

    property Cells: TDictionary<string, IExcelCell> read FCells;
    property ColumnWidths: TDictionary<string, Double> read FColumnWidths;
    property RowHeights: TDictionary<Integer, Double> read FRowHeights;
    property MergedRangesList: TList<string> read FMergedRanges;
  end;

  TExcelWorkbook = class(TInterfacedObject, IExcelWorkbook)
  private
    FSheets: TList<IExcelSheet>;
    FMetadata: TDocumentMetadata;
    FSharedStrings: TList<string>;
    FStyleBold: TList<Boolean>;
    FStyleItalic: TList<Boolean>;
    FStyleUnderline: TList<Boolean>;
    FStyleFontName: TList<string>;
    FStyleFontSize: TList<Double>;
    FStyleColors: TList<Cardinal>;
    FStyleBorderStyle: TList<TExcelBorderStyle>;
    FStyleBorderColor: TList<Cardinal>;
    FStyleHAlign: TList<TExcelHAlign>;
    FStyleVAlign: TList<TExcelVAlign>;
    FStyleWrapText: TList<Boolean>;

    procedure ParseWorkbook(const Xml: string);
    procedure ParseSharedStrings(const Xml: string);
    procedure ParseStyles(const Xml: string);
    procedure ParseSheet(const Sheet: TExcelSheet; const Xml: string);

    function BuildSharedStrings: TList<string>;
    function GetSharedStringIndex(const Strings: TList<string>; const Value: string): Integer;
    function EscapeXml(const Text: string): string;
    function AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;
    function GenerateContentTypes: string;
    function GenerateRels: string;
    function GenerateWorkbook: string;
    function GenerateWorkbookRels: string;
    function GenerateSheet(const Sheet: TExcelSheet; const SharedStrings: TList<string>; const StyleMap: TDictionary<string, Integer>): string;
    function GenerateSharedStrings(const Strings: TList<string>): string;
    function GenerateStyles(const StyleMap: TDictionary<string, Integer>): string;
    function BuildStyleMap: TDictionary<string, Integer>;
    function GetStyleKey(const Cell: TExcelCell): string;
    function GetCellStyleIndex(const Cell: TExcelCell; const StyleMap: TDictionary<string, Integer>): Integer;
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
  end;

implementation

uses
  System.Math,
  System.StrUtils,
  Office4D.Errors,
  Office4D.Relationships,
  Office4D.Types;

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

  ExcelDateOffset = 25569;
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
    Result := FFloatValue - ExcelDateOffset
  else
    Result := 0;
end;

procedure TExcelCell.SetAsDateTime(const Value: TDateTime);
begin
  FDateTimeValue := Value;
  FFloatValue := Value + ExcelDateOffset;
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

function TExcelCell.GetFontName: string;
begin
  Result := FFontName;
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

function TExcelCell.GetBorderStyle: TExcelBorderStyle;
begin
  Result := FBorderStyle;
end;

procedure TExcelCell.SetBorderStyle(const Value: TExcelBorderStyle);
begin
  FBorderStyle := Value;
end;

function TExcelCell.GetBorderColor: Cardinal;
begin
  Result := FBorderColor;
end;

procedure TExcelCell.SetBorderColor(const Value: Cardinal);
begin
  FBorderColor := Value;
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
  const HasFont = (FBold) or (FItalic) or (FUnderline) or (FFontName <> '') or (FFontSize <> 0);
  const HasFill = (FBackgroundColor <> 0);
  const HasFormat = (FCellType = TCellType.DateTime);
  const HasBorder = (FBorderStyle <> TExcelBorderStyle.None);
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
end;

destructor TExcelSheet.Destroy;
begin
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

procedure TExcelSheet.SetCellFormula(const Address, Formula, Value: string);
begin
  var Cell := GetCell(Address) as TExcelCell;
  Cell.FFormula := Formula;
  Cell.SetAsFloat(StrToFloatDef(Value, 0, TFormatSettings.Invariant));
  Cell.FStringValue := Value;
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
  FStyleBorderStyle := TList<TExcelBorderStyle>.Create;
  FStyleBorderColor := TList<Cardinal>.Create;
  FStyleHAlign := TList<TExcelHAlign>.Create;
  FStyleVAlign := TList<TExcelVAlign>.Create;
  FStyleWrapText := TList<Boolean>.Create;
end;

destructor TExcelWorkbook.Destroy;
begin
  FStyleWrapText.Free;
  FStyleVAlign.Free;
  FStyleHAlign.Free;
  FStyleBorderColor.Free;
  FStyleBorderStyle.Free;
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

      for var I := 0 to FSheets.Count - 1 do
      begin
        const SheetXml = TEncoding.UTF8.GetBytes(GenerateSheet(FSheets[I] as TExcelSheet, SharedStrings, StyleMap));
        Zip.Add(SheetXml, PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix);
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

      for var I := 0 to FSheets.Count - 1 do
      begin
        const SheetXml = TEncoding.UTF8.GetBytes(GenerateSheet(FSheets[I] as TExcelSheet, SharedStrings, StyleMap));
        Zip.Add(SheetXml, PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix);
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

function TExcelWorkbook.EscapeXml(const Text: string): string;
begin
  Result := Text;
  Result := StringReplace(Result, '&', '&amp;', [rfReplaceAll]);
  Result := StringReplace(Result, '<', '&lt;', [rfReplaceAll]);
  Result := StringReplace(Result, '>', '&gt;', [rfReplaceAll]);
  Result := StringReplace(Result, '"', '&quot;', [rfReplaceAll]);
  Result := StringReplace(Result, '''', '&apos;', [rfReplaceAll]);
end;

function TExcelWorkbook.GenerateContentTypes: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Types xmlns="' + ContentTypesNs + '">');
    SB.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    SB.Append('<Default Extension="xml" ContentType="application/xml"/>');
    SB.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    for var I := 0 to FSheets.Count - 1 do
      SB.Append('<Override PartName="/xl/worksheets/sheet' + IntToStr(I + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
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
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<workbook xmlns="' + SpreadsheetNs + '" xmlns:r="' + OfficeDocRelsNs + '">');
    SB.Append('<sheets>');
    for var I := 0 to FSheets.Count - 1 do
      SB.Append('<sheet name="' + FSheets[I].Name + '" sheetId="' + IntToStr(I + 1) + '" r:id="rId' + IntToStr(I + 1) + '"/>');
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
    SB.Append('<worksheet xmlns="' + SpreadsheetNs + '">');

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

    SB.Append('</worksheet>');
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
      SB.Append('<si><t>' + StringItem + '</t></si>');
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
  Result := Format('%d|%d|%d|%s|%d|%d|%s|%s|%d|%d|%d|%d|%d', [
    Ord(Cell.FBold),
    Cell.FBackgroundColor,
    DateFlag,
    Cell.FNumberFormat,
    Ord(Cell.FItalic),
    Ord(Cell.FUnderline),
    Cell.FFontName,
    FontSizeStr,
    Ord(Cell.FBorderStyle),
    Cell.FBorderColor,
    Ord(Cell.FHAlign),
    Ord(Cell.FVAlign),
    Ord(Cell.FWrapText)
  ]);
end;

function TExcelWorkbook.BuildStyleMap: TDictionary<string, Integer>;
begin
  Result := TDictionary<string, Integer>.Create;
  Result.Add('0|0|0||0|0|||0|0|0|0|0', 0);
  Result.Add('0|0|1||0|0|||0|0|0|0|0', 1);

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
      const FontKey = Format('%d|%d|%d|%s|%s', [
        Ord(Cell.FBold), Ord(Cell.FItalic), Ord(Cell.FUnderline), Cell.FFontName, FontSizeStr]);
      if (FontKey <> '0|0|0||') and (not FontKeys.Contains(FontKey)) then
        FontKeys.Add(FontKey);

      const BorderKey = Format('%d|%d', [Ord(Cell.FBorderStyle), Cell.FBorderColor]);
      if (BorderKey <> '0|0') and (not BorderKeys.Contains(BorderKey)) then
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
        SB.Append('<numFmt numFmtId="' + IntToStr(165 + I) + '" formatCode="' + EscapeXml(NumFormats[I]) + '"/>');
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
      SB.Append('<font>');
      if IsBold then SB.Append('<b/>');
      if IsItalic then SB.Append('<i/>');
      if IsUnderline then SB.Append('<u/>');
      if Size <> '' then
        SB.Append('<sz val="' + Size + '"/>')
      else
        SB.Append('<sz val="11"/>');
      if Name <> '' then
        SB.Append('<name val="' + EscapeXml(Name) + '"/>')
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
      const BorderStyle = TExcelBorderStyle(StrToIntDef(BorderParts[0], 0));
      const BorderColor = StrToIntDef(BorderParts[1], 0);
      const StyleStr = BorderStyleToString(BorderStyle);
      var ColorAttr := '';
      if BorderColor <> 0 then
        ColorAttr := '<color rgb="FF' + IntToHex(BorderColor, 6) + '"/>'
      else
        ColorAttr := '<color auto="1"/>';
      SB.Append('<border>');
      SB.Append('<left style="' + StyleStr + '">' + ColorAttr + '</left>');
      SB.Append('<right style="' + StyleStr + '">' + ColorAttr + '</right>');
      SB.Append('<top style="' + StyleStr + '">' + ColorAttr + '</top>');
      SB.Append('<bottom style="' + StyleStr + '">' + ColorAttr + '</bottom>');
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
        const CellBorderStyle = TExcelBorderStyle(StrToIntDef(Parts[8], 0));
        const CellBorderColor = StrToIntDef(Parts[9], 0);
        const CellHAlign = TExcelHAlign(StrToIntDef(Parts[10], 0));
        const CellVAlign = TExcelVAlign(StrToIntDef(Parts[11], 0));
        const CellWrapText = Parts[12] = '1';

        const FontKey = Format('%d|%d|%d|%s|%s', [
          Ord(IsBold), Ord(IsItalic), Ord(IsUnderline), CellFontName, CellFontSize]);
        var FontId := 0;
        if FontKey <> '0|0|0||' then
          FontId := 1 + FontKeys.IndexOf(FontKey);

        var FillId := 0;
        if BgColor <> 0 then
          FillId := 2 + Colors.IndexOf(BgColor);

        const BorderKey = Format('%d|%d', [Ord(CellBorderStyle), CellBorderColor]);
        var BorderId := 0;
        if BorderKey <> '0|0' then
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
  const Matches = TRegEx.Matches(Xml, '<sheet\s+name="([^"]+)"[^/>]*(?:/>|>)', [roIgnoreCase]);
  for var Match in Matches do
    if Match.Groups.Count > 1 then
      FSheets.Add(TExcelSheet.Create(Match.Groups[1].Value));
end;

procedure TExcelWorkbook.ParseSharedStrings(const Xml: string);
begin
  FSharedStrings.Clear;
  const Matches = TRegEx.Matches(Xml, '<si>.*?<t[^>]*>([^<]*)</t>.*?</si>', [roIgnoreCase, roSingleLine]);
  for var Match in Matches do
    if Match.Groups.Count > 1 then
      FSharedStrings.Add(Match.Groups[1].Value);
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

begin
  FStyleBold.Clear;
  FStyleItalic.Clear;
  FStyleUnderline.Clear;
  FStyleFontName.Clear;
  FStyleFontSize.Clear;
  FStyleColors.Clear;
  FStyleBorderStyle.Clear;
  FStyleBorderColor.Clear;
  FStyleHAlign.Clear;
  FStyleVAlign.Clear;
  FStyleWrapText.Clear;

  var FontsBold := TList<Boolean>.Create;
  var FontsItalic := TList<Boolean>.Create;
  var FontsUnderline := TList<Boolean>.Create;
  var FontsName := TList<string>.Create;
  var FontsSize := TList<Double>.Create;
  var Fills := TList<Cardinal>.Create;
  var BorderStyles := TList<TExcelBorderStyle>.Create;
  var BorderColors := TList<Cardinal>.Create;
  try
    const FontMatches = TRegEx.Matches(Xml, '<font>(.*?)</font>', [roIgnoreCase, roSingleLine]);
    for var Match in FontMatches do
    begin
      const FontXml = Match.Groups[1].Value;
      FontsBold.Add(Pos('<b/>', FontXml) > 0);
      FontsItalic.Add(Pos('<i/>', FontXml) > 0);
      FontsUnderline.Add(Pos('<u/>', FontXml) > 0);

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
      BorderStyles.Add(TExcelBorderStyle.None);
      BorderColors.Add(0);
    end;
    const BorderFullMatches = TRegEx.Matches(Xml, '<border>(.*?)</border>', [roIgnoreCase, roSingleLine]);
    for var Match in BorderFullMatches do
    begin
      const BorderXml = Match.Groups[1].Value;
      const LeftMatch = TRegEx.Match(BorderXml, '<left\s+style="([^"]*)"', [roIgnoreCase]);
      if LeftMatch.Success then
        BorderStyles.Add(StringToBorderStyle(LeftMatch.Groups[1].Value))
      else
        BorderStyles.Add(TExcelBorderStyle.None);

      const BColorMatch = TRegEx.Match(BorderXml, '<left[^>]*>.*?<color\s+rgb="FF([0-9A-Fa-f]{6})"', [roIgnoreCase, roSingleLine]);
      if BColorMatch.Success then
        BorderColors.Add(StrToInt64Def('$' + BColorMatch.Groups[1].Value, 0))
      else
        BorderColors.Add(0);
    end;

    const CellXfsMatch = TRegEx.Match(Xml, '<cellXfs[^>]*>(.*?)</cellXfs>', [roIgnoreCase, roSingleLine]);
    if CellXfsMatch.Success then
    begin
      const CellXfsXml = CellXfsMatch.Groups[1].Value;
      const XfMatches = TRegEx.Matches(CellXfsXml, '<xf\s[^/>]*(?:/>|(?:[^>]*>.*?</xf>))', [roIgnoreCase, roSingleLine]);
      for var Match in XfMatches do
      begin
        const XfXml = Match.Value;

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

        if BorderId < BorderStyles.Count then
          FStyleBorderStyle.Add(BorderStyles[BorderId])
        else
          FStyleBorderStyle.Add(TExcelBorderStyle.None);

        if BorderId < BorderColors.Count then
          FStyleBorderColor.Add(BorderColors[BorderId])
        else
          FStyleBorderColor.Add(0);

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
    BorderColors.Free;
    BorderStyles.Free;
    Fills.Free;
    FontsSize.Free;
    FontsName.Free;
    FontsUnderline.Free;
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
      // Convert column number back to letters
      var ColLetters := '';
      while ColNum > 0 do
      begin
        const Remainder = (ColNum - 1) mod 26;
        ColLetters := Chr(Ord('A') + Remainder) + ColLetters;
        ColNum := (ColNum - 1) div 26;
      end;
      NewColPart := ColLetters;
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

procedure TExcelWorkbook.ParseSheet(const Sheet: TExcelSheet; const Xml: string);
begin
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
          var ColLetters := '';
          var N := ColNum;
          while N > 0 do
          begin
            ColLetters := Chr(Ord('A') + (N - 1) mod 26) + ColLetters;
            N := (N - 1) div 26;
          end;
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
          if (StyleIdx < FStyleBorderStyle.Count) and (FStyleBorderStyle[StyleIdx] <> TExcelBorderStyle.None) then
            Cell.FBorderStyle := FStyleBorderStyle[StyleIdx];
          if (StyleIdx < FStyleBorderColor.Count) and (FStyleBorderColor[StyleIdx] <> 0) then
            Cell.FBorderColor := FStyleBorderColor[StyleIdx];
          if (StyleIdx < FStyleHAlign.Count) and (FStyleHAlign[StyleIdx] <> TExcelHAlign.None) then
            Cell.FHAlign := FStyleHAlign[StyleIdx];
          if (StyleIdx < FStyleVAlign.Count) and (FStyleVAlign[StyleIdx] <> TExcelVAlign.None) then
            Cell.FVAlign := FStyleVAlign[StyleIdx];
          if (StyleIdx < FStyleWrapText.Count) and (FStyleWrapText[StyleIdx]) then
            Cell.FWrapText := True;
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
        if (StyleIdx < FStyleBorderStyle.Count) and (FStyleBorderStyle[StyleIdx] <> TExcelBorderStyle.None) then
          Cell.FBorderStyle := FStyleBorderStyle[StyleIdx];
        if (StyleIdx < FStyleBorderColor.Count) and (FStyleBorderColor[StyleIdx] <> 0) then
          Cell.FBorderColor := FStyleBorderColor[StyleIdx];
        if (StyleIdx < FStyleHAlign.Count) and (FStyleHAlign[StyleIdx] <> TExcelHAlign.None) then
          Cell.FHAlign := FStyleHAlign[StyleIdx];
        if (StyleIdx < FStyleVAlign.Count) and (FStyleVAlign[StyleIdx] <> TExcelVAlign.None) then
          Cell.FVAlign := FStyleVAlign[StyleIdx];
        if (StyleIdx < FStyleWrapText.Count) and (FStyleWrapText[StyleIdx]) then
          Cell.FWrapText := True;
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
  Result := TExcelSheet.Create(Name);
  FSheets.Add(Result);
end;

function TExcelWorkbook.SheetByName(const Name: string): IExcelSheet;
begin
  for var Sheet in FSheets do
    if SameText(Sheet.Name, Name) then
      Exit(Sheet);
  Result := nil;
end;

end.
