unit Office4D.Excel.Reader;

interface

uses
  System.Generics.Collections,
  Office4D.Excel,
  Office4D.Package,
  Office4D.Relationships,
  Office4D.Xml,
  Office4D.Excel.Model;

type
  /// <summary>
  /// Fills a TExcelWorkbookContent from an opened .xlsx package. The shared
  /// strings and the style tables it parses first are scratch state: cells refer
  /// to them by index, so they only have to survive until every sheet is read.
  /// Elements are located by name in the namespace the part binds to
  /// SpreadsheetML and attributes are looked up by name, so attribute order,
  /// quote style and the prefix a writer picks make no difference.
  /// </summary>
  TExcelWorkbookReader = class
  private
    type
      TSharedFormulas = TDictionary<Integer, TPair<string, string>>;

      /// <summary>
      /// What one c element of a sheet says about its cell.
      /// </summary>
      TCellXml = record
        Address: string;
        StyleIndex: Integer;
        CellType: string;
        Formula: string;
        SharedFormulaIndex: Integer;
        Value: string;
        HasValue: Boolean;
        InlineText: string;

        function HoldsContent: Boolean;
      end;

  private
    FContent: TExcelWorkbookContent;
    FSheetRelationshipIds: TList<string>;
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
    function SheetVisibilityOf(const State: string): TExcelSheetVisibility;
    procedure ParseSharedStrings(const Xml: string);
    function RichTextOf(const RichTextXml, Prefix: string): string;
    function InlineStringOf(const CellXml, Prefix: string): string;
    procedure ParseStyles(const Xml: string);
    function ElementsIn(const Xml, ContainerName, ElementName: string): TArray<TXmlElement>;
    function ChildAttribute(const Xml, ElementName, AttributeName: string): string;
    function IsFontFlagSet(const FontXml, ElementName: string): Boolean;
    function ColorOf(const Xml, ElementName: string): Cardinal;
    function RgbOf(const ArgbValue: string): Cardinal;
    function FillColorOf(const FillXml, Prefix: string; const IndexedPalette: TList<Cardinal>): Cardinal;
    procedure ParseBorderSide(const BorderXml, Prefix, SideName: string; out SideStyle: TExcelBorderStyle;
                              out SideColor: Cardinal);
    function StringToBorderStyle(const Value: string): TExcelBorderStyle;
    procedure ParseSheet(const Sheet: TExcelSheet; const Xml: string);
    procedure ParseFrozenPane(const Sheet: TExcelSheet; const Xml, Prefix: string);
    function SplitCountOf(const Split: string): Integer;
    procedure ParseColumnWidths(const Sheet: TExcelSheet; const Xml, Prefix: string);
    procedure ParseSheetData(const Sheet: TExcelSheet; const Xml, Prefix: string);
    procedure ReadRows(const Sheet: TExcelSheet; const SheetDataXml, Prefix: string; const Cells: TList<TCellXml>);
    procedure ApplyRowHeight(const Sheet: TExcelSheet; const Row: TXmlElement; const RowNumber: Integer);
    function ColumnNumberOf(const Address: string): Integer;
    function CellXmlOf(const Element: TXmlElement; const Address, Prefix: string): TCellXml;
    procedure CollectSharedFormulas(const Cells: TList<TCellXml>; const SharedFormulas: TSharedFormulas);
    procedure ApplyCell(const Sheet: TExcelSheet; const CellXml: TCellXml; const SharedFormulas: TSharedFormulas);
    function WriteCellValue(const Sheet: TExcelSheet; const CellXml: TCellXml;
                            const SharedFormulas: TSharedFormulas): TExcelCell;
    function ResolveSharedFormula(const CellXml: TCellXml; const SharedFormulas: TSharedFormulas): string;
    function AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;
    procedure ApplyCellStyle(const Cell: TExcelCell; const StyleIdx: Integer);
    procedure ParseMergedCells(const Sheet: TExcelSheet; const Xml, Prefix: string);
    procedure ParseComments(const Sheet: TExcelSheet; const Xml: string);

    procedure ReadMetadata(const Package: TOXMLPackage);
    procedure ReadSheets(const Package: TOXMLPackage);
    function SheetPartPath(const Rels: TRelationships; const SheetIndex: Integer): string;
    procedure ReadSheetComments(const Package: TOXMLPackage; const Sheet: TExcelSheet; const SheetPath: string);
  public
    constructor Create(const Content: TExcelWorkbookContent);
    destructor Destroy; override;

    procedure Read(const Package: TOXMLPackage);
  end;

implementation

uses
  System.SysUtils,
  System.Math,
  System.StrUtils,
  System.RegularExpressions,
  Office4D.Metadata,
  Office4D.Types;

const
  InlineStringCellType = 'inlineStr';

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

  PartSharedStrings = 'xl/sharedStrings.xml';
  PartStyles = 'xl/styles.xml';
  PartWorkbook = 'xl/workbook.xml';
  PartWorkbookRels = 'xl/_rels/workbook.xml.rels';
  PartSheetPrefix = 'xl/worksheets/sheet';
  PartSheetSuffix = '.xml';

  RelTypeComments = NsOfficeDocumentRelationships + '/comments';
  RelationshipIdName = 'id';

{ TExcelWorkbookReader.TCellXml }

function TExcelWorkbookReader.TCellXml.HoldsContent: Boolean;
begin
  // A cell holds content when it carries <v>, a formula without a cached
  // value, or an inline string, which keeps its text in <is> instead of <v>.
  Result := HasValue or (Formula <> '') or (CellType = InlineStringCellType);
end;

{ TExcelWorkbookReader }

procedure TExcelWorkbookReader.ParseWorkbook(const Xml: string);
begin
  FSheetRelationshipIds.Clear;

  const Prefix = TXml.PrefixOf(Xml, NsSpreadsheetML);
  const RelationshipIdAttribute = TXml.PrefixOf(Xml, NsOfficeDocumentRelationships) + RelationshipIdName;

  for var Element in TXml.FindElements(Xml, Prefix + 'sheet') do
  begin
    const Sheet = FContent.CreateSheet(Element.Attribute('name'));
    Sheet.Visibility := SheetVisibilityOf(Element.Attribute('state'));
    FSheetRelationshipIds.Add(Element.Attribute(RelationshipIdAttribute));
  end;
end;

function TExcelWorkbookReader.SheetVisibilityOf(const State: string): TExcelSheetVisibility;
begin
  if SameText(State, 'hidden') then
    Exit(TExcelSheetVisibility.Hidden);

  if SameText(State, 'veryHidden') then
    Exit(TExcelSheetVisibility.VeryHidden);

  Result := TExcelSheetVisibility.Visible;
end;

procedure TExcelWorkbookReader.ParseSharedStrings(const Xml: string);
begin
  FSharedStrings.Clear;

  // Every <si> occupies exactly one index slot, including empty entries (<si><t/></si>)
  // and rich-text entries (<si><r><t>..</t></r><r><t>..</t></r></si>).
  const Prefix = TXml.PrefixOf(Xml, NsSpreadsheetML);
  for var SharedString in TXml.FindElements(Xml, Prefix + 'si') do
  begin
    const Text = RichTextOf(SharedString.Inner, Prefix);
    FSharedStrings.Add(Text);
  end;
end;

// The text of a shared string <si> or an inline string <is>: plain or split over
// rich text runs. Phonetic guide blocks carry their own <t> elements that are not
// part of the text.
function TExcelWorkbookReader.RichTextOf(const RichTextXml, Prefix: string): string;
begin
  const PhoneticGuides = TXml.FindElements(RichTextXml, Prefix + 'rPh');
  const TextXml = TXml.RemoveElements(RichTextXml, PhoneticGuides);

  Result := '';
  for var TextElement in TXml.FindElements(TextXml, Prefix + 't') do
  begin
    const Text = TXml.Unescape(TextElement.Inner);
    Result := Result + Text;
  end;
end;

function TExcelWorkbookReader.InlineStringOf(const CellXml, Prefix: string): string;
begin
  var InlineString: TXmlElement;
  if TXml.TryFindElement(CellXml, Prefix + 'is', InlineString) then
    Result := RichTextOf(InlineString.Inner, Prefix)
  else
    Result := '';
end;

procedure TExcelWorkbookReader.ParseStyles(const Xml: string);

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
  // The codes use Excel's own literal forms (the "_)" width-padding token
  // and the quoted "$" currency literal); Excel substitutes the locale
  // currency symbol for "$" on display, which is a rendering concern only.
  function BuiltInFormatCode(const NumFmtId: Integer): string;
  begin
    case NumFmtId of
      1: Result := '0';
      2: Result := '0.00';
      3: Result := '#,##0';
      4: Result := '#,##0.00';
      5: Result := '"$"#,##0_);("$"#,##0)';
      6: Result := '"$"#,##0_);[Red]("$"#,##0)';
      7: Result := '"$"#,##0.00_);("$"#,##0.00)';
      8: Result := '"$"#,##0.00_);[Red]("$"#,##0.00)';
      9: Result := '0%';
      10: Result := '0.00%';
      11: Result := '0.00E+00';
      12: Result := '# ?/?';
      13: Result := '# ??/??';
      37: Result := '#,##0_);(#,##0)';
      38: Result := '#,##0_);[Red](#,##0)';
      39: Result := '#,##0.00_);(#,##0.00)';
      40: Result := '#,##0.00_);[Red](#,##0.00)';
      41: Result := '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)';
      42: Result := '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)';
      43: Result := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
      44: Result := '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
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
    const Prefix = TXml.PrefixOf(Xml, NsSpreadsheetML);

    for var NumFmt in ElementsIn(Xml, Prefix + 'numFmts', Prefix + 'numFmt') do
    begin
      const NumFmtId = StrToIntDef(NumFmt.Attribute('numFmtId'), -1);
      CustomNumFmts.AddOrSetValue(NumFmtId, NumFmt.Attribute('formatCode'));
    end;

    for var Font in ElementsIn(Xml, Prefix + 'fonts', Prefix + 'font') do
    begin
      const FontXml = Font.Inner;
      const IsBold      = IsFontFlagSet(FontXml, Prefix + 'b');
      const IsItalic    = IsFontFlagSet(FontXml, Prefix + 'i');
      const IsUnderline = IsFontFlagSet(FontXml, Prefix + 'u');
      const IsStrikeout = IsFontFlagSet(FontXml, Prefix + 'strike');
      const FontColor   = ColorOf(FontXml, Prefix + 'color');
      const FontName    = ChildAttribute(FontXml, Prefix + 'name', 'val');
      const FontSize    = ChildAttribute(FontXml, Prefix + 'sz', 'val');

      FontsBold.Add(IsBold);
      FontsItalic.Add(IsItalic);
      FontsUnderline.Add(IsUnderline);
      FontsStrikeout.Add(IsStrikeout);
      FontsColor.Add(FontColor);
      FontsName.Add(FontName);
      FontsSize.Add(StrToFloatDef(FontSize, 0, TFormatSettings.Invariant));
    end;

    // Build the indexed colour palette. Start with the OOXML default palette.
    // If the styles XML contains a custom <indexedColors> block it replaces
    // the defaults entirely, so we rebuild the list from those entries instead.
    var IndexedPalette := TList<Cardinal>.Create;
    try
      for var I := 0 to High(OoxmlIndexedColors) do
        IndexedPalette.Add(OoxmlIndexedColors[I]);

      const CustomColors = ElementsIn(Xml, Prefix + 'indexedColors', Prefix + 'rgbColor');
      if Length(CustomColors) > 0 then
      begin
        IndexedPalette.Clear;
        for var CustomColor in CustomColors do
        begin
          const Rgb = RgbOf(CustomColor.Attribute('rgb'));
          IndexedPalette.Add(Rgb);
        end;
      end;

      for var Fill in ElementsIn(Xml, Prefix + 'fills', Prefix + 'fill') do
      begin
        const FillColor = FillColorOf(Fill.Inner, Prefix, IndexedPalette);
        Fills.Add(FillColor);
      end;
    finally
      IndexedPalette.Free;
    end;

    for var Border in ElementsIn(Xml, Prefix + 'borders', Prefix + 'border') do
    begin
      var SideStyle: TExcelBorderStyle;
      var SideColor: Cardinal;

      ParseBorderSide(Border.Inner, Prefix, 'top', SideStyle, SideColor);
      TopStyles.Add(SideStyle);
      TopColors.Add(SideColor);
      ParseBorderSide(Border.Inner, Prefix, 'right', SideStyle, SideColor);
      RightStyles.Add(SideStyle);
      RightColors.Add(SideColor);
      ParseBorderSide(Border.Inner, Prefix, 'bottom', SideStyle, SideColor);
      BottomStyles.Add(SideStyle);
      BottomColors.Add(SideColor);
      ParseBorderSide(Border.Inner, Prefix, 'left', SideStyle, SideColor);
      LeftStyles.Add(SideStyle);
      LeftColors.Add(SideColor);
    end;

    var CellXfs: TXmlElement;
    if TXml.TryFindElement(Xml, Prefix + 'cellXfs', CellXfs) then
    begin
      for var Xf in TXml.FindElements(CellXfs.Inner, Prefix + 'xf') do
      begin
        const NumFmtId = StrToIntDef(Xf.Attribute('numFmtId'), 0);

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

        const FontId   = StrToIntDef(Xf.Attribute('fontId'), 0);
        const FillId   = StrToIntDef(Xf.Attribute('fillId'), 0);
        const BorderId = StrToIntDef(Xf.Attribute('borderId'), 0);

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

        var Alignment: TXmlElement;
        if TXml.TryFindElement(Xf.Inner, Prefix + 'alignment', Alignment) then
        begin
          const Horizontal = Alignment.Attribute('horizontal');
          const Vertical   = Alignment.Attribute('vertical');
          const WrapText   = Alignment.Attribute('wrapText');

          FStyleHAlign.Add(StringToHAlign(Horizontal));
          FStyleVAlign.Add(StringToVAlign(Vertical));
          FStyleWrapText.Add(TXml.BooleanOf(WrapText, False));
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

// The ElementName children of the first ContainerName element. Scoping to the
// container keeps, for example, the fonts of differential formats (dxfs) out of
// the font table that cell formats index into.
function TExcelWorkbookReader.ElementsIn(const Xml, ContainerName, ElementName: string): TArray<TXmlElement>;
begin
  var Container: TXmlElement;
  if TXml.TryFindElement(Xml, ContainerName, Container) then
    Result := TXml.FindElements(Container.Inner, ElementName)
  else
    Result := [];
end;

function TExcelWorkbookReader.ChildAttribute(const Xml, ElementName, AttributeName: string): string;
begin
  var Element: TXmlElement;
  if TXml.TryFindElement(Xml, ElementName, Element) then
    Result := Element.Attribute(AttributeName)
  else
    Result := '';
end;

// A font flag such as <b/> is on when present, unless its val switches it off:
// openpyxl writes <b val="1"/>, and <u val="none"/> means no underline.
function TExcelWorkbookReader.IsFontFlagSet(const FontXml, ElementName: string): Boolean;
begin
  var Flag: TXmlElement;
  if not TXml.TryFindElement(FontXml, ElementName, Flag) then
    Exit(False);

  const Value = Flag.Attribute('val');
  Result := (Value <> 'none') and TXml.BooleanOf(Value, True);
end;

function TExcelWorkbookReader.ColorOf(const Xml, ElementName: string): Cardinal;
begin
  const Rgb = ChildAttribute(Xml, ElementName, 'rgb');
  Result := RgbOf(Rgb);
end;

// An ARGB value as RRGGBB. Excel ignores the alpha byte of cell colours, and
// writers differ in what they put there (Excel FF, openpyxl 00).
function TExcelWorkbookReader.RgbOf(const ArgbValue: string): Cardinal;
begin
  const RgbLength = 6;
  const IsTooShort = (Length(ArgbValue) < RgbLength);
  if IsTooShort then
    Exit(0);

  const RgbDigits = Copy(ArgbValue, Length(ArgbValue) - RgbLength + 1, RgbLength);
  Result := StrToInt64Def('$' + RgbDigits, 0);
end;

function TExcelWorkbookReader.FillColorOf(const FillXml, Prefix: string;
                                          const IndexedPalette: TList<Cardinal>): Cardinal;
begin
  var ForegroundColor: TXmlElement;
  if not TXml.TryFindElement(FillXml, Prefix + 'fgColor', ForegroundColor) then
    Exit(0);

  const Rgb = ForegroundColor.Attribute('rgb');
  if not Rgb.IsEmpty then
  begin
    Result := RgbOf(Rgb);
    Exit;
  end;

  const PaletteIndex = StrToIntDef(ForegroundColor.Attribute('indexed'), -1);
  const IsInPalette = (PaletteIndex >= 0) and (PaletteIndex < IndexedPalette.Count);
  if IsInPalette then
    Result := IndexedPalette[PaletteIndex]
  else
    Result := 0;
end;

procedure TExcelWorkbookReader.ParseBorderSide(const BorderXml, Prefix, SideName: string;
                                               out SideStyle: TExcelBorderStyle; out SideColor: Cardinal);
begin
  SideStyle := TExcelBorderStyle.None;
  SideColor := 0;

  var Side: TXmlElement;
  if not TXml.TryFindElement(BorderXml, Prefix + SideName, Side) then
    Exit;

  SideStyle := StringToBorderStyle(Side.Attribute('style'));
  SideColor := ColorOf(Side.Inner, Prefix + 'color');
end;

function TExcelWorkbookReader.StringToBorderStyle(const Value: string): TExcelBorderStyle;
begin
  if Value = 'thin' then Result := TExcelBorderStyle.Thin
  else if Value = 'medium' then Result := TExcelBorderStyle.Medium
  else if Value = 'thick' then Result := TExcelBorderStyle.Thick
  else if Value = 'dashed' then Result := TExcelBorderStyle.Dashed
  else if Value = 'dotted' then Result := TExcelBorderStyle.Dotted
  else if Value = 'double' then Result := TExcelBorderStyle.Double
  else Result := TExcelBorderStyle.None;
end;

function TExcelWorkbookReader.AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;
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

procedure TExcelWorkbookReader.ParseComments(const Sheet: TExcelSheet; const Xml: string);
begin
  // A note's text can be split across multiple <r><t>...</t></r> runs (e.g. if a
  // run has different formatting), which are concatenated. Rich per-run formatting
  // itself is not preserved -- out of scope for plain-text notes.
  const Prefix = TXml.PrefixOf(Xml, NsSpreadsheetML);
  for var Comment in TXml.FindElements(Xml, Prefix + 'comment') do
  begin
    const Text = RichTextOf(Comment.Inner, Prefix);
    Sheet.SetNote(Comment.Attribute('ref'), Text);
  end;
end;

procedure TExcelWorkbookReader.ParseSheet(const Sheet: TExcelSheet; const Xml: string);
begin
  const Prefix = TXml.PrefixOf(Xml, NsSpreadsheetML);

  ParseFrozenPane(Sheet, Xml, Prefix);
  ParseColumnWidths(Sheet, Xml, Prefix);
  ParseSheetData(Sheet, Xml, Prefix);
  ParseMergedCells(Sheet, Xml, Prefix);
end;

procedure TExcelWorkbookReader.ParseFrozenPane(const Sheet: TExcelSheet; const Xml, Prefix: string);
begin
  var Pane: TXmlElement;
  if not TXml.TryFindElement(Xml, Prefix + 'pane', Pane) then
    Exit;

  // Only a frozen pane encodes xSplit/ySplit as whole row/column counts. An unfrozen
  // "split" pane stores the split-bar position in twentieths of a point (e.g. 2160), so
  // reading that as a frozen count would be meaningless -- skip anything whose state is
  // not frozen or frozenSplit.
  const State = Pane.Attribute('state');
  const IsFrozen = (SameText(State, 'frozen') or SameText(State, 'frozenSplit'));
  if not IsFrozen then
    Exit;

  var Split: string;
  if TXml.TryGetAttribute(Pane.OpenTag, 'xSplit', Split) then
    Sheet.SetFrozenColumns(SplitCountOf(Split));
  if TXml.TryGetAttribute(Pane.OpenTag, 'ySplit', Split) then
    Sheet.SetFrozenRows(SplitCountOf(Split));
end;

function TExcelWorkbookReader.SplitCountOf(const Split: string): Integer;
begin
  const SplitPosition = StrToFloatDef(Split, 0, TFormatSettings.Invariant);
  Result := Trunc(SplitPosition);
end;

procedure TExcelWorkbookReader.ParseColumnWidths(const Sheet: TExcelSheet; const Xml, Prefix: string);
begin
  for var Column in TXml.FindElements(Xml, Prefix + 'col') do
  begin
    const IsCustomWidth = TXml.BooleanOf(Column.Attribute('customWidth'), False);
    const Width = StrToFloatDef(Column.Attribute('width'), 0, TFormatSettings.Invariant);
    const HasCustomWidth = (IsCustomWidth and (Width > 0));
    if not HasCustomWidth then
      Continue;

    const FirstColumn = StrToIntDef(Column.Attribute('min'), 0);
    const LastColumn = StrToIntDef(Column.Attribute('max'), 0);
    for var ColumnNumber := FirstColumn to LastColumn do
    begin
      const ColumnLetters = TExcelSheet.ColumnNumberToLetters(ColumnNumber);
      Sheet.SetColumnWidth(ColumnLetters, Width);
    end;
  end;
end;

// Cells are read in two passes: a dependent shared-formula cell only names the
// shared formula by its si index, so every master cell (<f t="shared" si="N"
// ref="...">formula</f>) has to be known before the dependents can be resolved.
procedure TExcelWorkbookReader.ParseSheetData(const Sheet: TExcelSheet; const Xml, Prefix: string);
begin
  var SheetData: TXmlElement;
  if not TXml.TryFindElement(Xml, Prefix + 'sheetData', SheetData) then
    Exit;

  var Cells := TList<TCellXml>.Create;
  var SharedFormulas := TSharedFormulas.Create;
  try
    ReadRows(Sheet, SheetData.Inner, Prefix, Cells);
    CollectSharedFormulas(Cells, SharedFormulas);

    for var CellXml in Cells do
    begin
      ApplyCell(Sheet, CellXml, SharedFormulas);
    end;
  finally
    SharedFormulas.Free;
    Cells.Free;
  end;
end;

// The r attributes of rows and cells are optional: a row without one follows
// the previous row, and a cell without one takes the next column of its row.
procedure TExcelWorkbookReader.ReadRows(const Sheet: TExcelSheet; const SheetDataXml, Prefix: string;
                                        const Cells: TList<TCellXml>);
begin
  var RowNumber := 0;
  for var Row in TXml.FindElements(SheetDataXml, Prefix + 'row') do
  begin
    RowNumber := StrToIntDef(Row.Attribute('r'), RowNumber + 1);
    ApplyRowHeight(Sheet, Row, RowNumber);

    var ColumnNumber := 0;
    for var CellElement in TXml.FindElements(Row.Inner, Prefix + 'c') do
    begin
      var Address := CellElement.Attribute('r');
      if Address.IsEmpty then
        Address := TExcelSheet.ColumnNumberToLetters(ColumnNumber + 1) + IntToStr(RowNumber);

      ColumnNumber := ColumnNumberOf(Address);
      Cells.Add(CellXmlOf(CellElement, Address, Prefix));
    end;
  end;
end;

procedure TExcelWorkbookReader.ApplyRowHeight(const Sheet: TExcelSheet; const Row: TXmlElement;
                                              const RowNumber: Integer);
begin
  const IsCustomHeight = TXml.BooleanOf(Row.Attribute('customHeight'), False);
  const Height = StrToFloatDef(Row.Attribute('ht'), 0, TFormatSettings.Invariant);
  const HasCustomHeight = (IsCustomHeight and (RowNumber > 0) and (Height > 0));
  if HasCustomHeight then
    Sheet.SetRowHeight(RowNumber, Height);
end;

function TExcelWorkbookReader.ColumnNumberOf(const Address: string): Integer;
begin
  var Letters := '';
  for var Symbol in Address do
  begin
    if CharInSet(Symbol, ['A'..'Z', 'a'..'z']) then
      Letters := Letters + Symbol;
  end;

  Result := TExcelSheet.ColumnLetterToNumber(Letters);
end;

function TExcelWorkbookReader.CellXmlOf(const Element: TXmlElement; const Address, Prefix: string): TCellXml;
begin
  Result.Address := Address;
  Result.StyleIndex := StrToIntDef(Element.Attribute('s'), 0);
  Result.CellType := Element.Attribute('t');
  Result.Formula := '';
  Result.SharedFormulaIndex := -1;
  Result.Value := '';
  Result.InlineText := '';

  var Formula: TXmlElement;
  if TXml.TryFindElement(Element.Inner, Prefix + 'f', Formula) then
  begin
    Result.Formula := TXml.Unescape(Formula.Inner);
    Result.SharedFormulaIndex := StrToIntDef(Formula.Attribute('si'), -1);
  end;

  var Value: TXmlElement;
  Result.HasValue := TXml.TryFindElement(Element.Inner, Prefix + 'v', Value);
  if Result.HasValue then
    Result.Value := TXml.Unescape(Value.Inner);

  if Result.CellType = InlineStringCellType then
    Result.InlineText := InlineStringOf(Element.Inner, Prefix);
end;

procedure TExcelWorkbookReader.CollectSharedFormulas(const Cells: TList<TCellXml>;
                                                     const SharedFormulas: TSharedFormulas);
begin
  for var CellXml in Cells do
  begin
    const IsSharedFormulaMaster = ((CellXml.SharedFormulaIndex >= 0) and (CellXml.Formula <> ''));
    if IsSharedFormulaMaster then
      SharedFormulas.AddOrSetValue(CellXml.SharedFormulaIndex,
                                   TPair<string, string>.Create(CellXml.Address, CellXml.Formula));
  end;
end;

// A cell without content still matters when it carries a style the workbook
// defines, such as a background colour on an otherwise empty cell.
procedure TExcelWorkbookReader.ApplyCell(const Sheet: TExcelSheet; const CellXml: TCellXml;
                                         const SharedFormulas: TSharedFormulas);
begin
  const IsStyled = ((CellXml.StyleIndex > 0) and (CellXml.StyleIndex < FStyleBold.Count));

  if not CellXml.HoldsContent then
  begin
    if IsStyled then
    begin
      const EmptyCell = Sheet.GetCell(CellXml.Address) as TExcelCell;
      ApplyCellStyle(EmptyCell, CellXml.StyleIndex);
    end;
    Exit;
  end;

  const Cell = WriteCellValue(Sheet, CellXml, SharedFormulas);
  if (Cell <> nil) and IsStyled then
    ApplyCellStyle(Cell, CellXml.StyleIndex);
end;

function TExcelWorkbookReader.WriteCellValue(const Sheet: TExcelSheet; const CellXml: TCellXml;
                                             const SharedFormulas: TSharedFormulas): TExcelCell;
begin
  const Address = CellXml.Address;
  const StyleIdx = CellXml.StyleIndex;
  const CellType = CellXml.CellType;
  const Value = CellXml.Value;

  var Formula := CellXml.Formula;
  const IsSharedFormulaDependent = (Formula.IsEmpty and (CellXml.SharedFormulaIndex >= 0));
  if IsSharedFormulaDependent then
    Formula := ResolveSharedFormula(CellXml, SharedFormulas);

  Result := nil;

  if Formula <> '' then
  begin
    const ValueType = TCachedValueType.FromCellTypeAttribute(CellType);
    Sheet.SetCellFormula(Address, Formula, Value, ValueType);
    Result := Sheet.GetCell(Address) as TExcelCell;
  end
  else if CellType = 's' then
  begin
    const StrIndex = StrToIntDef(Value, -1);
    if (StrIndex >= 0) and (StrIndex < FSharedStrings.Count) then
    begin
      Sheet.SetCellValue(Address, FSharedStrings[StrIndex], True);
      Result := Sheet.GetCell(Address) as TExcelCell;
    end;
  end
  else if CellType = InlineStringCellType then
  begin
    Sheet.SetCellValue(Address, CellXml.InlineText, True);
    Result := Sheet.GetCell(Address) as TExcelCell;
  end
  else if CellType = 'b' then
  begin
    Sheet.SetBooleanValue(Address, Value = '1');
    Result := Sheet.GetCell(Address) as TExcelCell;
  end
  else
  begin
    if (StyleIdx > 0) and (StyleIdx < FStyleIsDate.Count) and FStyleIsDate[StyleIdx] then
      Sheet.SetDateTimeValue(Address, StrToFloatDef(Value, 0, TFormatSettings.Invariant))
    else
      Sheet.SetCellValue(Address, Value, False);
    Result := Sheet.GetCell(Address) as TExcelCell;
  end;
end;

// Dependent shared-formula cells have an empty formula text but carry a si index.
// The formula of the master cell is moved by the offset from master to dependent.
function TExcelWorkbookReader.ResolveSharedFormula(const CellXml: TCellXml;
                                                   const SharedFormulas: TSharedFormulas): string;
begin
  var Master: TPair<string, string>;
  if not SharedFormulas.TryGetValue(CellXml.SharedFormulaIndex, Master) then
    Exit('');

  var MasterColumn, MasterRow: Integer;
  var DependentColumn, DependentRow: Integer;
  TExcelSheet.ParseCellAddress(Master.Key, MasterColumn, MasterRow);
  TExcelSheet.ParseCellAddress(CellXml.Address, DependentColumn, DependentRow);

  Result := AdjustFormulaRefs(Master.Value, DependentRow - MasterRow, DependentColumn - MasterColumn);
end;

procedure TExcelWorkbookReader.ApplyCellStyle(const Cell: TExcelCell; const StyleIdx: Integer);
begin
  if (StyleIdx < FStyleBold.Count) and (FStyleBold[StyleIdx]) then
    Cell.SetBold(True);
  if (StyleIdx < FStyleItalic.Count) and (FStyleItalic[StyleIdx]) then
    Cell.SetItalic(True);
  if (StyleIdx < FStyleUnderline.Count) and (FStyleUnderline[StyleIdx]) then
    Cell.SetUnderline(True);
  if (StyleIdx < FStyleFontName.Count) and (FStyleFontName[StyleIdx] <> '') then
    Cell.SetFontName(FStyleFontName[StyleIdx]);
  if (StyleIdx < FStyleFontSize.Count) and (FStyleFontSize[StyleIdx] <> 0) then
    Cell.SetFontSize(FStyleFontSize[StyleIdx]);
  if (StyleIdx < FStyleColors.Count) and (FStyleColors[StyleIdx] <> 0) then
    Cell.SetBackgroundColor(FStyleColors[StyleIdx]);
  if (StyleIdx < FStyleFontColor.Count) and (FStyleFontColor[StyleIdx] <> 0) then
    Cell.SetFontColor(FStyleFontColor[StyleIdx]);
  if (StyleIdx < FStyleStrikeout.Count) and (FStyleStrikeout[StyleIdx]) then
    Cell.SetStrikeout(True);
  if (StyleIdx < FStyleBorderTopStyle.Count) and (FStyleBorderTopStyle[StyleIdx] <> TExcelBorderStyle.None) then
    Cell.SetBorderStyle([TExcelBorderSide.Top], FStyleBorderTopStyle[StyleIdx]);
  if (StyleIdx < FStyleBorderTopColor.Count) and (FStyleBorderTopColor[StyleIdx] <> 0) then
    Cell.SetBorderColor([TExcelBorderSide.Top], FStyleBorderTopColor[StyleIdx]);
  if (StyleIdx < FStyleBorderRightStyle.Count) and (FStyleBorderRightStyle[StyleIdx] <> TExcelBorderStyle.None) then
    Cell.SetBorderStyle([TExcelBorderSide.Right], FStyleBorderRightStyle[StyleIdx]);
  if (StyleIdx < FStyleBorderRightColor.Count) and (FStyleBorderRightColor[StyleIdx] <> 0) then
    Cell.SetBorderColor([TExcelBorderSide.Right], FStyleBorderRightColor[StyleIdx]);
  if (StyleIdx < FStyleBorderBottomStyle.Count) and (FStyleBorderBottomStyle[StyleIdx] <> TExcelBorderStyle.None) then
    Cell.SetBorderStyle([TExcelBorderSide.Bottom], FStyleBorderBottomStyle[StyleIdx]);
  if (StyleIdx < FStyleBorderBottomColor.Count) and (FStyleBorderBottomColor[StyleIdx] <> 0) then
    Cell.SetBorderColor([TExcelBorderSide.Bottom], FStyleBorderBottomColor[StyleIdx]);
  if (StyleIdx < FStyleBorderLeftStyle.Count) and (FStyleBorderLeftStyle[StyleIdx] <> TExcelBorderStyle.None) then
    Cell.SetBorderStyle([TExcelBorderSide.Left], FStyleBorderLeftStyle[StyleIdx]);
  if (StyleIdx < FStyleBorderLeftColor.Count) and (FStyleBorderLeftColor[StyleIdx] <> 0) then
    Cell.SetBorderColor([TExcelBorderSide.Left], FStyleBorderLeftColor[StyleIdx]);
  if (StyleIdx < FStyleHAlign.Count) and (FStyleHAlign[StyleIdx] <> TExcelHAlign.None) then
    Cell.SetHAlign(FStyleHAlign[StyleIdx]);
  if (StyleIdx < FStyleVAlign.Count) and (FStyleVAlign[StyleIdx] <> TExcelVAlign.None) then
    Cell.SetVAlign(FStyleVAlign[StyleIdx]);
  if (StyleIdx < FStyleWrapText.Count) and (FStyleWrapText[StyleIdx]) then
    Cell.SetWrapText(True);
  if (StyleIdx < FStyleNumberFormat.Count) and (FStyleNumberFormat[StyleIdx] <> '') then
    Cell.SetNumberFormat(FStyleNumberFormat[StyleIdx]);
end;

procedure TExcelWorkbookReader.ParseMergedCells(const Sheet: TExcelSheet; const Xml, Prefix: string);
begin
  for var MergeCell in TXml.FindElements(Xml, Prefix + 'mergeCell') do
  begin
    const Range = MergeCell.Attribute('ref');
    if not Range.IsEmpty then
      Sheet.MergeCells(Range);
  end;
end;

constructor TExcelWorkbookReader.Create(const Content: TExcelWorkbookContent);
begin
  inherited Create;
  FContent := Content;
  FSheetRelationshipIds := TList<string>.Create;
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

destructor TExcelWorkbookReader.Destroy;
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
  FSheetRelationshipIds.Free;
  inherited;
end;

procedure TExcelWorkbookReader.Read(const Package: TOXMLPackage);
begin
  ReadMetadata(Package);

  if Package.PartExists(PartSharedStrings) then
    ParseSharedStrings(Package.GetPartContent(PartSharedStrings));

  if Package.PartExists(PartStyles) then
    ParseStyles(Package.GetPartContent(PartStyles));

  if Package.PartExists(PartWorkbook) then
    ParseWorkbook(Package.GetPartContent(PartWorkbook));

  ReadSheets(Package);
end;

procedure TExcelWorkbookReader.ReadMetadata(const Package: TOXMLPackage);
begin
  FContent.Metadata := TMetadataParser.ParsePackage(Package);
end;

procedure TExcelWorkbookReader.ReadSheets(const Package: TOXMLPackage);
begin
  var Rels := TRelationships.Create;
  try
    if Package.PartExists(PartWorkbookRels) then
      Rels.LoadFromXml(Package.GetPartContent(PartWorkbookRels));

    for var I := 0 to FContent.Sheets.Count - 1 do
    begin
      const SheetPath = SheetPartPath(Rels, I);
      const Sheet = FContent.Sheets[I] as TExcelSheet;

      if Package.PartExists(SheetPath) then
        ParseSheet(Sheet, Package.GetPartContent(SheetPath));

      ReadSheetComments(Package, Sheet, SheetPath);
    end;
  finally
    Rels.Free;
  end;
end;

// A sheet's part is the target of the relationship its r:id names. Part names
// are the writer's choice: after sheets are reordered in Excel, the first sheet
// can live in sheet3.xml, and openpyxl states its targets as absolute paths.
// Without a resolvable relationship the conventional name is the best guess.
function TExcelWorkbookReader.SheetPartPath(const Rels: TRelationships; const SheetIndex: Integer): string;
begin
  var Relationship: TRelationship;
  const HasRelationship = ((SheetIndex < FSheetRelationshipIds.Count) and
                           Rels.TryGetById(FSheetRelationshipIds[SheetIndex], Relationship));
  if HasRelationship then
  begin
    Result := TRelationships.ResolveTarget(PartWorkbook, Relationship.Target);
    Exit;
  end;

  Result := PartSheetPrefix + IntToStr(SheetIndex + 1) + PartSheetSuffix;
end;

procedure TExcelWorkbookReader.ReadSheetComments(const Package: TOXMLPackage; const Sheet: TExcelSheet;
                                                 const SheetPath: string);
begin
  const SheetRelsPath = TRelationships.RelationshipsPartOf(SheetPath);
  if not Package.PartExists(SheetRelsPath) then
    Exit;

  var SheetRels := TRelationships.Create;
  try
    SheetRels.LoadFromXml(Package.GetPartContent(SheetRelsPath));
    const CommentsTarget = SheetRels.GetTargetByType(RelTypeComments);
    if CommentsTarget = '' then
      Exit;

    const CommentsPath = TRelationships.ResolveTarget(SheetPath, CommentsTarget);
    if Package.PartExists(CommentsPath) then
      ParseComments(Sheet, Package.GetPartContent(CommentsPath));
  finally
    SheetRels.Free;
  end;
end;

end.
