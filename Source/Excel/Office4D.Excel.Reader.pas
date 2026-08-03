unit Office4D.Excel.Reader;

interface

uses
  System.Generics.Collections,
  Office4D.Excel,
  Office4D.Package,
  Office4D.Excel.Model;

type
  /// <summary>
  /// Fills a TExcelWorkbookContent from an opened .xlsx package. The shared
  /// strings and the style tables it parses first are scratch state: cells refer
  /// to them by index, so they only have to survive until every sheet is read.
  /// </summary>
  TExcelWorkbookReader = class
  private
    FContent: TExcelWorkbookContent;
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
    function AdjustFormulaRefs(const Formula: string; const RowDelta, ColDelta: Integer): string;

    procedure ReadMetadata(const Package: TOXMLPackage);
    procedure ReadSheets(const Package: TOXMLPackage);
    procedure ReadSheetComments(const Package: TOXMLPackage; const Sheet: TExcelSheet; SheetIndex: Integer);
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

  PartSharedStrings = 'xl/sharedStrings.xml';
  PartStyles = 'xl/styles.xml';
  PartWorkbook = 'xl/workbook.xml';
  PartWorkbookRels = 'xl/_rels/workbook.xml.rels';
  PartSheetPrefix = 'xl/worksheets/sheet';
  PartSheetSuffix = '.xml';

{ TExcelWorkbookReader }

procedure TExcelWorkbookReader.ParseWorkbook(const Xml: string);
begin
  const Matches = TRegEx.Matches(Xml, '<sheet\s+name="([^"]+)"([^>]*?)(?:/>|>)', [roIgnoreCase]);
  for var Match in Matches do
    if Match.Groups.Count > 1 then
    begin
      const Sheet = FContent.CreateSheet(Match.Groups[1].Value);
      const StateMatch = TRegEx.Match(Match.Groups[2].Value, 'state="([^"]+)"', [roIgnoreCase]);
      if StateMatch.Success then
      begin
        if SameText(StateMatch.Groups[1].Value, 'hidden') then
          Sheet.Visibility := TExcelSheetVisibility.Hidden
        else if SameText(StateMatch.Groups[1].Value, 'veryHidden') then
          Sheet.Visibility := TExcelSheetVisibility.VeryHidden;
      end;

    end;
end;

procedure TExcelWorkbookReader.ParseSharedStrings(const Xml: string);
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

procedure TExcelWorkbookReader.ParseStyles(const Xml: string);

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

procedure TExcelWorkbookReader.ParseSheet(const Sheet: TExcelSheet; const Xml: string);
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
        Sheet.SetFrozenColumns(Trunc(StrToFloatDef(XSplitMatch.Groups[1].Value, 0, TFormatSettings.Invariant)));
      if YSplitMatch.Success then
        Sheet.SetFrozenRows(Trunc(StrToFloatDef(YSplitMatch.Groups[1].Value, 0, TFormatSettings.Invariant)));
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
    end;

  const MergeMatches = TRegEx.Matches(Xml, '<mergeCell\s+ref="([^"]+)"', [roIgnoreCase]);
  for var MergeMatch in MergeMatches do
    if MergeMatch.Groups.Count > 1 then
      Sheet.MergeCells(MergeMatch.Groups[1].Value);
end;

constructor TExcelWorkbookReader.Create(const Content: TExcelWorkbookContent);
begin
  inherited Create;
  FContent := Content;
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
      const SheetPath = PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix;
      const Sheet = FContent.Sheets[I] as TExcelSheet;

      if Package.PartExists(SheetPath) then
        ParseSheet(Sheet, Package.GetPartContent(SheetPath));

      ReadSheetComments(Package, Sheet, I);
    end;
  finally
    Rels.Free;
  end;
end;

procedure TExcelWorkbookReader.ReadSheetComments(const Package: TOXMLPackage; const Sheet: TExcelSheet;
  SheetIndex: Integer);
begin
  const SheetRelsPath = 'xl/worksheets/_rels/sheet' + IntToStr(SheetIndex + 1) + '.xml.rels';
  if not Package.PartExists(SheetRelsPath) then
    Exit;

  var SheetRels := TRelationships.Create;
  try
    SheetRels.LoadFromXml(Package.GetPartContent(SheetRelsPath));
    const CommentsTarget = SheetRels.GetTargetByType(NsOfficeDocumentRelationships + '/comments');
    if CommentsTarget = '' then
      Exit;

    // Targets are written relative to xl/worksheets/ (e.g. '../comments1.xml').
    // This resolves that specific convention rather than implementing full OPC
    // relative-URI resolution -- consistent with this library's other pragmatic,
    // regex-based parsing rather than a general-purpose XML/URI toolchain.
    var ResolvedPath := CommentsTarget;
    if ResolvedPath.StartsWith('../') then
      ResolvedPath := 'xl/' + Copy(ResolvedPath, 4, MaxInt);
    if Package.PartExists(ResolvedPath) then
      ParseComments(Sheet, Package.GetPartContent(ResolvedPath));
  finally
    SheetRels.Free;
  end;
end;

end.
