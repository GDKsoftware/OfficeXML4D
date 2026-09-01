unit Office4D.Excel.Writer;

interface

uses
  System.Generics.Collections,
  System.Zip,
  Office4D.Excel,
  Office4D.Excel.Model;

type
  /// <summary>
  /// Turns a TExcelWorkbookContent into the parts of an .xlsx and adds them to
  /// a zip. Cell text and cell formatting are written by reference: strings go
  /// into a shared table and formats into a style table, both built up front so
  /// every sheet can point at the same indexes.
  /// </summary>
  TExcelWorkbookWriter = class
  private
    FContent: TExcelWorkbookContent;

    // Position of every shared string within the list BuildSharedStrings returns.
    // The list stays the ordered form the sharedStrings part is written from;
    // membership and index lookups go through this map, so neither is a scan.
    FSharedStringIndex: TDictionary<string, Integer>;

    function BuildSharedStrings: TList<string>;
    function GetSharedStringIndex(const Value: string): Integer;
    function BuildStyleMap: TDictionary<string, Integer>;
    function GetStyleKey(const Cell: TExcelCell): string;
    function GetCellStyleIndex(const Cell: TExcelCell; const StyleMap: TDictionary<string, Integer>): Integer;

    function GenerateContentTypes: string;
    function GenerateRels: string;
    function GenerateWorkbook: string;
    function GenerateWorkbookRels: string;
    function GenerateSheet(const Sheet: TExcelSheet;
      const StyleMap: TDictionary<string, Integer>): string;
    function GenerateComments(const Sheet: TExcelSheet): string;
    function GenerateVmlDrawing(const Sheet: TExcelSheet): string;
    function GenerateSheetRels(const CommentsFileIndex: Integer): string;
    function GenerateSharedStrings(const Strings: TList<string>): string;
    function GenerateStyles(const StyleMap: TDictionary<string, Integer>): string;

    procedure WriteSheets(const Zip: TZipFile;
      const StyleMap: TDictionary<string, Integer>);
  public
    constructor Create(const Content: TExcelWorkbookContent);
    destructor Destroy; override;

    procedure WriteParts(const Zip: TZipFile);
  end;

implementation

uses
  System.SysUtils,
  System.Math,
  System.StrUtils,
  System.RegularExpressions,
  System.Generics.Defaults,
  Office4D.Errors,
  Office4D.Types,
  Office4D.Xml;

const
  SpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  PartSharedStrings = 'xl/sharedStrings.xml';
  PartStyles = 'xl/styles.xml';
  PartWorkbook = 'xl/workbook.xml';
  PartWorkbookRels = 'xl/_rels/workbook.xml.rels';
  PartSheetPrefix = 'xl/worksheets/sheet';
  PartSheetSuffix = '.xml';

{ TExcelWorkbookWriter }

constructor TExcelWorkbookWriter.Create(const Content: TExcelWorkbookContent);
begin
  inherited Create;
  FContent := Content;
  FSharedStringIndex := TDictionary<string, Integer>.Create;
end;

destructor TExcelWorkbookWriter.Destroy;
begin
  FSharedStringIndex.Free;
  inherited;
end;

procedure TExcelWorkbookWriter.WriteParts(const Zip: TZipFile);
begin
  var SharedStrings := BuildSharedStrings;
  var StyleMap := BuildStyleMap;
  try
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateContentTypes), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateRels), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateWorkbook), PartWorkbook);
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateWorkbookRels), PartWorkbookRels);

    WriteSheets(Zip, StyleMap);

    if SharedStrings.Count > 0 then
      Zip.Add(TEncoding.UTF8.GetBytes(GenerateSharedStrings(SharedStrings)), PartSharedStrings);

    Zip.Add(TEncoding.UTF8.GetBytes(GenerateStyles(StyleMap)), PartStyles);
  finally
    StyleMap.Free;
    SharedStrings.Free;
  end;
end;

procedure TExcelWorkbookWriter.WriteSheets(const Zip: TZipFile;
  const StyleMap: TDictionary<string, Integer>);
begin
  var CommentsFileIndex := 0;

  for var I := 0 to FContent.Sheets.Count - 1 do
  begin
    var ExcelSheet := FContent.Sheets[I] as TExcelSheet;
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSheet(ExcelSheet, StyleMap)),
      PartSheetPrefix + IntToStr(I + 1) + PartSheetSuffix);

    if not ExcelSheet.HasNotes then
      Continue;

    Inc(CommentsFileIndex);

    Zip.Add(TEncoding.UTF8.GetBytes(GenerateComments(ExcelSheet)),
      'xl/comments' + IntToStr(CommentsFileIndex) + '.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateVmlDrawing(ExcelSheet)),
      'xl/drawings/vmlDrawing' + IntToStr(CommentsFileIndex) + '.vml');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSheetRels(CommentsFileIndex)),
      'xl/worksheets/_rels/sheet' + IntToStr(I + 1) + '.xml.rels');
  end;
end;

function TExcelWorkbookWriter.BuildSharedStrings: TList<string>;
begin
  Result := TList<string>.Create;
  FSharedStringIndex.Clear;

  for var Sheet in FContent.Sheets do
  begin
    var ExcelSheet := Sheet as TExcelSheet;
    for var Pair in ExcelSheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      if (Cell.IsString) and (Cell.GetAsString <> '') then
      begin
        const Value = Cell.GetAsString;
        // The map decides whether the string is already known, and remembers
        // where it landed so GetSharedStringIndex never has to search for it.
        if not FSharedStringIndex.ContainsKey(Value) then
        begin
          FSharedStringIndex.Add(Value, Result.Count);
          Result.Add(Value);
        end;
      end;
    end;
  end;
end;

function TExcelWorkbookWriter.GetSharedStringIndex(const Value: string): Integer;
begin
  if not FSharedStringIndex.TryGetValue(Value, Result) then
    Result := -1;
end;

function TExcelWorkbookWriter.GenerateContentTypes: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Types xmlns="' + NsContentTypes + '">');
    SB.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    SB.Append('<Default Extension="xml" ContentType="application/xml"/>');

    var HasAnyNotes := False;
    for var Sheet in FContent.Sheets do
      if (Sheet as TExcelSheet).HasNotes then
        HasAnyNotes := True;
    if HasAnyNotes then
      SB.Append('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');

    SB.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>');
    for var I := 0 to FContent.Sheets.Count - 1 do
      SB.Append('<Override PartName="/xl/worksheets/sheet' + IntToStr(I + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');

    if HasAnyNotes then
    begin
      var CommentsFileIndex := 0;
      for var Sheet in FContent.Sheets do
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

function TExcelWorkbookWriter.GenerateRels: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + NsPackageRelationships + '">');
    SB.Append('<Relationship Id="rId1" Type="' + NsOfficeDocumentRelationships + '/officeDocument" Target="xl/workbook.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbookWriter.GenerateWorkbook: string;
begin
  var HasVisibleSheet := False;
  for var Sheet in FContent.Sheets do
    if Sheet.Visibility = TExcelSheetVisibility.Visible then
    begin
      HasVisibleSheet := True;
      Break;
    end;
  // Excel refuses to open a workbook in which every sheet is hidden.
  if (FContent.Sheets.Count > 0) and not HasVisibleSheet then
    raise EExcelWorkbookException.Create('A workbook must contain at least one visible sheet');

  var HasFormulas := False;
  for var Sheet in FContent.Sheets do
  begin
    if (Sheet as TExcelSheet).HasFormulas then
    begin
      HasFormulas := True;
      Break;
    end;
  end;

  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<workbook xmlns="' + SpreadsheetNs + '" xmlns:r="' + NsOfficeDocumentRelationships + '">');
    SB.Append('<sheets>');
    for var I := 0 to FContent.Sheets.Count - 1 do
    begin
      var StateAttr := '';
      if FContent.Sheets[I].Visibility = TExcelSheetVisibility.Hidden then
        StateAttr := ' state="hidden"'
      else if FContent.Sheets[I].Visibility = TExcelSheetVisibility.VeryHidden then
        StateAttr := ' state="veryHidden"';
      SB.Append('<sheet name="' + TXml.Escape(FContent.Sheets[I].Name) + '" sheetId="' + IntToStr(I + 1) + '"' + StateAttr + ' r:id="rId' + IntToStr(I + 1) + '"/>');
    end;
    SB.Append('</sheets>');
    // Formulas are not evaluated here, so a formula cell's cached <v> is stale
    // as soon as anything it depends on is written. Ask Excel to recalculate on
    // load, but only when there is a formula to recalculate: a workbook without
    // formulas carries no calcPr at all. The recalculation makes Excel prompt to
    // save on close, so the caller can opt out via RecalculateOnLoad.
    if HasFormulas and FContent.RecalculateOnLoad then
      SB.Append('<calcPr calcId="0" fullCalcOnLoad="1"/>');
    SB.Append('</workbook>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbookWriter.GenerateWorkbookRels: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + NsPackageRelationships + '">');
    for var I := 0 to FContent.Sheets.Count - 1 do
      SB.Append('<Relationship Id="rId' + IntToStr(I + 1) + '" Type="' + NsOfficeDocumentRelationships + '/worksheet" Target="worksheets/sheet' + IntToStr(I + 1) + '.xml"/>');
    SB.Append('<Relationship Id="rId' + IntToStr(FContent.Sheets.Count + 1) + '" Type="' + NsOfficeDocumentRelationships + '/sharedStrings" Target="sharedStrings.xml"/>');
    SB.Append('<Relationship Id="rId' + IntToStr(FContent.Sheets.Count + 2) + '" Type="' + NsOfficeDocumentRelationships + '/styles" Target="styles.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbookWriter.GenerateSheet(const Sheet: TExcelSheet; const StyleMap: TDictionary<string, Integer>): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<worksheet xmlns="' + SpreadsheetNs + '" xmlns:r="' + NsOfficeDocumentRelationships + '">');

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
              SB.Append('<c r="' + CellPair.Key + '"' + StyleAttr + '><f>' + TXml.Escape(Cell.GetFormula) + '</f><v>' + FloatVal + '</v></c>');
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
                    const StrIdx = GetSharedStringIndex(Cell.GetAsString);
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

function TExcelWorkbookWriter.GenerateComments(const Sheet: TExcelSheet): string;
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

function TExcelWorkbookWriter.GenerateVmlDrawing(const Sheet: TExcelSheet): string;
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

function TExcelWorkbookWriter.GenerateSheetRels(const CommentsFileIndex: Integer): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + NsPackageRelationships + '">');
    SB.Append('<Relationship Id="rId1" Type="' + NsOfficeDocumentRelationships + '/vmlDrawing" Target="../drawings/vmlDrawing' +
      IntToStr(CommentsFileIndex) + '.vml"/>');
    SB.Append('<Relationship Id="rId2" Type="' + NsOfficeDocumentRelationships + '/comments" Target="../comments' +
      IntToStr(CommentsFileIndex) + '.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TExcelWorkbookWriter.GenerateSharedStrings(const Strings: TList<string>): string;
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

function TExcelWorkbookWriter.GetStyleKey(const Cell: TExcelCell): string;
begin
  var FontSizeStr := '';
  if Cell.GetFontSize <> 0 then
    FontSizeStr := FormatFloat('0.##', Cell.GetFontSize, TFormatSettings.Invariant);
  // 0 = no date format, 1 = date only (built-in numFmtId 14), 2 = date with time (built-in numFmtId 22)
  var DateFlag := 0;
  if Cell.CellType = TCellType.DateTime then
    if Frac(Cell.GetAsDateTime) <> 0 then
      DateFlag := 2
    else
      DateFlag := 1;
  Result := Format('%d|%d|%d|%s|%d|%d|%s|%s|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d|%d', [
    Ord(Cell.GetBold),
    Cell.GetBackgroundColor,
    DateFlag,
    Cell.GetNumberFormat,
    Ord(Cell.GetItalic),
    Ord(Cell.GetUnderline),
    Cell.GetFontName,
    FontSizeStr,
    Ord(Cell.GetBorderStyle([TExcelBorderSide.Top])),             // 8
    Cell.GetBorderColor([TExcelBorderSide.Top]),                  // 9
    Ord(Cell.GetBorderStyle([TExcelBorderSide.Right])),           // 10
    Cell.GetBorderColor([TExcelBorderSide.Right]),                // 11
    Ord(Cell.GetBorderStyle([TExcelBorderSide.Bottom])),          // 12
    Cell.GetBorderColor([TExcelBorderSide.Bottom]),               // 13
    Ord(Cell.GetBorderStyle([TExcelBorderSide.Left])),            // 14
    Cell.GetBorderColor([TExcelBorderSide.Left]),                 // 15
    Ord(Cell.GetHAlign),
    Ord(Cell.GetVAlign),
    Ord(Cell.GetWrapText),
    Cell.GetFontColor,
    Ord(Cell.GetStrikeout)
  ]);
end;

function TExcelWorkbookWriter.BuildStyleMap: TDictionary<string, Integer>;
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
  for var Sheet in FContent.Sheets do
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

function TExcelWorkbookWriter.GetCellStyleIndex(const Cell: TExcelCell; const StyleMap: TDictionary<string, Integer>): Integer;
begin
  const Key = GetStyleKey(Cell);
  if StyleMap.ContainsKey(Key) then
    Result := StyleMap[Key]
  else
    Result := 0;
end;

function TExcelWorkbookWriter.GenerateStyles(const StyleMap: TDictionary<string, Integer>): string;

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

  for var Sheet in FContent.Sheets do
  begin
    var ExcelSheet := Sheet as TExcelSheet;
    for var Pair in ExcelSheet.Cells do
    begin
      var Cell := Pair.Value as TExcelCell;
      if (Cell.GetBackgroundColor <> 0) and (not Colors.Contains(Cell.GetBackgroundColor)) then
        Colors.Add(Cell.GetBackgroundColor);
      if (Cell.GetNumberFormat <> '') and (not NumFormats.Contains(Cell.GetNumberFormat)) then
        NumFormats.Add(Cell.GetNumberFormat);

      var FontSizeStr := '';
      if Cell.GetFontSize <> 0 then
        FontSizeStr := FormatFloat('0.##', Cell.GetFontSize, TFormatSettings.Invariant);
      const FontKey = Format('%d|%d|%d|%s|%s|%d|%d', [
        Ord(Cell.GetBold), Ord(Cell.GetItalic), Ord(Cell.GetUnderline), Cell.GetFontName, FontSizeStr, Cell.GetFontColor, Ord(Cell.GetStrikeout)]);
      if (FontKey <> '0|0|0|||0|0') and (not FontKeys.Contains(FontKey)) then
        FontKeys.Add(FontKey);

      const BorderKey = Format('%d|%d|%d|%d|%d|%d|%d|%d', [
        Ord(Cell.GetBorderStyle([TExcelBorderSide.Top])), Cell.GetBorderColor([TExcelBorderSide.Top]),
        Ord(Cell.GetBorderStyle([TExcelBorderSide.Right])), Cell.GetBorderColor([TExcelBorderSide.Right]),
        Ord(Cell.GetBorderStyle([TExcelBorderSide.Bottom])), Cell.GetBorderColor([TExcelBorderSide.Bottom]),
        Ord(Cell.GetBorderStyle([TExcelBorderSide.Left])), Cell.GetBorderColor([TExcelBorderSide.Left])]);
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

end.
