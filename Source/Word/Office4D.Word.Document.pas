unit Office4D.Word.Document;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections,
  System.RegularExpressions,
  System.Zip,
  Office4D.Word,
  Office4D.Metadata,
  Office4D.Package,
  Office4D.Relationships,
  Office4D.Xml;

type
  TWordRun = class(TInterfacedObject, IWordRun)
  private
    FText: string;
    FBold: Boolean;
    FItalic: Boolean;
    FUnderline: Boolean;
    FHyperlink: string;
    FFontName: string;
    FFontSize: Integer;
    FFontColor: string;
    FImage: TWordImage;
    FHasImage: Boolean;
  public
    function GetText: string;
    procedure SetText(const Value: string);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetHyperlink: string;
    procedure SetHyperlink(const Value: string);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Integer;
    procedure SetFontSize(const Value: Integer);
    function GetFontColor: string;
    procedure SetFontColor(const Value: string);
    function GetImage: TWordImage;
    function HasImage: Boolean;
    procedure AddImage(const AData: TBytes; const AExtension: string; AWidthPx, AHeightPx: Integer);
  end;

  TWordParagraph = class(TInterfacedObject, IWordParagraph)
  private
    FRuns: TList<IWordRun>;
    FListStyle: TListStyle;
    FAlignment: TParagraphAlignment;
    FLineSpacing: TLineSpacing;
    FIndent: TParagraphIndent;
  public
    constructor Create;
    destructor Destroy; override;

    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IWordRun;
    function AddRun(const Text: string): IWordRun;
    procedure AddLineBreak;
    procedure AddTab;
    procedure AddPageBreak;
    function GetListStyle: TListStyle;
    procedure SetListStyle(const Value: TListStyle);
    function GetAlignment: TParagraphAlignment;
    procedure SetAlignment(const Value: TParagraphAlignment);
    function GetLineSpacing: TLineSpacing;
    procedure SetLineSpacing(const Value: TLineSpacing);
    function GetIndent: TParagraphIndent;
    procedure SetIndent(const Value: TParagraphIndent);
  end;

  TWordTableCell = class(TInterfacedObject, IWordTableCell)
  private
    FText: string;
    FShading: string;
    FWidth: Integer;
  public
    function GetText: string;
    procedure SetText(const Value: string);
    function GetShading: string;
    procedure SetShading(const Value: string);
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);
  end;

  TWordTable = class(TInterfacedObject, IWordTable)
  private
    FRowCount: Integer;
    FColCount: Integer;
    FCells: TList<TList<IWordTableCell>>;
    FBorders: TTableBorders;
    FColumnWidths: TArray<Integer>;
  public
    constructor Create(const Rows, Cols: Integer);
    destructor Destroy; override;

    function GetRowCount: Integer;
    function GetColCount: Integer;
    function GetCell(Row, Col: Integer): IWordTableCell;
    function GetBorders: TTableBorders;
    procedure SetBorders(const Value: TTableBorders);
    procedure SetColumnWidths(const Widths: array of Integer);
  end;

  TWordHeaderFooter = class(TInterfacedObject, IWordHeaderFooter)
  private
    FText: string;
  public
    function GetText: string;
    procedure SetText(const Value: string);
  end;

  TWordDocument = class(TInterfacedObject, IWordDocument)
  private
    FParagraphs: TList<IWordParagraph>;
    FTables: TList<IWordTable>;
    FMetadata: TDocumentMetadata;
    FPackage: TOXMLPackage;
    FPageOrientation: TPageOrientation;
    FPageMargins: TPageMargins;
    FHeader: IWordHeaderFooter;
    FFooter: IWordHeaderFooter;

    procedure ResetContent;
    procedure LoadFromPackage;
    procedure WriteParts(const Zip: TZipFile);
    procedure ParseDocumentXml(const XmlContent: string; const HyperlinkMap: TDictionary<string, string>);
    procedure ParseParagraphs(const Xml: string; const HyperlinkMap: TDictionary<string, string>);
    procedure ParseParagraph(const ParagraphXml: string; const HyperlinkMap: TDictionary<string, string>);
    procedure ParseParagraphContent(const Para: TWordParagraph; const ParagraphXml: string;
      const HyperlinkMap: TDictionary<string, string>);
    procedure ParseHyperlink(const Para: TWordParagraph; const Hyperlink: TXmlElement;
      const HyperlinkMap: TDictionary<string, string>);
    procedure ParseRun(const Para: TWordParagraph; const RunXml, Hyperlink: string);
    procedure ApplyListStyle(const Para: TWordParagraph; const ParagraphXml: string);
    procedure ParseDocumentRels(const XmlContent: string; const Hyperlinks: TDictionary<string, string>);
    procedure ParseHeaderFooterXml(const XmlContent: string; const Target: IWordHeaderFooter);
    function GenerateContentTypesXml: string;
    function GenerateRootRelsXml: string;
    function GenerateDocumentXml: string;
    function GenerateDocumentRelsXml: string;
    function GenerateHeaderXml: string;
    function GenerateFooterXml: string;
    function CollectHyperlinks: TList<string>;
    function GetHyperlinkId(const Url: string; const Hyperlinks: TList<string>): Integer;
    function HasListParagraphs: Boolean;
    function GenerateNumberingXml: string;
    function HasHeader: Boolean;
    function HasFooter: Boolean;
    function GenerateBorderXml(const ElementName: string; const Border: TTableBorder): string;
    function CollectImages: TList<TWordRun>;
    function GetImageContentType(const Extension: string): string;
    function GenerateDrawingXml(const ImageRun: TWordRun; ImageIndex: Integer; RelIdOffset: Integer): string;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;

    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IWordParagraph;
    function AddParagraph: IWordParagraph;

    function GetTableCount: Integer;
    function GetTable(Index: Integer): IWordTable;
    function AddTable(const Rows, Cols: Integer): IWordTable;

    function GetPageOrientation: TPageOrientation;
    procedure SetPageOrientation(const Value: TPageOrientation);
    function GetPageMargins: TPageMargins;
    procedure SetPageMargins(const Value: TPageMargins);
    function GetHeader: IWordHeaderFooter;
    function GetFooter: IWordHeaderFooter;
  end;

implementation

uses
  Office4D.Errors,
  Office4D.Types;

const
  XmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  RelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  ContentTypesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
  WordprocessingNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  DrawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  WpDrawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  PictureNs = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  RelationshipsDocNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  PartDocumentRels = 'word/_rels/document.xml.rels';
  PartRootRels = '_rels/.rels';
  PartCoreProps = 'docProps/core.xml';
  PartWordPrefix = 'word/';
  KeyHeader = '__header__';
  KeyFooter = '__footer__';

  EmuPerPixel = 9525;

  ElementParagraph = 'w:p';
  ElementRun = 'w:r';
  ElementHyperlink = 'w:hyperlink';

  TextElementPattern = '<w:t(?:\s[^>]*)?>([^<]*)</w:t>';

  /// Matches a w:t element with its text, or a standalone w:br or w:tab, so the
  /// children of a run can be walked in document order.
  RunContentPattern = TextElementPattern + '|<w:t\s*/>|<w:br(?:\s[^>]*)?/?>|<w:tab(?:\s[^>]*)?/?>';

{ TWordRun }

function TWordRun.GetText: string;
begin
  Result := FText;
end;

procedure TWordRun.SetText(const Value: string);
begin
  FText := Value;
end;

function TWordRun.GetBold: Boolean;
begin
  Result := FBold;
end;

procedure TWordRun.SetBold(const Value: Boolean);
begin
  FBold := Value;
end;

function TWordRun.GetItalic: Boolean;
begin
  Result := FItalic;
end;

procedure TWordRun.SetItalic(const Value: Boolean);
begin
  FItalic := Value;
end;

function TWordRun.GetUnderline: Boolean;
begin
  Result := FUnderline;
end;

procedure TWordRun.SetUnderline(const Value: Boolean);
begin
  FUnderline := Value;
end;

function TWordRun.GetHyperlink: string;
begin
  Result := FHyperlink;
end;

procedure TWordRun.SetHyperlink(const Value: string);
begin
  FHyperlink := Value;
end;

function TWordRun.GetFontName: string;
begin
  Result := FFontName;
end;

procedure TWordRun.SetFontName(const Value: string);
begin
  FFontName := Value;
end;

function TWordRun.GetFontSize: Integer;
begin
  Result := FFontSize;
end;

procedure TWordRun.SetFontSize(const Value: Integer);
begin
  FFontSize := Value;
end;

function TWordRun.GetFontColor: string;
begin
  Result := FFontColor;
end;

procedure TWordRun.SetFontColor(const Value: string);
begin
  FFontColor := Value;
end;

function TWordRun.GetImage: TWordImage;
begin
  Result := FImage;
end;

function TWordRun.HasImage: Boolean;
begin
  Result := FHasImage;
end;

procedure TWordRun.AddImage(const AData: TBytes; const AExtension: string; AWidthPx, AHeightPx: Integer);
begin
  FImage := TWordImage.Create(AData, AExtension, AWidthPx * EmuPerPixel, AHeightPx * EmuPerPixel);
  FHasImage := True;
  FText := '';
end;

{ TWordParagraph }

constructor TWordParagraph.Create;
begin
  inherited Create;
  FRuns := TList<IWordRun>.Create;
end;

destructor TWordParagraph.Destroy;
begin
  FRuns.Free;
  inherited;
end;

function TWordParagraph.GetText: string;
begin
  Result := '';
  for var Run in FRuns do
    Result := Result + Run.Text;
end;

function TWordParagraph.GetRunCount: Integer;
begin
  Result := FRuns.Count;
end;

function TWordParagraph.GetRun(Index: Integer): IWordRun;
begin
  Result := FRuns[Index];
end;

function TWordParagraph.AddRun(const Text: string): IWordRun;
begin
  var Run := TWordRun.Create;
  Run.FText := Text;
  FRuns.Add(Run);
  Result := Run;
end;

procedure TWordParagraph.AddLineBreak;
begin
  var Run := TWordRun.Create;
  Run.FText := sLineBreak;
  FRuns.Add(Run);
end;

procedure TWordParagraph.AddTab;
begin
  var Run := TWordRun.Create;
  Run.FText := #9;
  FRuns.Add(Run);
end;

function TWordParagraph.GetListStyle: TListStyle;
begin
  Result := FListStyle;
end;

procedure TWordParagraph.SetListStyle(const Value: TListStyle);
begin
  FListStyle := Value;
end;

function TWordParagraph.GetAlignment: TParagraphAlignment;
begin
  Result := FAlignment;
end;

procedure TWordParagraph.SetAlignment(const Value: TParagraphAlignment);
begin
  FAlignment := Value;
end;

procedure TWordParagraph.AddPageBreak;
begin
  var Run := TWordRun.Create;
  Run.FText := #12;
  FRuns.Add(Run);
end;

function TWordParagraph.GetLineSpacing: TLineSpacing;
begin
  Result := FLineSpacing;
end;

procedure TWordParagraph.SetLineSpacing(const Value: TLineSpacing);
begin
  FLineSpacing := Value;
end;

function TWordParagraph.GetIndent: TParagraphIndent;
begin
  Result := FIndent;
end;

procedure TWordParagraph.SetIndent(const Value: TParagraphIndent);
begin
  FIndent := Value;
end;

{ TWordTableCell }

function TWordTableCell.GetText: string;
begin
  Result := FText;
end;

procedure TWordTableCell.SetText(const Value: string);
begin
  FText := Value;
end;

function TWordTableCell.GetShading: string;
begin
  Result := FShading;
end;

procedure TWordTableCell.SetShading(const Value: string);
begin
  FShading := Value;
end;

function TWordTableCell.GetWidth: Integer;
begin
  Result := FWidth;
end;

procedure TWordTableCell.SetWidth(const Value: Integer);
begin
  FWidth := Value;
end;

{ TWordTable }

constructor TWordTable.Create(const Rows, Cols: Integer);
begin
  inherited Create;
  FRowCount := Rows;
  FColCount := Cols;
  FCells := TList<TList<IWordTableCell>>.Create;

  for var RowIndex := 0 to Rows - 1 do
  begin
    var Row := TList<IWordTableCell>.Create;
    for var ColIndex := 0 to Cols - 1 do
    begin
      Row.Add(TWordTableCell.Create);
    end;
    FCells.Add(Row);
  end;
end;

destructor TWordTable.Destroy;
begin
  for var Row in FCells do
  begin
    Row.Free;
  end;
  FCells.Free;
  inherited;
end;

function TWordTable.GetRowCount: Integer;
begin
  Result := FRowCount;
end;

function TWordTable.GetColCount: Integer;
begin
  Result := FColCount;
end;

function TWordTable.GetCell(Row, Col: Integer): IWordTableCell;
begin
  Result := FCells[Row][Col];
end;

function TWordTable.GetBorders: TTableBorders;
begin
  Result := FBorders;
end;

procedure TWordTable.SetBorders(const Value: TTableBorders);
begin
  FBorders := Value;
end;

procedure TWordTable.SetColumnWidths(const Widths: array of Integer);
begin
  SetLength(FColumnWidths, Length(Widths));
  for var I := 0 to High(Widths) do
    FColumnWidths[I] := Widths[I];
end;

{ TWordHeaderFooter }

function TWordHeaderFooter.GetText: string;
begin
  Result := FText;
end;

procedure TWordHeaderFooter.SetText(const Value: string);
begin
  FText := Value;
end;

{ TWordDocument }

constructor TWordDocument.Create;
begin
  inherited Create;
  FParagraphs := TList<IWordParagraph>.Create;
  FTables := TList<IWordTable>.Create;
  FMetadata.Clear;
  FHeader := TWordHeaderFooter.Create;
  FFooter := TWordHeaderFooter.Create;
end;

destructor TWordDocument.Destroy;
begin
  FTables.Free;
  FParagraphs.Free;
  FreeAndNil(FPackage);
  inherited;
end;

procedure TWordDocument.ParseDocumentXml(const XmlContent: string; const HyperlinkMap: TDictionary<string, string>);
begin
  FParagraphs.Clear;

  ParseParagraphs(TXml.ReduceAlternateContent(XmlContent), HyperlinkMap);
end;

procedure TWordDocument.ParseParagraphs(const Xml: string; const HyperlinkMap: TDictionary<string, string>);
begin
  for var Element in TXml.FindElements(Xml, ElementParagraph) do
  begin
    var ParagraphXml := Element.Inner;
    var NestedParagraphs := TXml.FindElements(ParagraphXml, ElementParagraph);

    if Length(NestedParagraphs) > 0 then
    begin
      // A paragraph nested in a shape's text box is a paragraph in its own
      // right. It is read first, and removing it here keeps the runs that
      // follow the shape inside the hosting paragraph.
      ParseParagraphs(ParagraphXml, HyperlinkMap);
      ParagraphXml := TXml.RemoveElements(ParagraphXml, NestedParagraphs);
    end;

    ParseParagraph(ParagraphXml, HyperlinkMap);
  end;
end;

procedure TWordDocument.ParseParagraph(const ParagraphXml: string; const HyperlinkMap: TDictionary<string, string>);
begin
  var Para := TWordParagraph.Create;

  ApplyListStyle(Para, ParagraphXml);
  ParseParagraphContent(Para, ParagraphXml, HyperlinkMap);

  if Para.GetRunCount > 0 then
    FParagraphs.Add(Para)
  else
    Para.Free;
end;

procedure TWordDocument.ApplyListStyle(const Para: TWordParagraph; const ParagraphXml: string);
begin
  var NumIdPattern := '<w:numId\s+w:val="(\d+)"';
  var NumIdMatch := TRegEx.Match(ParagraphXml, NumIdPattern, [roIgnoreCase]);
  if not (NumIdMatch.Success and (NumIdMatch.Groups.Count > 1)) then
    Exit;

  var NumId := StrToIntDef(NumIdMatch.Groups[1].Value, 0);
  if NumId = 1 then
    Para.FListStyle := TListStyle.Bullet
  else if NumId = 2 then
    Para.FListStyle := TListStyle.Numbered;
end;

procedure TWordDocument.ParseParagraphContent(const Para: TWordParagraph; const ParagraphXml: string;
  const HyperlinkMap: TDictionary<string, string>);
begin
  var Hyperlinks := TXml.FindElements(ParagraphXml, ElementHyperlink);
  var Runs := TXml.FindElements(ParagraphXml, ElementRun);
  var NextHyperlink := 0;

  for var Run in Runs do
  begin
    while (NextHyperlink < Length(Hyperlinks)) and (Hyperlinks[NextHyperlink].StartPos < Run.StartPos) do
    begin
      ParseHyperlink(Para, Hyperlinks[NextHyperlink], HyperlinkMap);
      Inc(NextHyperlink);
    end;

    var IsHyperlinkRun := (NextHyperlink > 0) and (Run.EndPos <= Hyperlinks[NextHyperlink - 1].EndPos);
    if not IsHyperlinkRun then
      ParseRun(Para, Run.Inner, '');
  end;

  while NextHyperlink < Length(Hyperlinks) do
  begin
    ParseHyperlink(Para, Hyperlinks[NextHyperlink], HyperlinkMap);
    Inc(NextHyperlink);
  end;
end;

procedure TWordDocument.ParseHyperlink(const Para: TWordParagraph; const Hyperlink: TXmlElement;
  const HyperlinkMap: TDictionary<string, string>);
begin
  var Url := '';

  var RelIdMatch := TRegEx.Match(Hyperlink.OpenTag, 'r:id="([^"]*)"', [roIgnoreCase]);
  if RelIdMatch.Success and (RelIdMatch.Groups.Count > 1) then
    HyperlinkMap.TryGetValue(RelIdMatch.Groups[1].Value, Url);

  for var Run in TXml.FindElements(Hyperlink.Inner, ElementRun) do
    ParseRun(Para, Run.Inner, Url);
end;

procedure TWordDocument.ParseRun(const Para: TWordParagraph; const RunXml, Hyperlink: string);
begin
  for var Match in TRegEx.Matches(RunXml, RunContentPattern, [roIgnoreCase]) do
  begin
    if Match.Value.StartsWith('<w:br') then
      Para.AddLineBreak
    else if Match.Value.StartsWith('<w:tab') then
      Para.AddTab
    else if Match.Groups.Count > 1 then
    begin
      var TextValue := TXml.Unescape(Match.Groups[1].Value);
      if TextValue <> '' then
        Para.AddRun(TextValue).Hyperlink := Hyperlink;
    end;
  end;
end;

procedure TWordDocument.LoadFromFile(const FileName: string);
begin
  ResetContent;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(FileName);

  LoadFromPackage;
end;

procedure TWordDocument.LoadFromStream(const Stream: TStream);
begin
  ResetContent;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(Stream);

  LoadFromPackage;
end;

procedure TWordDocument.ResetContent;
begin
  FreeAndNil(FPackage);
  FParagraphs.Clear;
  FTables.Clear;
  FMetadata.Clear;
  FHeader.Text := '';
  FFooter.Text := '';
end;

procedure TWordDocument.LoadFromPackage;
begin
  var HyperlinkMap := TDictionary<string, string>.Create;
  try
    if FPackage.PartExists(PartDocumentRels) then
    begin
      var DocRelsXml := FPackage.GetPartContent(PartDocumentRels);
      ParseDocumentRels(DocRelsXml, HyperlinkMap);

      if HyperlinkMap.ContainsKey(KeyHeader) then
      begin
        const HeaderPath = PartWordPrefix + HyperlinkMap[KeyHeader];
        if FPackage.PartExists(HeaderPath) then
          ParseHeaderFooterXml(FPackage.GetPartContent(HeaderPath), FHeader);
      end;

      if HyperlinkMap.ContainsKey(KeyFooter) then
      begin
        const FooterPath = PartWordPrefix + HyperlinkMap[KeyFooter];
        if FPackage.PartExists(FooterPath) then
          ParseHeaderFooterXml(FPackage.GetPartContent(FooterPath), FFooter);
      end;
    end;

    var RelsXml := FPackage.GetPartContent(PartRootRels);
    var Rels := TRelationships.Create;
    try
      Rels.LoadFromXml(RelsXml);
      var DocumentPath := Rels.GetTargetByType(RelTypeOfficeDocument);

      if DocumentPath <> '' then
      begin
        var DocumentXml := FPackage.GetPartContent(DocumentPath);
        ParseDocumentXml(DocumentXml, HyperlinkMap);
      end;
    finally
      Rels.Free;
    end;
  finally
    HyperlinkMap.Free;
  end;

  if FPackage.PartExists(PartCoreProps) then
  begin
    var CoreXml := FPackage.GetPartContent(PartCoreProps);
    var MetaParser := TMetadataParser.Create;
    try
      FMetadata := MetaParser.Parse(CoreXml);
    finally
      MetaParser.Free;
    end;
  end;
end;

function TWordDocument.CollectHyperlinks: TList<string>;
begin
  Result := TList<string>.Create;
  for var Para in FParagraphs do
  begin
    for var RunIndex := 0 to Para.RunCount - 1 do
    begin
      var Run := Para.Runs[RunIndex];
      if (Run.Hyperlink <> '') and (not Result.Contains(Run.Hyperlink)) then
        Result.Add(Run.Hyperlink);
    end;
  end;
end;

function TWordDocument.GetHyperlinkId(const Url: string; const Hyperlinks: TList<string>): Integer;
begin
  Result := Hyperlinks.IndexOf(Url) + 1;
end;

procedure TWordDocument.ParseDocumentRels(const XmlContent: string; const Hyperlinks: TDictionary<string, string>);
begin
  var Pattern := '<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]*hyperlink[^"]*"[^>]+Target="([^"]+)"';
  var Matches := TRegEx.Matches(XmlContent, Pattern, [roIgnoreCase, roSingleLine]);
  for var Match in Matches do
  begin
    if Match.Groups.Count > 2 then
    begin
      var RelId := Match.Groups[1].Value;
      var Target := Match.Groups[2].Value;
      Hyperlinks.AddOrSetValue(RelId, Target);
    end;
  end;

  var HeaderPattern := '<Relationship[^>]+Type="[^"]*header[^"]*"[^>]+Target="([^"]+)"';
  var HeaderMatch := TRegEx.Match(XmlContent, HeaderPattern, [roIgnoreCase, roSingleLine]);
  if HeaderMatch.Success and (HeaderMatch.Groups.Count > 1) then
    Hyperlinks.AddOrSetValue(KeyHeader, HeaderMatch.Groups[1].Value);

  var FooterPattern := '<Relationship[^>]+Type="[^"]*footer[^"]*"[^>]+Target="([^"]+)"';
  var FooterMatch := TRegEx.Match(XmlContent, FooterPattern, [roIgnoreCase, roSingleLine]);
  if FooterMatch.Success and (FooterMatch.Groups.Count > 1) then
    Hyperlinks.AddOrSetValue(KeyFooter, FooterMatch.Groups[1].Value);
end;

function TWordDocument.HasListParagraphs: Boolean;
begin
  for var Para in FParagraphs do
    if Para.ListStyle <> TListStyle.None then
      Exit(True);
  Result := False;
end;

function TWordDocument.GenerateNumberingXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:numbering xmlns:w="' + WordprocessingNs + '">' +
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0">' +
    '<w:start w:val="1"/>' +
    '<w:numFmt w:val="bullet"/>' +
    '<w:lvlText w:val=""/>' +
    '<w:lvlJc w:val="left"/>' +
    '</w:lvl>' +
    '</w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1">' +
    '<w:lvl w:ilvl="0">' +
    '<w:start w:val="1"/>' +
    '<w:numFmt w:val="decimal"/>' +
    '<w:lvlText w:val="%1."/>' +
    '<w:lvlJc w:val="left"/>' +
    '</w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '</w:numbering>';
end;

function TWordDocument.GenerateContentTypesXml: string;
begin
  var NumberingOverride := '';
  if HasListParagraphs then
    NumberingOverride := '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';

  var HeaderOverride := '';
  if HasHeader then
    HeaderOverride := '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';

  var FooterOverride := '';
  if HasFooter then
    FooterOverride := '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';

  var ImageDefaults := '';
  var Images := CollectImages;
  try
    var Extensions := TList<string>.Create;
    try
      for var ImgRun in Images do
      begin
        var Ext := LowerCase(ImgRun.GetImage.Extension);
        if not Extensions.Contains(Ext) then
        begin
          Extensions.Add(Ext);
          ImageDefaults := ImageDefaults + '<Default Extension="' + Ext + '" ContentType="' + GetImageContentType(Ext) + '"/>';
        end;
      end;
    finally
      Extensions.Free;
    end;
  finally
    Images.Free;
  end;

  Result :=
    XmlDeclaration + sLineBreak +
    '<Types xmlns="' + ContentTypesNs + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    ImageDefaults +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    NumberingOverride +
    HeaderOverride +
    FooterOverride +
    '</Types>';
end;

function TWordDocument.GenerateRootRelsXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';
end;

function TWordDocument.GenerateDocumentRelsXml: string;
begin
  var Hyperlinks := CollectHyperlinks;
  var Images := CollectImages;
  try
    var RelsContent := '';
    var NextId := 1;

    if HasListParagraphs then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" ' +
        'Target="numbering.xml"/>';
      Inc(NextId);
    end;

    if HasHeader then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ' +
        'Target="header1.xml"/>';
      Inc(NextId);
    end;

    if HasFooter then
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" ' +
        'Target="footer1.xml"/>';
      Inc(NextId);
    end;

    for var I := 0 to Hyperlinks.Count - 1 do
    begin
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId + I) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
        'Target="' + TXml.Escape(Hyperlinks[I]) + '" TargetMode="External"/>';
    end;
    NextId := NextId + Hyperlinks.Count;

    for var I := 0 to Images.Count - 1 do
    begin
      var ImgRun := Images[I];
      var Ext := LowerCase(ImgRun.GetImage.Extension);
      RelsContent := RelsContent +
        '<Relationship Id="rId' + IntToStr(NextId + I) + '" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="media/image' + IntToStr(I + 1) + '.' + Ext + '"/>';
    end;

    Result :=
      XmlDeclaration + sLineBreak +
      '<Relationships xmlns="' + RelationshipsNs + '">' +
      RelsContent +
      '</Relationships>';
  finally
    Hyperlinks.Free;
    Images.Free;
  end;
end;

function TWordDocument.GenerateDocumentXml: string;
begin
  var Hyperlinks := CollectHyperlinks;
  var Images := CollectImages;
  try
    var Body := '';
    var HyperlinkIdOffset := 0;
    if HasListParagraphs then
      Inc(HyperlinkIdOffset);
    if HasHeader then
      Inc(HyperlinkIdOffset);
    if HasFooter then
      Inc(HyperlinkIdOffset);

    var ImageIdOffset := HyperlinkIdOffset + Hyperlinks.Count;
    var ImageIndex := 0;

    for var I := 0 to FParagraphs.Count - 1 do
    begin
      var Para := FParagraphs[I];
      Body := Body + '<w:p>';

      const HasSpacing = (Para.LineSpacing.Line > 0) or (Para.LineSpacing.Before > 0) or (Para.LineSpacing.After > 0);
      const HasIndent = (Para.Indent.Left > 0) or (Para.Indent.Right > 0) or (Para.Indent.FirstLine <> 0);
      const NeedPPr = (Para.ListStyle <> TListStyle.None) or (Para.Alignment <> TParagraphAlignment.Left) or HasSpacing or HasIndent;
      if NeedPPr then
      begin
        Body := Body + '<w:pPr>';
        if Para.ListStyle <> TListStyle.None then
        begin
          var NumId := '1';
          if Para.ListStyle = TListStyle.Numbered then
            NumId := '2';
          Body := Body + '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + NumId + '"/></w:numPr>';
        end;
        if HasSpacing then
        begin
          var SpacingAttrs := '';
          if Para.LineSpacing.Before > 0 then
            SpacingAttrs := SpacingAttrs + ' w:before="' + IntToStr(Para.LineSpacing.Before) + '"';
          if Para.LineSpacing.After > 0 then
            SpacingAttrs := SpacingAttrs + ' w:after="' + IntToStr(Para.LineSpacing.After) + '"';
          if Para.LineSpacing.Line > 0 then
          begin
            SpacingAttrs := SpacingAttrs + ' w:line="' + IntToStr(Para.LineSpacing.Line) + '"';
            var RuleValue := 'auto';
            case Para.LineSpacing.Rule of
              TLineSpacingRule.Exact: RuleValue := 'exact';
              TLineSpacingRule.AtLeast: RuleValue := 'atLeast';
            end;
            SpacingAttrs := SpacingAttrs + ' w:lineRule="' + RuleValue + '"';
          end;
          Body := Body + '<w:spacing' + SpacingAttrs + '/>';
        end;
        if HasIndent then
        begin
          var IndentAttrs := '';
          if Para.Indent.Left > 0 then
            IndentAttrs := IndentAttrs + ' w:left="' + IntToStr(Para.Indent.Left) + '"';
          if Para.Indent.Right > 0 then
            IndentAttrs := IndentAttrs + ' w:right="' + IntToStr(Para.Indent.Right) + '"';
          if Para.Indent.FirstLine > 0 then
            IndentAttrs := IndentAttrs + ' w:firstLine="' + IntToStr(Para.Indent.FirstLine) + '"'
          else if Para.Indent.FirstLine < 0 then
            IndentAttrs := IndentAttrs + ' w:hanging="' + IntToStr(-Para.Indent.FirstLine) + '"';
          Body := Body + '<w:ind' + IndentAttrs + '/>';
        end;
        if Para.Alignment <> TParagraphAlignment.Left then
        begin
          var AlignValue := 'left';
          case Para.Alignment of
            TParagraphAlignment.Center: AlignValue := 'center';
            TParagraphAlignment.Right: AlignValue := 'right';
            TParagraphAlignment.Justify: AlignValue := 'both';
          end;
          Body := Body + '<w:jc w:val="' + AlignValue + '"/>';
        end;
        Body := Body + '</w:pPr>';
      end;

      for var J := 0 to Para.RunCount - 1 do
      begin
        var Run := Para.Runs[J];
        var WordRun := Run as TWordRun;
        var RunText := Run.Text;
        var HasHyperlink := Run.Hyperlink <> '';

        if WordRun.HasImage then
        begin
          Inc(ImageIndex);
          Body := Body + '<w:r>' + GenerateDrawingXml(WordRun, ImageIndex, ImageIdOffset) + '</w:r>';
        end
        else if RunText = sLineBreak then
          Body := Body + '<w:r><w:br/></w:r>'
        else if RunText = #9 then
          Body := Body + '<w:r><w:tab/></w:r>'
        else if RunText = #12 then
          Body := Body + '<w:r><w:br w:type="page"/></w:r>'
        else
        begin
          if HasHyperlink then
            Body := Body + '<w:hyperlink r:id="rId' + IntToStr(GetHyperlinkId(Run.Hyperlink, Hyperlinks) + HyperlinkIdOffset) + '">';

          Body := Body + '<w:r>';
          const NeedRPr = Run.Bold or Run.Italic or Run.Underline or
                         (Run.FontName <> '') or (Run.FontSize > 0) or (Run.FontColor <> '');
          if NeedRPr then
          begin
            Body := Body + '<w:rPr>';
            if Run.FontName <> '' then
              Body := Body + '<w:rFonts w:ascii="' + TXml.Escape(Run.FontName) + '" w:hAnsi="' + TXml.Escape(Run.FontName) + '"/>';
            if Run.Bold then
              Body := Body + '<w:b/>';
            if Run.Italic then
              Body := Body + '<w:i/>';
            if Run.Underline then
              Body := Body + '<w:u w:val="single"/>';
            if Run.FontSize > 0 then
              Body := Body + '<w:sz w:val="' + IntToStr(Run.FontSize) + '"/>';
            if Run.FontColor <> '' then
              Body := Body + '<w:color w:val="' + TXml.Escape(Run.FontColor) + '"/>';
            Body := Body + '</w:rPr>';
          end;
          Body := Body + '<w:t xml:space="preserve">' + TXml.Escape(RunText) + '</w:t></w:r>';

          if HasHyperlink then
            Body := Body + '</w:hyperlink>';
        end;
      end;

      Body := Body + '</w:p>';
    end;

    for var I := 0 to FTables.Count - 1 do
    begin
      var Table := FTables[I];
      var WordTable := Table as TWordTable;
      Body := Body + '<w:tbl>';

      Body := Body + '<w:tblPr>';
      Body := Body + '<w:tblW w:w="0" w:type="auto"/>';

      var Borders := WordTable.FBorders;
      const HasBorders = (Borders.Top.Style <> TBorderStyle.None) or
                         (Borders.Bottom.Style <> TBorderStyle.None) or
                         (Borders.Left.Style <> TBorderStyle.None) or
                         (Borders.Right.Style <> TBorderStyle.None) or
                         (Borders.InsideH.Style <> TBorderStyle.None) or
                         (Borders.InsideV.Style <> TBorderStyle.None);
      if HasBorders then
      begin
        Body := Body + '<w:tblBorders>';
        if Borders.Top.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:top', Borders.Top);
        if Borders.Left.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:left', Borders.Left);
        if Borders.Bottom.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:bottom', Borders.Bottom);
        if Borders.Right.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:right', Borders.Right);
        if Borders.InsideH.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:insideH', Borders.InsideH);
        if Borders.InsideV.Style <> TBorderStyle.None then
          Body := Body + GenerateBorderXml('w:insideV', Borders.InsideV);
        Body := Body + '</w:tblBorders>';
      end;
      Body := Body + '</w:tblPr>';

      if Length(WordTable.FColumnWidths) > 0 then
      begin
        Body := Body + '<w:tblGrid>';
        for var ColW in WordTable.FColumnWidths do
          Body := Body + '<w:gridCol w:w="' + IntToStr(ColW) + '"/>';
        Body := Body + '</w:tblGrid>';
      end;

      for var R := 0 to Table.RowCount - 1 do
      begin
        Body := Body + '<w:tr>';
        for var C := 0 to Table.ColCount - 1 do
        begin
          var Cell := Table.Cells[R, C];
          var WordCell := Cell as TWordTableCell;
          Body := Body + '<w:tc>';

          const HasCellPr = (WordCell.FShading <> '') or (WordCell.FWidth > 0);
          if HasCellPr then
          begin
            Body := Body + '<w:tcPr>';
            if WordCell.FWidth > 0 then
              Body := Body + '<w:tcW w:w="' + IntToStr(WordCell.FWidth) + '" w:type="dxa"/>';
            if WordCell.FShading <> '' then
              Body := Body + '<w:shd w:val="clear" w:fill="' + WordCell.FShading + '"/>';
            Body := Body + '</w:tcPr>';
          end;

          Body := Body + '<w:p><w:r><w:t>' + TXml.Escape(Cell.Text) + '</w:t></w:r></w:p>';
          Body := Body + '</w:tc>';
        end;
        Body := Body + '</w:tr>';
      end;

      Body := Body + '</w:tbl>';
    end;

    var SectPr := '';
    var NeedSectPr := (FPageOrientation = TPageOrientation.Landscape) or
       (FPageMargins.Top > 0) or (FPageMargins.Bottom > 0) or
       (FPageMargins.Left > 0) or (FPageMargins.Right > 0) or
       HasHeader or HasFooter;

    if NeedSectPr then
    begin
      SectPr := '<w:sectPr>';

      var HeaderFooterRelId := 1;
      if HasListParagraphs then
        Inc(HeaderFooterRelId);

      if HasHeader then
      begin
        SectPr := SectPr + '<w:headerReference w:type="default" r:id="rId' + IntToStr(HeaderFooterRelId) + '"/>';
        Inc(HeaderFooterRelId);
      end;

      if HasFooter then
        SectPr := SectPr + '<w:footerReference w:type="default" r:id="rId' + IntToStr(HeaderFooterRelId) + '"/>';

      if FPageOrientation = TPageOrientation.Landscape then
        SectPr := SectPr + '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
      else if NeedSectPr then
        SectPr := SectPr + '<w:pgSz w:w="12240" w:h="15840"/>';

      if (FPageMargins.Top > 0) or (FPageMargins.Bottom > 0) or
         (FPageMargins.Left > 0) or (FPageMargins.Right > 0) then
      begin
        SectPr := SectPr + '<w:pgMar w:top="' + IntToStr(FPageMargins.Top) + '" ' +
          'w:bottom="' + IntToStr(FPageMargins.Bottom) + '" ' +
          'w:left="' + IntToStr(FPageMargins.Left) + '" ' +
          'w:right="' + IntToStr(FPageMargins.Right) + '"/>';
      end;
      SectPr := SectPr + '</w:sectPr>';
    end;

    Result :=
      XmlDeclaration + sLineBreak +
      '<w:document xmlns:w="' + WordprocessingNs + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<w:body>' + Body + SectPr + '</w:body>' +
      '</w:document>';
  finally
    Hyperlinks.Free;
    Images.Free;
  end;
end;

procedure TWordDocument.SaveToFile(const FileName: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    WriteParts(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TWordDocument.SaveToStream(const Stream: TStream);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(Stream, zmWrite);
    WriteParts(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TWordDocument.WriteParts(const Zip: TZipFile);
begin
  var ContentTypesXml := GenerateContentTypesXml;
  var RelsXml := GenerateRootRelsXml;
  var DocumentXml := GenerateDocumentXml;
  var DocumentRelsXml := GenerateDocumentRelsXml;

  Zip.Add(TEncoding.UTF8.GetBytes(ContentTypesXml), '[Content_Types].xml');
  Zip.Add(TEncoding.UTF8.GetBytes(RelsXml), PartRootRels);
  Zip.Add(TEncoding.UTF8.GetBytes(DocumentXml), 'word/document.xml');
  Zip.Add(TEncoding.UTF8.GetBytes(DocumentRelsXml), PartDocumentRels);

  if HasListParagraphs then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateNumberingXml), 'word/numbering.xml');

  if HasHeader then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateHeaderXml), 'word/header1.xml');

  if HasFooter then
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateFooterXml), 'word/footer1.xml');

  var Images := CollectImages;
  try
    for var Idx := 0 to Images.Count - 1 do
    begin
      var Img := Images[Idx].GetImage;
      var MediaPath := 'word/media/image' + IntToStr(Idx + 1) + '.' + LowerCase(Img.Extension);
      Zip.Add(Img.Data, MediaPath);
    end;
  finally
    Images.Free;
  end;
end;

function TWordDocument.GetText: string;
begin
  var Lines := TStringList.Create;
  try
    for var Para in FParagraphs do
      Lines.Add(Para.Text);
    Result := Lines.Text;
  finally
    Lines.Free;
  end;
end;

function TWordDocument.GetMetadata: TDocumentMetadata;
begin
  Result := FMetadata;
end;

function TWordDocument.GetParagraphCount: Integer;
begin
  Result := FParagraphs.Count;
end;

function TWordDocument.GetParagraph(Index: Integer): IWordParagraph;
begin
  Result := FParagraphs[Index];
end;

function TWordDocument.AddParagraph: IWordParagraph;
begin
  var Para := TWordParagraph.Create;
  FParagraphs.Add(Para);
  Result := Para;
end;

function TWordDocument.GetTableCount: Integer;
begin
  Result := FTables.Count;
end;

function TWordDocument.GetTable(Index: Integer): IWordTable;
begin
  Result := FTables[Index];
end;

function TWordDocument.AddTable(const Rows, Cols: Integer): IWordTable;
begin
  var Table := TWordTable.Create(Rows, Cols);
  FTables.Add(Table);
  Result := Table;
end;

function TWordDocument.GetPageOrientation: TPageOrientation;
begin
  Result := FPageOrientation;
end;

procedure TWordDocument.SetPageOrientation(const Value: TPageOrientation);
begin
  FPageOrientation := Value;
end;

function TWordDocument.GetPageMargins: TPageMargins;
begin
  Result := FPageMargins;
end;

procedure TWordDocument.SetPageMargins(const Value: TPageMargins);
begin
  FPageMargins := Value;
end;

function TWordDocument.GetHeader: IWordHeaderFooter;
begin
  Result := FHeader;
end;

function TWordDocument.GetFooter: IWordHeaderFooter;
begin
  Result := FFooter;
end;

function TWordDocument.HasHeader: Boolean;
begin
  Result := FHeader.Text <> '';
end;

function TWordDocument.HasFooter: Boolean;
begin
  Result := FFooter.Text <> '';
end;

function TWordDocument.GenerateBorderXml(const ElementName: string; const Border: TTableBorder): string;
begin
  var StyleValue := 'single';
  case Border.Style of
    TBorderStyle.Single: StyleValue := 'single';
    TBorderStyle.Double: StyleValue := 'double';
    TBorderStyle.Dashed: StyleValue := 'dashed';
    TBorderStyle.Dotted: StyleValue := 'dotted';
    TBorderStyle.Thick: StyleValue := 'thick';
  end;
  Result := '<' + ElementName + ' w:val="' + StyleValue + '" w:sz="' + IntToStr(Border.Width) + '"';
  if Border.Color <> '' then
    Result := Result + ' w:color="' + Border.Color + '"';
  Result := Result + '/>';
end;

function TWordDocument.CollectImages: TList<TWordRun>;
begin
  Result := TList<TWordRun>.Create;
  for var Para in FParagraphs do
  begin
    for var RunIndex := 0 to Para.RunCount - 1 do
    begin
      var Run := Para.Runs[RunIndex];
      var WordRun := Run as TWordRun;
      if WordRun.HasImage then
        Result.Add(WordRun);
    end;
  end;
end;

function TWordDocument.GetImageContentType(const Extension: string): string;
begin
  var Ext := LowerCase(Extension);
  if (Ext = 'jpg') or (Ext = 'jpeg') then
    Result := 'image/jpeg'
  else if Ext = 'png' then
    Result := 'image/png'
  else if Ext = 'gif' then
    Result := 'image/gif'
  else if Ext = 'bmp' then
    Result := 'image/bmp'
  else
    Result := 'image/png';
end;

function TWordDocument.GenerateDrawingXml(const ImageRun: TWordRun; ImageIndex: Integer; RelIdOffset: Integer): string;
begin
  var Img := ImageRun.GetImage;
  var RelId := 'rId' + IntToStr(ImageIndex + RelIdOffset);
  var ImgName := 'Image ' + IntToStr(ImageIndex);
  var CxStr := IntToStr(Img.Width);
  var CyStr := IntToStr(Img.Height);

  Result :=
    '<w:drawing>' +
      '<wp:inline xmlns:wp="' + WpDrawingNs + '" distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="' + CxStr + '" cy="' + CyStr + '"/>' +
        '<wp:docPr id="' + IntToStr(ImageIndex) + '" name="' + ImgName + '"/>' +
        '<a:graphic xmlns:a="' + DrawingNs + '">' +
          '<a:graphicData uri="' + PictureNs + '">' +
            '<pic:pic xmlns:pic="' + PictureNs + '">' +
              '<pic:nvPicPr>' +
                '<pic:cNvPr id="' + IntToStr(ImageIndex) + '" name="' + ImgName + '"/>' +
                '<pic:cNvPicPr/>' +
              '</pic:nvPicPr>' +
              '<pic:blipFill>' +
                '<a:blip xmlns:r="' + RelationshipsDocNs + '" r:embed="' + RelId + '"/>' +
                '<a:stretch><a:fillRect/></a:stretch>' +
              '</pic:blipFill>' +
              '<pic:spPr>' +
                '<a:xfrm>' +
                  '<a:off x="0" y="0"/>' +
                  '<a:ext cx="' + CxStr + '" cy="' + CyStr + '"/>' +
                '</a:xfrm>' +
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
              '</pic:spPr>' +
            '</pic:pic>' +
          '</a:graphicData>' +
        '</a:graphic>' +
      '</wp:inline>' +
    '</w:drawing>';
end;

function TWordDocument.GenerateHeaderXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:hdr xmlns:w="' + WordprocessingNs + '">' +
    '<w:p><w:r><w:t>' + TXml.Escape(FHeader.Text) + '</w:t></w:r></w:p>' +
    '</w:hdr>';
end;

function TWordDocument.GenerateFooterXml: string;
begin
  Result :=
    XmlDeclaration + sLineBreak +
    '<w:ftr xmlns:w="' + WordprocessingNs + '">' +
    '<w:p><w:r><w:t>' + TXml.Escape(FFooter.Text) + '</w:t></w:r></w:p>' +
    '</w:ftr>';
end;

procedure TWordDocument.ParseHeaderFooterXml(const XmlContent: string; const Target: IWordHeaderFooter);
begin
  var Builder := TStringBuilder.Create;
  try
    for var Match in TRegEx.Matches(XmlContent, TextElementPattern, [roIgnoreCase]) do
      Builder.Append(TXml.Unescape(Match.Groups[1].Value));

    Target.Text := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

end.
