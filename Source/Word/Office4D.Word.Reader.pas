unit Office4D.Word.Reader;

interface

uses
  System.Generics.Collections,
  Office4D.Word,
  Office4D.Package,
  Office4D.Word.Model,
  Office4D.Xml;

type
  /// <summary>
  /// Fills a TWordDocumentContent from an opened .docx package. Reading is
  /// structural: elements are located with TXml.FindElements, which knows about
  /// nesting, so a w:p inside a shape's text box is a paragraph of its own and
  /// not the end of the paragraph hosting the shape.
  /// </summary>
  TWordDocumentReader = class
  private
    FPackage: TOXMLPackage;
    FContent: TWordDocumentContent;

    procedure ReadHeaderFooter(const HyperlinkMap: TDictionary<string, string>);
    procedure ReadDocument(const HyperlinkMap: TDictionary<string, string>);
    procedure ReadMetadata;

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
    procedure AddFirstPartPath(const Hyperlinks: TDictionary<string, string>; const Key, Target: string);
    procedure ParseHeaderFooterXml(const XmlContent: string; const Target: IWordHeaderFooter);
  public
    constructor Create(const Package: TOXMLPackage; const Content: TWordDocumentContent);

    procedure Read;
  end;

implementation

uses
  System.SysUtils,
  System.RegularExpressions,
  Office4D.Metadata,
  Office4D.Relationships,
  Office4D.Types;

const
  PartDocument = 'word/document.xml';
  PartDocumentRels = 'word/_rels/document.xml.rels';
  KeyHeader = '__header__';
  KeyFooter = '__footer__';

  ElementParagraph = 'w:p';
  ElementRun = 'w:r';
  ElementHyperlink = 'w:hyperlink';
  ElementNumId = 'w:numId';
  AttributeVal = 'w:val';
  AttributeRelationshipId = 'r:id';

  RelTypeSuffixHyperlink = '/hyperlink';
  RelTypeSuffixHeader = '/header';
  RelTypeSuffixFooter = '/footer';

  TextElementPattern = '<w:t(?:\s[^>]*)?>([^<]*)</w:t>';

  /// Matches a w:t element with its text, or a standalone w:br or w:tab, so the
  /// children of a run can be walked in document order.
  RunContentPattern = TextElementPattern + '|<w:t\s*/>|<w:br(?:\s[^>]*)?/?>|<w:tab(?:\s[^>]*)?/?>';

{ TWordDocumentReader }

constructor TWordDocumentReader.Create(const Package: TOXMLPackage; const Content: TWordDocumentContent);
begin
  inherited Create;
  FPackage := Package;
  FContent := Content;
end;

procedure TWordDocumentReader.Read;
begin
  var HyperlinkMap := TDictionary<string, string>.Create;
  try
    if FPackage.PartExists(PartDocumentRels) then
    begin
      ParseDocumentRels(FPackage.GetPartContent(PartDocumentRels), HyperlinkMap);
      ReadHeaderFooter(HyperlinkMap);
    end;

    ReadDocument(HyperlinkMap);
  finally
    HyperlinkMap.Free;
  end;

  ReadMetadata;
end;

procedure TWordDocumentReader.ReadHeaderFooter(const HyperlinkMap: TDictionary<string, string>);
begin
  if HyperlinkMap.ContainsKey(KeyHeader) then
  begin
    const HeaderPath = HyperlinkMap[KeyHeader];
    if FPackage.PartExists(HeaderPath) then
      ParseHeaderFooterXml(FPackage.GetPartContent(HeaderPath), FContent.Header);
  end;

  if HyperlinkMap.ContainsKey(KeyFooter) then
  begin
    const FooterPath = HyperlinkMap[KeyFooter];
    if FPackage.PartExists(FooterPath) then
      ParseHeaderFooterXml(FPackage.GetPartContent(FooterPath), FContent.Footer);
  end;
end;

procedure TWordDocumentReader.ReadDocument(const HyperlinkMap: TDictionary<string, string>);
begin
  var Rels := TRelationships.Create;
  try
    Rels.LoadFromXml(FPackage.GetPartContent(PartRootRels));
    const DocumentTarget = Rels.GetTargetByType(RelTypeOfficeDocument);
    if DocumentTarget.IsEmpty then
      Exit;

    const DocumentPath = TRelationships.ResolveTarget('', DocumentTarget);
    const DocumentXml = FPackage.GetPartContent(DocumentPath);
    ParseDocumentXml(DocumentXml, HyperlinkMap);
  finally
    Rels.Free;
  end;
end;

procedure TWordDocumentReader.ReadMetadata;
begin
  FContent.Metadata := TMetadataParser.ParsePackage(FPackage);
end;

procedure TWordDocumentReader.ParseDocumentXml(const XmlContent: string;
  const HyperlinkMap: TDictionary<string, string>);
begin
  FContent.Paragraphs.Clear;

  ParseParagraphs(TXml.ReduceAlternateContent(XmlContent), HyperlinkMap);
end;

procedure TWordDocumentReader.ParseParagraphs(const Xml: string; const HyperlinkMap: TDictionary<string, string>);
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

procedure TWordDocumentReader.ParseParagraph(const ParagraphXml: string;
  const HyperlinkMap: TDictionary<string, string>);
begin
  var Para := TWordParagraph.Create;

  ApplyListStyle(Para, ParagraphXml);
  ParseParagraphContent(Para, ParagraphXml, HyperlinkMap);

  if Para.GetRunCount > 0 then
    FContent.Paragraphs.Add(Para)
  else
    Para.Free;
end;

procedure TWordDocumentReader.ApplyListStyle(const Para: TWordParagraph; const ParagraphXml: string);
begin
  var NumIdElement: TXmlElement;
  if not TXml.TryFindElement(ParagraphXml, ElementNumId, NumIdElement) then
    Exit;

  var NumId := StrToIntDef(NumIdElement.Attribute(AttributeVal), 0);
  if NumId = 1 then
    Para.SetListStyle(TListStyle.Bullet)
  else if NumId = 2 then
    Para.SetListStyle(TListStyle.Numbered);
end;

procedure TWordDocumentReader.ParseParagraphContent(const Para: TWordParagraph; const ParagraphXml: string;
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

procedure TWordDocumentReader.ParseHyperlink(const Para: TWordParagraph; const Hyperlink: TXmlElement;
  const HyperlinkMap: TDictionary<string, string>);
begin
  var Url := '';

  const RelationshipId = Hyperlink.Attribute(AttributeRelationshipId);
  if not RelationshipId.IsEmpty then
    HyperlinkMap.TryGetValue(RelationshipId, Url);

  for var Run in TXml.FindElements(Hyperlink.Inner, ElementRun) do
    ParseRun(Para, Run.Inner, Url);
end;

procedure TWordDocumentReader.ParseRun(const Para: TWordParagraph; const RunXml, Hyperlink: string);
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

procedure TWordDocumentReader.ParseDocumentRels(const XmlContent: string;
  const Hyperlinks: TDictionary<string, string>);
begin
  var Rels := TRelationships.Create;
  try
    Rels.LoadFromXml(XmlContent);

    for var Relationship in Rels.Relationships do
    begin
      if Relationship.RelType.EndsWith(RelTypeSuffixHyperlink) then
        Hyperlinks.AddOrSetValue(Relationship.Id, Relationship.Target)
      else if Relationship.RelType.EndsWith(RelTypeSuffixHeader) then
        AddFirstPartPath(Hyperlinks, KeyHeader, Relationship.Target)
      else if Relationship.RelType.EndsWith(RelTypeSuffixFooter) then
        AddFirstPartPath(Hyperlinks, KeyFooter, Relationship.Target);
    end;
  finally
    Rels.Free;
  end;
end;

// Only the first header and footer are read. Their targets are relative to the
// document part or absolute from the package root, depending on the writer.
procedure TWordDocumentReader.AddFirstPartPath(const Hyperlinks: TDictionary<string, string>;
                                               const Key, Target: string);
begin
  if not Hyperlinks.ContainsKey(Key) then
    Hyperlinks.Add(Key, TRelationships.ResolveTarget(PartDocument, Target));
end;

procedure TWordDocumentReader.ParseHeaderFooterXml(const XmlContent: string; const Target: IWordHeaderFooter);
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
