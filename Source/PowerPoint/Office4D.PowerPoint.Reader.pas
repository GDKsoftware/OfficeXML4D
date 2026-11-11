unit Office4D.PowerPoint.Reader;

interface

uses
  Office4D.PowerPoint,
  Office4D.Package,
  Office4D.PowerPoint.Model,
  Office4D.Xml;

type
  /// <summary>
  /// Fills a TPowerPointContent from an opened .pptx package. Slides are read
  /// in the order the sldIdLst of presentation.xml states, not in part-name
  /// order.
  /// </summary>
  TPowerPointReader = class
  private
    FPackage: TOXMLPackage;
    FContent: TPowerPointContent;

    procedure ParseSlideXml(const Xml: string);
    function IsTitleShape(const ShapeXml: string): Boolean;
    function TextOf(const Xml: string): string;
    procedure ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);
    procedure ApplyRunProperties(const Run: IPowerPointRun; const RunProperties: TXmlElement);
  public
    constructor Create(const Package: TOXMLPackage; const Content: TPowerPointContent);

    procedure ParsePackage;
  end;

implementation

uses
  System.SysUtils,
  Office4D.Metadata,
  Office4D.Relationships,
  Office4D.Types;

const
  ElementSlideId = 'p:sldId';
  ElementShape = 'p:sp';
  ElementPlaceholder = 'p:ph';
  ElementParagraph = 'a:p';
  ElementParagraphProperties = 'a:pPr';
  ElementRun = 'a:r';
  ElementRunProperties = 'a:rPr';
  ElementText = 'a:t';
  ElementOutline = 'a:ln';
  ElementSolidFill = 'a:solidFill';
  ElementRgbColor = 'a:srgbClr';
  ElementLatinFont = 'a:latin';

  AttributeRelationshipId = 'r:id';

  PlaceholderTitle = 'title';
  PlaceholderCenteredTitle = 'ctrTitle';
  UnderlineSingle = 'sng';
  HundredthsPerPoint = 100;

{ TPowerPointReader }

constructor TPowerPointReader.Create(const Package: TOXMLPackage; const Content: TPowerPointContent);
begin
  inherited Create;
  FPackage := Package;
  FContent := Content;
end;

procedure TPowerPointReader.ParsePackage;
begin
  var RootRelsXml := FPackage.GetPartContent(PartRootRels);
  var RootRels := TRelationships.Create;
  try
    RootRels.LoadFromXml(RootRelsXml);
    const PresentationTarget = RootRels.GetTargetByType(RelTypeOfficeDocument);
    if PresentationTarget = '' then
      Exit;

    const PresentationPath = TRelationships.ResolveTarget('', PresentationTarget);
    const PresentationXml = FPackage.GetPartContent(PresentationPath);

    const RelsPath = TRelationships.RelationshipsPartOf(PresentationPath);
    if not FPackage.PartExists(RelsPath) then
      Exit;

    var PresRels := TRelationships.Create;
    try
      PresRels.LoadFromXml(FPackage.GetPartContent(RelsPath));

      // Slides are ordered by the sldIdLst in presentation.xml, not by part name.
      for var SlideId in TXml.FindElements(PresentationXml, ElementSlideId) do
      begin
        const RelId = SlideId.Attribute(AttributeRelationshipId);
        if RelId.IsEmpty then
          Continue;

        const SlideRelationship = PresRels.GetById(RelId);
        const SlidePath = TRelationships.ResolveTarget(PresentationPath, SlideRelationship.Target);
        const SlideXml = FPackage.GetPartContent(SlidePath);
        ParseSlideXml(SlideXml);
      end;
    finally
      PresRels.Free;
    end;
  finally
    RootRels.Free;
  end;

  FContent.Metadata := TMetadataParser.ParsePackage(FPackage);
end;

procedure TPowerPointReader.ParseSlideXml(const Xml: string);
begin
  const Slide = FContent.AddSlide;

  // A slide states shapes it cannot render everywhere twice, once per
  // mc:AlternateContent branch. Only the branch a consumer should read is kept,
  // so the shapes below are the shapes of the slide, each of them once.
  const SlideXml = TXml.ReduceAlternateContent(Xml);

  for var Shape in TXml.FindElements(SlideXml, ElementShape) do
  begin
    if IsTitleShape(Shape.Inner) then
    begin
      Slide.Title := TextOf(Shape.Inner);
    end
    else
    begin
      for var Paragraph in TXml.FindElements(Shape.Inner, ElementParagraph) do
      begin
        ParseParagraphXml(Paragraph.Inner, Slide);
      end;
    end;
  end;
end;

function TPowerPointReader.IsTitleShape(const ShapeXml: string): Boolean;
begin
  var Placeholder: TXmlElement;
  if not TXml.TryFindElement(ShapeXml, ElementPlaceholder, Placeholder) then
    Exit(False);

  const PlaceholderType = Placeholder.Attribute('type');
  Result := (PlaceholderType = PlaceholderTitle) or (PlaceholderType = PlaceholderCenteredTitle);
end;

function TPowerPointReader.TextOf(const Xml: string): string;
begin
  Result := '';
  for var TextElement in TXml.FindElements(Xml, ElementText) do
  begin
    const Text = TXml.Unescape(TextElement.Inner);
    Result := Result + Text;
  end;
end;

procedure TPowerPointReader.ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);
begin
  const Paragraph = Slide.AddParagraph;

  Paragraph.Bullet := (Pos('<a:buChar', Xml) > 0) or (Pos('<a:buAutoNum', Xml) > 0);

  var ParagraphProperties: TXmlElement;
  if TXml.TryFindElement(Xml, ElementParagraphProperties, ParagraphProperties) then
    Paragraph.IndentLevel := StrToIntDef(ParagraphProperties.Attribute('lvl'), 0);

  for var RunElement in TXml.FindElements(Xml, ElementRun) do
  begin
    var Text: TXmlElement;
    if not TXml.TryFindElement(RunElement.Inner, ElementText, Text) then
      Continue;

    const Run = Paragraph.AddRun(TXml.Unescape(Text.Inner));

    var RunProperties: TXmlElement;
    if TXml.TryFindElement(RunElement.Inner, ElementRunProperties, RunProperties) then
      ApplyRunProperties(Run, RunProperties);
  end;
end;

procedure TPowerPointReader.ApplyRunProperties(const Run: IPowerPointRun; const RunProperties: TXmlElement);
begin
  Run.Bold := TXml.BooleanOf(RunProperties.Attribute('b'), False);
  Run.Italic := TXml.BooleanOf(RunProperties.Attribute('i'), False);
  Run.Underline := (RunProperties.Attribute('u') = UnderlineSingle);

  var Size: string;
  if TXml.TryGetAttribute(RunProperties.OpenTag, 'sz', Size) then
    Run.FontSize := StrToIntDef(Size, 0) div HundredthsPerPoint;

  // An outline (a:ln) carries a solidFill of its own, which is not the text colour.
  const Outlines = TXml.FindElements(RunProperties.Inner, ElementOutline);
  const TextProperties = TXml.RemoveElements(RunProperties.Inner, Outlines);

  var Fill: TXmlElement;
  var Color: TXmlElement;
  const HasRgbColor = (TXml.TryFindElement(TextProperties, ElementSolidFill, Fill) and
                       TXml.TryFindElement(Fill.Inner, ElementRgbColor, Color));
  if HasRgbColor then
    Run.FontColor := Color.Attribute('val');

  var LatinFont: TXmlElement;
  if TXml.TryFindElement(TextProperties, ElementLatinFont, LatinFont) then
    Run.FontName := LatinFont.Attribute('typeface');
end;

end.
