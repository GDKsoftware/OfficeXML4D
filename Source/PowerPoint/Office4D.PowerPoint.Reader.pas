unit Office4D.PowerPoint.Reader;

interface

uses
  Office4D.PowerPoint,
  Office4D.Package,
  Office4D.PowerPoint.Model;

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
    procedure ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);
  public
    constructor Create(const Package: TOXMLPackage; const Content: TPowerPointContent);

    procedure ParsePackage;
  end;

implementation

uses
  System.SysUtils,
  System.RegularExpressions,
  Office4D.Metadata,
  Office4D.Relationships,
  Office4D.Types,
  Office4D.Xml;

const

  PartPresentationRels = 'ppt/_rels/presentation.xml.rels';
  PartPptPrefix = 'ppt/';

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
    var PresentationPath := RootRels.GetTargetByType(RelTypeOfficeDocument);
    if PresentationPath = '' then
      Exit;
    if PresentationPath.StartsWith('/') then
      PresentationPath := PresentationPath.Substring(1);

    const PresentationXml = FPackage.GetPartContent(PresentationPath);

    const RelsPath = PartPresentationRels;
    if not FPackage.PartExists(RelsPath) then
      Exit;

    var PresRels := TRelationships.Create;
    try
      PresRels.LoadFromXml(FPackage.GetPartContent(RelsPath));

      // Slides are ordered by the sldIdLst in presentation.xml, not by part name.
      const SlideIdMatches = TRegEx.Matches(PresentationXml, '<p:sldId\s[^>]*r:id="([^"]+)"', [roIgnoreCase]);
      for var Match in SlideIdMatches do
      begin
        if Match.Groups.Count > 1 then
        begin
          const RelId = Match.Groups[1].Value;
          var SlideTarget := PresRels.GetById(RelId).Target;
          if SlideTarget.StartsWith('/') then
            SlideTarget := SlideTarget.Substring(1)
          else
            SlideTarget := PartPptPrefix + SlideTarget;
          ParseSlideXml(FPackage.GetPartContent(SlideTarget));
        end;
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

  const SpMatches = TRegEx.Matches(SlideXml, '<p:sp>(.*?)</p:sp>', [roIgnoreCase, roSingleLine]);
  for var SpMatch in SpMatches do
  begin
    if SpMatch.Groups.Count > 1 then
    begin
      const SpXml = SpMatch.Groups[1].Value;
      const IsTitle = TRegEx.IsMatch(SpXml, '<p:ph\s[^>]*type="(?:title|ctrTitle)"', [roIgnoreCase]);
      if IsTitle then
      begin
        var Title := '';
        const TextMatches = TRegEx.Matches(SpXml, '<a:t>([^<]*)</a:t>', [roIgnoreCase]);
        for var TextMatch in TextMatches do
          Title := Title + TXml.Unescape(TextMatch.Groups[1].Value);
        Slide.Title := Title;
      end
      else
      begin
        const ParagraphMatches = TRegEx.Matches(SpXml, '<a:p>(.*?)</a:p>', [roIgnoreCase, roSingleLine]);
        for var ParagraphMatch in ParagraphMatches do
          if ParagraphMatch.Groups.Count > 1 then
            ParseParagraphXml(ParagraphMatch.Groups[1].Value, Slide);
      end;
    end;
  end;
end;

procedure TPowerPointReader.ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);
begin
  const Paragraph = Slide.AddParagraph;

  Paragraph.Bullet := (Pos('<a:buChar', Xml) > 0) or (Pos('<a:buAutoNum', Xml) > 0);
  const LevelMatch = TRegEx.Match(Xml, '<a:pPr[^>]*\slvl="(\d+)"', [roIgnoreCase]);
  if LevelMatch.Success then
    Paragraph.IndentLevel := StrToIntDef(LevelMatch.Groups[1].Value, 0);

  const RunMatches = TRegEx.Matches(Xml, '<a:r>(.*?)</a:r>', [roIgnoreCase, roSingleLine]);
  for var RunMatch in RunMatches do
  begin
    if RunMatch.Groups.Count > 1 then
    begin
      const RunXml = RunMatch.Groups[1].Value;
      const TextMatch = TRegEx.Match(RunXml, '<a:t>([^<]*)</a:t>', [roIgnoreCase]);
      if not TextMatch.Success then
        Continue;

      const Run = Paragraph.AddRun(TXml.Unescape(TextMatch.Groups[1].Value));

      const RPrMatch = TRegEx.Match(RunXml, '<a:rPr([^>]*?)(?:/>|>(.*?)</a:rPr>)', [roIgnoreCase, roSingleLine]);
      if RPrMatch.Success then
      begin
        const RPrAttrs = RPrMatch.Groups[1].Value;
        Run.Bold := TRegEx.IsMatch(RPrAttrs, '\bb="1"', [roIgnoreCase]);
        Run.Italic := TRegEx.IsMatch(RPrAttrs, '\bi="1"', [roIgnoreCase]);
        Run.Underline := TRegEx.IsMatch(RPrAttrs, '\bu="sng"', [roIgnoreCase]);

        const SizeMatch = TRegEx.Match(RPrAttrs, '\bsz="(\d+)"', [roIgnoreCase]);
        if SizeMatch.Success then
          Run.FontSize := StrToIntDef(SizeMatch.Groups[1].Value, 0) div 100;

        if RPrMatch.Groups.Count > 2 then
        begin
          const RPrChildren = RPrMatch.Groups[2].Value;
          const ColorMatch = TRegEx.Match(RPrChildren, '<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"', [roIgnoreCase]);
          if ColorMatch.Success then
            Run.FontColor := ColorMatch.Groups[1].Value;
          const FontMatch = TRegEx.Match(RPrChildren, '<a:latin typeface="([^"]*)"', [roIgnoreCase]);
          if FontMatch.Success then
            Run.FontName := FontMatch.Groups[1].Value;
        end;
      end;
    end;
  end;
end;

end.
