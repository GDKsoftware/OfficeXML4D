unit Office4D.PowerPoint.Presentation;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Zip,
  System.Generics.Collections,
  Office4D.PowerPoint,
  Office4D.Metadata,
  Office4D.Package;

type
  TPowerPointRun = class(TInterfacedObject, IPowerPointRun)
  private
    FText: string;
    FBold: Boolean;
    FItalic: Boolean;
    FUnderline: Boolean;
    FFontName: string;
    FFontSize: Integer;
    FFontColor: string;

    function GetText: string;
    procedure SetText(const Value: string);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Integer;
    procedure SetFontSize(const Value: Integer);
    function GetFontColor: string;
    procedure SetFontColor(const Value: string);
  end;

  TPowerPointParagraph = class(TInterfacedObject, IPowerPointParagraph)
  private
    FRuns: TList<IPowerPointRun>;
    FBullet: Boolean;
    FIndentLevel: Integer;

    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IPowerPointRun;
    function GetBullet: Boolean;
    procedure SetBullet(const Value: Boolean);
    function GetIndentLevel: Integer;
    procedure SetIndentLevel(const Value: Integer);
  public
    constructor Create;
    destructor Destroy; override;

    function AddRun(const Text: string): IPowerPointRun;
  end;

  TPowerPointSlide = class(TInterfacedObject, IPowerPointSlide)
  private
    FTitle: string;
    FParagraphs: TList<IPowerPointParagraph>;

    function GetTitle: string;
    procedure SetTitle(const Value: string);
    function GetText: string;
    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IPowerPointParagraph;
  public
    constructor Create;
    destructor Destroy; override;

    function AddParagraph: IPowerPointParagraph; overload;
    function AddParagraph(const Text: string): IPowerPointParagraph; overload;
  end;

  TPowerPointPresentation = class(TInterfacedObject, IPowerPointPresentation)
  private
    FSlides: TList<IPowerPointSlide>;
    FMetadata: TDocumentMetadata;
    FPackage: TOXMLPackage;

    procedure AddPartsToZip(const Zip: TZipFile);
    function GenerateContentTypesXml: string;
    function GenerateRootRelsXml: string;
    function GeneratePresentationXml: string;
    function GeneratePresentationRelsXml: string;
    function GenerateSlideMasterXml: string;
    function GenerateSlideMasterRelsXml: string;
    function GenerateSlideLayoutXml: string;
    function GenerateSlideLayoutRelsXml: string;
    function GenerateThemeXml: string;
    function GenerateSlideXml(const Slide: IPowerPointSlide): string;
    function GenerateSlideRelsXml: string;
    function GenerateParagraphXml(const Paragraph: IPowerPointParagraph): string;
    function GenerateRunXml(const Run: IPowerPointRun): string;

    procedure ParsePackage;
    procedure ParseSlideXml(const Xml: string);
    procedure ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;
    function GetSlideCount: Integer;
    function GetSlide(Index: Integer): IPowerPointSlide;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function AddSlide: IPowerPointSlide; overload;
    function AddSlide(const Title: string): IPowerPointSlide; overload;
  end;

implementation

uses
  System.RegularExpressions,
  Office4D.Relationships,
  Office4D.Xml;

const
  XmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  ContentTypesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
  RelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  PresentationMLNs = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  DrawingMLNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  DocRelationshipsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  RelTypeOfficeDocument = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RelTypeSlide = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  RelTypeSlideMaster = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
  RelTypeSlideLayout = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
  RelTypeTheme = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

  PartRootRels = '_rels/.rels';
  PartPresentation = 'ppt/presentation.xml';
  PartPresentationRels = 'ppt/_rels/presentation.xml.rels';
  PartSlideMaster = 'ppt/slideMasters/slideMaster1.xml';
  PartSlideMasterRels = 'ppt/slideMasters/_rels/slideMaster1.xml.rels';
  PartSlideLayout = 'ppt/slideLayouts/slideLayout1.xml';
  PartSlideLayoutRels = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
  PartTheme = 'ppt/theme/theme1.xml';
  PartCoreProps = 'docProps/core.xml';
  PartPptPrefix = 'ppt/';

  SlideXmlnsAttrs = ' xmlns:a="' + DrawingMLNs + '" xmlns:r="' + DocRelationshipsNs + '" xmlns:p="' + PresentationMLNs + '"';
  BulletChar = '&#8226;';

{ TPowerPointRun }

function TPowerPointRun.GetText: string;
begin
  Result := FText;
end;

procedure TPowerPointRun.SetText(const Value: string);
begin
  FText := Value;
end;

function TPowerPointRun.GetBold: Boolean;
begin
  Result := FBold;
end;

procedure TPowerPointRun.SetBold(const Value: Boolean);
begin
  FBold := Value;
end;

function TPowerPointRun.GetItalic: Boolean;
begin
  Result := FItalic;
end;

procedure TPowerPointRun.SetItalic(const Value: Boolean);
begin
  FItalic := Value;
end;

function TPowerPointRun.GetUnderline: Boolean;
begin
  Result := FUnderline;
end;

procedure TPowerPointRun.SetUnderline(const Value: Boolean);
begin
  FUnderline := Value;
end;

function TPowerPointRun.GetFontName: string;
begin
  Result := FFontName;
end;

procedure TPowerPointRun.SetFontName(const Value: string);
begin
  FFontName := Value;
end;

function TPowerPointRun.GetFontSize: Integer;
begin
  Result := FFontSize;
end;

procedure TPowerPointRun.SetFontSize(const Value: Integer);
begin
  FFontSize := Value;
end;

function TPowerPointRun.GetFontColor: string;
begin
  Result := FFontColor;
end;

procedure TPowerPointRun.SetFontColor(const Value: string);
begin
  FFontColor := Value;
end;

{ TPowerPointParagraph }

constructor TPowerPointParagraph.Create;
begin
  inherited Create;
  FRuns := TList<IPowerPointRun>.Create;
end;

destructor TPowerPointParagraph.Destroy;
begin
  FRuns.Free;
  inherited;
end;

function TPowerPointParagraph.GetText: string;
begin
  Result := '';
  for var Run in FRuns do
    Result := Result + Run.Text;
end;

function TPowerPointParagraph.GetRunCount: Integer;
begin
  Result := FRuns.Count;
end;

function TPowerPointParagraph.GetRun(Index: Integer): IPowerPointRun;
begin
  Result := FRuns[Index];
end;

function TPowerPointParagraph.AddRun(const Text: string): IPowerPointRun;
begin
  Result := TPowerPointRun.Create;
  Result.Text := Text;
  FRuns.Add(Result);
end;

function TPowerPointParagraph.GetBullet: Boolean;
begin
  Result := FBullet;
end;

procedure TPowerPointParagraph.SetBullet(const Value: Boolean);
begin
  FBullet := Value;
end;

function TPowerPointParagraph.GetIndentLevel: Integer;
begin
  Result := FIndentLevel;
end;

procedure TPowerPointParagraph.SetIndentLevel(const Value: Integer);
begin
  FIndentLevel := Value;
end;

{ TPowerPointSlide }

constructor TPowerPointSlide.Create;
begin
  inherited Create;
  FParagraphs := TList<IPowerPointParagraph>.Create;
end;

destructor TPowerPointSlide.Destroy;
begin
  FParagraphs.Free;
  inherited;
end;

function TPowerPointSlide.GetTitle: string;
begin
  Result := FTitle;
end;

procedure TPowerPointSlide.SetTitle(const Value: string);
begin
  FTitle := Value;
end;

function TPowerPointSlide.GetText: string;
begin
  Result := FTitle;
  for var Paragraph in FParagraphs do
  begin
    if Result <> '' then
      Result := Result + sLineBreak;
    Result := Result + Paragraph.Text;
  end;
end;

function TPowerPointSlide.GetParagraphCount: Integer;
begin
  Result := FParagraphs.Count;
end;

function TPowerPointSlide.GetParagraph(Index: Integer): IPowerPointParagraph;
begin
  Result := FParagraphs[Index];
end;

function TPowerPointSlide.AddParagraph: IPowerPointParagraph;
begin
  Result := TPowerPointParagraph.Create;
  FParagraphs.Add(Result);
end;

function TPowerPointSlide.AddParagraph(const Text: string): IPowerPointParagraph;
begin
  Result := AddParagraph;
  Result.AddRun(Text);
end;

{ TPowerPointPresentation }

constructor TPowerPointPresentation.Create;
begin
  inherited Create;
  FSlides := TList<IPowerPointSlide>.Create;
  FMetadata.Clear;
  FPackage := nil;
end;

destructor TPowerPointPresentation.Destroy;
begin
  FreeAndNil(FPackage);
  FSlides.Free;
  inherited;
end;

function TPowerPointPresentation.AddSlide: IPowerPointSlide;
begin
  Result := TPowerPointSlide.Create;
  FSlides.Add(Result);
end;

function TPowerPointPresentation.AddSlide(const Title: string): IPowerPointSlide;
begin
  Result := AddSlide;
  Result.Title := Title;
end;

function TPowerPointPresentation.GetSlideCount: Integer;
begin
  Result := FSlides.Count;
end;

function TPowerPointPresentation.GetSlide(Index: Integer): IPowerPointSlide;
begin
  Result := FSlides[Index];
end;

function TPowerPointPresentation.GetText: string;
begin
  Result := '';
  for var Slide in FSlides do
  begin
    if Result <> '' then
      Result := Result + sLineBreak;
    Result := Result + Slide.Text;
  end;
end;

function TPowerPointPresentation.GetMetadata: TDocumentMetadata;
begin
  Result := FMetadata;
end;

procedure TPowerPointPresentation.SaveToFile(const FileName: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    AddPartsToZip(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TPowerPointPresentation.SaveToStream(const Stream: TStream);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(Stream, zmWrite);
    AddPartsToZip(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TPowerPointPresentation.AddPartsToZip(const Zip: TZipFile);
begin
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateContentTypesXml), '[Content_Types].xml');
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateRootRelsXml), PartRootRels);
  Zip.Add(TEncoding.UTF8.GetBytes(GeneratePresentationXml), PartPresentation);
  Zip.Add(TEncoding.UTF8.GetBytes(GeneratePresentationRelsXml), PartPresentationRels);
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideMasterXml), PartSlideMaster);
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideMasterRelsXml), PartSlideMasterRels);
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideLayoutXml), PartSlideLayout);
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideLayoutRelsXml), PartSlideLayoutRels);
  Zip.Add(TEncoding.UTF8.GetBytes(GenerateThemeXml), PartTheme);

  for var I := 0 to FSlides.Count - 1 do
  begin
    const SlideName = 'slide' + IntToStr(I + 1);
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideXml(FSlides[I])), PartPptPrefix + 'slides/' + SlideName + '.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideRelsXml), PartPptPrefix + 'slides/_rels/' + SlideName + '.xml.rels');
  end;
end;

function TPowerPointPresentation.GenerateContentTypesXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Types xmlns="' + ContentTypesNs + '">');
    SB.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    SB.Append('<Default Extension="xml" ContentType="application/xml"/>');
    SB.Append('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>');
    SB.Append('<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>');
    SB.Append('<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>');
    SB.Append('<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
    for var I := 0 to FSlides.Count - 1 do
      SB.Append('<Override PartName="/ppt/slides/slide' + IntToStr(I + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
    SB.Append('</Types>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointPresentation.GenerateRootRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeOfficeDocument + '" Target="ppt/presentation.xml"/>' +
    '</Relationships>';
end;

function TPowerPointPresentation.GeneratePresentationXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<p:presentation' + SlideXmlnsAttrs + '>');
    SB.Append('<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>');
    if FSlides.Count > 0 then
    begin
      SB.Append('<p:sldIdLst>');
      for var I := 0 to FSlides.Count - 1 do
        SB.Append('<p:sldId id="' + IntToStr(256 + I) + '" r:id="rId' + IntToStr(2 + I) + '"/>');
      SB.Append('</p:sldIdLst>');
    end;
    SB.Append('<p:sldSz cx="12192000" cy="6858000"/>');
    SB.Append('<p:notesSz cx="6858000" cy="9144000"/>');
    SB.Append('</p:presentation>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointPresentation.GeneratePresentationRelsXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + RelationshipsNs + '">');
    SB.Append('<Relationship Id="rId1" Type="' + RelTypeSlideMaster + '" Target="slideMasters/slideMaster1.xml"/>');
    for var I := 0 to FSlides.Count - 1 do
      SB.Append('<Relationship Id="rId' + IntToStr(2 + I) + '" Type="' + RelTypeSlide + '" Target="slides/slide' + IntToStr(I + 1) + '.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointPresentation.GenerateSlideMasterXml: string;
begin
  Result := XmlDeclaration +
    '<p:sldMaster' + SlideXmlnsAttrs + '>' +
    '<p:cSld>' +
    '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' +
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '</p:spTree>' +
    '</p:cSld>' +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"' +
    ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
    '</p:sldMaster>';
end;

function TPowerPointPresentation.GenerateSlideMasterRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideLayout + '" Target="../slideLayouts/slideLayout1.xml"/>' +
    '<Relationship Id="rId2" Type="' + RelTypeTheme + '" Target="../theme/theme1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointPresentation.GenerateSlideLayoutXml: string;
begin
  Result := XmlDeclaration +
    '<p:sldLayout' + SlideXmlnsAttrs + ' type="obj">' +
    '<p:cSld name="Title and Content">' +
    '<p:spTree>' +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr/>' +
    '<p:sp>' +
    '<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
    '</p:sp>' +
    '<p:sp>' +
    '<p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>' +
    '</p:sp>' +
    '</p:spTree>' +
    '</p:cSld>' +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
    '</p:sldLayout>';
end;

function TPowerPointPresentation.GenerateSlideLayoutRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideMaster + '" Target="../slideMasters/slideMaster1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointPresentation.GenerateThemeXml: string;
begin
  Result := XmlDeclaration +
    '<a:theme xmlns:a="' + DrawingMLNs + '" name="Office Theme">' +
    '<a:themeElements>' +
    '<a:clrScheme name="Office">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
    '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
    '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
    '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
    '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
    '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
    '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
    '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
    '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
    '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
    '</a:clrScheme>' +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme>' +
    '<a:fmtScheme name="Office">' +
    '<a:fillStyleLst>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '</a:fillStyleLst>' +
    '<a:lnStyleLst>' +
    '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
    '</a:lnStyleLst>' +
    '<a:effectStyleLst>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '<a:effectStyle><a:effectLst/></a:effectStyle>' +
    '</a:effectStyleLst>' +
    '<a:bgFillStyleLst>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
    '</a:bgFillStyleLst>' +
    '</a:fmtScheme>' +
    '</a:themeElements>' +
    '</a:theme>';
end;

function TPowerPointPresentation.GenerateSlideXml(const Slide: IPowerPointSlide): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<p:sld' + SlideXmlnsAttrs + '>');
    SB.Append('<p:cSld>');
    SB.Append('<p:spTree>');
    SB.Append('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>');
    SB.Append('<p:grpSpPr/>');

    if Slide.Title <> '' then
    begin
      SB.Append('<p:sp>');
      SB.Append('<p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>');
      SB.Append('<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>');
      SB.Append('<p:spPr/>');
      SB.Append('<p:txBody><a:bodyPr/><a:lstStyle/>');
      SB.Append('<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>' + TXml.Escape(Slide.Title) + '</a:t></a:r></a:p>');
      SB.Append('</p:txBody>');
      SB.Append('</p:sp>');
    end;

    if Slide.ParagraphCount > 0 then
    begin
      SB.Append('<p:sp>');
      SB.Append('<p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>');
      SB.Append('<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>');
      SB.Append('<p:spPr/>');
      SB.Append('<p:txBody><a:bodyPr/><a:lstStyle/>');
      for var I := 0 to Slide.ParagraphCount - 1 do
        SB.Append(GenerateParagraphXml(Slide.Paragraphs[I]));
      SB.Append('</p:txBody>');
      SB.Append('</p:sp>');
    end;

    SB.Append('</p:spTree>');
    SB.Append('</p:cSld>');
    SB.Append('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>');
    SB.Append('</p:sld>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointPresentation.GenerateSlideRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideLayout + '" Target="../slideLayouts/slideLayout1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointPresentation.GenerateParagraphXml(const Paragraph: IPowerPointParagraph): string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append('<a:p>');

    // Bullets are always written explicitly (buChar or buNone), because PowerPoint
    // applies its own bulleted defaults to body placeholder text otherwise.
    var LevelAttr := '';
    if Paragraph.IndentLevel > 0 then
      LevelAttr := ' lvl="' + IntToStr(Paragraph.IndentLevel) + '"';
    SB.Append('<a:pPr' + LevelAttr + '>');
    if Paragraph.Bullet then
      SB.Append('<a:buFont typeface="Arial"/><a:buChar char="' + BulletChar + '"/>')
    else
      SB.Append('<a:buNone/>');
    SB.Append('</a:pPr>');

    if Paragraph.RunCount = 0 then
      SB.Append('<a:endParaRPr lang="en-US"/>')
    else
      for var I := 0 to Paragraph.RunCount - 1 do
        SB.Append(GenerateRunXml(Paragraph.Runs[I]));

    SB.Append('</a:p>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointPresentation.GenerateRunXml(const Run: IPowerPointRun): string;
begin
  var Attrs := ' lang="en-US"';
  if Run.Bold then
    Attrs := Attrs + ' b="1"';
  if Run.Italic then
    Attrs := Attrs + ' i="1"';
  if Run.Underline then
    Attrs := Attrs + ' u="sng"';
  if Run.FontSize > 0 then
    Attrs := Attrs + ' sz="' + IntToStr(Run.FontSize * 100) + '"';
  Attrs := Attrs + ' dirty="0"';

  var Children := '';
  if Run.FontColor <> '' then
    Children := Children + '<a:solidFill><a:srgbClr val="' + TXml.Escape(Run.FontColor) + '"/></a:solidFill>';
  if Run.FontName <> '' then
    Children := Children + '<a:latin typeface="' + TXml.Escape(Run.FontName) + '"/>';

  var RPr := '';
  if Children = '' then
    RPr := '<a:rPr' + Attrs + '/>'
  else
    RPr := '<a:rPr' + Attrs + '>' + Children + '</a:rPr>';

  Result := '<a:r>' + RPr + '<a:t>' + TXml.Escape(Run.Text) + '</a:t></a:r>';
end;

procedure TPowerPointPresentation.LoadFromFile(const FileName: string);
begin
  FreeAndNil(FPackage);
  FSlides.Clear;
  FMetadata.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(FileName);
  ParsePackage;
end;

procedure TPowerPointPresentation.LoadFromStream(const Stream: TStream);
begin
  FreeAndNil(FPackage);
  FSlides.Clear;
  FMetadata.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(Stream);
  ParsePackage;
end;

procedure TPowerPointPresentation.ParsePackage;
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

  if FPackage.PartExists(PartCoreProps) then
  begin
    var MetaParser := TMetadataParser.Create;
    try
      FMetadata := MetaParser.Parse(FPackage.GetPartContent(PartCoreProps));
    finally
      MetaParser.Free;
    end;
  end;
end;

procedure TPowerPointPresentation.ParseSlideXml(const Xml: string);
begin
  const Slide = AddSlide;

  const SpMatches = TRegEx.Matches(Xml, '<p:sp>(.*?)</p:sp>', [roIgnoreCase, roSingleLine]);
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

procedure TPowerPointPresentation.ParseParagraphXml(const Xml: string; const Slide: IPowerPointSlide);
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
