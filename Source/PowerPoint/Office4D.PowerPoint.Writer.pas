unit Office4D.PowerPoint.Writer;

interface

uses
  System.Zip,
  Office4D.PowerPoint,
  Office4D.PowerPoint.Model;

type
  /// <summary>
  /// Turns a TPowerPointContent into the parts of a .pptx and adds them to a
  /// zip. Besides the slides themselves, a presentation needs a slide master, a
  /// layout and a theme to be openable, so those are emitted here too.
  /// </summary>
  TPowerPointWriter = class
  private
    FContent: TPowerPointContent;

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
  public
    constructor Create(const Content: TPowerPointContent);

    procedure AddPartsToZip(const Zip: TZipFile);
  end;

implementation

uses
  System.SysUtils,
  Office4D.Types,
  Office4D.Xml;

const
  PresentationMLNs = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  DrawingMLNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  RelTypeSlide = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  RelTypeSlideMaster = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
  RelTypeSlideLayout = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
  RelTypeTheme = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme';

  PartPresentation = 'ppt/presentation.xml';
  PartPresentationRels = 'ppt/_rels/presentation.xml.rels';
  PartSlideMaster = 'ppt/slideMasters/slideMaster1.xml';
  PartSlideMasterRels = 'ppt/slideMasters/_rels/slideMaster1.xml.rels';
  PartSlideLayout = 'ppt/slideLayouts/slideLayout1.xml';
  PartSlideLayoutRels = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
  PartTheme = 'ppt/theme/theme1.xml';
  PartPptPrefix = 'ppt/';

  SlideXmlnsAttrs = ' xmlns:a="' + DrawingMLNs + '" xmlns:r="' + NsOfficeDocumentRelationships + '" xmlns:p="' + PresentationMLNs + '"';
  BulletChar = '&#8226;';

{ TPowerPointWriter }

constructor TPowerPointWriter.Create(const Content: TPowerPointContent);
begin
  inherited Create;
  FContent := Content;
end;

procedure TPowerPointWriter.AddPartsToZip(const Zip: TZipFile);
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

  for var I := 0 to FContent.Slides.Count - 1 do
  begin
    const SlideName = 'slide' + IntToStr(I + 1);
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideXml(FContent.Slides[I])), PartPptPrefix + 'slides/' + SlideName + '.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(GenerateSlideRelsXml), PartPptPrefix + 'slides/_rels/' + SlideName + '.xml.rels');
  end;
end;

function TPowerPointWriter.GenerateContentTypesXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Types xmlns="' + NsContentTypes + '">');
    SB.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>');
    SB.Append('<Default Extension="xml" ContentType="application/xml"/>');
    SB.Append('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>');
    SB.Append('<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>');
    SB.Append('<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>');
    SB.Append('<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
    for var I := 0 to FContent.Slides.Count - 1 do
      SB.Append('<Override PartName="/ppt/slides/slide' + IntToStr(I + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
    SB.Append('</Types>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointWriter.GenerateRootRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + NsPackageRelationships + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeOfficeDocument + '" Target="ppt/presentation.xml"/>' +
    '</Relationships>';
end;

function TPowerPointWriter.GeneratePresentationXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<p:presentation' + SlideXmlnsAttrs + '>');
    SB.Append('<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>');
    if FContent.Slides.Count > 0 then
    begin
      SB.Append('<p:sldIdLst>');
      for var I := 0 to FContent.Slides.Count - 1 do
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

function TPowerPointWriter.GeneratePresentationRelsXml: string;
begin
  var SB := TStringBuilder.Create;
  try
    SB.Append(XmlDeclaration);
    SB.Append('<Relationships xmlns="' + NsPackageRelationships + '">');
    SB.Append('<Relationship Id="rId1" Type="' + RelTypeSlideMaster + '" Target="slideMasters/slideMaster1.xml"/>');
    for var I := 0 to FContent.Slides.Count - 1 do
      SB.Append('<Relationship Id="rId' + IntToStr(2 + I) + '" Type="' + RelTypeSlide + '" Target="slides/slide' + IntToStr(I + 1) + '.xml"/>');
    SB.Append('</Relationships>');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TPowerPointWriter.GenerateSlideMasterXml: string;
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

function TPowerPointWriter.GenerateSlideMasterRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + NsPackageRelationships + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideLayout + '" Target="../slideLayouts/slideLayout1.xml"/>' +
    '<Relationship Id="rId2" Type="' + RelTypeTheme + '" Target="../theme/theme1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointWriter.GenerateSlideLayoutXml: string;
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

function TPowerPointWriter.GenerateSlideLayoutRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + NsPackageRelationships + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideMaster + '" Target="../slideMasters/slideMaster1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointWriter.GenerateThemeXml: string;
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

function TPowerPointWriter.GenerateSlideXml(const Slide: IPowerPointSlide): string;
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

function TPowerPointWriter.GenerateSlideRelsXml: string;
begin
  Result := XmlDeclaration +
    '<Relationships xmlns="' + NsPackageRelationships + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlideLayout + '" Target="../slideLayouts/slideLayout1.xml"/>' +
    '</Relationships>';
end;

function TPowerPointWriter.GenerateParagraphXml(const Paragraph: IPowerPointParagraph): string;
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

function TPowerPointWriter.GenerateRunXml(const Run: IPowerPointRun): string;
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

end.
