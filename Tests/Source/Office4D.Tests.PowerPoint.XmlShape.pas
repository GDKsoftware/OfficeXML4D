unit Office4D.Tests.PowerPoint.XmlShape;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.PowerPoint;

type
  /// <summary>
  /// A presentation reads the same whatever tool wrote it: attribute order,
  /// quote style, optional attributes and whitespace between elements carry no
  /// meaning, so they must not change the result.
  /// </summary>
  [TestFixture]
  TPowerPointXmlShapeTests = class
  private
    FPresentation: IPowerPointPresentation;
    FSourceFile: string;

    procedure LoadSlide(const SlideXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_RunPropertiesInAnyShape_ReadsFormatting;

    [Test]
    procedure LoadFromFile_ShapeWithAttributes_ReadsItsText;

    [Test]
    procedure LoadFromFile_TitlePlaceholderWithSingleQuotes_ReadsTitle;

    [Test]
    procedure LoadFromFile_SlideIdWithSingleQuotes_ReadsSlide;
  end;

implementation

uses
  Office4D.Tests.PackageBuilder;

const
  PresentationNs = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  DrawingNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  RelationshipsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  PackageRelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';

{ TPowerPointXmlShapeTests }

procedure TPowerPointXmlShapeTests.Setup;
begin
  FPresentation := TPowerPointPresentationFactory.CreatePresentation;
  FSourceFile := TPath.Combine(TPath.GetTempPath, 'pptxmlshape_' + TGUID.NewGuid.ToString + '.pptx');
end;

procedure TPowerPointXmlShapeTests.TearDown;
begin
  FPresentation := nil;
  if TFile.Exists(FSourceFile) then
    TFile.Delete(FSourceFile);
end;

procedure TPowerPointXmlShapeTests.LoadSlide(const SlideXml: string);
begin
  TTestPackage.WritePptx(FSourceFile, SlideXml);
  FPresentation.LoadFromFile(FSourceFile);
end;

procedure TPowerPointXmlShapeTests.LoadFromFile_RunPropertiesInAnyShape_ReadsFormatting;
begin
  LoadSlide('<p:sp><p:txBody><a:p><a:r>' +
            '<a:rPr lang="en-US" sz=''2400'' i = "1" b="true">' + sLineBreak +
            '  <a:solidFill>' + sLineBreak + '    <a:srgbClr val="FF0000"/>' + sLineBreak + '  </a:solidFill>' + sLineBreak +
            '  <a:latin pitchFamily="34" typeface="Arial"/>' + sLineBreak +
            '</a:rPr><a:t>Styled</a:t></a:r></a:p></p:txBody></p:sp>');

  const Run = FPresentation.Slides[0].Paragraphs[0].Runs[0];

  Assert.AreEqual('Styled', Run.Text);
  Assert.IsTrue(Run.Bold, 'b="true" is bold');
  Assert.IsTrue(Run.Italic, 'whitespace around = keeps italic');
  Assert.AreEqual(24, Run.FontSize);
  Assert.AreEqual('FF0000', Run.FontColor, 'whitespace between solidFill and srgbClr keeps the colour');
  Assert.AreEqual('Arial', Run.FontName, 'typeface need not be the first attribute');
end;

procedure TPowerPointXmlShapeTests.LoadFromFile_ShapeWithAttributes_ReadsItsText;
begin
  LoadSlide('<p:sp useBgFill="0"><p:txBody><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:txBody></p:sp>');

  Assert.AreEqual(1, FPresentation.Slides[0].ParagraphCount);
  Assert.AreEqual('Shape text', FPresentation.Slides[0].Paragraphs[0].Text);
end;

procedure TPowerPointXmlShapeTests.LoadFromFile_TitlePlaceholderWithSingleQuotes_ReadsTitle;
begin
  LoadSlide('<p:sp><p:nvSpPr><p:nvPr><p:ph idx=''0'' type=''title''/></p:nvPr></p:nvSpPr>' +
            '<p:txBody><a:p><a:r><a:t>Quarterly Report</a:t></a:r></a:p></p:txBody></p:sp>');

  Assert.AreEqual('Quarterly Report', FPresentation.Slides[0].Title);
end;

procedure TPowerPointXmlShapeTests.LoadFromFile_SlideIdWithSingleQuotes_ReadsSlide;
begin
  const RootRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Type="' + RelationshipsNs + '/officeDocument" Target="/ppt/presentation.xml" Id="rId1"/>' +
    '</Relationships>';
  const PresentationXml =
    '<p:presentation xmlns:p="' + PresentationNs + '" xmlns:r="' + RelationshipsNs + '">' +
    '<p:sldIdLst><p:sldId r:id = ''rId7'' id="256"/></p:sldIdLst></p:presentation>';
  const PresentationRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Target="slides/slide1.xml" Type="' + RelationshipsNs + '/slide" Id="rId7"/>' +
    '</Relationships>';
  const SlideXml =
    '<p:sld xmlns:a="' + DrawingNs + '" xmlns:p="' + PresentationNs + '"><p:cSld><p:spTree>' +
    '<p:sp><p:txBody><a:p><a:r><a:t>Only slide</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sld>';

  TTestPackage.WriteParts(FSourceFile,
    [TTestPart.Create('_rels/.rels', RootRelsXml),
     TTestPart.Create('ppt/presentation.xml', PresentationXml),
     TTestPart.Create('ppt/_rels/presentation.xml.rels', PresentationRelsXml),
     TTestPart.Create('ppt/slides/slide1.xml', SlideXml)]);
  FPresentation.LoadFromFile(FSourceFile);

  Assert.AreEqual(1, FPresentation.SlideCount);
  Assert.AreEqual('Only slide', FPresentation.Slides[0].Paragraphs[0].Text);
end;

end.
