unit Office4D.Tests.PowerPoint.AlternateContent;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.PowerPoint;

type
  /// <summary>
  ///   A slide states a shape it cannot render everywhere twice, once in
  ///   mc:Choice and once in mc:Fallback. Only one of the two branches is the
  ///   shape of the slide. Each test builds a minimal .pptx around the slide
  ///   XML it needs.
  /// </summary>
  [TestFixture]
  TPowerPointAlternateContentTests = class
  private
    FPresentation: IPowerPointPresentation;
    FTempFile: string;

    procedure LoadSlide(const SlideXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Slide_ShapeInChoiceAndFallback_ReadsShapeOnce;

    [Test]
    procedure Slide_ShapeInFallbackOnly_ReadsShape;
  end;

implementation

uses
  System.Zip;

const
  ContentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/ppt/presentation.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
    '<Override PartName="/ppt/slides/slide1.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
    '</Types>';

  RootRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
    'Target="ppt/presentation.xml"/>' +
    '</Relationships>';

  PresentationXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>' +
    '</p:presentation>';

  PresentationRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" ' +
    'Target="slides/slide1.xml"/>' +
    '</Relationships>';

  SlidePrefix =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<p:cSld><p:spTree>';

  SlideSuffix = '</p:spTree></p:cSld></p:sld>';

  ShapeXml =
    '<p:sp><p:txBody><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:txBody></p:sp>';

{ TPowerPointAlternateContentTests }

procedure TPowerPointAlternateContentTests.Setup;
begin
  FPresentation := TPowerPointPresentationFactory.CreatePresentation;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'pptaltcontent_test_' + TGUID.NewGuid.ToString + '.pptx');
end;

procedure TPowerPointAlternateContentTests.TearDown;
begin
  FPresentation := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TPowerPointAlternateContentTests.LoadSlide(const SlideXml: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FTempFile, zmWrite);
    Zip.Add(TEncoding.UTF8.GetBytes(ContentTypesXml), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(RootRelsXml), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(PresentationXml), 'ppt/presentation.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(PresentationRelsXml), 'ppt/_rels/presentation.xml.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(SlidePrefix + SlideXml + SlideSuffix), 'ppt/slides/slide1.xml');
    Zip.Close;
  finally
    Zip.Free;
  end;

  FPresentation.LoadFromFile(FTempFile);
end;

procedure TPowerPointAlternateContentTests.Slide_ShapeInChoiceAndFallback_ReadsShapeOnce;
begin
  LoadSlide('<mc:AlternateContent>' +
    '<mc:Choice Requires="a14">' + ShapeXml + '</mc:Choice>' +
    '<mc:Fallback>' + ShapeXml + '</mc:Fallback>' +
    '</mc:AlternateContent>');

  Assert.AreEqual(1, FPresentation.Slides[0].ParagraphCount);
  Assert.AreEqual('Shape text', FPresentation.Slides[0].Paragraphs[0].Text);
end;

procedure TPowerPointAlternateContentTests.Slide_ShapeInFallbackOnly_ReadsShape;
begin
  LoadSlide('<mc:AlternateContent><mc:Fallback>' + ShapeXml + '</mc:Fallback></mc:AlternateContent>');

  Assert.AreEqual(1, FPresentation.Slides[0].ParagraphCount);
  Assert.AreEqual('Shape text', FPresentation.Slides[0].Paragraphs[0].Text);
end;

initialization
  TDUnitX.RegisterTestFixture(TPowerPointAlternateContentTests);

end.
