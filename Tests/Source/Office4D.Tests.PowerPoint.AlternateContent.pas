unit Office4D.Tests.PowerPoint.AlternateContent;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.PowerPoint;

type
  /// <summary>
  /// A slide states a shape it cannot render everywhere twice, once in
  /// mc:Choice and once in mc:Fallback. Only one of the two branches is the
  /// shape of the slide. Each test builds a minimal .pptx around the slide
  /// XML it needs.
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
  Office4D.Tests.PackageBuilder;

const
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
  TTestPackage.WritePptx(FTempFile, SlideXml);
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
