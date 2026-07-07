unit Office4D.Tests.PowerPoint;

interface

uses
  System.SysUtils,
  System.IOUtils,
  System.Zip,
  DUnitX.TestFramework,
  Office4D.PowerPoint;

type
  [TestFixture]
  TPowerPointWriteTests = class
  private
    FPresentation: IPowerPointPresentation;
    FTempFile: string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NewPresentation_CreatesFile;

    [Test]
    procedure SaveToFile_ValidatesAsZip;

    [Test]
    procedure SaveToFile_ContainsRequiredParts;

    [Test]
    procedure SaveToFile_SlideCount_MatchesPresentationXml;

    [Test]
    procedure SaveToFile_TitleText_AppearsInSlideXml;

    [Test]
    procedure SaveToFile_BulletParagraph_WritesBuChar;

    [Test]
    procedure SaveToFile_PlainParagraph_WritesBuNone;

    [Test]
    procedure SaveToFile_BoldRun_WritesBoldAttribute;

    [Test]
    procedure SaveToFile_SpecialCharacters_AreEscaped;

    [Test]
    procedure SaveToStream_WritesToStream;
  end;

  [TestFixture]
  TPowerPointRoundTripTests = class
  private
    FPresentation: IPowerPointPresentation;
    FTempFile: string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure RoundTrip_SlideCount_IsPreserved;

    [Test]
    procedure RoundTrip_SlideOrder_IsPreserved;

    [Test]
    procedure RoundTrip_Title_IsPreserved;

    [Test]
    procedure RoundTrip_ParagraphTexts_ArePreserved;

    [Test]
    procedure RoundTrip_BulletAndIndentLevel_ArePreserved;

    [Test]
    procedure RoundTrip_RunFormatting_IsPreserved;

    [Test]
    procedure RoundTrip_GetText_ContainsTitleAndBody;

    [Test]
    procedure LoadFromStream_ReadsPresentation;

    [Test]
    procedure RoundTrip_SpecialCharacters_ArePreserved;
  end;

implementation

uses
  System.Classes,
  Office4D.Package;

{ TPowerPointWriteTests }

procedure TPowerPointWriteTests.Setup;
begin
  FPresentation := TPowerPointPresentationFactory.CreatePresentation;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'test_pptx_' + TGUID.NewGuid.ToString + '.pptx');
end;

procedure TPowerPointWriteTests.TearDown;
begin
  FPresentation := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TPowerPointWriteTests.SaveToFile_NewPresentation_CreatesFile;
begin
  FPresentation.AddSlide('Hello');

  FPresentation.SaveToFile(FTempFile);

  Assert.IsTrue(TFile.Exists(FTempFile), 'File should be created');
end;

procedure TPowerPointWriteTests.SaveToFile_ValidatesAsZip;
begin
  FPresentation.AddSlide('Hello');

  FPresentation.SaveToFile(FTempFile);

  var Zip := TZipFile.Create;
  try
    Zip.Open(FTempFile, zmRead);
    Assert.IsTrue(Zip.FileCount > 0, 'ZIP should contain files');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_ContainsRequiredParts;
begin
  FPresentation.AddSlide('Hello');

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('[Content_Types].xml'), 'Should contain content types');
    Assert.IsTrue(Package.PartExists('_rels/.rels'), 'Should contain root rels');
    Assert.IsTrue(Package.PartExists('ppt/presentation.xml'), 'Should contain presentation part');
    Assert.IsTrue(Package.PartExists('ppt/slideMasters/slideMaster1.xml'), 'Should contain slide master');
    Assert.IsTrue(Package.PartExists('ppt/slideLayouts/slideLayout1.xml'), 'Should contain slide layout');
    Assert.IsTrue(Package.PartExists('ppt/theme/theme1.xml'), 'Should contain theme');
    Assert.IsTrue(Package.PartExists('ppt/slides/slide1.xml'), 'Should contain slide part');
    Assert.IsTrue(Package.PartExists('ppt/slides/_rels/slide1.xml.rels'), 'Should contain slide rels');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_SlideCount_MatchesPresentationXml;
begin
  FPresentation.AddSlide('One');
  FPresentation.AddSlide('Two');
  FPresentation.AddSlide('Three');

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const PresentationXml = Package.GetPartContent('ppt/presentation.xml');
    Assert.IsTrue(Pos('<p:sldId id="256"', PresentationXml) > 0, 'Should contain first slide id');
    Assert.IsTrue(Pos('<p:sldId id="258"', PresentationXml) > 0, 'Should contain third slide id');
    Assert.IsTrue(Package.PartExists('ppt/slides/slide3.xml'), 'Should contain third slide part');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_TitleText_AppearsInSlideXml;
begin
  FPresentation.AddSlide('Quarterly Report');

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SlideXml = Package.GetPartContent('ppt/slides/slide1.xml');
    Assert.IsTrue(Pos('type="title"', SlideXml) > 0, 'Should contain title placeholder');
    Assert.IsTrue(Pos('<a:t>Quarterly Report</a:t>', SlideXml) > 0, 'Should contain title text');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_BulletParagraph_WritesBuChar;
begin
  const Slide = FPresentation.AddSlide('Agenda');
  const Paragraph = Slide.AddParagraph('First point');
  Paragraph.Bullet := True;

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SlideXml = Package.GetPartContent('ppt/slides/slide1.xml');
    Assert.IsTrue(Pos('<a:buChar', SlideXml) > 0, 'Bullet paragraph should contain buChar');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_PlainParagraph_WritesBuNone;
begin
  const Slide = FPresentation.AddSlide('Notes');
  Slide.AddParagraph('Plain text');

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SlideXml = Package.GetPartContent('ppt/slides/slide1.xml');
    Assert.IsTrue(Pos('<a:buNone/>', SlideXml) > 0, 'Plain paragraph should contain buNone');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_BoldRun_WritesBoldAttribute;
begin
  const Slide = FPresentation.AddSlide('Formatting');
  const Paragraph = Slide.AddParagraph;
  const Run = Paragraph.AddRun('Important');
  Run.Bold := True;
  Run.FontSize := 24;

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SlideXml = Package.GetPartContent('ppt/slides/slide1.xml');
    Assert.IsTrue(Pos('b="1"', SlideXml) > 0, 'Bold run should have b attribute');
    Assert.IsTrue(Pos('sz="2400"', SlideXml) > 0, 'Font size should be written in hundredths of a point');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToFile_SpecialCharacters_AreEscaped;
begin
  FPresentation.AddSlide('Q&A <Session>');

  FPresentation.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SlideXml = Package.GetPartContent('ppt/slides/slide1.xml');
    Assert.IsTrue(Pos('Q&amp;A &lt;Session&gt;', SlideXml) > 0, 'Special characters should be XML-escaped');
  finally
    Package.Free;
  end;
end;

procedure TPowerPointWriteTests.SaveToStream_WritesToStream;
begin
  FPresentation.AddSlide('Stream Test');

  var Stream := TMemoryStream.Create;
  try
    FPresentation.SaveToStream(Stream);
    Assert.IsTrue(Stream.Size > 0, 'Stream should contain data');
  finally
    Stream.Free;
  end;
end;

{ TPowerPointRoundTripTests }

procedure TPowerPointRoundTripTests.Setup;
begin
  FPresentation := TPowerPointPresentationFactory.CreatePresentation;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'test_pptx_' + TGUID.NewGuid.ToString + '.pptx');
end;

procedure TPowerPointRoundTripTests.TearDown;
begin
  FPresentation := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TPowerPointRoundTripTests.RoundTrip_SlideCount_IsPreserved;
begin
  FPresentation.AddSlide('One');
  FPresentation.AddSlide('Two');
  FPresentation.AddSlide('Three');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  Assert.AreEqual(3, Presentation2.SlideCount);
end;

procedure TPowerPointRoundTripTests.RoundTrip_SlideOrder_IsPreserved;
begin
  FPresentation.AddSlide('First');
  FPresentation.AddSlide('Second');
  FPresentation.AddSlide('Third');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  Assert.AreEqual('First', Presentation2.Slides[0].Title);
  Assert.AreEqual('Second', Presentation2.Slides[1].Title);
  Assert.AreEqual('Third', Presentation2.Slides[2].Title);
end;

procedure TPowerPointRoundTripTests.RoundTrip_Title_IsPreserved;
begin
  FPresentation.AddSlide('Quarterly Report');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  Assert.AreEqual('Quarterly Report', Presentation2.Slides[0].Title);
end;

procedure TPowerPointRoundTripTests.RoundTrip_ParagraphTexts_ArePreserved;
begin
  const Slide = FPresentation.AddSlide('Agenda');
  Slide.AddParagraph('First item');
  Slide.AddParagraph('Second item');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  const Slide2 = Presentation2.Slides[0];
  Assert.AreEqual(2, Slide2.ParagraphCount);
  Assert.AreEqual('First item', Slide2.Paragraphs[0].Text);
  Assert.AreEqual('Second item', Slide2.Paragraphs[1].Text);
end;

procedure TPowerPointRoundTripTests.RoundTrip_BulletAndIndentLevel_ArePreserved;
begin
  const Slide = FPresentation.AddSlide('Agenda');
  const Bullet = Slide.AddParagraph('Bullet point');
  Bullet.Bullet := True;
  const SubBullet = Slide.AddParagraph('Sub point');
  SubBullet.Bullet := True;
  SubBullet.IndentLevel := 1;
  Slide.AddParagraph('Plain closing line');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  const Slide2 = Presentation2.Slides[0];
  Assert.IsTrue(Slide2.Paragraphs[0].Bullet, 'First paragraph should be a bullet');
  Assert.AreEqual(0, Slide2.Paragraphs[0].IndentLevel);
  Assert.IsTrue(Slide2.Paragraphs[1].Bullet, 'Second paragraph should be a bullet');
  Assert.AreEqual(1, Slide2.Paragraphs[1].IndentLevel);
  Assert.IsFalse(Slide2.Paragraphs[2].Bullet, 'Third paragraph should not be a bullet');
end;

procedure TPowerPointRoundTripTests.RoundTrip_RunFormatting_IsPreserved;
begin
  const Slide = FPresentation.AddSlide('Formatting');
  const Paragraph = Slide.AddParagraph;
  const Run = Paragraph.AddRun('Styled text');
  Run.Bold := True;
  Run.Italic := True;
  Run.Underline := True;
  Run.FontSize := 28;
  Run.FontName := 'Arial';
  Run.FontColor := 'FF0000';

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  const Run2 = Presentation2.Slides[0].Paragraphs[0].Runs[0];
  Assert.AreEqual('Styled text', Run2.Text);
  Assert.IsTrue(Run2.Bold, 'Bold should be preserved');
  Assert.IsTrue(Run2.Italic, 'Italic should be preserved');
  Assert.IsTrue(Run2.Underline, 'Underline should be preserved');
  Assert.AreEqual(28, Run2.FontSize);
  Assert.AreEqual('Arial', Run2.FontName);
  Assert.AreEqual('FF0000', Run2.FontColor);
end;

procedure TPowerPointRoundTripTests.RoundTrip_GetText_ContainsTitleAndBody;
begin
  const Slide = FPresentation.AddSlide('My Title');
  Slide.AddParagraph('Body line');

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  Assert.IsTrue(Pos('My Title', Presentation2.Text) > 0, 'Text should contain the title');
  Assert.IsTrue(Pos('Body line', Presentation2.Text) > 0, 'Text should contain the body text');
end;

procedure TPowerPointRoundTripTests.LoadFromStream_ReadsPresentation;
begin
  FPresentation.AddSlide('Stream Round Trip');
  FPresentation.SaveToFile(FTempFile);

  var Stream := TFileStream.Create(FTempFile, fmOpenRead or fmShareDenyWrite);
  try
    const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
    Presentation2.LoadFromStream(Stream);
    Assert.AreEqual(1, Presentation2.SlideCount);
    Assert.AreEqual('Stream Round Trip', Presentation2.Slides[0].Title);
  finally
    Stream.Free;
  end;
end;

procedure TPowerPointRoundTripTests.RoundTrip_SpecialCharacters_ArePreserved;
begin
  const Special = 'Q&A <tag> "quote" ''apos'' 100% > 50%';
  const Slide = FPresentation.AddSlide(Special);
  Slide.AddParagraph(Special);

  FPresentation.SaveToFile(FTempFile);

  const Presentation2 = TPowerPointPresentationFactory.CreatePresentation;
  Presentation2.LoadFromFile(FTempFile);
  const Slide2 = Presentation2.Slides[0];
  Assert.AreEqual(Special, Slide2.Title, 'Title special characters should round-trip');
  Assert.AreEqual(Special, Slide2.Paragraphs[0].Text, 'Paragraph special characters should round-trip');
end;

initialization
  TDUnitX.RegisterTestFixture(TPowerPointWriteTests);
  TDUnitX.RegisterTestFixture(TPowerPointRoundTripTests);

end.
