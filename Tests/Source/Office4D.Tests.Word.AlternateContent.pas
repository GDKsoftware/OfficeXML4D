unit Office4D.Tests.Word.AlternateContent;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Word;

type
  /// <summary>
  /// Text extraction around runs that wrap mc:AlternateContent, the shape of a
  /// text box stated twice: once as DrawingML under mc:Choice and once as VML
  /// under mc:Fallback. Both branches carry the shape's text in a w:p of their
  /// own, nested inside the very paragraph that hosts them, and the paragraph
  /// continues with a plain run after the last shape.
  /// </summary>
  [TestFixture]
  TWordAlternateContentTests = class
  private
    FDoc: IWordDocument;
    FTempFile: string;

    function CountOccurrences(const Text, SubText: string): Integer;
    function ParagraphTextContaining(const SubText: string): string;
    function AlternateContentRun(const ShapeText: string): string;

    procedure LoadDocumentWithShapes(ShapeCount: Integer);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_ParagraphWithShapes_LoadsParagraphs;

    [Test]
    procedure GetText_RunAfterAlternateContent_IncludesSiblingRunText;

    [Test]
    procedure Paragraphs_RunAfterAlternateContent_IncludeSiblingRunText;

    [Test]
    procedure GetText_ShapeTextInChoiceAndFallback_IsNotDuplicated;

    [Test]
    procedure GetText_ManyShapesInOneParagraph_ReadsEveryShapeAndTheSiblingRun;
  end;

implementation

uses
  Office4D.Tests.PackageBuilder;

const
  /// The plain run that follows the shapes, and the text that went missing when
  /// the reader stopped at the closing tag of a paragraph nested in a shape.
  SiblingRunText = 'Heading after the shapes';

{ TWordAlternateContentTests }

procedure TWordAlternateContentTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'altcontent_test_' + TGUID.NewGuid.ToString + '.docx');
end;

procedure TWordAlternateContentTests.TearDown;
begin
  FDoc := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

function TWordAlternateContentTests.AlternateContentRun(const ShapeText: string): string;
begin
  const TextBoxContent = '<w:txbxContent><w:p><w:r><w:t>' + ShapeText + '</w:t></w:r></w:p></w:txbxContent>';

  Result :=
    '<w:r><mc:AlternateContent>' +
      '<mc:Choice Requires="wpg"><w:drawing><wps:wsp><wps:txbx>' + TextBoxContent +
        '</wps:txbx></wps:wsp></w:drawing></mc:Choice>' +
      '<mc:Fallback><w:pict><v:group><v:shape><v:textbox>' + TextBoxContent +
        '</v:textbox></v:shape></v:group></w:pict></mc:Fallback>' +
    '</mc:AlternateContent></w:r>';
end;

procedure TWordAlternateContentTests.LoadDocumentWithShapes(ShapeCount: Integer);
begin
  var Body := '<w:p>';

  for var I := 1 to ShapeCount do
    Body := Body + AlternateContentRun('Shape ' + IntToStr(I));

  Body := Body + '<w:r><w:t>' + SiblingRunText + '</w:t></w:r></w:p>';

  TTestPackage.WriteDocx(FTempFile, Body);
  FDoc.LoadFromFile(FTempFile);
end;

function TWordAlternateContentTests.CountOccurrences(const Text, SubText: string): Integer;
begin
  Result := 0;

  var Index := Pos(SubText, Text);
  while Index > 0 do
  begin
    Inc(Result);
    Index := Pos(SubText, Text, Index + Length(SubText));
  end;
end;

function TWordAlternateContentTests.ParagraphTextContaining(const SubText: string): string;
begin
  Result := '';

  for var I := 0 to FDoc.ParagraphCount - 1 do
  begin
    var ParagraphText := FDoc.Paragraphs[I].Text;
    if Pos(SubText, ParagraphText) > 0 then
      Exit(ParagraphText);
  end;
end;

procedure TWordAlternateContentTests.LoadFromFile_ParagraphWithShapes_LoadsParagraphs;
begin
  LoadDocumentWithShapes(1);

  Assert.IsTrue(FDoc.ParagraphCount > 0, 'Document should yield paragraphs');
end;

procedure TWordAlternateContentTests.GetText_RunAfterAlternateContent_IncludesSiblingRunText;
begin
  LoadDocumentWithShapes(1);

  Assert.Contains(FDoc.Text, SiblingRunText,
    'Text of the plain run following the AlternateContent runs is missing from the document text');
end;

procedure TWordAlternateContentTests.Paragraphs_RunAfterAlternateContent_IncludeSiblingRunText;
begin
  LoadDocumentWithShapes(1);

  Assert.IsNotEmpty(ParagraphTextContaining(SiblingRunText),
    'No paragraph exposes the run that follows the AlternateContent runs');
end;

procedure TWordAlternateContentTests.GetText_ShapeTextInChoiceAndFallback_IsNotDuplicated;
begin
  LoadDocumentWithShapes(1);

  Assert.AreEqual(1, CountOccurrences(FDoc.Text, 'Shape 1'),
    'Shape text present in both mc:Choice and mc:Fallback should be extracted once');
end;

procedure TWordAlternateContentTests.GetText_ManyShapesInOneParagraph_ReadsEveryShapeAndTheSiblingRun;
begin
  // The document that brought this to light held 13 shapes in a single
  // paragraph before its text run.
  LoadDocumentWithShapes(13);

  for var I := 1 to 13 do
    Assert.AreEqual(1, CountOccurrences(FDoc.Text, 'Shape ' + IntToStr(I) + sLineBreak),
      Format('Shape %d should be read exactly once', [I]));

  Assert.Contains(FDoc.Text, SiblingRunText);
end;

initialization
  TDUnitX.RegisterTestFixture(TWordAlternateContentTests);

end.
