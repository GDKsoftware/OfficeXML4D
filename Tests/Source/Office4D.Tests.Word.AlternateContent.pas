unit Office4D.Tests.Word.AlternateContent;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Word;

type
  /// <summary>
  /// Text extraction around runs that wrap mc:AlternateContent, the grouped
  /// VML/DrawingML shapes of a text box. The sample holds a paragraph with 13
  /// such runs followed by a plain sibling run carrying the visible heading,
  /// and states the shape text in both branches of every one of them.
  /// </summary>
  [TestFixture]
  TWordAlternateContentTests = class(TOffice4DTests)
  private
    FDoc: IWordDocument;

    function CountOccurrences(const Text, SubText: string): Integer;
    function ParagraphTextContaining(const SubText: string): string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_SampleWithAlternateContent_LoadsParagraphs;

    [Test]
    procedure GetText_RunAfterAlternateContent_IncludesSiblingRunText;

    [Test]
    procedure Paragraphs_RunAfterAlternateContent_IncludeSiblingRunText;

    [Test]
    procedure GetText_ShapeTextInChoiceAndFallback_IsNotDuplicated;
  end;

implementation

const
  /// Plain sibling run that directly follows the AlternateContent runs.
  /// Written without its leading accented character so this unit stays ASCII.
  SiblingRunText = 'NDICE GENERAL';

  /// Text living inside a shape, present both in mc:Choice (wps:txbx)
  /// and in mc:Fallback (v:textbox) and therefore only visible once.
  ShapeText = 'Mantenimiento';

{ TWordAlternateContentTests }

procedure TWordAlternateContentTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
  FDoc.LoadFromFile(GetWordAlternateContentSamplePath);
end;

procedure TWordAlternateContentTests.TearDown;
begin
  FDoc := nil;
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

procedure TWordAlternateContentTests.LoadFromFile_SampleWithAlternateContent_LoadsParagraphs;
begin
  Assert.IsTrue(FDoc.ParagraphCount > 0, 'Sample document should yield paragraphs');
end;

procedure TWordAlternateContentTests.GetText_RunAfterAlternateContent_IncludesSiblingRunText;
begin
  Assert.Contains(FDoc.Text, SiblingRunText,
    'Text of the plain run following the AlternateContent runs is missing from the document text');
end;

procedure TWordAlternateContentTests.Paragraphs_RunAfterAlternateContent_IncludeSiblingRunText;
begin
  Assert.IsNotEmpty(ParagraphTextContaining(SiblingRunText),
    'No paragraph exposes the run that follows the AlternateContent runs');
end;

procedure TWordAlternateContentTests.GetText_ShapeTextInChoiceAndFallback_IsNotDuplicated;
begin
  Assert.AreEqual(1, CountOccurrences(FDoc.Text, ShapeText),
    'Shape text present in both mc:Choice and mc:Fallback should be extracted once');
end;

initialization
  TDUnitX.RegisterTestFixture(TWordAlternateContentTests);

end.
