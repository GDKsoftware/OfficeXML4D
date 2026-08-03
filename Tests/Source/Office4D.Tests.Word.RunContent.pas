unit Office4D.Tests.Word.RunContent;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Word;

type
  /// <summary>
  /// Reading of run level content: every w:t of a run, breaks and tabs in
  /// document order, hyperlinks in place, and paragraph mark properties that
  /// must not be mistaken for a run. Each test builds a minimal .docx around
  /// the body XML it needs.
  /// </summary>
  [TestFixture]
  TWordRunContentTests = class
  private
    FDoc: IWordDocument;
    FTempFile: string;

    procedure LoadBody(const BodyXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Run_WithMultipleTextElements_ReadsEveryTextElement;

    [Test]
    procedure Run_WithBreakBetweenTextElements_KeepsDocumentOrder;

    [Test]
    procedure Run_WithTabBetweenTextElements_KeepsDocumentOrder;

    [Test]
    procedure Paragraph_WithParagraphMarkRunProperties_ReadsAllRuns;

    [Test]
    procedure Paragraph_WithHyperlinkBetweenRuns_KeepsDocumentOrder;
  end;

implementation

uses
  Office4D.Tests.PackageBuilder;

{ TWordRunContentTests }

procedure TWordRunContentTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'runcontent_test_' + TGUID.NewGuid.ToString + '.docx');
end;

procedure TWordRunContentTests.TearDown;
begin
  FDoc := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TWordRunContentTests.LoadBody(const BodyXml: string);
begin
  TTestPackage.WriteDocx(FTempFile, BodyXml);
  FDoc.LoadFromFile(FTempFile);
end;

procedure TWordRunContentTests.Run_WithMultipleTextElements_ReadsEveryTextElement;
begin
  LoadBody('<w:p><w:r><w:t>Hello</w:t><w:t xml:space="preserve"> World</w:t></w:r></w:p>');

  Assert.AreEqual('Hello World', FDoc.Paragraphs[0].Text);
end;

procedure TWordRunContentTests.Run_WithBreakBetweenTextElements_KeepsDocumentOrder;
begin
  LoadBody('<w:p><w:r><w:t>Line1</w:t><w:br/><w:t>Line2</w:t></w:r></w:p>');

  Assert.AreEqual('Line1' + sLineBreak + 'Line2', FDoc.Paragraphs[0].Text);
end;

procedure TWordRunContentTests.Run_WithTabBetweenTextElements_KeepsDocumentOrder;
begin
  LoadBody('<w:p><w:r><w:t>Col1</w:t><w:tab/><w:t>Col2</w:t></w:r></w:p>');

  Assert.AreEqual('Col1'#9'Col2', FDoc.Paragraphs[0].Text);
end;

procedure TWordRunContentTests.Paragraph_WithParagraphMarkRunProperties_ReadsAllRuns;
begin
  // The w:rPr of the paragraph mark sits in w:pPr and is not a run, so it must
  // neither add a run nor merge with the runs that follow it.
  LoadBody('<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>' +
    '<w:r><w:t>First</w:t></w:r><w:r><w:t xml:space="preserve"> Second</w:t></w:r></w:p>');

  Assert.AreEqual(2, FDoc.Paragraphs[0].RunCount);
  Assert.AreEqual('First Second', FDoc.Paragraphs[0].Text);
end;

procedure TWordRunContentTests.Paragraph_WithHyperlinkBetweenRuns_KeepsDocumentOrder;
begin
  LoadBody('<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
    '<w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
    '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>');

  Assert.AreEqual('Before link after', FDoc.Paragraphs[0].Text);
end;

initialization
  TDUnitX.RegisterTestFixture(TWordRunContentTests);

end.
