unit Office4D.Tests.Excel.InlineStrings;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelInlineStringTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FSourceFile: string;
    FSavedFile: string;

    procedure LoadInlineString(const InlineXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_InlineString_AddsTheCell;

    [Test]
    [TestCase('Plain', '<t>inline text</t>|inline text', '|')]
    [TestCase('Runs', '<r><t xml:space="preserve">Hello </t></r><r><rPr><b/></rPr><t>world</t></r>|Hello world', '|')]
    [TestCase('Escaped', '<t>A &amp; B</t>|A & B', '|')]
    [TestCase('Phonetic', '<t>Tokyo</t><rPh sb="0" eb="5"><t>toukyou</t></rPh>|Tokyo', '|')]
    procedure LoadFromFile_InlineString_ReadsItsText(const InlineXml, Expected: string);

    [Test]
    procedure LoadFromFile_InlineStringBeforeValueCell_KeepsBothValues;

    [Test]
    procedure RoundTrip_InlineString_KeepsText;
  end;

implementation

uses
  Office4D.Tests.PackageBuilder;

const
  InlineCellRowXml = '<row r="7"><c r="A7" t="inlineStr"><is>%s</is></c></row>';

procedure TExcelInlineStringTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FSourceFile := TPath.Combine(TPath.GetTempPath, 'inlinestring_source_' + TGUID.NewGuid.ToString + '.xlsx');
  FSavedFile := TPath.Combine(TPath.GetTempPath, 'inlinestring_saved_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelInlineStringTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FSourceFile) then
    TFile.Delete(FSourceFile);
  if TFile.Exists(FSavedFile) then
    TFile.Delete(FSavedFile);
end;

procedure TExcelInlineStringTests.LoadInlineString(const InlineXml: string);
begin
  TTestPackage.WriteXlsx(FSourceFile, Format(InlineCellRowXml, [InlineXml]));
  FWorkbook.LoadFromFile(FSourceFile);
end;

procedure TExcelInlineStringTests.LoadFromFile_InlineString_AddsTheCell;
begin
  LoadInlineString('<t>inline text</t>');

  const Sheet = FWorkbook.Sheets[0];

  Assert.IsTrue(Sheet.Cells.ContainsKey('A7'), 'The inline string cell must be read');
end;

procedure TExcelInlineStringTests.LoadFromFile_InlineString_ReadsItsText(const InlineXml, Expected: string);
begin
  LoadInlineString(InlineXml);

  const Actual = FWorkbook.Sheets[0].Cell['A7'].AsString;

  Assert.AreEqual(Expected, Actual);
end;

procedure TExcelInlineStringTests.LoadFromFile_InlineStringBeforeValueCell_KeepsBothValues;
begin
  TTestPackage.WriteXlsx(FSourceFile, '<row r="1"><c r="B1" t="inlineStr"><is><t>x</t></is></c><c r="C1"><v>5</v></c></row>');
  FWorkbook.LoadFromFile(FSourceFile);

  const Sheet = FWorkbook.Sheets[0];

  Assert.AreEqual('x', Sheet.Cell['B1'].AsString);
  Assert.AreEqual(Double(5), Sheet.Cell['C1'].AsFloat, 'The next cell keeps its own value');
end;

procedure TExcelInlineStringTests.RoundTrip_InlineString_KeepsText;
begin
  LoadInlineString('<t>inline text</t>');

  FWorkbook.SaveToFile(FSavedFile);
  const Reloaded = TExcelWorkbookFactory.Create;
  Reloaded.LoadFromFile(FSavedFile);

  Assert.AreEqual('inline text', Reloaded.Sheets[0].Cell['A7'].AsString);
end;

end.
