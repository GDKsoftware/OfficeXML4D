unit Office4D.Tests.Excel.FontStyle;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelFontStyleTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure GetFontStyle_NoFormatting_ReturnsEmptySet;

    [Test]
    procedure SetFontStyle_BoldAndStrikeout_SetsUnderlyingProperties;

    [Test]
    procedure SetFontStyle_IsAFullReplace_ClearsStylesNotInTheNewSet;

    [Test]
    procedure SetBold_DirectlyThenReadFontStyle_ReflectsTheChange;

    [Test]
    procedure SetStrikeout_DirectlyThenReadFontStyle_ReflectsTheChange;

    [Test]
    procedure SaveToFile_WithStrikeout_ContainsStrikeElement;

    [Test]
    procedure RoundTrip_FontStyle_PreservesAllFourFlags;

    [Test]
    procedure RoundTrip_CombinedFontStyleAndFontColor_PreservesBoth;

    [Test]
    procedure SaveToFile_StrikeoutOnlyCell_ResolvesToDistinctFontEntry;
  end;

implementation

uses
  Office4D.Package;

{ TExcelFontStyleTests }

procedure TExcelFontStyleTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'fontstyle_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelFontStyleTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelFontStyleTests.GetFontStyle_NoFormatting_ReturnsEmptySet;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];

  Assert.IsTrue(Cell.FontStyle = [], 'A cell with no formatting must report an empty FontStyle set');
end;

procedure TExcelFontStyleTests.SetFontStyle_BoldAndStrikeout_SetsUnderlyingProperties;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];

  Cell.FontStyle := [TExcelFontStyle.Bold, TExcelFontStyle.Strikeout];

  Assert.IsTrue(Cell.Bold, 'Bold property should reflect the FontStyle set');
  Assert.IsFalse(Cell.Italic, 'Italic was not in the set and must stay False');
  Assert.IsFalse(Cell.Underline, 'Underline was not in the set and must stay False');
  Assert.IsTrue(TExcelFontStyle.Bold in Cell.FontStyle, 'Bold should read back from FontStyle');
  Assert.IsTrue(TExcelFontStyle.Strikeout in Cell.FontStyle, 'Strikeout should read back from FontStyle');
  Assert.IsFalse(TExcelFontStyle.Italic in Cell.FontStyle, 'Italic should not appear in FontStyle');
end;

procedure TExcelFontStyleTests.SetFontStyle_IsAFullReplace_ClearsStylesNotInTheNewSet;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];

  // Establish Bold + Italic first via the existing per-flag properties.
  Cell.Bold := True;
  Cell.Italic := True;

  // A FontStyle assignment that omits Bold and Italic must clear them, not merge with them --
  // this is standard `set` assignment semantics, but worth asserting explicitly since it's
  // easy to accidentally implement FontStyle's setter as an OR instead of a replace.
  Cell.FontStyle := [TExcelFontStyle.Underline];

  Assert.IsFalse(Cell.Bold, 'Bold must be cleared, not merged, by a FontStyle assignment that omits it');
  Assert.IsFalse(Cell.Italic, 'Italic must be cleared, not merged, by a FontStyle assignment that omits it');
  Assert.IsTrue(Cell.Underline, 'Underline should be set from the new FontStyle value');
end;

procedure TExcelFontStyleTests.SetBold_DirectlyThenReadFontStyle_ReflectsTheChange;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];

  // FontStyle must be a live view over Bold/Italic/Underline/Strikeout, not separate
  // storage that can desync from them.
  Cell.Bold := True;

  Assert.IsTrue(TExcelFontStyle.Bold in Cell.FontStyle,
    'Setting Bold directly must be visible through FontStyle immediately, with no separate ' +
    'FontStyle state to fall out of sync');
end;

procedure TExcelFontStyleTests.SetStrikeout_DirectlyThenReadFontStyle_ReflectsTheChange;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const Cell = Sheet.Cell['A1'];

  // Strikeout must behave exactly like Bold/Italic/Underline here -- it's a standalone
  // property, not something only reachable through the FontStyle set.
  Cell.Strikeout := True;

  Assert.IsTrue(TExcelFontStyle.Strikeout in Cell.FontStyle,
    'Setting Strikeout directly must be visible through FontStyle immediately');
  Assert.IsTrue(Cell.Strikeout, 'The Strikeout property itself must read back True');
end;

procedure TExcelFontStyleTests.SaveToFile_WithStrikeout_ContainsStrikeElement;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Struck through';
  Sheet.Cell['A1'].FontStyle := [TExcelFontStyle.Strikeout];

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('<strike/>', StylesXml) > 0, 'Should contain strike element in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelFontStyleTests.RoundTrip_FontStyle_PreservesAllFourFlags;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Fully styled';
  Sheet.Cell['A1'].FontStyle := [TExcelFontStyle.Bold, TExcelFontStyle.Italic,
    TExcelFontStyle.Underline, TExcelFontStyle.Strikeout];
  Sheet.Cell['B1'].AsString := 'Plain';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  const CellA1 = Workbook2.Sheets[0].Cell['A1'];
  Assert.IsTrue(CellA1.FontStyle = [TExcelFontStyle.Bold, TExcelFontStyle.Italic,
    TExcelFontStyle.Underline, TExcelFontStyle.Strikeout], 'All four flags should survive round-trip');

  const CellB1 = Workbook2.Sheets[0].Cell['B1'];
  Assert.IsTrue(CellB1.FontStyle = [], 'Plain cell should round-trip with an empty FontStyle set');
end;

procedure TExcelFontStyleTests.RoundTrip_CombinedFontStyleAndFontColor_PreservesBoth;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red and struck through';
  Sheet.Cell['A1'].FontStyle := [TExcelFontStyle.Strikeout];
  Sheet.Cell['A1'].FontColor := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Cell = Workbook2.Sheets[0].Cell['A1'];

  Assert.IsTrue(TExcelFontStyle.Strikeout in Cell.FontStyle, 'Strikeout should survive alongside FontColor');
  Assert.AreEqual(Cardinal($FF0000), Cell.FontColor, 'FontColor should survive alongside Strikeout');
end;

procedure TExcelFontStyleTests.SaveToFile_StrikeoutOnlyCell_ResolvesToDistinctFontEntry;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  // A1 has strikeout only; B1 has identical bold/italic/underline/color/size/name (all
  // default) but no strikeout. If the FontKey field count/order ever drifts out of sync
  // between the collection loop and the cellXfs reconstruction -- the exact class of bug
  // that broke FontColor earlier -- these two would collapse onto the same font entry.
  Sheet.Cell['A1'].FontStyle := [TExcelFontStyle.Strikeout];
  Sheet.Cell['B1'].AsString := 'No strikeout';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    // Exactly one <strike/> should appear across the whole <fonts> block: the entry for
    // A1's font. B1 uses the default font (index 0), which never gets a <strike/>.
    var StrikeCount := 0;
    var SearchPos := 1;
    while True do
    begin
      const FoundPos = Pos('<strike/>', Copy(StylesXml, SearchPos, MaxInt));
      if FoundPos = 0 then Break;
      Inc(StrikeCount);
      SearchPos := SearchPos + FoundPos + Length('<strike/>') - 1;
    end;
    Assert.AreEqual(1, StrikeCount, 'Exactly one font entry should carry <strike/>');
  finally
    Package.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelFontStyleTests);

end.
