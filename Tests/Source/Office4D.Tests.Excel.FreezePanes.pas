unit Office4D.Tests.Excel.FreezePanes;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelFreezePanesTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NoFreeze_OmitsSheetViews;

    [Test]
    procedure FreezePanes_BothRowsAndColumns_SetsFrozenCounts;

    [Test]
    procedure FreezePanes_RowsOnly_LeavesColumnsAtZero;

    [Test]
    procedure FreezePanes_ColumnsOnly_LeavesRowsAtZero;

    [Test]
    procedure FreezePanes_TopLeftCorner_IsEquivalentToNoFreeze;

    [Test]
    procedure FreezePanes_InvalidAddress_RaisesException;

    [Test]
    procedure UnfreezePanes_AfterFreezing_ClearsFrozenState;

    [Test]
    procedure RoundTrip_FreezeBothRowsAndColumns_PreservesFrozenCounts;

    [Test]
    procedure SaveToFile_WithFreezePanes_ContainsCorrectPaneAttributes;

    [Test]
    procedure SaveToFile_FreezeRowsOnly_UsesBottomLeftActivePane;

    [Test]
    procedure SaveToFile_FreezeColumnsOnly_UsesTopRightActivePane;

    [Test]
    procedure SaveToFile_WithFreezePanesAndColumnWidths_SheetViewsPrecedesCols;

    [Test]
    procedure SaveToFile_FreezeBothAxes_EmitsThreeSelectionsAfterPane;

    [Test]
    procedure SaveToFile_FreezeRowsOnly_EmitsSingleBottomLeftSelection;

    [Test]
    procedure SaveToFile_FreezeColumnsOnly_EmitsSingleTopRightSelection;

    [Test]
    procedure LoadFromFile_SplitPane_IsNotReadAsFrozen;

    [Test]
    procedure LoadFromFile_FrozenSplitPane_ReadsFrozenCounts;
  end;

implementation

uses
  System.Zip,
  System.RegularExpressions,
  Office4D.Errors,
  Office4D.Package;

{ Rewrites xl/worksheets/sheet1.xml inside an existing .xlsx with NewSheetXml, copying every
  other part verbatim. Used to inject pane XML the library never writes itself (e.g. a "split"
  pane) so the read path can be exercised against it. }
procedure ReplaceSheet1Xml(const FileName, NewSheetXml: string);
const
  Sheet1Path = 'xl/worksheets/sheet1.xml';
var
  Names: TArray<string>;
  Contents: TArray<TBytes>;
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmRead);
    SetLength(Names, Zip.FileCount);
    SetLength(Contents, Zip.FileCount);
    for var I := 0 to Zip.FileCount - 1 do
    begin
      Names[I] := Zip.FileNames[I];
      if SameText(StringReplace(Names[I], '\', '/', [rfReplaceAll]), Sheet1Path) then
        Contents[I] := TEncoding.UTF8.GetBytes(NewSheetXml)
      else
        Zip.Read(I, Contents[I]);
    end;
  finally
    Zip.Free;
  end;

  TFile.Delete(FileName);

  var ZipOut := TZipFile.Create;
  try
    ZipOut.Open(FileName, zmWrite);
    for var I := 0 to High(Names) do
      ZipOut.Add(Contents[I], Names[I]);
  finally
    ZipOut.Free;
  end;
end;

{ TExcelFreezePanesTests }

procedure TExcelFreezePanesTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'freezepanes_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelFreezePanesTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelFreezePanesTests.SaveToFile_NoFreeze_OmitsSheetViews;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsFalse(Pos('<sheetViews>', SheetXml) > 0,
      'A sheet with no frozen panes must not emit a sheetViews element at all');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.FreezePanes_BothRowsAndColumns_SetsFrozenCounts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('C2');

  Assert.AreEqual(1, Sheet.FrozenRows, 'Row 1 should be frozen (topLeftCell row - 1)');
  Assert.AreEqual(2, Sheet.FrozenColumns, 'Columns A and B should be frozen (topLeftCell col - 1)');
end;

procedure TExcelFreezePanesTests.FreezePanes_RowsOnly_LeavesColumnsAtZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('A2');

  Assert.AreEqual(1, Sheet.FrozenRows, 'Row 1 should be frozen');
  Assert.AreEqual(0, Sheet.FrozenColumns, 'No columns should be frozen');
end;

procedure TExcelFreezePanesTests.FreezePanes_ColumnsOnly_LeavesRowsAtZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('B1');

  Assert.AreEqual(0, Sheet.FrozenRows, 'No rows should be frozen');
  Assert.AreEqual(1, Sheet.FrozenColumns, 'Column A should be frozen');
end;

procedure TExcelFreezePanesTests.FreezePanes_TopLeftCorner_IsEquivalentToNoFreeze;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  // Freezing at A1 (the top-left corner itself) has nothing above/left of it to freeze --
  // this matches Excel's own UI, where selecting A1 and choosing "Freeze Panes" is a no-op.
  Sheet.FreezePanes('A1');

  Assert.AreEqual(0, Sheet.FrozenRows);
  Assert.AreEqual(0, Sheet.FrozenColumns);
end;

procedure TExcelFreezePanesTests.FreezePanes_InvalidAddress_RaisesException;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.WillRaise(
    procedure
    begin
      Sheet.FreezePanes('NotACell');
    end,
    EExcelWorkbookException,
    'An unparseable cell address must raise rather than silently freezing nothing or defaulting'
  );
end;

procedure TExcelFreezePanesTests.UnfreezePanes_AfterFreezing_ClearsFrozenState;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.FreezePanes('C2');
  Sheet.UnfreezePanes;

  Assert.AreEqual(0, Sheet.FrozenRows);
  Assert.AreEqual(0, Sheet.FrozenColumns);
end;

procedure TExcelFreezePanesTests.RoundTrip_FreezeBothRowsAndColumns_PreservesFrozenCounts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Header';
  Sheet.FreezePanes('C2');

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  Assert.AreEqual(1, Sheet2.FrozenRows, 'Frozen row count should survive round-trip');
  Assert.AreEqual(2, Sheet2.FrozenColumns, 'Frozen column count should survive round-trip');
end;

procedure TExcelFreezePanesTests.SaveToFile_WithFreezePanesAndColumnWidths_SheetViewsPrecedesCols;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('B2');
  Sheet.SetColumnWidth('A', 20);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    const SheetViewsPos = Pos('<sheetViews>', SheetXml);
    const ColsPos = Pos('<cols>', SheetXml);

    Assert.IsTrue(SheetViewsPos > 0, 'sheetViews element should be present');
    Assert.IsTrue(ColsPos > 0, 'cols element should be present');
    // OOXML's schema is order-sensitive here -- sheetViews must precede cols, or the
    // file can render incorrectly (or not open) in strict readers even though it's
    // well-formed XML. This ordering bug has bitten other libraries in exactly this spot.
    Assert.IsTrue(SheetViewsPos < ColsPos, 'sheetViews must be written before cols in the worksheet XML');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_WithFreezePanes_ContainsCorrectPaneAttributes;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('C2');

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('xSplit="2"', SheetXml) > 0, 'xSplit should equal the frozen column count');
    Assert.IsTrue(Pos('ySplit="1"', SheetXml) > 0, 'ySplit should equal the frozen row count');
    Assert.IsTrue(Pos('topLeftCell="C2"', SheetXml) > 0, 'topLeftCell should be the first scrollable cell');
    Assert.IsTrue(Pos('state="frozen"', SheetXml) > 0, 'state should be frozen, not split');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeRowsOnly_UsesBottomLeftActivePane;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('A2'); // rows only

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    // Freezing rows only puts the active pane at bottom-left, not bottom-right -- getting
    // this wrong doesn't corrupt the file, but the scroll/selection behavior around the
    // freeze boundary misbehaves in Excel.
    Assert.IsTrue(Pos('activePane="bottomLeft"', SheetXml) > 0,
      'Row-only freeze must use activePane="bottomLeft"');
    Assert.IsFalse(Pos('activePane="bottomRight"', SheetXml) > 0,
      'Row-only freeze must not use the both-axes activePane value');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeColumnsOnly_UsesTopRightActivePane;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('B1'); // columns only

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('activePane="topRight"', SheetXml) > 0,
      'Column-only freeze must use activePane="topRight"');
    Assert.IsFalse(Pos('activePane="bottomRight"', SheetXml) > 0,
      'Column-only freeze must not use the both-axes activePane value');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeBothAxes_EmitsThreeSelectionsAfterPane;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('C2'); // both axes -> bottomRight active

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');

    // A both-axes freeze mirrors what Excel writes: one selection per pane, each pointing at
    // the top-left cell of its own region.
    Assert.AreEqual(3, TRegEx.Matches(SheetXml, '<selection\b').Count,
      'A both-axes freeze must emit exactly three selection elements');
    Assert.IsTrue(Pos('<selection pane="topRight" activeCell="C1" sqref="C1"/>', SheetXml) > 0,
      'topRight selection should sit on the first frozen-below row, first scrollable column');
    Assert.IsTrue(Pos('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', SheetXml) > 0,
      'bottomLeft selection should sit on column A, first scrollable row');
    Assert.IsTrue(Pos('<selection pane="bottomRight" activeCell="C2" sqref="C2"/>', SheetXml) > 0,
      'bottomRight selection should sit on the first fully scrollable cell');

    // CT_SheetView requires pane before selection; a selection ahead of the pane is invalid.
    Assert.IsTrue(Pos('<pane ', SheetXml) < Pos('<selection ', SheetXml),
      'The pane element must precede the selection elements');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeRowsOnly_EmitsSingleBottomLeftSelection;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('A2'); // rows only -> bottomLeft active

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');

    Assert.AreEqual(1, TRegEx.Matches(SheetXml, '<selection\b').Count,
      'A single-axis freeze must emit exactly one selection element');
    Assert.IsTrue(Pos('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', SheetXml) > 0,
      'Row-only freeze selection should live in the bottomLeft pane at the first scrollable row');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.SaveToFile_FreezeColumnsOnly_EmitsSingleTopRightSelection;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.FreezePanes('B1'); // columns only -> topRight active

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');

    Assert.AreEqual(1, TRegEx.Matches(SheetXml, '<selection\b').Count,
      'A single-axis freeze must emit exactly one selection element');
    Assert.IsTrue(Pos('<selection pane="topRight" activeCell="B1" sqref="B1"/>', SheetXml) > 0,
      'Column-only freeze selection should live in the topRight pane at the first scrollable column');
  finally
    Package.Free;
  end;
end;

procedure TExcelFreezePanesTests.LoadFromFile_SplitPane_IsNotReadAsFrozen;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  FWorkbook.SaveToFile(FTempFile);

  var OriginalXml := '';
  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    OriginalXml := Package.GetPartContent('xl/worksheets/sheet1.xml');
  finally
    Package.Free;
  end;

  // An unfrozen split stores the split-bar position in twentieths of a point (2160 = 108pt),
  // not a row/column count. Reading it as frozen would report 2160 frozen columns.
  const SplitViews =
    '<sheetViews><sheetView workbookViewId="0">' +
    '<pane xSplit="2160" ySplit="1440" topLeftCell="D5" activePane="bottomRight" state="split"/>' +
    '</sheetView></sheetViews>';
  ReplaceSheet1Xml(FTempFile, StringReplace(OriginalXml, '<sheetData>', SplitViews + '<sheetData>', []));

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  Assert.AreEqual(0, Sheet2.FrozenColumns, 'A split pane must not be interpreted as frozen columns');
  Assert.AreEqual(0, Sheet2.FrozenRows, 'A split pane must not be interpreted as frozen rows');
end;

procedure TExcelFreezePanesTests.LoadFromFile_FrozenSplitPane_ReadsFrozenCounts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  FWorkbook.SaveToFile(FTempFile);

  var OriginalXml := '';
  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    OriginalXml := Package.GetPartContent('xl/worksheets/sheet1.xml');
  finally
    Package.Free;
  end;

  // state="frozenSplit" still encodes xSplit/ySplit as whole counts, so it must be read back
  // like a plain frozen pane.
  const FrozenSplitViews =
    '<sheetViews><sheetView workbookViewId="0">' +
    '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozenSplit"/>' +
    '</sheetView></sheetViews>';
  ReplaceSheet1Xml(FTempFile, StringReplace(OriginalXml, '<sheetData>', FrozenSplitViews + '<sheetData>', []));

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  Assert.AreEqual(2, Sheet2.FrozenColumns, 'frozenSplit xSplit should be read as a frozen column count');
  Assert.AreEqual(1, Sheet2.FrozenRows, 'frozenSplit ySplit should be read as a frozen row count');
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelFreezePanesTests);

end.
