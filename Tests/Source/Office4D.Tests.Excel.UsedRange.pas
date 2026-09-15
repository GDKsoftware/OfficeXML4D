unit Office4D.Tests.Excel.UsedRange;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelUsedRangeTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

    function ReadSheet1Xml(const FileName: string): string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure UsedRange_EmptySheet_IsEmptyString;

    [Test]
    procedure LastRow_EmptySheet_IsZero;

    [Test]
    procedure LastColumn_EmptySheet_IsZero;

    [Test]
    procedure UsedRange_SingleCell_IsThatCell;

    [Test]
    procedure UsedRange_CellsStartingPastA1_StartsAtFirstUsedCell;

    [Test]
    procedure LastRow_BlankRowInsideData_IsNotTruncated;

    [Test]
    procedure LastColumn_MultiLetterColumn_IsParsedCorrectly;

    [Test]
    procedure UsedRange_AfterDeleteRow_Shrinks;

    [Test]
    procedure UsedRange_AfterClearColumn_Shrinks;

    [Test]
    procedure UsedRange_LoadedSample_MatchesCells;

    [Test]
    procedure SaveToFile_WithCells_WritesDimension;

    [Test]
    procedure SaveToFile_EmptySheet_OmitsDimension;

    [Test]
    procedure RoundTrip_UsedRange_IsPreserved;
  end;

implementation

uses
  Office4D.Package;

{ TExcelUsedRangeTests }

procedure TExcelUsedRangeTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'usedrange_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelUsedRangeTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

function TExcelUsedRangeTests.ReadSheet1Xml(const FileName: string): string;
begin
  var Package := TOXMLPackage.Create;
  try
    Package.Open(FileName);
    Result := Package.GetPartContent('xl/worksheets/sheet1.xml');
  finally
    Package.Free;
  end;
end;

procedure TExcelUsedRangeTests.UsedRange_EmptySheet_IsEmptyString;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.AreEqual('', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.LastRow_EmptySheet_IsZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.AreEqual(0, Sheet.LastRow);
end;

procedure TExcelUsedRangeTests.LastColumn_EmptySheet_IsZero;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.AreEqual(0, Sheet.LastColumn);
end;

procedure TExcelUsedRangeTests.UsedRange_SingleCell_IsThatCell;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['C7'].AsString := 'x';

  Assert.AreEqual('C7:C7', Sheet.UsedRange);
  Assert.AreEqual(7, Sheet.LastRow);
  Assert.AreEqual(3, Sheet.LastColumn);
end;

procedure TExcelUsedRangeTests.UsedRange_CellsStartingPastA1_StartsAtFirstUsedCell;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B2'].AsString := 'top left';
  Sheet.Cell['D5'].AsString := 'bottom right';

  Assert.AreEqual('B2:D5', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.LastRow_BlankRowInsideData_IsNotTruncated;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'first';
  Sheet.Cell['A3'].AsString := 'after a blank row';

  Assert.AreEqual(3, Sheet.LastRow, 'A blank row inside the data must not end the used range');
end;

procedure TExcelUsedRangeTests.LastColumn_MultiLetterColumn_IsParsedCorrectly;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'a';
  Sheet.Cell['AB3'].AsString := 'ab';

  Assert.AreEqual(28, Sheet.LastColumn, 'AB is column 28');
  Assert.AreEqual('A1:AB3', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.UsedRange_AfterDeleteRow_Shrinks;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'a';
  Sheet.Cell['A2'].AsString := 'b';
  Sheet.Cell['A3'].AsString := 'c';

  Sheet.DeleteRow(2);

  Assert.AreEqual(2, Sheet.LastRow);
  Assert.AreEqual('A1:A2', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.UsedRange_AfterClearColumn_Shrinks;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'a';
  Sheet.Cell['C1'].AsString := 'c';

  Sheet.ClearColumn('C');

  Assert.AreEqual(1, Sheet.LastColumn);
  Assert.AreEqual('A1:A1', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.UsedRange_LoadedSample_MatchesCells;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);
  const Sheet = FWorkbook.Sheets[0];

  const HasCells = (Sheet.Cells.Count > 0);
  Assert.IsTrue(HasCells, 'The sample sheet is expected to contain cells');
  Assert.IsTrue(Sheet.LastRow > 0);
  Assert.IsTrue(Sheet.LastColumn > 0);
  Assert.AreNotEqual('', Sheet.UsedRange);
end;

procedure TExcelUsedRangeTests.SaveToFile_WithCells_WritesDimension;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B2'].AsString := 'x';
  Sheet.Cell['E9'].AsFloat := 1;

  FWorkbook.SaveToFile(FTempFile);

  const SheetXml = ReadSheet1Xml(FTempFile);
  const HasDimension = (Pos('<dimension ref="B2:E9"/>', SheetXml) > 0);
  Assert.IsTrue(HasDimension, 'sheet1.xml must carry the used range as its dimension');
end;

procedure TExcelUsedRangeTests.SaveToFile_EmptySheet_OmitsDimension;
begin
  FWorkbook.AddSheet('Sheet1');

  FWorkbook.SaveToFile(FTempFile);

  const SheetXml = ReadSheet1Xml(FTempFile);
  const HasDimension = (Pos('<dimension', SheetXml) > 0);
  Assert.IsFalse(HasDimension, 'An empty sheet has no used range and must not emit a dimension');
end;

procedure TExcelUsedRangeTests.RoundTrip_UsedRange_IsPreserved;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B2'].AsString := 'x';
  Sheet.Cell['AA40'].AsFloat := 2.5;
  FWorkbook.SaveToFile(FTempFile);

  const Loaded = TExcelWorkbookFactory.Create;
  Loaded.LoadFromFile(FTempFile);

  Assert.AreEqual('B2:AA40', Loaded.Sheets[0].UsedRange);
  Assert.AreEqual(40, Loaded.Sheets[0].LastRow);
  Assert.AreEqual(27, Loaded.Sheets[0].LastColumn);
end;

end.
