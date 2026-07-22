unit Office4D.Tests.Excel.ClearColumnRow;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelClearColumnRowTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FSheet: IExcelSheet;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure ClearColumn_RemovesColumnCells_KeepsNeighbours;

    [Test]
    procedure ClearRow_RemovesRowCells_KeepsNeighbours;

    [Test]
    procedure ClearColumn_KeepsColumnWidthAndMerges;

    [Test]
    procedure ClearColumn_NonExisting_IsNoOp;

    [Test]
    procedure ClearRow_NonExisting_IsNoOp;

    [Test]
    procedure ClearColumn_RemovesNoteOnThatColumn;
  end;

implementation

{ TExcelClearColumnRowTests }

procedure TExcelClearColumnRowTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FSheet := FWorkbook.AddSheet('Sheet1');
end;

procedure TExcelClearColumnRowTests.TearDown;
begin
  FSheet := nil;
  FWorkbook := nil;
end;

procedure TExcelClearColumnRowTests.ClearColumn_RemovesColumnCells_KeepsNeighbours;
begin
  FSheet.Cell['A1'].AsString := 'a1';
  FSheet.Cell['B1'].AsString := 'b1';
  FSheet.Cell['B2'].AsString := 'b2';
  FSheet.Cell['C1'].AsString := 'c1';

  FSheet.ClearColumn('B');

  Assert.IsFalse(FSheet.Cells.ContainsKey('B1'), 'B1 must be cleared');
  Assert.IsFalse(FSheet.Cells.ContainsKey('B2'), 'B2 must be cleared');
  Assert.IsTrue(FSheet.Cells.ContainsKey('A1'), 'The neighbouring column A must be untouched');
  Assert.IsTrue(FSheet.Cells.ContainsKey('C1'), 'The neighbouring column C must be untouched');
end;

procedure TExcelClearColumnRowTests.ClearRow_RemovesRowCells_KeepsNeighbours;
begin
  FSheet.Cell['A1'].AsString := 'a1';
  FSheet.Cell['B1'].AsString := 'b1';
  FSheet.Cell['A2'].AsString := 'a2';
  FSheet.Cell['A3'].AsString := 'a3';

  FSheet.ClearRow(2);

  Assert.IsFalse(FSheet.Cells.ContainsKey('A2'), 'A2 must be cleared');
  Assert.IsTrue(FSheet.Cells.ContainsKey('A1'), 'Row 1 must be untouched');
  Assert.IsTrue(FSheet.Cells.ContainsKey('B1'), 'Row 1 must be untouched');
  Assert.IsTrue(FSheet.Cells.ContainsKey('A3'), 'Row 3 must be untouched');
end;

procedure TExcelClearColumnRowTests.ClearColumn_KeepsColumnWidthAndMerges;
begin
  FSheet.Cell['B1'].AsString := 'b1';
  FSheet.SetColumnWidth('B', 25);
  FSheet.MergeCells('B1:B3');

  FSheet.ClearColumn('B');

  // Clear removes contents, not structure: width and merges survive.
  Assert.AreEqual(Double(25), FSheet.GetColumnWidth('B'), 0.001, 'Clearing contents must keep the column width');
  Assert.AreEqual(1, Integer(Length(FSheet.GetMergedRanges)), 'Clearing contents must keep merged ranges');
end;

procedure TExcelClearColumnRowTests.ClearColumn_NonExisting_IsNoOp;
begin
  FSheet.Cell['A1'].AsString := 'a1';

  // Clearing an empty column must be a no-op; a raised exception would fail the test.
  FSheet.ClearColumn('Z');

  Assert.IsTrue(FSheet.Cells.ContainsKey('A1'));
end;

procedure TExcelClearColumnRowTests.ClearRow_NonExisting_IsNoOp;
begin
  FSheet.Cell['A1'].AsString := 'a1';

  // Clearing an empty row must be a no-op; a raised exception would fail the test.
  FSheet.ClearRow(99);

  Assert.IsTrue(FSheet.Cells.ContainsKey('A1'));
end;

procedure TExcelClearColumnRowTests.ClearColumn_RemovesNoteOnThatColumn;
begin
  FSheet.Cell['B1'].AsString := 'b1';
  FSheet.Note['B1'] := 'note on B';
  FSheet.Note['A1'] := 'note on A';

  FSheet.ClearColumn('B');

  Assert.AreEqual('', FSheet.Note['B1'], 'Clearing a column also clears its notes');
  Assert.AreEqual('note on A', FSheet.Note['A1'], 'A note on a neighbouring column is untouched');
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelClearColumnRowTests);

end.
