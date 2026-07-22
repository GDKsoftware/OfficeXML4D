unit Office4D.Tests.Excel.DeleteColumnRow;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelDeleteColumnRowTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FSheet: IExcelSheet;
    FTempFile: string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    // --- Reflow of cells and structure ---
    [Test]
    procedure DeleteColumn_ShiftsCellsLeftAndDropsDeleted;

    [Test]
    procedure DeleteColumn_MovesCellStyleWithTheCell;

    [Test]
    procedure DeleteColumn_ShiftsColumnWidths;

    [Test]
    procedure DeleteRow_ShiftsCellsUpAndDropsDeleted;

    [Test]
    procedure DeleteRow_ShiftsRowHeights;

    [Test]
    procedure DeleteColumn_ShrinksFrozenColumns_LeavesRows;

    [Test]
    procedure DeleteColumn_MergedRangeShrinks;

    [Test]
    procedure DeleteColumn_MergedRangeFullyOnColumn_IsDropped;

    [Test]
    procedure RoundTrip_DeleteColumn_Persists;

    // --- Formula reference rewriting ---
    [Test]
    procedure DeleteColumn_Formula_RefAfterDeleted_ShiftsLeft;

    [Test]
    procedure DeleteColumn_Formula_RefBeforeDeleted_Unchanged;

    [Test]
    procedure DeleteColumn_Formula_RefOnDeleted_BecomesRefError;

    [Test]
    procedure DeleteColumn_Formula_AbsoluteRef_IsAlsoShifted;

    [Test]
    procedure DeleteColumn_Formula_RangeShrinks;

    [Test]
    procedure DeleteColumn_Formula_RangeEndOnDeleted_ClampsNotRefError;

    [Test]
    procedure DeleteColumn_Formula_RangeFullyOnColumn_BecomesRefError;

    [Test]
    procedure DeleteColumn_Formula_CrossSheetRef_IsShifted;

    [Test]
    procedure DeleteColumn_Formula_RefToOtherSheet_Untouched;

    [Test]
    procedure DeleteColumn_Formula_StringLiteral_Untouched;

    [Test]
    procedure DeleteRow_Formula_RefBelowDeleted_ShiftsUp;
  end;

implementation

{ TExcelDeleteColumnRowTests }

procedure TExcelDeleteColumnRowTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FSheet := FWorkbook.AddSheet('Sheet1');
  FTempFile := TPath.Combine(TPath.GetTempPath, 'deletecolrow_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelDeleteColumnRowTests.TearDown;
begin
  FSheet := nil;
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_ShiftsCellsLeftAndDropsDeleted;
begin
  FSheet.Cell['A1'].AsString := 'a';
  FSheet.Cell['B1'].AsString := 'b';
  FSheet.Cell['C1'].AsString := 'c';
  FSheet.Cell['D1'].AsString := 'd';

  FSheet.DeleteColumn('B');

  Assert.AreEqual('a', FSheet.Cell['A1'].AsString, 'Column before the deleted one stays put');
  Assert.AreEqual('c', FSheet.Cell['B1'].AsString, 'Column C shifts into B');
  Assert.AreEqual('d', FSheet.Cell['C1'].AsString, 'Column D shifts into C');
  Assert.IsFalse(FSheet.Cells.ContainsKey('D1'), 'The trailing column must be empty after the shift');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_MovesCellStyleWithTheCell;
begin
  FSheet.Cell['C1'].AsString := 'x';
  FSheet.Cell['C1'].Bold := True;

  FSheet.DeleteColumn('B');

  Assert.AreEqual('x', FSheet.Cell['B1'].AsString, 'C1 moves to B1');
  Assert.IsTrue(FSheet.Cell['B1'].Bold, 'Styling travels with the cell');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_ShiftsColumnWidths;
begin
  FSheet.SetColumnWidth('B', 10);
  FSheet.SetColumnWidth('C', 20);
  FSheet.SetColumnWidth('D', 30);

  FSheet.DeleteColumn('C');

  Assert.AreEqual(Double(10), FSheet.GetColumnWidth('B'), 0.001, 'Width before the deleted column is unchanged');
  Assert.AreEqual(Double(30), FSheet.GetColumnWidth('C'), 0.001, 'Width of column D shifts into C');
end;

procedure TExcelDeleteColumnRowTests.DeleteRow_ShiftsCellsUpAndDropsDeleted;
begin
  FSheet.Cell['A1'].AsString := 'r1';
  FSheet.Cell['A2'].AsString := 'r2';
  FSheet.Cell['A3'].AsString := 'r3';

  FSheet.DeleteRow(2);

  Assert.AreEqual('r1', FSheet.Cell['A1'].AsString);
  Assert.AreEqual('r3', FSheet.Cell['A2'].AsString, 'Row 3 shifts up into row 2');
  Assert.IsFalse(FSheet.Cells.ContainsKey('A3'), 'The trailing row must be empty after the shift');
end;

procedure TExcelDeleteColumnRowTests.DeleteRow_ShiftsRowHeights;
begin
  FSheet.SetRowHeight(1, 10);
  FSheet.SetRowHeight(2, 20);
  FSheet.SetRowHeight(3, 30);

  FSheet.DeleteRow(2);

  Assert.AreEqual(Double(10), FSheet.GetRowHeight(1), 0.001);
  Assert.AreEqual(Double(30), FSheet.GetRowHeight(2), 0.001, 'Height of row 3 shifts up into row 2');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_ShrinksFrozenColumns_LeavesRows;
begin
  FSheet.FreezePanes('C2'); // 2 frozen columns, 1 frozen row

  FSheet.DeleteColumn('A'); // inside the frozen region

  Assert.AreEqual(1, FSheet.FrozenColumns, 'Deleting a frozen column shrinks the frozen column count');
  Assert.AreEqual(1, FSheet.FrozenRows, 'The frozen row count is unaffected by a column delete');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_MergedRangeShrinks;
begin
  FSheet.MergeCells('B1:D1');

  FSheet.DeleteColumn('C');

  Assert.AreEqual(1, Integer(Length(FSheet.GetMergedRanges)));
  Assert.AreEqual('B1:C1', FSheet.GetMergedRanges[0], 'The merged range shrinks by one column');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_MergedRangeFullyOnColumn_IsDropped;
begin
  FSheet.MergeCells('C1:C3');

  FSheet.DeleteColumn('C');

  Assert.AreEqual(0, Integer(Length(FSheet.GetMergedRanges)), 'A merge entirely on the deleted column disappears');
end;

procedure TExcelDeleteColumnRowTests.RoundTrip_DeleteColumn_Persists;
begin
  FSheet.Cell['A1'].AsString := 'a';
  FSheet.Cell['B1'].AsString := 'b';
  FSheet.Cell['C1'].AsString := 'c';

  FSheet.DeleteColumn('B');
  FWorkbook.SaveToFile(FTempFile);

  const Reloaded = TExcelWorkbookFactory.Create;
  Reloaded.LoadFromFile(FTempFile);
  const Sheet = Reloaded.Sheets[0];

  Assert.AreEqual('a', Sheet.Cell['A1'].AsString);
  Assert.AreEqual('c', Sheet.Cell['B1'].AsString, 'The shifted layout must survive a save/load round-trip');
  Assert.IsFalse(Sheet.Cells.ContainsKey('C1'));
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RefAfterDeleted_ShiftsLeft;
begin
  FSheet.Cell['A1'].SetFormula('D1', 0);

  FSheet.DeleteColumn('C');

  Assert.AreEqual('C1', FSheet.Cell['A1'].Formula, 'A reference past the deleted column shifts left');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RefBeforeDeleted_Unchanged;
begin
  FSheet.Cell['A1'].SetFormula('B1', 0);

  FSheet.DeleteColumn('C');

  Assert.AreEqual('B1', FSheet.Cell['A1'].Formula, 'A reference before the deleted column is untouched');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RefOnDeleted_BecomesRefError;
begin
  FSheet.Cell['A1'].SetFormula('C1', 0);

  FSheet.DeleteColumn('C');

  Assert.AreEqual('#REF!', FSheet.Cell['A1'].Formula, 'A reference to the deleted column becomes #REF!');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_AbsoluteRef_IsAlsoShifted;
begin
  FSheet.Cell['A1'].SetFormula('$D$1', 0);

  FSheet.DeleteColumn('C');

  // Unlike a copy, a delete shifts absolute references too.
  Assert.AreEqual('$C$1', FSheet.Cell['A1'].Formula, 'An absolute reference is shifted on delete');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RangeShrinks;
begin
  FSheet.Cell['A1'].SetFormula('SUM(B1:D1)', 0);

  FSheet.DeleteColumn('C');

  Assert.AreEqual('SUM(B1:C1)', FSheet.Cell['A1'].Formula, 'A range spanning the deleted column shrinks');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RangeEndOnDeleted_ClampsNotRefError;
begin
  FSheet.Cell['A1'].SetFormula('SUM(A2:C2)', 0);

  FSheet.DeleteColumn('C');

  // The far edge sat on the deleted column, so the range clamps to the column before it.
  Assert.AreEqual('SUM(A2:B2)', FSheet.Cell['A1'].Formula);
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RangeFullyOnColumn_BecomesRefError;
begin
  FSheet.Cell['A1'].SetFormula('SUM(C1:C5)', 0);

  FSheet.DeleteColumn('C');

  // Excel collapses the range itself to #REF! but keeps the surrounding function.
  Assert.AreEqual('SUM(#REF!)', FSheet.Cell['A1'].Formula, 'A range entirely on the deleted column becomes #REF!');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_CrossSheetRef_IsShifted;
begin
  const Sheet2 = FWorkbook.AddSheet('Sheet2');
  FSheet.Cell['A1'].SetFormula('Sheet2!D1', 0);

  Sheet2.DeleteColumn('C');

  Assert.AreEqual('Sheet2!C1', FSheet.Cell['A1'].Formula, 'A cross-sheet reference to the edited sheet shifts too');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_RefToOtherSheet_Untouched;
begin
  FWorkbook.AddSheet('Sheet2');
  FSheet.Cell['A1'].SetFormula('Sheet2!D1', 0);

  FSheet.DeleteColumn('C'); // delete on Sheet1, not Sheet2

  Assert.AreEqual('Sheet2!D1', FSheet.Cell['A1'].Formula, 'A reference to a different sheet is left alone');
end;

procedure TExcelDeleteColumnRowTests.DeleteColumn_Formula_StringLiteral_Untouched;
begin
  FSheet.Cell['A1'].SetFormula('CONCATENATE("see C3",D1)', 0);

  FSheet.DeleteColumn('C');

  // The literal text "see C3" must not be treated as a reference; D1 still shifts to C1.
  Assert.AreEqual('CONCATENATE("see C3",C1)', FSheet.Cell['A1'].Formula);
end;

procedure TExcelDeleteColumnRowTests.DeleteRow_Formula_RefBelowDeleted_ShiftsUp;
begin
  FSheet.Cell['A1'].SetFormula('A5', 0);

  FSheet.DeleteRow(3);

  Assert.AreEqual('A4', FSheet.Cell['A1'].Formula, 'A reference below the deleted row shifts up');
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelDeleteColumnRowTests);

end.
