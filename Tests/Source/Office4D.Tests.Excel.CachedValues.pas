unit Office4D.Tests.Excel.CachedValues;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelCachedValueTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

    Function ReadSheet1Xml: string;
    Function SaveAndLoad: IExcelWorkbook;
  public
    [Setup]
    Procedure Setup;

    [TearDown]
    Procedure TearDown;

    [Test]
    Procedure SetFormula_HasCachedValue_IsTrue;

    [Test]
    Procedure SetFormulaWithoutValue_HasCachedValue_IsFalse;

    [Test]
    Procedure ValueCell_HasCachedValue_IsFalse;

    [Test]
    Procedure SetCachedValue_FormulaCell_ReplacesValueAndKeepsFormula;

    [Test]
    Procedure SetCachedValue_CellWithoutFormula_Raises;

    [Test]
    Procedure ClearCachedValue_FormulaCell_KeepsFormula;

    [Test]
    Procedure ClearCachedValue_CellWithoutFormula_Raises;

    [Test]
    Procedure SetCachedValue_AfterClear_RestoresCachedValue;

    [Test]
    Procedure AssigningValue_FormulaCell_HasCachedValueIsFalse;

    [Test]
    Procedure SheetClearCachedValues_LeavesValueCellsIntact;

    [Test]
    Procedure WorkbookClearCachedValues_ClearsEverySheet;

    [Test]
    Procedure SaveToFile_ClearedCachedValue_OmitsValueElement;

    [Test]
    Procedure RoundTrip_ClearedCachedValue_KeepsFormulaWithoutValue;

    [Test]
    Procedure RoundTrip_ClearedCachedValue_KeepsStyle;

    [Test]
    Procedure RoundTrip_SetCachedValue_IsPreserved;
  end;

implementation

uses
  Office4D.Errors,
  Office4D.Package;

Procedure TExcelCachedValueTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'cachedvalue_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

Procedure TExcelCachedValueTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

Function TExcelCachedValueTests.ReadSheet1Xml: string;
begin
  FWorkbook.SaveToFile(FTempFile);
  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Result := Package.GetPartContent('xl/worksheets/sheet1.xml');
  finally
    Package.Free;
  end;
end;

Function TExcelCachedValueTests.SaveAndLoad: IExcelWorkbook;
begin
  FWorkbook.SaveToFile(FTempFile);
  Result := TExcelWorkbookFactory.Create;
  Result.LoadFromFile(FTempFile);
end;

Procedure TExcelCachedValueTests.SetFormula_HasCachedValue_IsTrue;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);

  Assert.IsTrue(Sheet.Cell['B1'].HasCachedValue);
end;

Procedure TExcelCachedValueTests.SetFormulaWithoutValue_HasCachedValue_IsFalse;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].AsString := 'Replaced';

  Sheet.Cell['B1'].SetFormula('A1*2');

  Assert.IsTrue(Sheet.Cell['B1'].HasFormula);
  Assert.AreEqual('A1*2', Sheet.Cell['B1'].Formula);
  Assert.IsFalse(Sheet.Cell['B1'].HasCachedValue);
  Assert.AreEqual('', Sheet.Cell['B1'].AsString, 'The previous value must not linger');
end;

Procedure TExcelCachedValueTests.ValueCell_HasCachedValue_IsFalse;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 10;

  Assert.IsFalse(Sheet.Cell['A1'].HasCachedValue, 'Only a formula cell carries a cached value');
end;

Procedure TExcelCachedValueTests.SetCachedValue_FormulaCell_ReplacesValueAndKeepsFormula;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);

  Sheet.Cell['B1'].SetCachedValue(42.5);

  Assert.AreEqual(Double(42.5), Sheet.Cell['B1'].AsFloat);
  Assert.AreEqual('A1*2', Sheet.Cell['B1'].Formula);
  Assert.IsTrue(Sheet.Cell['B1'].HasCachedValue);
end;

Procedure TExcelCachedValueTests.SetCachedValue_CellWithoutFormula_Raises;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 10;

  Assert.WillRaise(
    Procedure
    begin
      Sheet.Cell['A1'].SetCachedValue(42);
    end,
    EExcelWorkbookException,
    'A cell without a formula must refuse a cached value');
  Assert.AreEqual(Double(10), Sheet.Cell['A1'].AsFloat, 'The refused call must leave the value alone');
end;

Procedure TExcelCachedValueTests.ClearCachedValue_FormulaCell_KeepsFormula;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);

  Sheet.Cell['B1'].ClearCachedValue;

  Assert.IsFalse(Sheet.Cell['B1'].HasCachedValue);
  Assert.IsTrue(Sheet.Cell['B1'].HasFormula);
  Assert.AreEqual('A1*2', Sheet.Cell['B1'].Formula);
  Assert.AreEqual(Double(0), Sheet.Cell['B1'].AsFloat);
end;

Procedure TExcelCachedValueTests.ClearCachedValue_CellWithoutFormula_Raises;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 10;

  Assert.WillRaise(
    Procedure
    begin
      Sheet.Cell['A1'].ClearCachedValue;
    end,
    EExcelWorkbookException,
    'A cell without a formula has no cached value to clear');
  Assert.AreEqual(Double(10), Sheet.Cell['A1'].AsFloat, 'The refused call must leave the value alone');
end;

Procedure TExcelCachedValueTests.SetCachedValue_AfterClear_RestoresCachedValue;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B1'].ClearCachedValue;

  Sheet.Cell['B1'].SetCachedValue(30);

  Assert.IsTrue(Sheet.Cell['B1'].HasCachedValue);
  Assert.AreEqual(Double(30), Sheet.Cell['B1'].AsFloat);
end;

Procedure TExcelCachedValueTests.AssigningValue_FormulaCell_HasCachedValueIsFalse;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);

  Sheet.Cell['B1'].AsFloat := 7;

  Assert.IsFalse(Sheet.Cell['B1'].HasFormula);
  Assert.IsFalse(Sheet.Cell['B1'].HasCachedValue, 'A value replaces the formula and its cached value');
end;

Procedure TExcelCachedValueTests.SheetClearCachedValues_LeavesValueCellsIntact;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 10;
  Sheet.Cell['A2'].AsString := 'Label';
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B2'].SetFormula('A1+1', 11);

  Sheet.ClearCachedValues;

  Assert.IsFalse(Sheet.Cell['B1'].HasCachedValue, 'B1');
  Assert.IsFalse(Sheet.Cell['B2'].HasCachedValue, 'B2');
  Assert.AreEqual(Double(10), Sheet.Cell['A1'].AsFloat);
  Assert.AreEqual('Label', Sheet.Cell['A2'].AsString);
end;

Procedure TExcelCachedValueTests.WorkbookClearCachedValues_ClearsEverySheet;
begin
  const First = FWorkbook.AddSheet('First');
  const Second = FWorkbook.AddSheet('Second');
  First.Cell['A1'].SetFormula('1+1', 2);
  Second.Cell['A1'].SetFormula('First!A1*2', 4);

  FWorkbook.ClearCachedValues;

  Assert.IsFalse(First.Cell['A1'].HasCachedValue, 'First!A1');
  Assert.IsFalse(Second.Cell['A1'].HasCachedValue, 'Second!A1');
end;

Procedure TExcelCachedValueTests.SaveToFile_ClearedCachedValue_OmitsValueElement;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B1'].ClearCachedValue;

  const SheetXml = ReadSheet1Xml;

  const HasBareFormula = (Pos('<c r="B1"><f>A1*2</f></c>', SheetXml) > 0);
  Assert.IsTrue(HasBareFormula, 'A formula without a cached value is written without <v>');
end;

Procedure TExcelCachedValueTests.RoundTrip_ClearedCachedValue_KeepsFormulaWithoutValue;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B1'].ClearCachedValue;
  Sheet.Cell['C1'].AsFloat := 5;

  const Loaded = SaveAndLoad;

  const LoadedSheet = Loaded.Sheets[0];
  Assert.IsTrue(LoadedSheet.Cell['B1'].HasFormula, 'The formula must survive a reload');
  Assert.AreEqual('A1*2', LoadedSheet.Cell['B1'].Formula);
  Assert.IsFalse(LoadedSheet.Cell['B1'].HasCachedValue);
  Assert.AreEqual(Double(5), LoadedSheet.Cell['C1'].AsFloat, 'The next cell keeps its own value');
end;

Procedure TExcelCachedValueTests.RoundTrip_ClearedCachedValue_KeepsStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B1'].Bold := True;
  Sheet.Cell['B1'].ClearCachedValue;

  const Loaded = SaveAndLoad;

  Assert.IsTrue(Loaded.Sheets[0].Cell['B1'].Bold);
end;

Procedure TExcelCachedValueTests.RoundTrip_SetCachedValue_IsPreserved;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].SetFormula('A1*2', 20);
  Sheet.Cell['B1'].SetCachedValue(42.5);

  const Loaded = SaveAndLoad;

  Assert.IsTrue(Loaded.Sheets[0].Cell['B1'].HasCachedValue);
  Assert.AreEqual(Double(42.5), Loaded.Sheets[0].Cell['B1'].AsFloat);
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelCachedValueTests);

end.
