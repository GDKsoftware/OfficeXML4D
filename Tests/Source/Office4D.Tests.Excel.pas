unit Office4D.Tests.Excel;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelReadTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_ValidXlsx_LoadsWorkbook;

    [Test]
    procedure LoadFromFile_NonExistentFile_RaisesException;

    [Test]
    procedure SheetCount_AfterLoad_ReturnsCount;

    [Test]
    procedure SheetByIndex_ValidIndex_ReturnsSheet;

    [Test]
    procedure SheetByName_ValidName_ReturnsSheet;

    [Test]
    procedure Sheet_GetName_ReturnsSheetName;

    [Test]
    procedure Cell_StringValue_ReturnsString;

    [Test]
    procedure Cell_NumberValue_ReturnsNumber;

    [Test]
    procedure Metadata_LastModifiedBy_ReturnsAuthor;
  end;

  [TestFixture]
  TExcelDOMTests = class
  private
    FWorkbook: IExcelWorkbook;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure AddSheet_Empty_AddsToSheets;

    [Test]
    procedure AddSheet_WithName_SetsName;

    [Test]
    procedure SetCellValue_String_StoresValue;

    [Test]
    procedure SetCellValue_Number_StoresValue;

    [Test]
    procedure GetCellValue_EmptyCell_ReturnsEmpty;
  end;

  [TestFixture]
  TExcelAdvancedTests = class
  private
    FWorkbook: IExcelWorkbook;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Cell_BooleanTrue_ReturnsTrue;

    [Test]
    procedure Cell_BooleanFalse_ReturnsFalse;

    [Test]
    procedure Cell_AsBoolean_ReturnsBoolean;

    [Test]
    procedure Cell_DateValue_ReturnsDateTime;

    [Test]
    procedure Cell_AsDateTime_ReturnsDate;

    [Test]
    procedure Cell_DateRoundTrip_PreservesDate;

    [Test]
    procedure Cell_HasFormula_WhenNoFormula_ReturnsFalse;

    [Test]
    procedure Cell_Formula_WhenNoFormula_ReturnsEmpty;
  end;

  [TestFixture]
  TExcelFormulaTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Cell_WithFormula_HasFormulaIsTrue;

    [Test]
    procedure Cell_WithFormula_ReturnsFormulaString;

    [Test]
    procedure Cell_WithFormula_ReturnsCalculatedValue;

    [Test]
    procedure Cell_WithFormula_CrossSheetReference;

    [Test]
    procedure RoundTrip_Formula_PreservesFormula;

    [Test]
    procedure SetFormula_CreatesFormulaCell;

    [Test]
    procedure SetFormula_WithValue_StoresCalculatedValue;

    [Test]
    procedure SetFormula_RoundTrip_PreservesFormula;
  end;

implementation

uses
  Office4D.Errors;

{ TExcelReadTests }

procedure TExcelReadTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
end;

procedure TExcelReadTests.TearDown;
begin
  FWorkbook := nil;
end;

procedure TExcelReadTests.LoadFromFile_ValidXlsx_LoadsWorkbook;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  Assert.IsTrue(FWorkbook.SheetCount > 0);
end;

procedure TExcelReadTests.LoadFromFile_NonExistentFile_RaisesException;
begin
  Assert.WillRaise(
    procedure
    begin
      FWorkbook.LoadFromFile('C:\nonexistent\file.xlsx');
    end,
    EPackageNotFound
  );
end;

procedure TExcelReadTests.SheetCount_AfterLoad_ReturnsCount;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  Assert.AreEqual(1, FWorkbook.SheetCount);
end;

procedure TExcelReadTests.SheetByIndex_ValidIndex_ReturnsSheet;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  var Sheet := FWorkbook.Sheets[0];

  Assert.IsNotNull(Sheet);
end;

procedure TExcelReadTests.SheetByName_ValidName_ReturnsSheet;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  var Sheet := FWorkbook.SheetByName('Sheet1');

  Assert.IsNotNull(Sheet);
end;

procedure TExcelReadTests.Sheet_GetName_ReturnsSheetName;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  Assert.AreEqual('Sheet1', FWorkbook.Sheets[0].Name);
end;

procedure TExcelReadTests.Cell_StringValue_ReturnsString;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  var Value := FWorkbook.Sheets[0].Cell['A1'].AsString;

  Assert.AreEqual('Hello', Value);
end;

procedure TExcelReadTests.Cell_NumberValue_ReturnsNumber;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  var Value: Double := FWorkbook.Sheets[0].Cell['B1'].AsFloat;

  Assert.AreEqual(Double(42), Value);
end;

procedure TExcelReadTests.Metadata_LastModifiedBy_ReturnsAuthor;
begin
  FWorkbook.LoadFromFile(GetExcelSimpleSamplePath);

  Assert.AreEqual('Marco Geuze', FWorkbook.Metadata.LastModifiedBy);
end;

{ TExcelDOMTests }

procedure TExcelDOMTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
end;

procedure TExcelDOMTests.TearDown;
begin
  FWorkbook := nil;
end;

procedure TExcelDOMTests.AddSheet_Empty_AddsToSheets;
begin
  FWorkbook.AddSheet('Sheet1');

  Assert.AreEqual(1, FWorkbook.SheetCount);
end;

procedure TExcelDOMTests.AddSheet_WithName_SetsName;
begin
  var Sheet := FWorkbook.AddSheet('MySheet');

  Assert.AreEqual('MySheet', Sheet.Name);
end;

procedure TExcelDOMTests.SetCellValue_String_StoresValue;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Hello';

  Assert.AreEqual('Hello', Sheet.Cell['A1'].AsString);
end;

procedure TExcelDOMTests.SetCellValue_Number_StoresValue;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B1'].AsFloat := 123.45;

  Assert.AreEqual(Double(123.45), Sheet.Cell['B1'].AsFloat);
end;

procedure TExcelDOMTests.GetCellValue_EmptyCell_ReturnsEmpty;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');

  Assert.IsEmpty(Sheet.Cell['A1'].AsString);
end;

{ TExcelAdvancedTests }

procedure TExcelAdvancedTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
end;

procedure TExcelAdvancedTests.TearDown;
begin
  FWorkbook := nil;
end;

procedure TExcelAdvancedTests.Cell_BooleanTrue_ReturnsTrue;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsBoolean := True;

  Assert.IsTrue(Sheet.Cell['A1'].AsBoolean);
end;

procedure TExcelAdvancedTests.Cell_BooleanFalse_ReturnsFalse;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsBoolean := False;

  Assert.IsFalse(Sheet.Cell['A1'].AsBoolean);
end;

procedure TExcelAdvancedTests.Cell_AsBoolean_ReturnsBoolean;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsBoolean := True;
  Sheet.Cell['A2'].AsBoolean := False;

  Assert.AreEqual(True, Sheet.Cell['A1'].AsBoolean);
  Assert.AreEqual(False, Sheet.Cell['A2'].AsBoolean);
end;

procedure TExcelAdvancedTests.Cell_DateValue_ReturnsDateTime;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  var TestDate: TDateTime := EncodeDate(2024, 6, 15);
  Sheet.Cell['A1'].AsDateTime := TestDate;

  Assert.AreEqual(TestDate, Sheet.Cell['A1'].AsDateTime);
end;

procedure TExcelAdvancedTests.Cell_AsDateTime_ReturnsDate;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  var TestDate := EncodeDate(2024, 1, 1);
  Sheet.Cell['A1'].AsDateTime := TestDate;

  var Year, Month, Day: Word;
  DecodeDate(Sheet.Cell['A1'].AsDateTime, Year, Month, Day);

  Assert.AreEqual(Word(2024), Year);
  Assert.AreEqual(Word(1), Month);
  Assert.AreEqual(Word(1), Day);
end;

procedure TExcelAdvancedTests.Cell_DateRoundTrip_PreservesDate;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  var TestDate := EncodeDate(2024, 12, 25) + EncodeTime(14, 30, 0, 0);
  Sheet.Cell['A1'].AsDateTime := TestDate;

  Assert.AreEqual(TestDate, Sheet.Cell['A1'].AsDateTime, 0.0001);
end;

procedure TExcelAdvancedTests.Cell_HasFormula_WhenNoFormula_ReturnsFalse;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 42;

  Assert.IsFalse(Sheet.Cell['A1'].HasFormula);
end;

procedure TExcelAdvancedTests.Cell_Formula_WhenNoFormula_ReturnsEmpty;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 42;

  Assert.IsEmpty(Sheet.Cell['A1'].Formula);
end;

{ TExcelFormulaTests }

procedure TExcelFormulaTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'test_formula_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelFormulaTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelFormulaTests.Cell_WithFormula_HasFormulaIsTrue;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);

  var Sheet := FWorkbook.SheetByName('Formula');
  Assert.IsNotNull(Sheet);

  Assert.IsTrue(Sheet.Cell['A3'].HasFormula,'Formula!A3');
  Assert.IsTrue(Sheet.Cell['A10'].HasFormula,'Formula!A10');
end;

procedure TExcelFormulaTests.Cell_WithFormula_ReturnsFormulaString;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);

  var Sheet := FWorkbook.SheetByName('Formula');
  Assert.IsNotNull(Sheet);

  Assert.AreEqual('A1+A2', Sheet.Cell['A3'].Formula);
end;

procedure TExcelFormulaTests.Cell_WithFormula_ReturnsCalculatedValue;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);

  var Sheet := FWorkbook.SheetByName('Formula');
  Assert.IsNotNull(Sheet);

  Assert.AreEqual(Double(1), Sheet.Cell['A3'].AsFloat,'Formula!A3');
  Assert.AreEqual(Double(34), Sheet.Cell['A10'].AsFloat,'Formula!A10');
end;

procedure TExcelFormulaTests.Cell_WithFormula_CrossSheetReference;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);

  var FormulaSheet := FWorkbook.SheetByName('Formula');
  Assert.IsNotNull(FormulaSheet);

  Assert.IsTrue(FormulaSheet.Cell['C1'].HasFormula,'Formula!C1');
  Assert.IsTrue(FormulaSheet.Cell['D1'].HasFormula,'Formula!D1');
  Assert.AreEqual('Simple!A1', FormulaSheet.Cell['C1'].Formula,'Formula!C1');
  Assert.AreEqual('Simple!B1', FormulaSheet.Cell['D1'].Formula,'Formula!D1');

  var SimpleSheet := FWorkbook.SheetByName('Simple');
  Assert.IsNotNull(SimpleSheet);

  Assert.AreEqual(SimpleSheet.Cell['A1'].AsString, FormulaSheet.Cell['C1'].AsString,'Formula!C1');
  Assert.AreEqual(SimpleSheet.Cell['B1'].AsString, FormulaSheet.Cell['D1'].AsString,'Formula!D1');
end;

procedure TExcelFormulaTests.RoundTrip_Formula_PreservesFormula;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);
  FWorkbook.SaveToFile(FTempFile);

  var Workbook2 := TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  var Sheet := Workbook2.SheetByName('Formula');
  Assert.IsNotNull(Sheet);

  Assert.IsTrue(Sheet.Cell['A3'].HasFormula,'Formula!A3');
  Assert.IsTrue(Sheet.Cell['A10'].HasFormula,'Formula!A10');
  Assert.AreEqual('A1+A2', Sheet.Cell['A3'].Formula,'Formula!A3');
  Assert.AreEqual('A8+A9', Sheet.Cell['A10'].Formula,'Formula!A10');
end;

procedure TExcelFormulaTests.SetFormula_CreatesFormulaCell;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 10;
  Sheet.Cell['B1'].AsFloat := 20;
  Sheet.Cell['C1'].SetFormula('A1+B1', 30);

  Assert.IsTrue(Sheet.Cell['C1'].HasFormula);
  Assert.AreEqual('A1+B1', Sheet.Cell['C1'].Formula);
end;

procedure TExcelFormulaTests.SetFormula_WithValue_StoresCalculatedValue;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['C1'].SetFormula('SUM(A1:B1)', 100);

  Assert.AreEqual(Double(100), Sheet.Cell['C1'].AsFloat);
end;

procedure TExcelFormulaTests.SetFormula_RoundTrip_PreservesFormula;
begin
  var Sheet := FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 5;
  Sheet.Cell['B1'].AsFloat := 15;
  Sheet.Cell['C1'].SetFormula('A1*B1', 75);
  FWorkbook.SaveToFile(FTempFile);

  var Workbook2 := TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.IsTrue(Workbook2.Sheets[0].Cell['C1'].HasFormula);
  Assert.AreEqual('A1*B1', Workbook2.Sheets[0].Cell['C1'].Formula);
  Assert.AreEqual(Double(75), Workbook2.Sheets[0].Cell['C1'].AsFloat);
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelReadTests);
  TDUnitX.RegisterTestFixture(TExcelDOMTests);
  TDUnitX.RegisterTestFixture(TExcelAdvancedTests);
  TDUnitX.RegisterTestFixture(TExcelFormulaTests);

end.
