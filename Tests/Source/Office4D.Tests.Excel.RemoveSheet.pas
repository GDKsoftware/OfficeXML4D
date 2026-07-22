unit Office4D.Tests.Excel.RemoveSheet;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelRemoveSheetTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure RemoveSheet_ByIndex_DecrementsCountAndKeepsOthers;

    [Test]
    procedure RemoveSheetByName_RemovesTheNamedSheet;

    [Test]
    procedure RemoveSheet_IndexOutOfRange_Raises;

    [Test]
    procedure RemoveSheetByName_UnknownName_Raises;

    [Test]
    procedure RemoveSheet_LastRemainingSheet_Raises;

    [Test]
    procedure RoundTrip_AfterRemove_SheetsRenumberedSequentially;
  end;

implementation

uses
  Office4D.Errors;

{ TExcelRemoveSheetTests }

procedure TExcelRemoveSheetTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'removesheet_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelRemoveSheetTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelRemoveSheetTests.RemoveSheet_ByIndex_DecrementsCountAndKeepsOthers;
begin
  FWorkbook.AddSheet('Alpha');
  FWorkbook.AddSheet('Beta');
  FWorkbook.AddSheet('Gamma');

  FWorkbook.RemoveSheet(1); // Beta

  Assert.AreEqual(2, FWorkbook.SheetCount, 'Removing a sheet must decrement the count');
  Assert.AreEqual('Alpha', FWorkbook.Sheets[0].Name);
  Assert.AreEqual('Gamma', FWorkbook.Sheets[1].Name, 'The sheet after the removed one must slide into its place');
end;

procedure TExcelRemoveSheetTests.RemoveSheetByName_RemovesTheNamedSheet;
begin
  FWorkbook.AddSheet('Alpha');
  FWorkbook.AddSheet('Beta');
  FWorkbook.AddSheet('Gamma');

  FWorkbook.RemoveSheetByName('Beta');

  Assert.AreEqual(2, FWorkbook.SheetCount);
  Assert.IsNull(FWorkbook.SheetByName('Beta'), 'The named sheet must be gone');
  Assert.IsNotNull(FWorkbook.SheetByName('Alpha'));
  Assert.IsNotNull(FWorkbook.SheetByName('Gamma'));
end;

procedure TExcelRemoveSheetTests.RemoveSheet_IndexOutOfRange_Raises;
begin
  FWorkbook.AddSheet('Alpha');
  FWorkbook.AddSheet('Beta');

  Assert.WillRaise(
    procedure
    begin
      FWorkbook.RemoveSheet(5);
    end,
    EExcelWorkbookException,
    'An out-of-range index must raise rather than silently doing nothing');
end;

procedure TExcelRemoveSheetTests.RemoveSheetByName_UnknownName_Raises;
begin
  FWorkbook.AddSheet('Alpha');
  FWorkbook.AddSheet('Beta');

  Assert.WillRaise(
    procedure
    begin
      FWorkbook.RemoveSheetByName('DoesNotExist');
    end,
    EExcelWorkbookException,
    'Removing an unknown sheet name must raise');
end;

procedure TExcelRemoveSheetTests.RemoveSheet_LastRemainingSheet_Raises;
begin
  FWorkbook.AddSheet('Only');

  // Excel refuses to open a workbook with no sheets, so the last one cannot be removed.
  Assert.WillRaise(
    procedure
    begin
      FWorkbook.RemoveSheet(0);
    end,
    EExcelWorkbookException,
    'Removing the last remaining sheet must raise');
end;

procedure TExcelRemoveSheetTests.RoundTrip_AfterRemove_SheetsRenumberedSequentially;
begin
  FWorkbook.AddSheet('Alpha').Cell['A1'].AsString := 'a';
  FWorkbook.AddSheet('Beta').Cell['A1'].AsString := 'b';
  FWorkbook.AddSheet('Gamma').Cell['A1'].AsString := 'g';

  FWorkbook.RemoveSheet(1); // Beta
  FWorkbook.SaveToFile(FTempFile);

  const Reloaded = TExcelWorkbookFactory.Create;
  Reloaded.LoadFromFile(FTempFile);

  Assert.AreEqual(2, Reloaded.SheetCount, 'The saved workbook must contain only the surviving sheets');
  Assert.AreEqual('Alpha', Reloaded.Sheets[0].Name);
  Assert.AreEqual('Gamma', Reloaded.Sheets[1].Name);
  Assert.AreEqual('a', Reloaded.Sheets[0].Cell['A1'].AsString, 'Surviving sheet content must be intact');
  Assert.AreEqual('g', Reloaded.Sheets[1].Cell['A1'].AsString);
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelRemoveSheetTests);

end.
