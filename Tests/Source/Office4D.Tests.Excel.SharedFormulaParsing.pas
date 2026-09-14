unit Office4D.Tests.Excel.SharedFormulaParsing;

interface

uses
  System.SysUtils,
  System.IOUtils,
  System.Diagnostics,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelSharedFormulaParsingTests = class(TOffice4DTests)
  private
    FTempFile: string;

    function BuildMinimalXlsx(const SheetDataXml: string): TBytes;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_ManySelfClosingCellsNoFormulas_CompletesQuickly;

    [Test]
    procedure LoadFromFile_RealWorldSparseSheetShape_LoadsWithoutError;

    [Test]
    procedure LoadFromFile_SharedFormulaAlongsideSelfClosingCells_StillResolvesCorrectly;
  end;

implementation

uses
  System.Classes,
  System.Zip;

{ TExcelSharedFormulaParsingTests }

procedure TExcelSharedFormulaParsingTests.Setup;
begin
  FTempFile := TPath.Combine(TPath.GetTempPath, 'sharedformula_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelSharedFormulaParsingTests.TearDown;
begin
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

function TExcelSharedFormulaParsingTests.BuildMinimalXlsx(const SheetDataXml: string): TBytes;
const
  ContentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>';
  RelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>';
  WorkbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';
var
  SheetXml: string;
  Zip: TZipFile;
  Stream: TBytesStream;
begin
  SheetXml :=
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>' + SheetDataXml + '</sheetData></worksheet>';

  Stream := TBytesStream.Create;
  try
    Zip := TZipFile.Create;
    try
      Zip.Open(Stream, zmWrite);
      Zip.Add(TEncoding.UTF8.GetBytes(ContentTypesXml), '[Content_Types].xml');
      Zip.Add(TEncoding.UTF8.GetBytes(RelsXml), '_rels/.rels');
      Zip.Add(TEncoding.UTF8.GetBytes(WorkbookXml), 'xl/workbook.xml');
      Zip.Add(TEncoding.UTF8.GetBytes(SheetXml), 'xl/worksheets/sheet1.xml');
      Zip.Close;
    finally
      Zip.Free;
    end;
    Result := Stream.Bytes;
    SetLength(Result, Stream.Size);
  finally
    Stream.Free;
  end;
end;

procedure TExcelSharedFormulaParsingTests.LoadFromFile_ManySelfClosingCellsNoFormulas_CompletesQuickly;
var
  SB: TStringBuilder;
  i: Integer;
  Stopwatch: TStopwatch;
  Workbook: IExcelWorkbook;
begin
  // Deliberately adversarial for the pre-fix regex: a long unbroken run of self-closing,
  // styled-but-valueless cells with no <f> element anywhere in the sheet. Verified
  // empirically (outside Delphi) that the unguarded pattern shows clear polynomial
  // blowup on inputs shaped like this -- this is the regression guard for that fix.
  SB := TStringBuilder.Create;
  try
    SB.Append('<row r="1">');
    for i := 1 to 3000 do
      SB.Append('<c r="A' + IntToStr(i) + '" s="1"/>');
    SB.Append('</row>');

    TFile.WriteAllBytes(FTempFile, BuildMinimalXlsx(SB.ToString));
  finally
    SB.Free;
  end;

  Stopwatch := TStopwatch.StartNew;
  Workbook := TExcelWorkbookFactory.Create;
  Workbook.LoadFromFile(FTempFile);
  Stopwatch.Stop;

  // Generous bound -- this is a regression guard against catastrophic (super-linear)
  // blowup, not a tight performance benchmark. 3000 cells should load near-instantly;
  // if this fix regresses, expect this to time out or fail long before 5 seconds.
  Assert.IsTrue(Stopwatch.ElapsedMilliseconds < 5000,
    Format('LoadFromFile took %dms for 3000 self-closing cells -- possible regression ' +
      'of the SharedMasterMatches self-closing-cell guard', [Stopwatch.ElapsedMilliseconds]));
end;

procedure TExcelSharedFormulaParsingTests.LoadFromFile_RealWorldSparseSheetShape_LoadsWithoutError;
var
  Workbook: IExcelWorkbook;
begin
  // Mirrors the actual shape that surfaced this bug: short runs of self-closing styled
  // cells interleaved with real valued cells, no formulas anywhere in the sheet.
  const SheetData =
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2" s="1"/></row>' +
    '<row r="3"><c r="C3" s="3"/><c r="D3" s="2"/><c r="E3" s="1"/></row>' +
    '<row r="4"><c r="C4" s="3"/><c r="D4" s="2"/><c r="E4" s="1"/></row>' +
    '<row r="5"><c r="A5"><v>5</v></c></row>';

  TFile.WriteAllBytes(FTempFile, BuildMinimalXlsx(SheetData));

  Workbook := TExcelWorkbookFactory.Create;
  Workbook.LoadFromFile(FTempFile); // must not hang or raise

  Assert.AreEqual(Double(1), Workbook.Sheets[0].Cell['A2'].AsFloat);
  Assert.AreEqual(Double(5), Workbook.Sheets[0].Cell['A5'].AsFloat);
end;

procedure TExcelSharedFormulaParsingTests.LoadFromFile_SharedFormulaAlongsideSelfClosingCells_StillResolvesCorrectly;
var
  Workbook: IExcelWorkbook;
begin
  // Confirms the fix doesn't break legitimate shared-formula master/dependent resolution
  // when self-closing cells are also present in the same sheet -- a realistic mixed case,
  // not just the pathological all-self-closing stress case above.
  const SheetData =
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/></row>' +
    '<row r="2">' +
    '<c r="A2"><f t="shared" ref="A2:A3" si="0">B1+1</f><v>2</v></c>' +
    '<c r="B2" s="1"/>' +
    '</row>' +
    '<row r="3">' +
    '<c r="A3"><f t="shared" si="0"/><v>2</v></c>' +
    '<c r="B3" s="1"/>' +
    '</row>';

  TFile.WriteAllBytes(FTempFile, BuildMinimalXlsx(SheetData));

  Workbook := TExcelWorkbookFactory.Create;
  Workbook.LoadFromFile(FTempFile);

  Assert.AreEqual('B1+1', Workbook.Sheets[0].Cell['A2'].Formula, 'Master formula should parse correctly');
  Assert.IsTrue(Workbook.Sheets[0].Cell['A3'].HasFormula, 'Dependent cell should have a resolved formula');
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelSharedFormulaParsingTests);

end.
