unit Office4D.Tests.Excel.SharedFormulaParsing;

interface

uses
  System.SysUtils,
  System.Zip,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelSharedFormulaParsingTests = class(TOffice4DTests)
  private
    function LoadWorkbook(const SheetDataXml: string): IExcelWorkbook;
    function BuildSelfClosingCellsRow(const CellCount: Integer): string;
    function BuildMinimalXlsx(const SheetDataXml: string): TBytes;
    procedure AddXmlPart(const Zip: TZipFile; const PartName: string; const Xml: string);

  public
    [Test]
    procedure LoadFromStream_ManySelfClosingCellsNoFormulas_CompletesQuickly;

    [Test]
    procedure LoadFromStream_SparseSheetWithSelfClosingCells_LoadsValues;

    [Test]
    procedure LoadFromStream_SharedFormulaAlongsideSelfClosingCells_ResolvesDependentFormula;
  end;

implementation

uses
  System.Classes,
  System.Diagnostics;

const
  SelfClosingCellCount = 3000;
  MaxLoadTimeMs = 5000;

{ TExcelSharedFormulaParsingTests }

procedure TExcelSharedFormulaParsingTests.LoadFromStream_ManySelfClosingCellsNoFormulas_CompletesQuickly;
begin
  const SheetData = BuildSelfClosingCellsRow(SelfClosingCellCount);

  var Stopwatch := TStopwatch.StartNew;
  const Workbook = LoadWorkbook(SheetData);
  Stopwatch.Stop;

  Assert.IsNotNull(Workbook.Sheets[0]);
  const LoadedQuickly = (Stopwatch.ElapsedMilliseconds < MaxLoadTimeMs);
  const FailureMessage = Format('Loading %d self-closing cells took %d ms; the self-closing-cell guard in SharedMasterMatches may have regressed',
                                [SelfClosingCellCount, Stopwatch.ElapsedMilliseconds]);
  Assert.IsTrue(LoadedQuickly, FailureMessage);
end;

procedure TExcelSharedFormulaParsingTests.LoadFromStream_SparseSheetWithSelfClosingCells_LoadsValues;
begin
  const SheetData =
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2" s="1"/></row>' +
    '<row r="3"><c r="C3" s="3"/><c r="D3" s="2"/><c r="E3" s="1"/></row>' +
    '<row r="4"><c r="C4" s="3"/><c r="D4" s="2"/><c r="E4" s="1"/></row>' +
    '<row r="5"><c r="A5"><v>5</v></c></row>';

  const Workbook = LoadWorkbook(SheetData);
  const Sheet = Workbook.Sheets[0];

  Assert.AreEqual(Double(1), Sheet.Cell['A2'].AsFloat);
  Assert.AreEqual(Double(5), Sheet.Cell['A5'].AsFloat);
end;

procedure TExcelSharedFormulaParsingTests.LoadFromStream_SharedFormulaAlongsideSelfClosingCells_ResolvesDependentFormula;
begin
  const SheetData =
    '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/></row>' +
    '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B1+1</f><v>2</v></c><c r="B2" s="1"/></row>' +
    '<row r="3"><c r="A3"><f t="shared" si="0"/><v>2</v></c><c r="B3" s="1"/></row>';

  const Workbook = LoadWorkbook(SheetData);
  const Sheet = Workbook.Sheets[0];

  Assert.AreEqual('B1+1', Sheet.Cell['A2'].Formula, 'Master cell keeps its own formula');
  Assert.AreEqual('B2+1', Sheet.Cell['A3'].Formula, 'Dependent cell gets the shared formula shifted one row down');
end;

function TExcelSharedFormulaParsingTests.LoadWorkbook(const SheetDataXml: string): IExcelWorkbook;
begin
  const Xlsx = BuildMinimalXlsx(SheetDataXml);
  const Stream = TBytesStream.Create(Xlsx);
  try
    Result := TExcelWorkbookFactory.Create;
    Result.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

function TExcelSharedFormulaParsingTests.BuildSelfClosingCellsRow(const CellCount: Integer): string;
begin
  const RowXml = TStringBuilder.Create;
  try
    RowXml.Append('<row r="1">');
    for var CellIndex := 1 to CellCount do
    begin
      const CellXml = Format('<c r="A%d" s="1"/>', [CellIndex]);
      RowXml.Append(CellXml);
    end;
    RowXml.Append('</row>');

    Result := RowXml.ToString;
  finally
    RowXml.Free;
  end;
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
  SheetXmlTemplate =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData>%s</sheetData></worksheet>';
begin
  const SheetXml = Format(SheetXmlTemplate, [SheetDataXml]);

  const Stream = TBytesStream.Create;
  try
    const Zip = TZipFile.Create;
    try
      Zip.Open(Stream, TZipMode.zmWrite);
      AddXmlPart(Zip, '[Content_Types].xml', ContentTypesXml);
      AddXmlPart(Zip, '_rels/.rels', RelsXml);
      AddXmlPart(Zip, 'xl/workbook.xml', WorkbookXml);
      AddXmlPart(Zip, 'xl/worksheets/sheet1.xml', SheetXml);
      Zip.Close;
    finally
      Zip.Free;
    end;

    Result := Copy(Stream.Bytes, 0, Stream.Size);
  finally
    Stream.Free;
  end;
end;

procedure TExcelSharedFormulaParsingTests.AddXmlPart(const Zip: TZipFile; const PartName: string; const Xml: string);
begin
  const PartContent = TEncoding.UTF8.GetBytes(Xml);
  Zip.Add(PartContent, PartName);
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelSharedFormulaParsingTests);

end.
