unit Office4D.Tests.Excel.XmlShape;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Tests.PackageBuilder,
  Office4D.Excel;

type
  /// <summary>
  /// A workbook reads the same whatever tool wrote it: attribute order, quote
  /// style, whitespace around '=', the namespace prefix and the part names a
  /// writer picks carry no meaning in XML, so they must not change the result.
  /// </summary>
  [TestFixture]
  TExcelXmlShapeTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FSourceFile: string;

    procedure Load(const Parts: TArray<TTestPart>);
    procedure LoadSheets(const SheetsXml: string);
    procedure LoadSheetData(const SheetDataXml: string);
    procedure LoadStyles(const StyleElementsXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    [TestCase('OpenpyxlOrder', '<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Sales" sheetId="1" state="visible" r:id="rId1"/>')]
    [TestCase('NameLast', '<sheet state="visible" sheetId="1" r:id="rId1" name="Sales"/>')]
    [TestCase('SingleQuotes', '<sheet name=''Sales'' sheetId=''1'' r:id=''rId1''/>')]
    [TestCase('SpacesAroundEquals', '<sheet name = "Sales" sheetId = "1" r:id = "rId1"/>')]
    [TestCase('OpenAndCloseTag', '<sheet sheetId="1" name="Sales" r:id="rId1"></sheet>')]
    procedure LoadFromFile_SheetElementInAnyShape_ReadsTheSheet(const SheetXml: string);

    [Test]
    [TestCase('Hidden', 'hidden,1')]
    [TestCase('VeryHidden', 'veryHidden,2')]
    procedure LoadFromFile_StateBeforeName_ReadsVisibility(const State: string; const Expected: Integer);

    [Test]
    procedure LoadFromFile_PrefixedSpreadsheetNamespace_ReadsSheetsStringsAndStyles;

    [Test]
    procedure LoadFromFile_SheetOrderDiffersFromPartNames_ReadsThePartOfEachSheet;

    [Test]
    [TestCase('TypeFirst', '<c t="s" r="A1" s="1"><v>0</v></c>')]
    [TestCase('StyleFirst', '<c s="1" t="s" r="A1"><v>0</v></c>')]
    [TestCase('SingleQuotes', '<c r=''A1'' s=''1'' t=''s''><v>0</v></c>')]
    [TestCase('InlineStringTypeFirst', '<c t="inlineStr" s="1" r="A1"><is><t>Apples</t></is></c>')]
    procedure LoadFromFile_CellAttributesInAnyOrder_ReadsValueAndStyle(const CellXml: string);

    [Test]
    [TestCase('SelfClosing', '<c r="A1" s="1"/>')]
    [TestCase('OpenAndCloseTag', '<c r="A1" s="1"></c>')]
    [TestCase('StyleFirst', '<c s="1" r="A1"/>')]
    procedure LoadFromFile_StyledEmptyCellInAnyShape_KeepsItsStyle(const CellXml: string);

    [Test]
    procedure LoadFromFile_CellsWithoutReference_TakeTheirPlaceInTheRow;

    [Test]
    [TestCase('Plain', '<b/>|True', '|')]
    [TestCase('ValueOne', '<b val="1"/>|True', '|')]
    [TestCase('ValueTrue', '<b val="true"/>|True', '|')]
    [TestCase('ValueZero', '<b val="0"/>|False', '|')]
    procedure LoadFromFile_BoldElementWithValue_ReadsBold(const BoldXml: string; const Expected: Boolean);

    [Test]
    procedure LoadFromFile_ColorWithZeroAlpha_ReadsColor;

    [Test]
    procedure LoadFromFile_NumFmtWithFormatCodeFirst_ReadsFormat;

    [Test]
    procedure LoadFromFile_BorderWithAttributes_KeepsBorderIndexes;

    [Test]
    procedure LoadFromFile_OpenpyxlWorkbook_ReadsSheetsValuesAndStyles;
  end;

implementation

const
  SpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  RelationshipsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  PackageRelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  WorksheetType = RelationshipsNs + '/worksheet';

  RootRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Type="' + RelationshipsNs + '/officeDocument" Target="xl/workbook.xml" Id="rId1"/>' +
    '</Relationships>';

  WorkbookXml =
    '<workbook xmlns="' + SpreadsheetNs + '" xmlns:r="' + RelationshipsNs + '"><sheets>%s</sheets></workbook>';

  WorkbookRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Type="' + WorksheetType + '" Target="/xl/worksheets/sheet1.xml" Id="rId1"/>' +
    '</Relationships>';

  SheetXml = '<worksheet xmlns="' + SpreadsheetNs + '"><sheetData>%s</sheetData></worksheet>';

  SharedStringsXml = '<sst xmlns="' + SpreadsheetNs + '"><si><t>Apples</t></si></sst>';

  StylesXml =
    '<styleSheet xmlns="' + SpreadsheetNs + '">%s</styleSheet>';

  BoldStyleXml =
    '<fonts><font><sz val="11"/></font><font><b/></font></fonts>' +
    '<cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs>';

  ApplesRowXml = '<row r="1"><c r="A1" t="inlineStr"><is><t>Apples</t></is></c></row>';
  DefaultSheetXml = '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>';
  StyledCellRowXml = '<row r="1"><c r="A1" s="1"><v>1</v></c></row>';

  PartRootRels = '_rels/.rels';
  PartWorkbook = 'xl/workbook.xml';
  PartWorkbookRels = 'xl/_rels/workbook.xml.rels';
  PartSheet1 = 'xl/worksheets/sheet1.xml';
  PartSheet2 = 'xl/worksheets/sheet2.xml';
  PartSharedStrings = 'xl/sharedStrings.xml';
  PartStyles = 'xl/styles.xml';

{ TExcelXmlShapeTests }

procedure TExcelXmlShapeTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FSourceFile := TPath.Combine(TPath.GetTempPath, 'xmlshape_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelXmlShapeTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FSourceFile) then
    TFile.Delete(FSourceFile);
end;

procedure TExcelXmlShapeTests.Load(const Parts: TArray<TTestPart>);
begin
  TTestPackage.WriteParts(FSourceFile, Parts);
  FWorkbook.LoadFromFile(FSourceFile);
end;

procedure TExcelXmlShapeTests.LoadSheets(const SheetsXml: string);
begin
  Load([TTestPart.Create(PartRootRels, RootRelsXml),
        TTestPart.Create(PartWorkbook, Format(WorkbookXml, [SheetsXml])),
        TTestPart.Create(PartWorkbookRels, WorkbookRelsXml),
        TTestPart.Create(PartSheet1, Format(SheetXml, [ApplesRowXml]))]);
end;

procedure TExcelXmlShapeTests.LoadSheetData(const SheetDataXml: string);
begin
  Load([TTestPart.Create(PartRootRels, RootRelsXml),
        TTestPart.Create(PartWorkbook, Format(WorkbookXml, [DefaultSheetXml])),
        TTestPart.Create(PartWorkbookRels, WorkbookRelsXml),
        TTestPart.Create(PartSharedStrings, SharedStringsXml),
        TTestPart.Create(PartStyles, Format(StylesXml, [BoldStyleXml])),
        TTestPart.Create(PartSheet1, Format(SheetXml, [SheetDataXml]))]);
end;

procedure TExcelXmlShapeTests.LoadStyles(const StyleElementsXml: string);
begin
  Load([TTestPart.Create(PartRootRels, RootRelsXml),
        TTestPart.Create(PartWorkbook, Format(WorkbookXml, [DefaultSheetXml])),
        TTestPart.Create(PartWorkbookRels, WorkbookRelsXml),
        TTestPart.Create(PartStyles, Format(StylesXml, [StyleElementsXml])),
        TTestPart.Create(PartSheet1, Format(SheetXml, [StyledCellRowXml]))]);
end;

procedure TExcelXmlShapeTests.LoadFromFile_SheetElementInAnyShape_ReadsTheSheet(const SheetXml: string);
begin
  LoadSheets(SheetXml);

  Assert.AreEqual(1, FWorkbook.SheetCount, 'The sheet must be found');
  Assert.AreEqual('Sales', FWorkbook.Sheets[0].Name);
  Assert.AreEqual('Apples', FWorkbook.Sheets[0].Cell['A1'].AsString, 'The sheet part must be read');
end;

procedure TExcelXmlShapeTests.LoadFromFile_StateBeforeName_ReadsVisibility(const State: string;
                                                                            const Expected: Integer);
begin
  LoadSheets(Format('<sheet state="%s" name="Sales" sheetId="1" r:id="rId1"/>', [State]));

  const Actual = FWorkbook.Sheets[0].Visibility;

  Assert.AreEqual<TExcelSheetVisibility>(TExcelSheetVisibility(Expected), Actual);
end;

procedure TExcelXmlShapeTests.LoadFromFile_PrefixedSpreadsheetNamespace_ReadsSheetsStringsAndStyles;
begin
  const PrefixedWorkbookXml =
    '<x:workbook xmlns:x="' + SpreadsheetNs + '" xmlns:r="' + RelationshipsNs + '">' +
    '<x:sheets><x:sheet name="Sales" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>';
  const PrefixedSheetXml =
    '<x:worksheet xmlns:x="' + SpreadsheetNs + '"><x:sheetData><x:row r="1">' +
    '<x:c r="A1" s="1" t="s"><x:v>0</x:v></x:c><x:c r="B1"><x:v>12.5</x:v></x:c>' +
    '</x:row></x:sheetData></x:worksheet>';
  const PrefixedSharedStringsXml =
    '<x:sst xmlns:x="' + SpreadsheetNs + '"><x:si><x:t>Apples</x:t></x:si></x:sst>';
  const PrefixedStylesXml =
    '<x:styleSheet xmlns:x="' + SpreadsheetNs + '">' +
    '<x:fonts><x:font><x:sz val="11"/></x:font><x:font><x:b/></x:font></x:fonts>' +
    '<x:cellXfs><x:xf numFmtId="0" fontId="0"/><x:xf numFmtId="0" fontId="1"/></x:cellXfs>' +
    '</x:styleSheet>';

  Load([TTestPart.Create(PartRootRels, RootRelsXml),
        TTestPart.Create(PartWorkbook, PrefixedWorkbookXml),
        TTestPart.Create(PartWorkbookRels, WorkbookRelsXml),
        TTestPart.Create(PartSharedStrings, PrefixedSharedStringsXml),
        TTestPart.Create(PartStyles, PrefixedStylesXml),
        TTestPart.Create(PartSheet1, PrefixedSheetXml)]);

  const Sheet = FWorkbook.Sheets[0];

  Assert.AreEqual('Sales', Sheet.Name);
  Assert.AreEqual('Apples', Sheet.Cell['A1'].AsString);
  Assert.IsTrue(Sheet.Cell['A1'].Bold, 'The cell style must be read');
  Assert.AreEqual(Double(12.5), Sheet.Cell['B1'].AsFloat);
end;

procedure TExcelXmlShapeTests.LoadFromFile_SheetOrderDiffersFromPartNames_ReadsThePartOfEachSheet;
begin
  const SheetsXml =
    '<sheet name="Summary" sheetId="2" r:id="rId2"/>' +
    '<sheet name="Data" sheetId="1" r:id="rId1"/>';
  const TwoSheetRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + WorksheetType + '" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="' + WorksheetType + '" Target="worksheets/sheet2.xml"/>' +
    '</Relationships>';

  Load([TTestPart.Create(PartRootRels, RootRelsXml),
        TTestPart.Create(PartWorkbook, Format(WorkbookXml, [SheetsXml])),
        TTestPart.Create(PartWorkbookRels, TwoSheetRelsXml),
        TTestPart.Create(PartSheet1, Format(SheetXml, ['<row r="1"><c r="A1"><v>1</v></c></row>'])),
        TTestPart.Create(PartSheet2, Format(SheetXml, ['<row r="1"><c r="A1"><v>2</v></c></row>']))]);

  Assert.AreEqual('Summary', FWorkbook.Sheets[0].Name);
  Assert.AreEqual(Double(2), FWorkbook.Sheets[0].Cell['A1'].AsFloat, 'Summary is stored in sheet2.xml');
  Assert.AreEqual(Double(1), FWorkbook.Sheets[1].Cell['A1'].AsFloat, 'Data is stored in sheet1.xml');
end;

procedure TExcelXmlShapeTests.LoadFromFile_CellAttributesInAnyOrder_ReadsValueAndStyle(const CellXml: string);
begin
  LoadSheetData(Format('<row r="1">%s</row>', [CellXml]));

  const Cell = FWorkbook.Sheets[0].Cell['A1'];

  Assert.AreEqual('Apples', Cell.AsString);
  Assert.IsTrue(Cell.Bold, 'The style index must be read whatever its position');
end;

procedure TExcelXmlShapeTests.LoadFromFile_StyledEmptyCellInAnyShape_KeepsItsStyle(const CellXml: string);
begin
  LoadSheetData(Format('<row r="1">%s</row>', [CellXml]));

  const Sheet = FWorkbook.Sheets[0];

  Assert.IsTrue(Sheet.Cells.ContainsKey('A1'), 'A styled empty cell must be kept');
  Assert.IsTrue(Sheet.Cell['A1'].Bold);
end;

procedure TExcelXmlShapeTests.LoadFromFile_CellsWithoutReference_TakeTheirPlaceInTheRow;
begin
  LoadSheetData('<row r="2"><c t="inlineStr"><is><t>Pears</t></is></c><c><v>6</v></c></row>' +
                '<row><c><v>7</v></c></row>');

  const Sheet = FWorkbook.Sheets[0];

  Assert.AreEqual('Pears', Sheet.Cell['A2'].AsString);
  Assert.AreEqual(Double(6), Sheet.Cell['B2'].AsFloat);
  Assert.AreEqual(Double(7), Sheet.Cell['A3'].AsFloat, 'A row without r follows the previous row');
end;

procedure TExcelXmlShapeTests.LoadFromFile_BoldElementWithValue_ReadsBold(const BoldXml: string;
                                                                         const Expected: Boolean);
begin
  LoadStyles(Format('<fonts><font/><font>%s</font></fonts>' +
                    '<cellXfs><xf fontId="0"/><xf fontId="1"/></cellXfs>', [BoldXml]));

  const Actual = FWorkbook.Sheets[0].Cell['A1'].Bold;

  Assert.AreEqual(Expected, Actual);
end;

procedure TExcelXmlShapeTests.LoadFromFile_ColorWithZeroAlpha_ReadsColor;
begin
  LoadStyles('<fonts><font/><font><color rgb="00FF0000"/></font></fonts>' +
             '<fills><fill/><fill/><fill><patternFill patternType="solid"><fgColor rgb="0000FF00"/></patternFill></fill></fills>' +
             '<cellXfs><xf fontId="0"/><xf fontId="1" fillId="2"/></cellXfs>');

  const Cell = FWorkbook.Sheets[0].Cell['A1'];

  Assert.AreEqual(Cardinal($FF0000), Cell.FontColor);
  Assert.AreEqual(Cardinal($00FF00), Cell.BackgroundColor);
end;

procedure TExcelXmlShapeTests.LoadFromFile_NumFmtWithFormatCodeFirst_ReadsFormat;
begin
  LoadStyles('<numFmts><numFmt formatCode="0.000" numFmtId="164"/></numFmts>' +
             '<cellXfs><xf numFmtId="0"/><xf fontId="0" numFmtId="164"/></cellXfs>');

  const Actual = FWorkbook.Sheets[0].Cell['A1'].NumberFormat;

  Assert.AreEqual('0.000', Actual);
end;

procedure TExcelXmlShapeTests.LoadFromFile_BorderWithAttributes_KeepsBorderIndexes;
begin
  LoadStyles('<borders><border/>' +
             '<border diagonalUp="0"><left/><right/><top style="thin"><color rgb="FF0000FF"/></top><bottom/></border>' +
             '</borders>' +
             '<cellXfs><xf borderId="0"/><xf borderId="1"/></cellXfs>');

  const Cell = FWorkbook.Sheets[0].Cell['A1'];

  Assert.AreEqual<TExcelBorderStyle>(TExcelBorderStyle.Thin, Cell.BorderStyle[[TExcelBorderSide.Top]]);
  Assert.AreEqual(Cardinal($0000FF), Cell.BorderColor[[TExcelBorderSide.Top]]);
end;

procedure TExcelXmlShapeTests.LoadFromFile_OpenpyxlWorkbook_ReadsSheetsValuesAndStyles;
begin
  FWorkbook.LoadFromFile(TPath.Combine(GetSamplesPath, 'Excel\openpyxl.xlsx'));

  const Sales = FWorkbook.Sheets[0];
  const Archive = FWorkbook.Sheets[1];

  Assert.AreEqual(2, FWorkbook.SheetCount);
  Assert.AreEqual('Sales', Sales.Name);
  Assert.AreEqual('Archive', Archive.Name);
  Assert.AreEqual<TExcelSheetVisibility>(TExcelSheetVisibility.Hidden, Archive.Visibility);
  Assert.AreEqual('Old', Archive.Cell['A1'].AsString);

  Assert.AreEqual('Product', Sales.Cell['A1'].AsString);
  Assert.IsTrue(Sales.Cell['A1'].Bold, 'openpyxl writes bold as <b val="1"/>');
  Assert.AreEqual(Cardinal($FF0000), Sales.Cell['A1'].FontColor, 'openpyxl writes colours with alpha 00');
  Assert.AreEqual(Double(12.5), Sales.Cell['B2'].AsFloat);
  Assert.AreEqual('0.000', Sales.Cell['B2'].NumberFormat);
  Assert.AreEqual<TExcelBorderStyle>(TExcelBorderStyle.Thin, Sales.Cell['A3'].BorderStyle[[TExcelBorderSide.Top]]);
  Assert.AreEqual(Cardinal($00FF00), Sales.Cell['B3'].BackgroundColor);
  Assert.AreEqual('SUM(B2:B3)', Sales.Cell['B4'].Formula);
  Assert.AreEqual(1, Sales.FrozenRows);
end;

end.
