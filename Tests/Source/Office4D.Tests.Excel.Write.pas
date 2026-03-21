unit Office4D.Tests.Excel.Write;

interface

uses
  System.SysUtils,
  System.IOUtils,
  System.Classes,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelWriteTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NewWorkbook_CreatesFile;

    [Test]
    procedure SaveToFile_WithData_ContainsData;

    [Test]
    procedure RoundTrip_LoadModifySave_PreservesContent;

    [Test]
    procedure SaveToFile_ValidatesAsZip;

    [Test]
    procedure SaveToFile_ContainsContentTypes;

    [Test]
    procedure SaveToFile_ContainsWorkbookXml;

    [Test]
    procedure RoundTrip_BooleanValues_PreservesBoolean;

    [Test]
    procedure RoundTrip_DateTimeValues_PreservesDateTime;

    [Test]
    procedure RoundTrip_MixedCellTypes_PreservesAllTypes;

    [Test]
    procedure SaveToFile_FloatWithDecimals_PreservesPrecision;

    [Test]
    procedure SaveToFile_CellsInColumnOrder_ProducesValidXml;

    [Test]
    procedure SaveToFile_DateTimeCells_HaveDateStyle;

    [Test]
    procedure SaveToFile_ContainsStylesXml;

    [Test]
    procedure SaveToStream_WritesToStream;

    [Test]
    procedure LoadFromStream_ReadsFromStream;

    [Test]
    procedure RoundTrip_StreamBased_PreservesContent;

    [Test]
    procedure SaveToFile_WithBoldCell_ContainsBoldStyle;

    [Test]
    procedure RoundTrip_BoldCell_PreservesBold;

    [Test]
    procedure SaveToFile_WithBackgroundColor_ContainsColorStyle;

    [Test]
    procedure RoundTrip_BackgroundColor_PreservesColor;

    [Test]
    procedure SaveToFile_WithCurrencyFormat_ContainsFormat;

    [Test]
    procedure SaveToFile_WithPercentFormat_ContainsFormat;

    [Test]
    procedure SaveToFile_WithColumnWidth_ContainsColsXml;

    [Test]
    procedure SaveToFile_WithMergedCells_ContainsMergeCellsXml;

    [Test]
    procedure SaveToFile_MultipleColumnWidths_ContainsAllCols;

    [Test]
    procedure SaveToFile_MultipleMergedRanges_ContainsAllMerges;

    [Test]
    procedure SaveToFile_WithRowHeight_ContainsRowHeightXml;

    [Test]
    procedure RoundTrip_RowHeight_PreservesRowHeight;

    [Test]
    procedure SaveToFile_WithItalicCell_ContainsItalicStyle;

    [Test]
    procedure RoundTrip_ItalicCell_PreservesItalic;

    [Test]
    procedure SaveToFile_WithUnderlineCell_ContainsUnderlineStyle;

    [Test]
    procedure RoundTrip_UnderlineCell_PreservesUnderline;

    [Test]
    procedure SaveToFile_WithFontName_ContainsFontNameStyle;

    [Test]
    procedure RoundTrip_FontName_PreservesFontName;

    [Test]
    procedure SaveToFile_WithFontSize_ContainsFontSizeStyle;

    [Test]
    procedure RoundTrip_FontSize_PreservesFontSize;

    [Test]
    procedure SaveToFile_WithBorderThin_ContainsBorderStyle;

    [Test]
    procedure RoundTrip_BorderThin_PreservesBorder;

    [Test]
    procedure SaveToFile_WithBorderColor_ContainsBorderColor;

    [Test]
    procedure RoundTrip_BorderColor_PreservesBorderColor;

    [Test]
    procedure SaveToFile_WithHAlign_ContainsAlignmentXml;

    [Test]
    procedure RoundTrip_HAlign_PreservesAlignment;

    [Test]
    procedure SaveToFile_WithVAlign_ContainsAlignmentXml;

    [Test]
    procedure RoundTrip_VAlign_PreservesAlignment;

    [Test]
    procedure SaveToFile_WithWrapText_ContainsWrapTextXml;

    [Test]
    procedure RoundTrip_WrapText_PreservesWrapText;

    [Test]
    procedure RoundTrip_CombinedFormatting_PreservesAll;
  end;

implementation

uses
  System.Zip,
  Office4D.Package;

{ TExcelWriteTests }

procedure TExcelWriteTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'test_output_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelWriteTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelWriteTests.SaveToFile_NewWorkbook_CreatesFile;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  Assert.IsTrue(TFile.Exists(FTempFile), 'File should be created');
end;

procedure TExcelWriteTests.SaveToFile_WithData_ContainsData;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Hello from Office4D';
  Sheet.Cell['B1'].AsFloat := 123;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('Hello from Office4D', Workbook2.Sheets[0].Cell['A1'].AsString);
  Assert.AreEqual(Double(123), Workbook2.Sheets[0].Cell['B1'].AsFloat);
end;

procedure TExcelWriteTests.RoundTrip_LoadModifySave_PreservesContent;
begin
  FWorkbook.LoadFromFile(GetExcelSamplePath);
  FWorkbook.Sheets[0].Cell['C1'].AsString := 'Added by test';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('Hello', Workbook2.Sheets[0].Cell['A1'].AsString);
  Assert.AreEqual('Added by test', Workbook2.Sheets[0].Cell['C1'].AsString);
end;

procedure TExcelWriteTests.SaveToFile_ValidatesAsZip;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Zip := TZipFile.Create;
  try
    Zip.Open(FTempFile, zmRead);
    Assert.IsTrue(Zip.FileCount > 0, 'ZIP should contain files');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_ContainsContentTypes;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('[Content_Types].xml'), 'Should contain [Content_Types].xml');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_ContainsWorkbookXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('xl/workbook.xml'), 'Should contain xl/workbook.xml');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_BooleanValues_PreservesBoolean;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsBoolean := True;
  Sheet.Cell['A2'].AsBoolean := False;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.IsTrue(Workbook2.Sheets[0].Cell['A1'].AsBoolean, 'A1 should be True');
  Assert.IsFalse(Workbook2.Sheets[0].Cell['A2'].AsBoolean, 'A2 should be False');
end;

procedure TExcelWriteTests.RoundTrip_DateTimeValues_PreservesDateTime;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const TestDate = EncodeDate(2024, 6, 15);
  Sheet.Cell['A1'].AsDateTime := TestDate;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(TestDate, Workbook2.Sheets[0].Cell['A1'].AsDateTime, 0.0001, 'DateTime should be preserved');
end;

procedure TExcelWriteTests.RoundTrip_MixedCellTypes_PreservesAllTypes;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Text';
  Sheet.Cell['B1'].AsFloat := 123.45;
  Sheet.Cell['C1'].AsBoolean := True;
  Sheet.Cell['D1'].AsDateTime := EncodeDate(2024, 1, 1);

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('Text', Workbook2.Sheets[0].Cell['A1'].AsString);
  Assert.AreEqual(Double(123.45), Workbook2.Sheets[0].Cell['B1'].AsFloat, 0.001);
  Assert.IsTrue(Workbook2.Sheets[0].Cell['C1'].AsBoolean);
  Assert.AreEqual(EncodeDate(2024, 1, 1), Workbook2.Sheets[0].Cell['D1'].AsDateTime, 0.0001);
end;

procedure TExcelWriteTests.SaveToFile_FloatWithDecimals_PreservesPrecision;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 123.456789;
  Sheet.Cell['A2'].AsFloat := 0.00001;
  Sheet.Cell['A3'].AsFloat := 999999.99;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Double(123.456789), Workbook2.Sheets[0].Cell['A1'].AsFloat, 0.000001);
  Assert.AreEqual(Double(0.00001), Workbook2.Sheets[0].Cell['A2'].AsFloat, 0.000001);
  Assert.AreEqual(Double(999999.99), Workbook2.Sheets[0].Cell['A3'].AsFloat, 0.01);
end;

procedure TExcelWriteTests.SaveToFile_CellsInColumnOrder_ProducesValidXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['E1'].AsString := 'E';
  Sheet.Cell['A1'].AsString := 'A';
  Sheet.Cell['C1'].AsString := 'C';
  Sheet.Cell['B1'].AsString := 'B';
  Sheet.Cell['D1'].AsString := 'D';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');

    const PosA = Pos('r="A1"', SheetXml);
    const PosB = Pos('r="B1"', SheetXml);
    const PosC = Pos('r="C1"', SheetXml);
    const PosD = Pos('r="D1"', SheetXml);
    const PosE = Pos('r="E1"', SheetXml);

    Assert.IsTrue(PosA < PosB, 'A1 should come before B1');
    Assert.IsTrue(PosB < PosC, 'B1 should come before C1');
    Assert.IsTrue(PosC < PosD, 'C1 should come before D1');
    Assert.IsTrue(PosD < PosE, 'D1 should come before E1');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_DateTimeCells_HaveDateStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := EncodeDate(2024, 6, 15);
  Sheet.Cell['B1'].AsFloat := 123.45;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('r="A1" s="1"', SheetXml) > 0, 'DateTime cell should have style s="1"');
    Assert.IsTrue(Pos('r="B1"', SheetXml) > 0, 'Number cell should exist');
    Assert.IsFalse(Pos('r="B1" s="1"', SheetXml) > 0, 'Number cell should not have date style');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_ContainsStylesXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('xl/styles.xml'), 'Should contain xl/styles.xml');

    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('numFmt', StylesXml) > 0, 'Should contain number format');
    Assert.IsTrue(Pos('yyyy-mm-dd', StylesXml) > 0, 'Should contain date format code');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToStream_WritesToStream;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Stream Test';

  var Stream := TMemoryStream.Create;
  try
    FWorkbook.SaveToStream(Stream);
    Assert.IsTrue(Stream.Size > 0, 'Stream should contain data');
  finally
    Stream.Free;
  end;
end;

procedure TExcelWriteTests.LoadFromStream_ReadsFromStream;
begin
  var Stream := TFileStream.Create(GetExcelSamplePath, fmOpenRead or fmShareDenyWrite);
  try
    const Workbook = TExcelWorkbookFactory.Create;
    Workbook.LoadFromStream(Stream);
    Assert.AreEqual(3, Workbook.SheetCount, 'Should have 3 sheets');
  finally
    Stream.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_StreamBased_PreservesContent;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Stream Round Trip';
  Sheet.Cell['B1'].AsFloat := 456.78;

  var Stream := TMemoryStream.Create;
  try
    FWorkbook.SaveToStream(Stream);
    Stream.Position := 0;

    const Workbook2 = TExcelWorkbookFactory.Create;
    Workbook2.LoadFromStream(Stream);

    Assert.AreEqual('Stream Round Trip', Workbook2.Sheets[0].Cell['A1'].AsString);
    Assert.AreEqual(Double(456.78), Workbook2.Sheets[0].Cell['B1'].AsFloat, 0.001);
  finally
    Stream.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_WithBoldCell_ContainsBoldStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Bold text';
  Sheet.Cell['A1'].Bold := True;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('<b/>', StylesXml) > 0, 'Should contain bold element in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_BoldCell_PreservesBold;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Bold text';
  Sheet.Cell['A1'].Bold := True;
  Sheet.Cell['B1'].AsString := 'Normal text';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.IsTrue(Workbook2.Sheets[0].Cell['A1'].Bold, 'A1 should be bold');
  Assert.IsFalse(Workbook2.Sheets[0].Cell['B1'].Bold, 'B1 should not be bold');
end;

procedure TExcelWriteTests.SaveToFile_WithBackgroundColor_ContainsColorStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Colored cell';
  Sheet.Cell['A1'].BackgroundColor := $FFFF00;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('FFFF00', StylesXml) > 0, 'Should contain color in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_BackgroundColor_PreservesColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Yellow';
  Sheet.Cell['A1'].BackgroundColor := $FFFF00;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Cardinal($FFFF00), Workbook2.Sheets[0].Cell['A1'].BackgroundColor);
end;

procedure TExcelWriteTests.SaveToFile_WithCurrencyFormat_ContainsFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 1234.56;
  Sheet.Cell['A1'].NumberFormat := '"$"#,##0.00';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('#,##0.00', StylesXml) > 0, 'Should contain currency format');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_WithPercentFormat_ContainsFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 0.75;
  Sheet.Cell['A1'].NumberFormat := '0%';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('0%', StylesXml) > 0, 'Should contain percent format');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_WithColumnWidth_ContainsColsXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.SetColumnWidth('A', 20);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('<cols>', SheetXml) > 0, 'Should contain cols element');
    Assert.IsTrue(Pos('width="20"', SheetXml) > 0, 'Should contain width value');
    Assert.IsTrue(Pos('customWidth="1"', SheetXml) > 0, 'Should contain customWidth');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_WithMergedCells_ContainsMergeCellsXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Merged';
  Sheet.MergeCells('A1:C1');

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('<mergeCells', SheetXml) > 0, 'Should contain mergeCells element');
    Assert.IsTrue(Pos('ref="A1:C1"', SheetXml) > 0, 'Should contain merge range');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_MultipleColumnWidths_ContainsAllCols;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'A';
  Sheet.Cell['B1'].AsString := 'B';
  Sheet.Cell['C1'].AsString := 'C';
  Sheet.SetColumnWidth('A', 15);
  Sheet.SetColumnWidth('B', 20);
  Sheet.SetColumnWidth('C', 25);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('min="1"', SheetXml) > 0, 'Should contain column A (1)');
    Assert.IsTrue(Pos('min="2"', SheetXml) > 0, 'Should contain column B (2)');
    Assert.IsTrue(Pos('min="3"', SheetXml) > 0, 'Should contain column C (3)');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_MultipleMergedRanges_ContainsAllMerges;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Header 1';
  Sheet.Cell['D1'].AsString := 'Header 2';
  Sheet.MergeCells('A1:C1');
  Sheet.MergeCells('D1:F1');

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('count="2"', SheetXml) > 0, 'Should have count of 2 merged ranges');
    Assert.IsTrue(Pos('ref="A1:C1"', SheetXml) > 0, 'Should contain first merge range');
    Assert.IsTrue(Pos('ref="D1:F1"', SheetXml) > 0, 'Should contain second merge range');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_WithRowHeight_ContainsRowHeightXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Tall row';
  Sheet.SetRowHeight(1, 30);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('ht="30"', SheetXml) > 0, 'Should contain row height');
    Assert.IsTrue(Pos('customHeight="1"', SheetXml) > 0, 'Should contain customHeight');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_RowHeight_PreservesRowHeight;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Test';
  Sheet.SetRowHeight(1, 25);

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Double(25), Workbook2.Sheets[0].GetRowHeight(1), 0.01, 'Row height should be preserved');
end;

procedure TExcelWriteTests.SaveToFile_WithItalicCell_ContainsItalicStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Italic text';
  Sheet.Cell['A1'].Italic := True;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('<i/>', StylesXml) > 0, 'Should contain italic element in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_ItalicCell_PreservesItalic;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Italic text';
  Sheet.Cell['A1'].Italic := True;
  Sheet.Cell['B1'].AsString := 'Normal text';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.IsTrue(Workbook2.Sheets[0].Cell['A1'].Italic, 'A1 should be italic');
  Assert.IsFalse(Workbook2.Sheets[0].Cell['B1'].Italic, 'B1 should not be italic');
end;

procedure TExcelWriteTests.SaveToFile_WithUnderlineCell_ContainsUnderlineStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Underlined text';
  Sheet.Cell['A1'].Underline := True;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('<u/>', StylesXml) > 0, 'Should contain underline element in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_UnderlineCell_PreservesUnderline;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Underlined';
  Sheet.Cell['A1'].Underline := True;
  Sheet.Cell['B1'].AsString := 'Normal';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.IsTrue(Workbook2.Sheets[0].Cell['A1'].Underline, 'A1 should be underlined');
  Assert.IsFalse(Workbook2.Sheets[0].Cell['B1'].Underline, 'B1 should not be underlined');
end;

procedure TExcelWriteTests.SaveToFile_WithFontName_ContainsFontNameStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Arial text';
  Sheet.Cell['A1'].FontName := 'Arial';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('val="Arial"', StylesXml) > 0, 'Should contain font name in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_FontName_PreservesFontName;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Arial text';
  Sheet.Cell['A1'].FontName := 'Arial';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('Arial', Workbook2.Sheets[0].Cell['A1'].FontName);
end;

procedure TExcelWriteTests.SaveToFile_WithFontSize_ContainsFontSizeStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Big text';
  Sheet.Cell['A1'].FontSize := 16;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('val="16"', StylesXml) > 0, 'Should contain font size in styles');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_FontSize_PreservesFontSize;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Big text';
  Sheet.Cell['A1'].FontSize := 16;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Double(16), Workbook2.Sheets[0].Cell['A1'].FontSize, 0.01);
end;

procedure TExcelWriteTests.SaveToFile_WithBorderThin_ContainsBorderStyle;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Bordered';
  Sheet.Cell['A1'].BorderStyle := TExcelBorderStyle.Thin;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('style="thin"', StylesXml) > 0, 'Should contain thin border style');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_BorderThin_PreservesBorder;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Bordered';
  Sheet.Cell['A1'].BorderStyle := TExcelBorderStyle.Thin;
  Sheet.Cell['B1'].AsString := 'No border';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Ord(TExcelBorderStyle.Thin), Ord(Workbook2.Sheets[0].Cell['A1'].BorderStyle));
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Workbook2.Sheets[0].Cell['B1'].BorderStyle));
end;

procedure TExcelWriteTests.SaveToFile_WithBorderColor_ContainsBorderColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red border';
  Sheet.Cell['A1'].BorderStyle := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('FF0000', StylesXml) > 0, 'Should contain border color');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_BorderColor_PreservesBorderColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red border';
  Sheet.Cell['A1'].BorderStyle := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Cardinal($FF0000), Workbook2.Sheets[0].Cell['A1'].BorderColor);
end;

procedure TExcelWriteTests.SaveToFile_WithHAlign_ContainsAlignmentXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Centered';
  Sheet.Cell['A1'].HAlign := TExcelHAlign.Center;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('horizontal="center"', StylesXml) > 0, 'Should contain horizontal alignment');
    Assert.IsTrue(Pos('applyAlignment="1"', StylesXml) > 0, 'Should contain applyAlignment');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_HAlign_PreservesAlignment;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Centered';
  Sheet.Cell['A1'].HAlign := TExcelHAlign.Center;
  Sheet.Cell['B1'].AsString := 'Normal';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Ord(TExcelHAlign.Center), Ord(Workbook2.Sheets[0].Cell['A1'].HAlign));
  Assert.AreEqual(Ord(TExcelHAlign.None), Ord(Workbook2.Sheets[0].Cell['B1'].HAlign));
end;

procedure TExcelWriteTests.SaveToFile_WithVAlign_ContainsAlignmentXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Top aligned';
  Sheet.Cell['A1'].VAlign := TExcelVAlign.Top;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('vertical="top"', StylesXml) > 0, 'Should contain vertical alignment');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_VAlign_PreservesAlignment;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Top aligned';
  Sheet.Cell['A1'].VAlign := TExcelVAlign.Top;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Ord(TExcelVAlign.Top), Ord(Workbook2.Sheets[0].Cell['A1'].VAlign));
end;

procedure TExcelWriteTests.SaveToFile_WithWrapText_ContainsWrapTextXml;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Long text that wraps';
  Sheet.Cell['A1'].WrapText := True;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('wrapText="1"', StylesXml) > 0, 'Should contain wrapText');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_WrapText_PreservesWrapText;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Wrapped text';
  Sheet.Cell['A1'].WrapText := True;
  Sheet.Cell['B1'].AsString := 'Normal';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.IsTrue(Workbook2.Sheets[0].Cell['A1'].WrapText, 'A1 should have wrapText');
  Assert.IsFalse(Workbook2.Sheets[0].Cell['B1'].WrapText, 'B1 should not have wrapText');
end;

procedure TExcelWriteTests.RoundTrip_CombinedFormatting_PreservesAll;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Fully formatted';
  Sheet.Cell['A1'].Bold := True;
  Sheet.Cell['A1'].Italic := True;
  Sheet.Cell['A1'].Underline := True;
  Sheet.Cell['A1'].FontName := 'Arial';
  Sheet.Cell['A1'].FontSize := 14;
  Sheet.Cell['A1'].BackgroundColor := $FFFF00;
  Sheet.Cell['A1'].BorderStyle := TExcelBorderStyle.Medium;
  Sheet.Cell['A1'].BorderColor := $0000FF;
  Sheet.Cell['A1'].HAlign := TExcelHAlign.Center;
  Sheet.Cell['A1'].VAlign := TExcelVAlign.Bottom;
  Sheet.Cell['A1'].WrapText := True;
  Sheet.SetRowHeight(1, 30);

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  const Cell = Workbook2.Sheets[0].Cell['A1'];
  Assert.AreEqual('Fully formatted', Cell.AsString);
  Assert.IsTrue(Cell.Bold, 'Should be bold');
  Assert.IsTrue(Cell.Italic, 'Should be italic');
  Assert.IsTrue(Cell.Underline, 'Should be underlined');
  Assert.AreEqual('Arial', Cell.FontName);
  Assert.AreEqual(Double(14), Cell.FontSize, 0.01);
  Assert.AreEqual(Cardinal($FFFF00), Cell.BackgroundColor);
  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle));
  Assert.AreEqual(Cardinal($0000FF), Cell.BorderColor);
  Assert.AreEqual(Ord(TExcelHAlign.Center), Ord(Cell.HAlign));
  Assert.AreEqual(Ord(TExcelVAlign.Bottom), Ord(Cell.VAlign));
  Assert.IsTrue(Cell.WrapText, 'Should have wrapText');
  Assert.AreEqual(Double(30), Workbook2.Sheets[0].GetRowHeight(1), 0.01);
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelWriteTests);

end.
