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
    procedure RoundTrip_SpecialCharacters_ArePreserved;

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
    procedure SaveToFile_DateTimeCell_WritesCorrectExcelValue;

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

    [Test]
    procedure SaveToFile_EmptyStringCell_WritesNoInvalidSharedStringIndex;

    [Test]
    procedure RoundTrip_EmptyStringCell_PreservesAdjacentValues;

    [Test]
    procedure SaveToFile_DateCell_UsesLocaleAwareShortDateFormat;

    [Test]
    procedure SaveToFile_DateTimeCellWithTime_UsesLocaleAwareDateTimeFormat;

    [Test]
    procedure SaveToFile_DateTimeCellWithNumberFormat_UsesCustomFormat;

    [Test]
    procedure RoundTrip_DateTimeWithTime_PreservesValue;

    [Test]
    procedure SaveToFile_WithFontColor_ContainsFontColor;

    [Test]
    procedure RoundTrip_FontColor_PreservesFontColor;

    [Test]
    procedure AddSheet_NewSheet_IsVisible;

    [Test]
    procedure SaveToFile_HiddenSheet_ContainsStateHidden;

    [Test]
    procedure SaveToFile_VeryHiddenSheet_ContainsStateVeryHidden;

    [Test]
    procedure RoundTrip_SheetVisibility_PreservesVisibility;

    [Test]
    procedure SaveToFile_AllSheetsHidden_RaisesException;

    [Test]
    procedure RoundTrip_CustomNumberFormat_PreservesFormat;

    [Test]
    procedure RoundTrip_CustomDateFormat_PreservesFormat;

    [Test]
    procedure LoadFromFile_ExcelAuthoredNumberFormats_ArePreserved;
  end;

implementation

uses
  System.Zip,
  Office4D.Errors,
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
    Assert.IsTrue(Pos('<styleSheet', StylesXml) > 0, 'Should contain styleSheet element');
    Assert.IsTrue(Pos('<cellXfs', StylesXml) > 0, 'Should contain cellXfs element');
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

procedure TExcelWriteTests.SaveToFile_WithFontColor_ContainsFontColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red font color';
  Sheet.Cell['A1'].FontColor := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('FF0000', StylesXml) > 0, 'Should contain font color');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_FontColor_PreservesFontColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red font color';
  Sheet.Cell['A1'].FontColor := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Cardinal($FF0000), Workbook2.Sheets[0].Cell['A1'].FontColor);
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
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Thin;

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
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Thin;
  Sheet.Cell['B1'].AsString := 'No border';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Ord(TExcelBorderStyle.Thin), Ord(Workbook2.Sheets[0].Cell['A1'].BorderStyle[AllBorderSides]));
  Assert.AreEqual(Ord(TExcelBorderStyle.None), Ord(Workbook2.Sheets[0].Cell['B1'].BorderStyle[AllBorderSides]));
end;

procedure TExcelWriteTests.SaveToFile_WithBorderColor_ContainsBorderColor;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Red border';
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor[AllBorderSides] := $FF0000;

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
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Thin;
  Sheet.Cell['A1'].BorderColor[AllBorderSides] := $FF0000;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Cardinal($FF0000), Workbook2.Sheets[0].Cell['A1'].BorderColor[AllBorderSides]);
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
  Sheet.Cell['A1'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Medium;
  Sheet.Cell['A1'].BorderColor[AllBorderSides] := $0000FF;
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
  Assert.AreEqual(Ord(TExcelBorderStyle.Medium), Ord(Cell.BorderStyle[AllBorderSides]));
  Assert.AreEqual(Cardinal($0000FF), Cell.BorderColor[AllBorderSides]);
  Assert.AreEqual(Ord(TExcelHAlign.Center), Ord(Cell.HAlign));
  Assert.AreEqual(Ord(TExcelVAlign.Bottom), Ord(Cell.VAlign));
  Assert.IsTrue(Cell.WrapText, 'Should have wrapText');
  Assert.AreEqual(Double(30), Workbook2.Sheets[0].GetRowHeight(1), 0.01);
end;

procedure TExcelWriteTests.SaveToFile_EmptyStringCell_WritesNoInvalidSharedStringIndex;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Filled';
  Sheet.Cell['B1'].AsString := '';
  Sheet.Cell['C1'].AsString := '';
  Sheet.Cell['C1'].Bold := True;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsFalse(Pos('<v>-1</v>', SheetXml) > 0, 'Empty string cells should not reference sharedStrings index -1');
    Assert.IsFalse(Pos('r="B1"', SheetXml) > 0, 'Unstyled empty string cell should be omitted');
    Assert.IsTrue(Pos('r="C1"', SheetXml) > 0, 'Styled empty string cell should be written');
    Assert.IsFalse(Pos('r="C1" s="1" t="s"', SheetXml) > 0, 'Styled empty string cell should not be a shared string cell');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_EmptyStringCell_PreservesAdjacentValues;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Left';
  Sheet.Cell['B1'].AsString := '';
  Sheet.Cell['C1'].AsString := 'Right';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('Left', Workbook2.Sheets[0].Cell['A1'].AsString);
  Assert.AreEqual('', Workbook2.Sheets[0].Cell['B1'].AsString);
  Assert.AreEqual('Right', Workbook2.Sheets[0].Cell['C1'].AsString);
end;

procedure TExcelWriteTests.SaveToFile_DateCell_UsesLocaleAwareShortDateFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := EncodeDate(2024, 6, 15);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('numFmtId="14"', StylesXml) > 0, 'Date cell should use built-in locale-aware short date format 14');
    Assert.IsFalse(Pos('yyyy-mm-dd', StylesXml) > 0, 'Should not contain a hardcoded date format code');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_DateTimeCellWithTime_UsesLocaleAwareDateTimeFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := EncodeDate(2024, 6, 15) + EncodeTime(14, 30, 0, 0);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('numFmtId="22"', StylesXml) > 0, 'Date cell with time should use built-in locale-aware date/time format 22');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_DateTimeCell_WritesCorrectExcelValue;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := EncodeDate(2024, 6, 15);

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('<c r="A1" s="1"><v>45458</v></c>', SheetXml) > 0, 'A1 should contain the literal Excel value 45458 for 2024-06-15, got: ' + SheetXml);
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_DateTimeCellWithNumberFormat_UsesCustomFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := EncodeDate(2024, 6, 15) + EncodeTime(14, 30, 0, 0);
  Sheet.Cell['A1'].NumberFormat := 'dd.mm.yyyy hh:mm';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('formatCode="dd.mm.yyyy hh:mm"', StylesXml) > 0, 'Custom number format should be written');
    Assert.IsTrue(Pos('numFmtId="165"', StylesXml) > 0, 'Custom format should be referenced by the cell style');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_DateTimeWithTime_PreservesValue;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  const TestValue = EncodeDate(2024, 6, 15) + EncodeTime(14, 30, 45, 0);
  Sheet.Cell['A1'].AsDateTime := TestValue;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Double(TestValue), Double(Workbook2.Sheets[0].Cell['A1'].AsDateTime), 1E-8);
end;

procedure TExcelWriteTests.RoundTrip_SpecialCharacters_ArePreserved;
begin
  const Special = 'R&D <tag> "q" ''a'' 5>3';
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := Special;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(Special, Workbook2.Sheets[0].Cell['A1'].AsString, 'Cell special characters should round-trip');
end;

procedure TExcelWriteTests.AddSheet_NewSheet_IsVisible;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');

  Assert.AreEqual(Ord(TExcelSheetVisibility.Visible), Ord(Sheet.Visibility));
end;

procedure TExcelWriteTests.SaveToFile_HiddenSheet_ContainsStateHidden;
begin
  FWorkbook.AddSheet('Visible').Cell['A1'].AsString := 'Shown';
  FWorkbook.AddSheet('Hidden').Visibility := TExcelSheetVisibility.Hidden;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const WorkbookXml = Package.GetPartContent('xl/workbook.xml');
    const StateCount = (Length(WorkbookXml) - Length(StringReplace(WorkbookXml, 'state="', '', [rfReplaceAll]))) div Length('state="');
    Assert.IsTrue(Pos('state="hidden"', WorkbookXml) > 0, 'Hidden sheet should have state="hidden"');
    Assert.AreEqual(1, StateCount, 'Only the hidden sheet should carry a state attribute');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.SaveToFile_VeryHiddenSheet_ContainsStateVeryHidden;
begin
  FWorkbook.AddSheet('Visible').Cell['A1'].AsString := 'Shown';
  FWorkbook.AddSheet('VeryHidden').Visibility := TExcelSheetVisibility.VeryHidden;

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const WorkbookXml = Package.GetPartContent('xl/workbook.xml');
    Assert.IsTrue(Pos('state="veryHidden"', WorkbookXml) > 0, 'Very hidden sheet should have state="veryHidden"');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_SheetVisibility_PreservesVisibility;
begin
  FWorkbook.AddSheet('Shown').Cell['A1'].AsString := 'Data';
  FWorkbook.AddSheet('Concealed').Visibility := TExcelSheetVisibility.Hidden;
  FWorkbook.AddSheet('Buried').Visibility := TExcelSheetVisibility.VeryHidden;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual(3, Workbook2.SheetCount);
  Assert.AreEqual(Ord(TExcelSheetVisibility.Visible), Ord(Workbook2.Sheets[0].Visibility));
  Assert.AreEqual(Ord(TExcelSheetVisibility.Hidden), Ord(Workbook2.Sheets[1].Visibility));
  Assert.AreEqual(Ord(TExcelSheetVisibility.VeryHidden), Ord(Workbook2.Sheets[2].Visibility));
end;

procedure TExcelWriteTests.SaveToFile_AllSheetsHidden_RaisesException;
begin
  FWorkbook.AddSheet('Sheet1').Visibility := TExcelSheetVisibility.Hidden;

  Assert.WillRaise(
    procedure
    begin
      FWorkbook.SaveToFile(FTempFile);
    end,
    EExcelWorkbookException
  );
end;

procedure TExcelWriteTests.RoundTrip_CustomNumberFormat_PreservesFormat;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsFloat := 1.5;
  Sheet.Cell['A1'].NumberFormat := '0.000';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('0.000', Workbook2.Sheets[0].Cell['A1'].NumberFormat, 'Custom number format should survive a load');

  // Save the reloaded workbook again: the format must also survive a full
  // load-save cycle, not just live on the in-memory cell.
  Workbook2.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const StylesXml = Package.GetPartContent('xl/styles.xml');
    Assert.IsTrue(Pos('formatCode="0.000"', StylesXml) > 0, 'Resaved styles.xml should still contain the custom format');
  finally
    Package.Free;
  end;
end;

procedure TExcelWriteTests.RoundTrip_CustomDateFormat_PreservesFormat;
begin
  const TestValue = EncodeDate(2024, 6, 15) + EncodeTime(14, 30, 0, 0);
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsDateTime := TestValue;
  Sheet.Cell['A1'].NumberFormat := 'dd.mm.yyyy hh:mm';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('dd.mm.yyyy hh:mm', Workbook2.Sheets[0].Cell['A1'].NumberFormat, 'Custom date format should survive a load');
  Assert.AreEqual(Double(TestValue), Double(Workbook2.Sheets[0].Cell['A1'].AsDateTime), 1E-8, 'Date value should be preserved');
end;

procedure TExcelWriteTests.LoadFromFile_ExcelAuthoredNumberFormats_ArePreserved;

  procedure AddPart(const Zip: TZipFile; const PartName, Content: string);
  begin
    Zip.Add(TEncoding.UTF8.GetBytes(Content), PartName);
  end;

const
  SpreadsheetNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  OfficeDocRelsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
begin
  // Excel-authored files reference built-in formats (like 9 = "0%") by id only
  // and start custom numFmt ids at 164, one below the 165 this library uses,
  // so this file cannot be produced through the library's own save path.
  const SourceFile = TPath.Combine(TPath.GetTempPath, 'excel_authored_' + TGUID.NewGuid.ToString + '.xlsx');
  var Zip := TZipFile.Create;
  try
    Zip.Open(SourceFile, zmWrite);
    AddPart(Zip, 'xl/workbook.xml',
      '<workbook xmlns="' + SpreadsheetNs + '" xmlns:r="' + OfficeDocRelsNs + '">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    AddPart(Zip, 'xl/styles.xml',
      '<styleSheet xmlns="' + SpreadsheetNs + '">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellXfs count="3">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>' +
      '</cellXfs></styleSheet>');
    AddPart(Zip, 'xl/worksheets/sheet1.xml',
      '<worksheet xmlns="' + SpreadsheetNs + '"><sheetData>' +
      '<row r="1"><c r="A1" s="1"><v>1.5</v></c><c r="B1" s="2"><v>0.25</v></c></row>' +
      '</sheetData></worksheet>');
    Zip.Close;

    FWorkbook.LoadFromFile(SourceFile);
    Assert.AreEqual('0.000', FWorkbook.Sheets[0].Cell['A1'].NumberFormat, 'Custom numFmt 164 should be applied to the cell');
    Assert.AreEqual('0%', FWorkbook.Sheets[0].Cell['B1'].NumberFormat, 'Built-in numFmt 9 should map to its format code');

    FWorkbook.SaveToFile(FTempFile);

    const Workbook2 = TExcelWorkbookFactory.Create;
    Workbook2.LoadFromFile(FTempFile);
    Assert.AreEqual('0.000', Workbook2.Sheets[0].Cell['A1'].NumberFormat, 'Custom format should survive a resave');
    Assert.AreEqual('0%', Workbook2.Sheets[0].Cell['B1'].NumberFormat, 'Built-in format should survive a resave');
  finally
    Zip.Free;
    if TFile.Exists(SourceFile) then
      TFile.Delete(SourceFile);
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelWriteTests);

end.
