unit Office4D.Tests.Word.Write;

interface

uses
  System.SysUtils,
  System.IOUtils,
  System.Classes,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Word;

type
  [TestFixture]
  TWordWriteTests = class(TOffice4DTests)
  private
    FDoc: IWordDocument;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NewDocument_CreatesFile;

    [Test]
    procedure SaveToFile_WithParagraph_ContainsText;

    [Test]
    procedure RoundTrip_LoadModifySave_PreservesContent;

    [Test]
    procedure RoundTrip_SpecialCharacters_ArePreserved;

    [Test]
    procedure SaveToFile_ValidatesAsZip;

    [Test]
    procedure SaveToFile_ContainsContentTypes;

    [Test]
    procedure SaveToFile_ContainsDocumentXml;

    [Test]
    procedure SaveToStream_WritesToStream;

    [Test]
    procedure LoadFromStream_ReadsFromStream;

    [Test]
    procedure RoundTrip_StreamBased_PreservesContent;

    [Test]
    procedure SaveToFile_WithBoldRun_ContainsBoldXml;

    [Test]
    procedure SaveToFile_WithItalicRun_ContainsItalicXml;

    [Test]
    procedure SaveToFile_WithUnderlineRun_ContainsUnderlineXml;

    [Test]
    procedure SaveToFile_WithAllFormatting_ContainsAllXml;

    [Test]
    procedure SaveToFile_WithHyperlink_ContainsHyperlinkXml;

    [Test]
    procedure SaveToFile_WithHyperlink_ContainsRelationship;

    [Test]
    procedure RoundTrip_Hyperlink_PreservesUrl;

    [Test]
    procedure SaveToFile_WithBulletList_ContainsNumberingXml;

    [Test]
    procedure SaveToFile_WithNumberedList_ContainsNumberingXml;

    [Test]
    procedure RoundTrip_BulletList_PreservesStyle;

    [Test]
    procedure SaveToFile_WithLandscape_ContainsOrientation;

    [Test]
    procedure SaveToFile_WithMargins_ContainsPgMar;

    [Test]
    procedure SaveToFile_WithHeader_ContainsHeaderXml;

    [Test]
    procedure SaveToFile_WithFooter_ContainsFooterXml;

    [Test]
    procedure RoundTrip_Header_PreservesText;

    [Test]
    procedure RoundTrip_Footer_PreservesText;

    [Test]
    procedure SaveToFile_WithFontName_ContainsFontXml;

    [Test]
    procedure SaveToFile_WithFontSize_ContainsSizeXml;

    [Test]
    procedure SaveToFile_WithFontColor_ContainsColorXml;

    [Test]
    procedure SaveToFile_WithAlignment_ContainsJcXml;

    [Test]
    procedure SaveToFile_WithCenterAlignment_ContainsCenterXml;

    [Test]
    procedure SaveToFile_WithRightAlignment_ContainsRightXml;

    [Test]
    procedure SaveToFile_WithJustifyAlignment_ContainsBothXml;

    [Test]
    procedure SaveToFile_WithPageBreak_ContainsBreakXml;

    [Test]
    procedure SaveToFile_WithDoubleSpacing_ContainsSpacingXml;

    [Test]
    procedure SaveToFile_WithSpaceBefore_ContainsBeforeAttr;

    [Test]
    procedure SaveToFile_WithLeftIndent_ContainsIndXml;

    [Test]
    procedure SaveToFile_WithFirstLineIndent_ContainsFirstLineAttr;

    [Test]
    procedure SaveToFile_WithTableBorders_ContainsBordersXml;

    [Test]
    procedure SaveToFile_WithCellShading_ContainsShdXml;

    [Test]
    procedure SaveToFile_WithColumnWidths_ContainsGridXml;

    [Test]
    procedure SaveToFile_WithImage_ContainsMediaFolder;

    [Test]
    procedure SaveToFile_WithImage_ContainsDrawingXml;

    [Test]
    procedure SaveToFile_WithImage_ContainsImageRelationship;

    [Test]
    procedure SaveToFile_WithPngAndJpeg_ContainsBothContentTypes;
  end;

implementation

uses
  System.Zip,
  Office4D.Package;

{ TWordWriteTests }

procedure TWordWriteTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'test_output_' + TGUID.NewGuid.ToString + '.docx');
end;

procedure TWordWriteTests.TearDown;
begin
  FDoc := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TWordWriteTests.SaveToFile_NewDocument_CreatesFile;
begin
  FDoc.AddParagraph.AddRun('Test');

  FDoc.SaveToFile(FTempFile);

  Assert.IsTrue(TFile.Exists(FTempFile), 'File should be created');
end;

procedure TWordWriteTests.SaveToFile_WithParagraph_ContainsText;
begin
  FDoc.AddParagraph.AddRun('Hello from Office4D');
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.Contains(Doc2.Text, 'Hello from Office4D');
end;

procedure TWordWriteTests.RoundTrip_LoadModifySave_PreservesContent;
begin
  FDoc.LoadFromFile(GetWordSamplePath);
  FDoc.AddParagraph.AddRun('Added by test');
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.Contains(Doc2.Text, 'Hello');
  Assert.Contains(Doc2.Text, 'Added by test');
end;

procedure TWordWriteTests.SaveToFile_ValidatesAsZip;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.SaveToFile(FTempFile);

  var Zip := TZipFile.Create;
  try
    Zip.Open(FTempFile, zmRead);
    Assert.IsTrue(Zip.FileCount > 0, 'ZIP should contain files');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_ContainsContentTypes;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('[Content_Types].xml'), 'Should contain [Content_Types].xml');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_ContainsDocumentXml;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/document.xml'), 'Should contain word/document.xml');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToStream_WritesToStream;
begin
  FDoc.AddParagraph.AddRun('Stream Test');

  var Stream := TMemoryStream.Create;
  try
    FDoc.SaveToStream(Stream);
    Assert.IsTrue(Stream.Size > 0, 'Stream should contain data');
  finally
    Stream.Free;
  end;
end;

procedure TWordWriteTests.LoadFromStream_ReadsFromStream;
begin
  var Stream := TFileStream.Create(GetWordSamplePath, fmOpenRead or fmShareDenyWrite);
  try
    var Doc := TWordDocumentFactory.CreateDocument;
    Doc.LoadFromStream(Stream);
    Assert.Contains(Doc.Text, 'Hello');
  finally
    Stream.Free;
  end;
end;

procedure TWordWriteTests.RoundTrip_StreamBased_PreservesContent;
begin
  FDoc.AddParagraph.AddRun('Stream Round Trip');

  var Stream := TMemoryStream.Create;
  try
    FDoc.SaveToStream(Stream);
    Stream.Position := 0;

    var Doc2 := TWordDocumentFactory.CreateDocument;
    Doc2.LoadFromStream(Stream);

    Assert.Contains(Doc2.Text, 'Stream Round Trip');
  finally
    Stream.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithBoldRun_ContainsBoldXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Bold Text');
  Run.Bold := True;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:b/>', DocXml) > 0, 'Should contain bold element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithItalicRun_ContainsItalicXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Italic Text');
  Run.Italic := True;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:i/>', DocXml) > 0, 'Should contain italic element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithUnderlineRun_ContainsUnderlineXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Underline Text');
  Run.Underline := True;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:u w:val="single"/>', DocXml) > 0, 'Should contain underline element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithAllFormatting_ContainsAllXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Formatted Text');
  Run.Bold := True;
  Run.Italic := True;
  Run.Underline := True;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:rPr>', DocXml) > 0, 'Should contain run properties');
    Assert.IsTrue(Pos('<w:b/>', DocXml) > 0, 'Should contain bold');
    Assert.IsTrue(Pos('<w:i/>', DocXml) > 0, 'Should contain italic');
    Assert.IsTrue(Pos('<w:u w:val="single"/>', DocXml) > 0, 'Should contain underline');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithHyperlink_ContainsHyperlinkXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Click here');
  Run.Hyperlink := 'https://www.example.com';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:hyperlink', DocXml) > 0, 'Should contain hyperlink element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithHyperlink_ContainsRelationship;
begin
  var Run := FDoc.AddParagraph.AddRun('Click here');
  Run.Hyperlink := 'https://www.example.com';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var RelsXml := Package.GetPartContent('word/_rels/document.xml.rels');
    Assert.IsTrue(Pos('https://www.example.com', RelsXml) > 0, 'Should contain hyperlink URL in relationships');
    Assert.IsTrue(Pos('hyperlink', RelsXml) > 0, 'Should contain hyperlink relationship type');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.RoundTrip_Hyperlink_PreservesUrl;
begin
  var Run := FDoc.AddParagraph.AddRun('Visit site');
  Run.Hyperlink := 'https://www.test.com/page';
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.AreEqual('https://www.test.com/page', Doc2.Paragraphs[0].Runs[0].Hyperlink);
end;

procedure TWordWriteTests.SaveToFile_WithBulletList_ContainsNumberingXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('First item');
  Para.ListStyle := TListStyle.Bullet;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/numbering.xml'), 'Should contain numbering.xml');
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:numPr>', DocXml) > 0, 'Should contain numPr element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithNumberedList_ContainsNumberingXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('First item');
  Para.ListStyle := TListStyle.Numbered;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/numbering.xml'), 'Should contain numbering.xml');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.RoundTrip_BulletList_PreservesStyle;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Bullet item');
  Para.ListStyle := TListStyle.Bullet;
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.AreEqual(TListStyle.Bullet, Doc2.Paragraphs[0].ListStyle);
end;

procedure TWordWriteTests.SaveToFile_WithLandscape_ContainsOrientation;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.PageOrientation := TPageOrientation.Landscape;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:orient="landscape"', DocXml) > 0, 'Should contain landscape orientation');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithMargins_ContainsPgMar;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.PageMargins := TPageMargins.Create(1440, 1440, 1440, 1440);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:pgMar', DocXml) > 0, 'Should contain page margins');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithHeader_ContainsHeaderXml;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.Header.Text := 'My Header';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/header1.xml'), 'Should contain header1.xml');
    var HeaderXml := Package.GetPartContent('word/header1.xml');
    Assert.IsTrue(Pos('My Header', HeaderXml) > 0, 'Header should contain text');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithFooter_ContainsFooterXml;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.Footer.Text := 'My Footer';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/footer1.xml'), 'Should contain footer1.xml');
    var FooterXml := Package.GetPartContent('word/footer1.xml');
    Assert.IsTrue(Pos('My Footer', FooterXml) > 0, 'Footer should contain text');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.RoundTrip_Header_PreservesText;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.Header.Text := 'Header Text';
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.AreEqual('Header Text', Doc2.Header.Text);
end;

procedure TWordWriteTests.RoundTrip_Footer_PreservesText;
begin
  FDoc.AddParagraph.AddRun('Test');
  FDoc.Footer.Text := 'Footer Text';
  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);

  Assert.AreEqual('Footer Text', Doc2.Footer.Text);
end;

procedure TWordWriteTests.SaveToFile_WithFontName_ContainsFontXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Arial Text');
  Run.FontName := 'Arial';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:rFonts w:ascii="Arial"', DocXml) > 0, 'Should contain font name');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithFontSize_ContainsSizeXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Large Text');
  Run.FontSize := 24;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:sz w:val="24"/>', DocXml) > 0, 'Should contain font size');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithFontColor_ContainsColorXml;
begin
  var Run := FDoc.AddParagraph.AddRun('Red Text');
  Run.FontColor := 'FF0000';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:color w:val="FF0000"/>', DocXml) > 0, 'Should contain font color');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithAlignment_ContainsJcXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Centered Text');
  Para.Alignment := TParagraphAlignment.Center;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:jc w:val="center"/>', DocXml) > 0, 'Should contain alignment element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithCenterAlignment_ContainsCenterXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Center');
  Para.Alignment := TParagraphAlignment.Center;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:val="center"', DocXml) > 0, 'Should contain center alignment');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithRightAlignment_ContainsRightXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Right');
  Para.Alignment := TParagraphAlignment.Right;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:val="right"', DocXml) > 0, 'Should contain right alignment');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithJustifyAlignment_ContainsBothXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Justified text');
  Para.Alignment := TParagraphAlignment.Justify;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:val="both"', DocXml) > 0, 'Should contain both (justify) alignment');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithPageBreak_ContainsBreakXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Before page break');
  Para.AddPageBreak;

  var Para2 := FDoc.AddParagraph;
  Para2.AddRun('After page break');
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:br w:type="page"/>', DocXml) > 0, 'Should contain page break element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithDoubleSpacing_ContainsSpacingXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Double spaced paragraph');
  Para.LineSpacing := TLineSpacing.Double;
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:spacing', DocXml) > 0, 'Should contain spacing element');
    Assert.IsTrue(Pos('w:line="480"', DocXml) > 0, 'Should contain double spacing value (480)');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithSpaceBefore_ContainsBeforeAttr;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Paragraph with space before');
  Para.LineSpacing := TLineSpacing.Create(240, TLineSpacingRule.Auto, 200, 100);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:before="200"', DocXml) > 0, 'Should contain before attribute');
    Assert.IsTrue(Pos('w:after="100"', DocXml) > 0, 'Should contain after attribute');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithLeftIndent_ContainsIndXml;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Indented paragraph');
  Para.Indent := TParagraphIndent.Create(720, 0, 0);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:ind', DocXml) > 0, 'Should contain ind element');
    Assert.IsTrue(Pos('w:left="720"', DocXml) > 0, 'Should contain left indent value');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithFirstLineIndent_ContainsFirstLineAttr;
begin
  var Para := FDoc.AddParagraph;
  Para.AddRun('Paragraph with first line indent');
  Para.Indent := TParagraphIndent.Create(0, 0, 720);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('w:firstLine="720"', DocXml) > 0, 'Should contain firstLine attribute');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithTableBorders_ContainsBordersXml;
begin
  FDoc.AddParagraph.AddRun('Table with borders:');
  var Table := FDoc.AddTable(2, 2);
  Table.Borders := TTableBorders.All(TTableBorder.Create(TBorderStyle.Single, 8, '000000'));
  Table.Cells[0, 0].Text := 'A1';
  Table.Cells[0, 1].Text := 'B1';
  Table.Cells[1, 0].Text := 'A2';
  Table.Cells[1, 1].Text := 'B2';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:tblBorders>', DocXml) > 0, 'Should contain tblBorders element');
    Assert.IsTrue(Pos('<w:top w:val="single"', DocXml) > 0, 'Should contain top border');
    Assert.IsTrue(Pos('<w:insideH w:val="single"', DocXml) > 0, 'Should contain inside horizontal border');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithCellShading_ContainsShdXml;
begin
  FDoc.AddParagraph.AddRun('Table with shading:');
  var Table := FDoc.AddTable(2, 2);
  Table.Cells[0, 0].Shading := '4472C4';
  Table.Cells[0, 0].Text := 'Blue cell';
  Table.Cells[0, 1].Text := 'Normal';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:shd w:val="clear" w:fill="4472C4"', DocXml) > 0, 'Should contain shd element with fill color');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithColumnWidths_ContainsGridXml;
begin
  FDoc.AddParagraph.AddRun('Table with custom column widths:');
  var Table := FDoc.AddTable(2, 3);
  Table.SetColumnWidths([2000, 3000, 1000]);
  Table.Cells[0, 0].Text := 'Narrow';
  Table.Cells[0, 1].Text := 'Wide';
  Table.Cells[0, 2].Text := 'Tiny';
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:tblGrid>', DocXml) > 0, 'Should contain tblGrid element');
    Assert.IsTrue(Pos('<w:gridCol w:w="2000"/>', DocXml) > 0, 'Should contain first column width');
    Assert.IsTrue(Pos('<w:gridCol w:w="3000"/>', DocXml) > 0, 'Should contain second column width');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithImage_ContainsMediaFolder;
begin
  var Para := FDoc.AddParagraph;
  var Run := Para.AddRun('');
  var PngData: TBytes := [$89, $50, $4E, $47, $0D, $0A, $1A, $0A, $00, $00, $00, $0D,
    $49, $48, $44, $52, $00, $00, $00, $01, $00, $00, $00, $01, $08, $02, $00, $00,
    $00, $90, $77, $53, $DE, $00, $00, $00, $0C, $49, $44, $41, $54, $08, $D7, $63,
    $F8, $CF, $C0, $00, $00, $00, $03, $00, $01, $00, $18, $DD, $8D, $B4, $00, $00,
    $00, $00, $49, $45, $4E, $44, $AE, $42, $60, $82];
  Run.AddImage(PngData, 'png', 100, 100);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('word/media/image1.png'), 'Should contain image in media folder');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithImage_ContainsDrawingXml;
begin
  var Para := FDoc.AddParagraph;
  var Run := Para.AddRun('');
  var PngData: TBytes := [$89, $50, $4E, $47, $0D, $0A, $1A, $0A, $00, $00, $00, $0D,
    $49, $48, $44, $52, $00, $00, $00, $01, $00, $00, $00, $01, $08, $02, $00, $00,
    $00, $90, $77, $53, $DE, $00, $00, $00, $0C, $49, $44, $41, $54, $08, $D7, $63,
    $F8, $CF, $C0, $00, $00, $00, $03, $00, $01, $00, $18, $DD, $8D, $B4, $00, $00,
    $00, $00, $49, $45, $4E, $44, $AE, $42, $60, $82];
  Run.AddImage(PngData, 'png', 100, 100);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var DocXml := Package.GetPartContent('word/document.xml');
    Assert.IsTrue(Pos('<w:drawing>', DocXml) > 0, 'Should contain drawing element');
    Assert.IsTrue(Pos('<wp:inline', DocXml) > 0, 'Should contain inline element');
    Assert.IsTrue(Pos('<a:blip', DocXml) > 0, 'Should contain blip element');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithImage_ContainsImageRelationship;
begin
  var Para := FDoc.AddParagraph;
  var Run := Para.AddRun('');
  var PngData: TBytes := [$89, $50, $4E, $47, $0D, $0A, $1A, $0A, $00, $00, $00, $0D,
    $49, $48, $44, $52, $00, $00, $00, $01, $00, $00, $00, $01, $08, $02, $00, $00,
    $00, $90, $77, $53, $DE, $00, $00, $00, $0C, $49, $44, $41, $54, $08, $D7, $63,
    $F8, $CF, $C0, $00, $00, $00, $03, $00, $01, $00, $18, $DD, $8D, $B4, $00, $00,
    $00, $00, $49, $45, $4E, $44, $AE, $42, $60, $82];
  Run.AddImage(PngData, 'png', 100, 100);
  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var RelsXml := Package.GetPartContent('word/_rels/document.xml.rels');
    Assert.IsTrue(Pos('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"', RelsXml) > 0,
      'Should contain image relationship type');
    Assert.IsTrue(Pos('Target="media/image1.png"', RelsXml) > 0, 'Should contain image target');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.SaveToFile_WithPngAndJpeg_ContainsBothContentTypes;
begin
  var Para := FDoc.AddParagraph;
  var PngData: TBytes := [$89, $50, $4E, $47, $0D, $0A, $1A, $0A, $00, $00, $00, $0D,
    $49, $48, $44, $52, $00, $00, $00, $01, $00, $00, $00, $01, $08, $02, $00, $00,
    $00, $90, $77, $53, $DE, $00, $00, $00, $0C, $49, $44, $41, $54, $08, $D7, $63,
    $F8, $CF, $C0, $00, $00, $00, $03, $00, $01, $00, $18, $DD, $8D, $B4, $00, $00,
    $00, $00, $49, $45, $4E, $44, $AE, $42, $60, $82];
  var JpegData: TBytes := [$FF, $D8, $FF, $E0, $00, $10, $4A, $46, $49, $46, $00, $01,
    $01, $00, $00, $01, $00, $01, $00, $00, $FF, $DB, $00, $43, $00, $08, $06, $06,
    $07, $06, $05, $08, $07, $07, $07, $09, $09, $FF, $D9];

  var Run1 := Para.AddRun('');
  Run1.AddImage(PngData, 'png', 50, 50);

  var Run2 := Para.AddRun('');
  Run2.AddImage(JpegData, 'jpeg', 50, 50);

  FDoc.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    var ContentTypes := Package.GetPartContent('[Content_Types].xml');
    Assert.IsTrue(Pos('Extension="png" ContentType="image/png"', ContentTypes) > 0, 'Should contain PNG content type');
    Assert.IsTrue(Pos('Extension="jpeg" ContentType="image/jpeg"', ContentTypes) > 0, 'Should contain JPEG content type');
  finally
    Package.Free;
  end;
end;

procedure TWordWriteTests.RoundTrip_SpecialCharacters_ArePreserved;
begin
  const Special = 'R&D <tag> "q" ''a'' 5>3';
  FDoc.AddParagraph.AddRun(Special);

  FDoc.SaveToFile(FTempFile);

  var Doc2 := TWordDocumentFactory.CreateDocument;
  Doc2.LoadFromFile(FTempFile);
  Assert.AreEqual(Special, Doc2.Paragraphs[0].Runs[0].Text, 'Run special characters should round-trip');
end;

initialization
  TDUnitX.RegisterTestFixture(TWordWriteTests);

end.
