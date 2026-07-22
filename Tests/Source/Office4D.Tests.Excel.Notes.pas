unit Office4D.Tests.Excel.Notes;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Excel;

type
  [TestFixture]
  TExcelNotesTests = class(TOffice4DTests)
  private
    FWorkbook: IExcelWorkbook;
    FTempFile: string;

  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure SaveToFile_NoNotes_OmitsCommentsAndVmlParts;

    [Test]
    procedure SaveToFile_WithNote_CreatesCommentsPart;

    [Test]
    procedure SaveToFile_WithNote_CreatesVmlDrawingPart;

    [Test]
    procedure SaveToFile_WithNote_CreatesSheetRelsWithBothRelationships;

    [Test]
    procedure SaveToFile_WithNote_WorksheetReferencesLegacyDrawing;

    [Test]
    procedure SaveToFile_WithNote_LegacyDrawingFollowsMergeCells;

    [Test]
    procedure RoundTrip_Note_PreservesText;

    [Test]
    procedure RoundTrip_MultipleNotesOnDifferentCells_PreservesEachIndependently;

    [Test]
    procedure RoundTrip_SettingNoteToEmptyString_RemovesNote;

    [Test]
    procedure RoundTrip_NoteWithSpecialCharacters_ArePreserved;

    [Test]
    procedure SaveToFile_OnlySecondSheetHasNotes_NumbersCommentsFileCorrectly;

    [Test]
    procedure GetNote_UnsetCell_ReturnsEmptyString;

    [Test]
    procedure SaveToFile_WithNote_VmlAnchorsNoteToItsCell;
  end;

implementation

uses
  Office4D.Package;

{ TExcelNotesTests }

procedure TExcelNotesTests.Setup;
begin
  FWorkbook := TExcelWorkbookFactory.Create;
  FTempFile := TPath.Combine(TPath.GetTempPath, 'notes_test_' + TGUID.NewGuid.ToString + '.xlsx');
end;

procedure TExcelNotesTests.TearDown;
begin
  FWorkbook := nil;
  if TFile.Exists(FTempFile) then
    TFile.Delete(FTempFile);
end;

procedure TExcelNotesTests.GetNote_UnsetCell_ReturnsEmptyString;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Assert.AreEqual('', Sheet.Note['A1'], 'A cell with no note set must return an empty string, not raise');
end;

procedure TExcelNotesTests.SaveToFile_NoNotes_OmitsCommentsAndVmlParts;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Plain';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsFalse(Package.PartExists('xl/comments1.xml'),
      'A workbook with no notes at all must not emit any comments part');
    Assert.IsFalse(Package.PartExists('xl/drawings/vmlDrawing1.vml'),
      'A workbook with no notes at all must not emit any VML drawing part');
    Assert.IsFalse(Package.PartExists('xl/worksheets/_rels/sheet1.xml.rels'),
      'A sheet with no notes must not emit a .rels file at all');

    const ContentTypesXml = Package.GetPartContent('[Content_Types].xml');
    Assert.IsFalse(Pos('Extension="vml"', ContentTypesXml) > 0,
      'Content types must not declare the vml extension when nothing uses it');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.SaveToFile_WithNote_CreatesCommentsPart;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'This is a note';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('xl/comments1.xml'), 'comments1.xml should be created');

    const CommentsXml = Package.GetPartContent('xl/comments1.xml');
    Assert.IsTrue(Pos('ref="A1"', CommentsXml) > 0, 'The comment should be keyed to the correct cell address');
    Assert.IsTrue(Pos('This is a note', CommentsXml) > 0, 'The note text should appear in the comments part');

    const ContentTypesXml = Package.GetPartContent('[Content_Types].xml');
    Assert.IsTrue(Pos('PartName="/xl/comments1.xml"', ContentTypesXml) > 0,
      'Content types must declare an Override for the comments part');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.SaveToFile_WithNote_CreatesVmlDrawingPart;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B3'].AsString := 'Data';
  Sheet.Note['B3'] := 'Note text';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('xl/drawings/vmlDrawing1.vml'), 'vmlDrawing1.vml should be created');

    const VmlXml = Package.GetPartContent('xl/drawings/vmlDrawing1.vml');
    // B3 is column 2 (0-based col 1), row 3 (0-based row 2) -- x:Row/x:Column are 0-based,
    // unlike the rest of this library's 1-based convention.
    Assert.IsTrue(Pos('<x:Row>2</x:Row>', VmlXml) > 0, 'x:Row should be the 0-based row index');
    Assert.IsTrue(Pos('<x:Column>1</x:Column>', VmlXml) > 0, 'x:Column should be the 0-based column index');

    const ContentTypesXml = Package.GetPartContent('[Content_Types].xml');
    Assert.IsTrue(Pos('Extension="vml"', ContentTypesXml) > 0,
      'Content types must declare the vml Default extension when a note exists');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.SaveToFile_WithNote_CreatesSheetRelsWithBothRelationships;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'Note';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsTrue(Package.PartExists('xl/worksheets/_rels/sheet1.xml.rels'),
      'The sheet must get its own .rels file -- new package capability this feature required');

    const RelsXml = Package.GetPartContent('xl/worksheets/_rels/sheet1.xml.rels');
    Assert.IsTrue(Pos('Target="../drawings/vmlDrawing1.vml"', RelsXml) > 0,
      'Relative target must resolve correctly from xl/worksheets/ to xl/drawings/');
    Assert.IsTrue(Pos('Target="../comments1.xml"', RelsXml) > 0,
      'Relative target must resolve correctly from xl/worksheets/ to xl/');
    Assert.IsTrue(Pos('/relationships/vmlDrawing"', RelsXml) > 0, 'vmlDrawing relationship type must be present');
    Assert.IsTrue(Pos('/relationships/comments"', RelsXml) > 0, 'comments relationship type must be present');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.SaveToFile_WithNote_WorksheetReferencesLegacyDrawing;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'Note';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    Assert.IsTrue(Pos('<legacyDrawing r:id="rId1"/>', SheetXml) > 0,
      'The worksheet must reference the VML drawing via legacyDrawing');
    // legacyDrawing's r: prefix requires the worksheet root element to declare xmlns:r --
    // easy to miss, since no earlier feature needed a relationship-id attribute here.
    Assert.IsTrue(Pos('xmlns:r=', SheetXml) > 0,
      'The worksheet root element must declare the xmlns:r namespace for r:id to be valid');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.SaveToFile_WithNote_LegacyDrawingFollowsMergeCells;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'Note';
  Sheet.MergeCells('A1:B1');

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const SheetXml = Package.GetPartContent('xl/worksheets/sheet1.xml');
    const MergeCellsPos = Pos('<mergeCells', SheetXml);
    const LegacyDrawingPos = Pos('<legacyDrawing', SheetXml);

    Assert.IsTrue(MergeCellsPos > 0, 'mergeCells should be present');
    Assert.IsTrue(LegacyDrawingPos > 0, 'legacyDrawing should be present');
    // CT_Worksheet's schema order places legacyDrawing well after mergeCells -- getting
    // this backwards is the same class of ordering bug as sheetViews-before-cols.
    Assert.IsTrue(MergeCellsPos < LegacyDrawingPos, 'mergeCells must precede legacyDrawing');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.RoundTrip_Note_PreservesText;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'Please double-check this figure';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual('Please double-check this figure', Workbook2.Sheets[0].Note['A1']);
end;

procedure TExcelNotesTests.RoundTrip_MultipleNotesOnDifferentCells_PreservesEachIndependently;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'X';
  Sheet.Cell['B5'].AsString := 'Y';
  Sheet.Cell['C10'].AsString := 'Z';
  Sheet.Note['A1'] := 'First note';
  Sheet.Note['B5'] := 'Second note';
  Sheet.Note['C10'] := 'Third note';

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  const Sheet2 = Workbook2.Sheets[0];

  Assert.AreEqual('First note', Sheet2.Note['A1']);
  Assert.AreEqual('Second note', Sheet2.Note['B5']);
  Assert.AreEqual('Third note', Sheet2.Note['C10']);
end;

procedure TExcelNotesTests.RoundTrip_SettingNoteToEmptyString_RemovesNote;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := 'Temporary note';
  Sheet.Note['A1'] := '';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    Assert.IsFalse(Package.PartExists('xl/comments1.xml'),
      'Clearing the only note on a sheet must leave it with zero notes, so no comments part should be written');
  finally
    Package.Free;
  end;
end;

procedure TExcelNotesTests.RoundTrip_NoteWithSpecialCharacters_ArePreserved;
begin
  const Special = 'R&D <tag> "q" ''a'' 5>3';
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['A1'].AsString := 'Data';
  Sheet.Note['A1'] := Special;

  FWorkbook.SaveToFile(FTempFile);

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);

  Assert.AreEqual(Special, Workbook2.Sheets[0].Note['A1'], 'Note text special characters should round-trip');
end;

procedure TExcelNotesTests.SaveToFile_OnlySecondSheetHasNotes_NumbersCommentsFileCorrectly;
begin
  const Sheet1 = FWorkbook.AddSheet('Sheet1');
  Sheet1.Cell['A1'].AsString := 'No notes here';

  const Sheet2 = FWorkbook.AddSheet('Sheet2');
  Sheet2.Cell['A1'].AsString := 'Has a note';
  Sheet2.Note['A1'] := 'Note on sheet 2';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    // Comments files are numbered sequentially among sheets that actually have notes,
    // not by sheet index -- sheet1 has none, so sheet2's rels/comments must still be
    // numbered "1", not "2".
    Assert.IsFalse(Package.PartExists('xl/worksheets/_rels/sheet1.xml.rels'),
      'sheet1 has no notes and must not get a .rels file');
    Assert.IsTrue(Package.PartExists('xl/worksheets/_rels/sheet2.xml.rels'),
      'sheet2 has a note and must get a .rels file');

    const Sheet2RelsXml = Package.GetPartContent('xl/worksheets/_rels/sheet2.xml.rels');
    Assert.IsTrue(Pos('comments1.xml', Sheet2RelsXml) > 0,
      'The first (and only) sheet with notes should reference comments1.xml, not comments2.xml');
  finally
    Package.Free;
  end;

  const Workbook2 = TExcelWorkbookFactory.Create;
  Workbook2.LoadFromFile(FTempFile);
  Assert.AreEqual('', Workbook2.Sheets[0].Note['A1'], 'Sheet1 should still have no note after round-trip');
  Assert.AreEqual('Note on sheet 2', Workbook2.Sheets[1].Note['A1'], 'Sheet2''s note should round-trip correctly');
end;

procedure TExcelNotesTests.SaveToFile_WithNote_VmlAnchorsNoteToItsCell;
begin
  const Sheet = FWorkbook.AddSheet('Sheet1');
  Sheet.Cell['B3'].AsString := 'Data';
  Sheet.Note['B3'] := 'Note text';

  FWorkbook.SaveToFile(FTempFile);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FTempFile);
    const VmlXml = Package.GetPartContent('xl/drawings/vmlDrawing1.vml');
    // B3 is column 2 (0-based 1), row 3 (0-based 2). The anchor starts one column to the
    // right of the cell and spans two columns / four rows, so each note box sits next to
    // its own cell instead of every box landing on the same fixed position.
    Assert.IsTrue(Pos('<x:Anchor>2, 15, 2, 2, 4, 15, 6, 4</x:Anchor>', VmlXml) > 0,
      'Each note must carry an x:Anchor positioned relative to its own cell');
  finally
    Package.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TExcelNotesTests);

end.
