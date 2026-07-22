# Office4D

Pure Delphi library for reading and writing Microsoft Office Open XML documents (.docx, .xlsx, .pptx).

## Features

- **No external dependencies** - Uses only Delphi built-in units (System.Zip, Xml.XMLDoc)
- **Full read/write support** - Create new documents or modify existing ones
- **Interface-based API** - Clean, testable design
- **Rich formatting** - Bold, italic, underline, fonts, colors, alignment, hyperlinks, lists, tables
- **Cell styling** - Background colors, number formats, formulas, column widths, merged cells

## Supported Formats

| Format | Extension | Read | Write |
|--------|-----------|------|-------|
| Word | .docx | Yes | Yes |
| Excel | .xlsx | Yes | Yes |
| PowerPoint | .pptx | Text extraction | Yes (basic) |

## Installation

1. Clone the repository
2. Add `Source` folder and subfolders to your Delphi library path
3. Add units to your uses clause

## Usage

### Reading Word Documents

```pascal
uses
  Office4D.Word;

var
  Doc: IWordDocument;
begin
  Doc := TWordDocumentFactory.CreateDocument;
  Doc.LoadFromFile('report.docx');

  // Get plain text
  Memo1.Text := Doc.Text;

  // Get metadata
  ShowMessage('Author: ' + Doc.Metadata.Author);
  ShowMessage('Title: ' + Doc.Metadata.Title);

  // Access paragraphs
  for var I := 0 to Doc.ParagraphCount - 1 do
    WriteLn(Doc.Paragraphs[I].Text);
end;
```

### Creating Word Documents

```pascal
uses
  Office4D.Word;

var
  Doc: IWordDocument;
  Para: IWordParagraph;
  Run: IWordRun;
begin
  Doc := TWordDocumentFactory.CreateDocument;

  // Page settings
  Doc.PageOrientation := TPageOrientation.Portrait;
  Doc.PageMargins := TPageMargins.Create(1440, 1440, 1440, 1440); // 1 inch

  // Header and footer
  Doc.Header.Text := 'Document Header';
  Doc.Footer.Text := 'Page 1';

  // Add formatted paragraph
  Para := Doc.AddParagraph;
  Run := Para.AddRun('Bold text');
  Run.Bold := True;
  Para.AddRun(' and ');
  Run := Para.AddRun('italic text');
  Run.Italic := True;

  // Font formatting
  Para := Doc.AddParagraph;
  Run := Para.AddRun('Custom Font');
  Run.FontName := 'Arial';
  Run.FontSize := 24;  // half-points (24 = 12pt)
  Run.FontColor := 'FF0000';  // Red

  // Paragraph alignment
  Para := Doc.AddParagraph;
  Para.Alignment := TParagraphAlignment.Center;
  Para.AddRun('Centered text');

  // Add hyperlink
  Para := Doc.AddParagraph;
  Run := Para.AddRun('Click here');
  Run.Hyperlink := 'https://example.com';
  Run.Underline := True;

  // Add bullet list
  Para := Doc.AddParagraph;
  Para.ListStyle := TListStyle.Bullet;
  Para.AddRun('First item');

  Para := Doc.AddParagraph;
  Para.ListStyle := TListStyle.Bullet;
  Para.AddRun('Second item');

  // Add table
  var Table := Doc.AddTable(3, 2);
  Table.Cells[0, 0].Text := 'Header 1';
  Table.Cells[0, 1].Text := 'Header 2';
  Table.Cells[1, 0].Text := 'Data 1';
  Table.Cells[1, 1].Text := 'Data 2';

  Doc.SaveToFile('output.docx');
end;
```

### Reading Excel Workbooks

```pascal
uses
  Office4D.Excel;

var
  Workbook: IExcelWorkbook;
  Sheet: IExcelSheet;
begin
  Workbook := TExcelWorkbookFactory.Create;
  Workbook.LoadFromFile('data.xlsx');

  // Access sheets
  for var I := 0 to Workbook.SheetCount - 1 do
  begin
    Sheet := Workbook.Sheets[I];
    WriteLn('Sheet: ' + Sheet.Name);
  end;

  // Get sheet by name
  Sheet := Workbook.SheetByName('Sheet1');

  // Read cell values
  WriteLn(Sheet.Cell['A1'].AsString);
  WriteLn(FloatToStr(Sheet.Cell['B1'].AsFloat));
  WriteLn(BoolToStr(Sheet.Cell['C1'].AsBoolean, True));
  WriteLn(DateTimeToStr(Sheet.Cell['D1'].AsDateTime));

  // Check for formula
  if Sheet.Cell['E1'].HasFormula then
    WriteLn('Formula: ' + Sheet.Cell['E1'].Formula);
end;
```

### Creating Excel Workbooks

```pascal
uses
  Office4D.Excel;

var
  Workbook: IExcelWorkbook;
  Sheet: IExcelSheet;
begin
  Workbook := TExcelWorkbookFactory.Create;

  // Add sheet
  Sheet := Workbook.AddSheet('Products');

  // Header row with styling
  Sheet.Cell['A1'].AsString := 'Product';
  Sheet.Cell['A1'].Bold := True;
  Sheet.Cell['A1'].BackgroundColor := $4472C4; // Blue

  Sheet.Cell['B1'].AsString := 'Price';
  Sheet.Cell['B1'].Bold := True;
  Sheet.Cell['B1'].BackgroundColor := $4472C4;

  // Data rows
  Sheet.Cell['A2'].AsString := 'Laptop';
  Sheet.Cell['B2'].AsFloat := 999.00;
  Sheet.Cell['B2'].NumberFormat := '"$"#,##0.00';

  Sheet.Cell['A3'].AsString := 'Mouse';
  Sheet.Cell['B3'].AsFloat := 29.95;
  Sheet.Cell['B3'].NumberFormat := '"$"#,##0.00';

  // Formula with calculated value
  Sheet.Cell['B4'].SetFormula('SUM(B2:B3)', 1028.95);
  Sheet.Cell['B4'].Bold := True;
  Sheet.Cell['B4'].NumberFormat := '"$"#,##0.00';

  // Boolean and DateTime
  Sheet.Cell['C2'].AsBoolean := True;
  Sheet.Cell['D2'].AsDateTime := Now;

  // Column widths
  Sheet.SetColumnWidth('A', 20);
  Sheet.SetColumnWidth('B', 15);

  // Merged cells
  Sheet.Cell['E1'].AsString := 'Merged Header';
  Sheet.MergeCells('E1:G1');

  // Font, border and cell styling
  Sheet.Cell['A4'].AsString := 'Computer';
  Sheet.Cell['A4'].Italic := True;
  Sheet.Cell['A4'].FontStyle := Sheet.Cell['A4'].FontStyle + [TExcelFontStyle.Strikeout];
  Sheet.Cell['A4'].BorderStyle[AllBorderSides] := TExcelBorderStyle.Thin;
  Sheet.Cell['B5'].AsFloat := -999.00;
  Sheet.Cell['B5'].FontColor := $FF0000;
  Sheet.Cell['B5'].NumberFormat := '"$"#,##0.00';
  Sheet.Cell['B5'].BorderColor[[TExcelBorderSide.Bottom, TExcelBorderSide.Top]] := $00FF00;
  Sheet.Cell['B5'].BorderStyle[[TExcelBorderSide.Top]] := TExcelBorderStyle.Medium;
  Sheet.Cell['B5'].BorderStyle[[TExcelBorderSide.Bottom]] := TExcelBorderStyle.Double;

  Workbook.SaveToFile('output.xlsx');
end;
```

### Editing Sheets, Columns and Rows

```pascal
uses
  Office4D.Excel;

var
  Workbook: IExcelWorkbook;
  Sheet: IExcelSheet;
begin
  Workbook := TExcelWorkbookFactory.Create;
  Workbook.LoadFromFile('report.xlsx');

  // Remove a whole sheet (by index or by name)
  Workbook.RemoveSheetByName('Scratch');

  Sheet := Workbook.SheetByName('Data');

  // Clear a column or row: removes cell contents and styling, but leaves the
  // structure (column widths, row heights, merged ranges) in place.
  Sheet.ClearColumn('D');
  Sheet.ClearRow(5);

  // Delete a column or row: everything past it shifts back by one. Column widths,
  // row heights, merged ranges and frozen panes move with it, and every formula
  // reference is rewritten (absolute refs and ranges included). A reference to a
  // deleted line becomes #REF!, exactly as Excel does.
  Sheet.DeleteColumn('C');
  Sheet.DeleteRow(2);

  Workbook.SaveToFile('report.xlsx');
end;
```

### Stream Support

Both Word and Excel support stream-based I/O:

```pascal
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    // Save to stream
    Doc.SaveToStream(Stream);

    // Load from stream
    Stream.Position := 0;
    Doc.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;
```

## Building

```batch
# Build and run tests
./build.bat Release Win64 Tests

# Build library only
./build.bat Release Win64
```

## Project Structure

```
Source/
  Core/        - ZIP package handling, relationships, types
  Common/      - Metadata, shared interfaces
  Word/        - Word document implementation
  Excel/       - Excel workbook implementation
  PowerPoint/  - PowerPoint presentation implementation
Tests/
  Source/      - DUnitX test suite
  Samples/     - Sample documents for testing
Examples/      - Demo application
```

## Word Features

- Paragraphs with multiple runs
- Text formatting: bold, italic, underline
- Font formatting: name, size, color
- Paragraph alignment: left, center, right, justify
- Hyperlinks
- Bullet and numbered lists
- Tables
- Headers and footers
- Page orientation and margins
- Metadata (author, title, dates)

## Excel Features

- Multiple worksheets
- Sheet visibility (visible, hidden, very hidden)
- Add and remove sheets
- Clear a column or row (contents only, structure preserved)
- Delete a column or row with reflow and formula-reference rewriting (#REF! on deleted lines)
- Cell data types: string, number, boolean, datetime
- Cell styling: background color, border style, border color
- Text formatting: bold, italic, underline, strikeout
- Number formats (currency, percentage, custom)
- Formulas with calculated values
- Column widths
- Merged cells
- Freeze panes
- Shared strings optimization
- Metadata (author, title, dates)

## PowerPoint Features

- Creating presentations with multiple slides
- Slide titles and body text with paragraphs
- Bullet lists with indent levels
- Run formatting: bold, italic, underline, font name, size, color
- Reading text back from presentations (title, paragraphs, formatting)
- Metadata (author, title, dates)

```pascal
uses Office4D.PowerPoint;

var Presentation := TPowerPointPresentationFactory.CreatePresentation;
var Slide := Presentation.AddSlide('Quarterly Report');
var Bullet := Slide.AddParagraph('Revenue up 12%');
Bullet.Bullet := True;
Presentation.SaveToFile('report.pptx');
```

## Requirements

- Delphi 10.3 Rio or later (for inline variable declarations)
- No additional components or libraries required

## License

MIT License - see [LICENSE](LICENSE) file

## Contributing

1. Fork the repository
2. Create a feature branch
3. Write tests for new features
4. Submit a pull request
