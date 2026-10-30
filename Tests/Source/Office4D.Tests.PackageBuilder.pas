unit Office4D.Tests.PackageBuilder;

interface

type
  /// <summary>
  /// Writes the smallest package a reader accepts around a piece of body XML,
  /// so a test can state the construct it is about and nothing else. What the
  /// test cares about is the XML it passes in; everything else here is the
  /// wrapping an OPC package needs to be openable at all.
  /// </summary>
  TTestPackage = record
  public
    /// Writes a .docx whose w:body holds BodyXml.
    class procedure WriteDocx(const FileName, BodyXml: string); static;

    /// Writes a single slide .pptx whose p:spTree holds SlideXml.
    class procedure WritePptx(const FileName, SlideXml: string); static;

    /// Writes a single sheet .xlsx whose sheetData holds SheetDataXml.
    class procedure WriteXlsx(const FileName, SheetDataXml: string); static;
  end;

implementation

uses
  System.SysUtils,
  System.Zip;

const
  RelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  ContentTypesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
  RelTypeOfficeDocument = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RelTypeSlide = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  RelTypeWorksheet = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

  XmlDeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  DocxContentTypesXml =
    XmlDeclaration +
    '<Types xmlns="' + ContentTypesNs + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';

  DocxRootRelsXml =
    XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeOfficeDocument + '" Target="word/document.xml"/>' +
    '</Relationships>';

  DocumentPrefix =
    XmlDeclaration +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
    'xmlns:v="urn:schemas-microsoft-com:vml"><w:body>';

  DocumentSuffix = '</w:body></w:document>';

  PptxContentTypesXml =
    XmlDeclaration +
    '<Types xmlns="' + ContentTypesNs + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/ppt/presentation.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
    '<Override PartName="/ppt/slides/slide1.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
    '</Types>';

  PptxRootRelsXml =
    XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeOfficeDocument + '" Target="ppt/presentation.xml"/>' +
    '</Relationships>';

  PresentationXml =
    XmlDeclaration +
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>' +
    '</p:presentation>';

  PresentationRelsXml =
    XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeSlide + '" Target="slides/slide1.xml"/>' +
    '</Relationships>';

  SlidePrefix =
    XmlDeclaration +
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<p:cSld><p:spTree>';

  SlideSuffix = '</p:spTree></p:cSld></p:sld>';

  XlsxContentTypesXml =
    XmlDeclaration +
    '<Types xmlns="' + ContentTypesNs + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>';

  XlsxRootRelsXml =
    XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeOfficeDocument + '" Target="xl/workbook.xml"/>' +
    '</Relationships>';

  WorkbookXml =
    XmlDeclaration +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';

  WorkbookRelsXml =
    XmlDeclaration +
    '<Relationships xmlns="' + RelationshipsNs + '">' +
    '<Relationship Id="rId1" Type="' + RelTypeWorksheet + '" Target="worksheets/sheet1.xml"/>' +
    '</Relationships>';

  SheetPrefix =
    XmlDeclaration +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>';

  SheetSuffix = '</sheetData></worksheet>';

{ TTestPackage }

class procedure TTestPackage.WriteDocx(const FileName, BodyXml: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    Zip.Add(TEncoding.UTF8.GetBytes(DocxContentTypesXml), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(DocxRootRelsXml), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(DocumentPrefix + BodyXml + DocumentSuffix), 'word/document.xml');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

class procedure TTestPackage.WritePptx(const FileName, SlideXml: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    Zip.Add(TEncoding.UTF8.GetBytes(PptxContentTypesXml), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(PptxRootRelsXml), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(PresentationXml), 'ppt/presentation.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(PresentationRelsXml), 'ppt/_rels/presentation.xml.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(SlidePrefix + SlideXml + SlideSuffix), 'ppt/slides/slide1.xml');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

class procedure TTestPackage.WriteXlsx(const FileName, SheetDataXml: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    Zip.Add(TEncoding.UTF8.GetBytes(XlsxContentTypesXml), '[Content_Types].xml');
    Zip.Add(TEncoding.UTF8.GetBytes(XlsxRootRelsXml), '_rels/.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(WorkbookXml), 'xl/workbook.xml');
    Zip.Add(TEncoding.UTF8.GetBytes(WorkbookRelsXml), 'xl/_rels/workbook.xml.rels');
    Zip.Add(TEncoding.UTF8.GetBytes(SheetPrefix + SheetDataXml + SheetSuffix), 'xl/worksheets/sheet1.xml');
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

end.
