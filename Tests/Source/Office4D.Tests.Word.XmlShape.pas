unit Office4D.Tests.Word.XmlShape;

interface

uses
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Word;

type
  /// <summary>
  /// A document reads the same whatever tool wrote its relationships: the
  /// order of the attributes and an absolute or relative target carry no
  /// meaning, so they must not change the result.
  /// </summary>
  [TestFixture]
  TWordXmlShapeTests = class
  private
    FDoc: IWordDocument;
    FSourceFile: string;

    procedure LoadDocument(const BodyXml, RelsXml, HeaderPartXml: string);
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromFile_HyperlinkRelationshipWithIdLast_ReadsHyperlink;

    [Test]
    procedure LoadFromFile_HeaderWithAbsoluteTarget_ReadsHeader;
  end;

implementation

uses
  Office4D.Tests.PackageBuilder;

const
  RelationshipsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  PackageRelationshipsNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
  WordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  RootRelsXml =
    '<Relationships xmlns="' + PackageRelationshipsNs + '">' +
    '<Relationship Type="' + RelationshipsNs + '/officeDocument" Target="word/document.xml" Id="rId1"/>' +
    '</Relationships>';

  DocumentXml =
    '<w:document xmlns:w="' + WordNs + '" xmlns:r="' + RelationshipsNs + '"><w:body>%s</w:body></w:document>';

  DocumentRelsXml = '<Relationships xmlns="' + PackageRelationshipsNs + '">%s</Relationships>';

  HeaderXml = '<w:hdr xmlns:w="' + WordNs + '"><w:p><w:r><w:t>Page header</w:t></w:r></w:p></w:hdr>';

  PlainParagraphXml = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>';

{ TWordXmlShapeTests }

procedure TWordXmlShapeTests.Setup;
begin
  FDoc := TWordDocumentFactory.CreateDocument;
  FSourceFile := TPath.Combine(TPath.GetTempPath, 'wordxmlshape_' + TGUID.NewGuid.ToString + '.docx');
end;

procedure TWordXmlShapeTests.TearDown;
begin
  FDoc := nil;
  if TFile.Exists(FSourceFile) then
    TFile.Delete(FSourceFile);
end;

procedure TWordXmlShapeTests.LoadDocument(const BodyXml, RelsXml, HeaderPartXml: string);
begin
  TTestPackage.WriteParts(FSourceFile,
    [TTestPart.Create('_rels/.rels', RootRelsXml),
     TTestPart.Create('word/document.xml', Format(DocumentXml, [BodyXml])),
     TTestPart.Create('word/_rels/document.xml.rels', RelsXml),
     TTestPart.Create('word/header1.xml', HeaderPartXml)]);
  FDoc.LoadFromFile(FSourceFile);
end;

procedure TWordXmlShapeTests.LoadFromFile_HyperlinkRelationshipWithIdLast_ReadsHyperlink;
begin
  const Relationship =
    '<Relationship Type="' + RelationshipsNs + '/hyperlink" ' +
    'Target="https://example.com/?a=1&amp;b=2" TargetMode="External" Id="rId5"/>';

  LoadDocument('<w:p><w:hyperlink r:id="rId5"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>',
               Format(DocumentRelsXml, [Relationship]),
               HeaderXml);

  const Actual = FDoc.Paragraphs[0].Runs[0].Hyperlink;

  Assert.AreEqual('https://example.com/?a=1&b=2', Actual);
end;

procedure TWordXmlShapeTests.LoadFromFile_HeaderWithAbsoluteTarget_ReadsHeader;
begin
  const Relationship =
    '<Relationship Target="/word/header1.xml" Type="' + RelationshipsNs + '/header" Id="rId3"/>';

  LoadDocument(PlainParagraphXml, Format(DocumentRelsXml, [Relationship]), HeaderXml);

  const Actual = FDoc.Header.Text;

  Assert.AreEqual('Page header', Actual);
end;

end.
