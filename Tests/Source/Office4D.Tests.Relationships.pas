unit Office4D.Tests.Relationships;

interface

uses
  System.SysUtils,
  DUnitX.TestFramework,
  Office4D.Relationships,
  Office4D.Types;

type
  [TestFixture]
  TRelationshipsTests = class
  private
    FRels: TRelationships;

    const SampleRelsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>';
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure LoadFromXml_ValidXml_ParsesRelationships;

    [Test]
    procedure Count_AfterLoad_ReturnsCorrectCount;

    [Test]
    procedure GetById_ExistingId_ReturnsRelationship;

    [Test]
    procedure GetByType_OfficeDocument_ReturnsRelationship;

    [Test]
    procedure GetByType_CoreProperties_ReturnsRelationship;

    [Test]
    procedure GetTargetByType_OfficeDocument_ReturnsPath;

    [Test]
    procedure GetTargetByType_NonExistent_ReturnsEmpty;

    [Test]
    procedure HasType_ExistingType_ReturnsTrue;

    [Test]
    procedure HasType_NonExistentType_ReturnsFalse;

    [Test]
    [TestCase('TypeTargetId', '<Relationship Type="t" Target="a.xml" Id="rId7"/>')]
    [TestCase('SingleQuotes', '<Relationship Id=''rId7'' Type=''t'' Target=''a.xml''/>')]
    [TestCase('SpacesAroundEquals', '<Relationship Target = "a.xml" Id = "rId7" Type = "t"/>')]
    procedure LoadFromXml_RelationshipInAnyShape_ReadsIt(const RelationshipXml: string);

    [Test]
    procedure LoadFromXml_PrefixedNamespace_ReadsRelationships;

    [Test]
    procedure LoadFromXml_EscapedTarget_IsDecoded;

    [Test]
    procedure TryGetById_UnknownId_ReturnsFalse;

    [Test]
    [TestCase('Relative', 'xl/workbook.xml,worksheets/sheet1.xml,xl/worksheets/sheet1.xml')]
    [TestCase('Absolute', 'xl/workbook.xml,/xl/worksheets/sheet1.xml,xl/worksheets/sheet1.xml')]
    [TestCase('ParentFolder', 'xl/worksheets/sheet1.xml,../comments1.xml,xl/comments1.xml')]
    [TestCase('CurrentFolder', 'xl/workbook.xml,./worksheets/sheet1.xml,xl/worksheets/sheet1.xml')]
    procedure ResolveTarget_Target_ReturnsPartName(const SourcePart, Target, Expected: string);

    [Test]
    [TestCase('Nested', 'xl/worksheets/sheet1.xml,xl/worksheets/_rels/sheet1.xml.rels')]
    [TestCase('Root', 'workbook.xml,_rels/workbook.xml.rels')]
    procedure RelationshipsPartOf_PartName_ReturnsRelsPart(const PartName, Expected: string);
  end;

implementation

{ TRelationshipsTests }

procedure TRelationshipsTests.Setup;
begin
  FRels := TRelationships.Create;
end;

procedure TRelationshipsTests.TearDown;
begin
  FRels.Free;
end;

procedure TRelationshipsTests.LoadFromXml_ValidXml_ParsesRelationships;
begin
  FRels.LoadFromXml(SampleRelsXml);

  Assert.IsTrue(FRels.Count > 0);
end;

procedure TRelationshipsTests.Count_AfterLoad_ReturnsCorrectCount;
begin
  FRels.LoadFromXml(SampleRelsXml);

  Assert.AreEqual(3, FRels.Count);
end;

procedure TRelationshipsTests.GetById_ExistingId_ReturnsRelationship;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Rel := FRels.GetById('rId1');

  Assert.AreEqual('rId1', Rel.Id);
  Assert.AreEqual('word/document.xml', Rel.Target);
end;

procedure TRelationshipsTests.GetByType_OfficeDocument_ReturnsRelationship;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Rel := FRels.GetByType(RelTypeOfficeDocument);

  Assert.AreEqual('rId1', Rel.Id);
  Assert.AreEqual('word/document.xml', Rel.Target);
end;

procedure TRelationshipsTests.GetByType_CoreProperties_ReturnsRelationship;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Rel := FRels.GetByType(RelTypeCoreProperties);

  Assert.AreEqual('rId2', Rel.Id);
  Assert.AreEqual('docProps/core.xml', Rel.Target);
end;

procedure TRelationshipsTests.GetTargetByType_OfficeDocument_ReturnsPath;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Target := FRels.GetTargetByType(RelTypeOfficeDocument);

  Assert.AreEqual('word/document.xml', Target);
end;

procedure TRelationshipsTests.GetTargetByType_NonExistent_ReturnsEmpty;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Target := FRels.GetTargetByType('http://nonexistent/type');

  Assert.IsEmpty(Target);
end;

procedure TRelationshipsTests.HasType_ExistingType_ReturnsTrue;
begin
  FRels.LoadFromXml(SampleRelsXml);

  Assert.IsTrue(FRels.HasType(RelTypeOfficeDocument));
  Assert.IsTrue(FRels.HasType(RelTypeCoreProperties));
end;

procedure TRelationshipsTests.HasType_NonExistentType_ReturnsFalse;
begin
  FRels.LoadFromXml(SampleRelsXml);

  Assert.IsFalse(FRels.HasType('http://nonexistent/type'));
end;

procedure TRelationshipsTests.LoadFromXml_RelationshipInAnyShape_ReadsIt(const RelationshipXml: string);
begin
  FRels.LoadFromXml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    RelationshipXml + '</Relationships>');

  var Rel := FRels.GetById('rId7');

  Assert.AreEqual('t', Rel.RelType);
  Assert.AreEqual('a.xml', Rel.Target);
end;

procedure TRelationshipsTests.LoadFromXml_PrefixedNamespace_ReadsRelationships;
begin
  FRels.LoadFromXml('<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<pr:Relationship Target="a.xml" Type="t" Id="rId7"/></pr:Relationships>');

  Assert.AreEqual('a.xml', FRels.GetById('rId7').Target);
end;

procedure TRelationshipsTests.LoadFromXml_EscapedTarget_IsDecoded;
begin
  FRels.LoadFromXml('<Relationships><Relationship Id="rId1" Type="t" ' +
                    'Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/></Relationships>');

  Assert.AreEqual('https://example.com/?a=1&b=2', FRels.GetById('rId1').Target);
end;

procedure TRelationshipsTests.TryGetById_UnknownId_ReturnsFalse;
begin
  FRels.LoadFromXml(SampleRelsXml);

  var Rel: TRelationship;
  const Found = FRels.TryGetById('rId99', Rel);

  Assert.IsFalse(Found);
end;

procedure TRelationshipsTests.ResolveTarget_Target_ReturnsPartName(const SourcePart, Target, Expected: string);
begin
  const Actual = TRelationships.ResolveTarget(SourcePart, Target);

  Assert.AreEqual(Expected, Actual);
end;

procedure TRelationshipsTests.RelationshipsPartOf_PartName_ReturnsRelsPart(const PartName, Expected: string);
begin
  const Actual = TRelationships.RelationshipsPartOf(PartName);

  Assert.AreEqual(Expected, Actual);
end;

initialization
  TDUnitX.RegisterTestFixture(TRelationshipsTests);

end.
