unit Office4D.Tests.Xml;

interface

uses
  System.SysUtils,
  DUnitX.TestFramework,
  Office4D.Xml;

type
  [TestFixture]
  TXmlFindElementsTests = class
  public
    [Test]
    procedure FindElements_SiblingElements_ReturnsEachInDocumentOrder;

    [Test]
    procedure FindElements_NestedSameName_ReturnsOuterElementOnly;

    [Test]
    procedure FindElements_NestedSameName_KeepsContentAfterNestedElement;

    [Test]
    procedure FindElements_LongerElementName_IsNotMatched;

    [Test]
    procedure FindElements_SelfClosingElement_ReturnsEmptyInner;

    [Test]
    procedure FindElements_AttributeContainingAngleBracket_FindsElement;

    [Test]
    procedure FindElements_OpenTag_ContainsAttributes;

    [Test]
    procedure FindElements_UnknownElement_ReturnsNothing;

    [Test]
    procedure RemoveElements_NestedElements_CutsThemOut;

    [Test]
    procedure RemoveElements_Nothing_ReturnsInput;

    [Test]
    procedure ReplaceElements_Replacement_SplicesItIn;

    [Test]
    procedure FindElements_WhitespaceInCloseTag_FindsElement;
  end;

  [TestFixture]
  TXmlAttributeTests = class
  public
    [Test]
    [TestCase('First', '<sheet name="Sales" sheetId="1"/>|Sales', '|')]
    [TestCase('Last', '<sheet sheetId="1" state="hidden" name="Sales"/>|Sales', '|')]
    [TestCase('SingleQuotes', '<sheet name=''Sales''/>|Sales', '|')]
    [TestCase('SpacesAroundEquals', '<sheet name = "Sales"/>|Sales', '|')]
    [TestCase('OpenTag', '<sheet name="Sales">|Sales', '|')]
    [TestCase('Escaped', '<sheet name="A &amp; B"/>|A & B', '|')]
    [TestCase('OtherQuoteInside', '<sheet name=''Say "hi"''/>|Say "hi"', '|')]
    [TestCase('NameInsideOtherValue', '<sheet title=''name="Wrong"'' name="Right"/>|Right', '|')]
    procedure AttributeValue_AnyShape_ReturnsValue(const StartTag, Expected: string);

    [Test]
    procedure AttributeValue_LongerOrPrefixedName_IsNotMatched;

    [Test]
    procedure AttributeValue_PrefixedName_MatchesWithPrefix;

    [Test]
    procedure TryGetAttribute_EmptyValue_ReturnsTrue;

    [Test]
    procedure TryGetAttribute_MissingAttribute_ReturnsFalse;

    [Test]
    [TestCase('Prefixed', '<x:workbook xmlns:x="urn:main"/>|x:', '|')]
    [TestCase('SingleQuotes', '<x:workbook xmlns:x=''urn:main''/>|x:', '|')]
    [TestCase('OtherNamespaceFirst', '<x:workbook xmlns="urn:other" xmlns:x="urn:main"/>|x:', '|')]
    procedure PrefixOf_Declaration_ReturnsPrefix(const Xml, Expected: string);

    [Test]
    [TestCase('Default', '<workbook xmlns="urn:main"/>')]
    [TestCase('Undeclared', '<workbook xmlns="urn:other"/>')]
    procedure PrefixOf_DefaultOrUndeclared_ReturnsEmpty(const Xml: string);

    [Test]
    [TestCase('One', '1,False,True')]
    [TestCase('True', 'true,False,True')]
    [TestCase('Zero', '0,True,False')]
    [TestCase('False', 'false,True,False')]
    [TestCase('Other', 'yes,True,True')]
    procedure BooleanOf_Value_ReturnsBoolean(const Value: string; const Default, Expected: Boolean);

    [Test]
    procedure BooleanOf_Absent_ReturnsDefault;
  end;

implementation

{ TXmlFindElementsTests }

procedure TXmlFindElementsTests.FindElements_SiblingElements_ReturnsEachInDocumentOrder;
begin
  var Elements := TXml.FindElements('<w:p><w:r>A</w:r><w:r>B</w:r></w:p>', 'w:r');

  Assert.AreEqual(2, Integer(Length(Elements)));
  Assert.AreEqual('A', Elements[0].Inner);
  Assert.AreEqual('B', Elements[1].Inner);
end;

procedure TXmlFindElementsTests.FindElements_NestedSameName_ReturnsOuterElementOnly;
begin
  var Elements := TXml.FindElements('<w:body><w:p>outer<w:p>inner</w:p></w:p></w:body>', 'w:p');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('outer<w:p>inner</w:p>', Elements[0].Inner);
end;

procedure TXmlFindElementsTests.FindElements_NestedSameName_KeepsContentAfterNestedElement;
begin
  var Elements := TXml.FindElements('<w:p><w:p>inner</w:p>tail</w:p>', 'w:p');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('<w:p>inner</w:p>tail', Elements[0].Inner);
end;

procedure TXmlFindElementsTests.FindElements_LongerElementName_IsNotMatched;
begin
  var Elements := TXml.FindElements('<w:rPr><w:b/></w:rPr><w:r>text</w:r>', 'w:r');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('text', Elements[0].Inner);
end;

procedure TXmlFindElementsTests.FindElements_SelfClosingElement_ReturnsEmptyInner;
begin
  var Elements := TXml.FindElements('<w:p><w:r/></w:p>', 'w:r');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('', Elements[0].Inner);
end;

procedure TXmlFindElementsTests.FindElements_AttributeContainingAngleBracket_FindsElement;
begin
  var Elements := TXml.FindElements('<v:shape style="a>b"><w:r>text</w:r></v:shape>', 'v:shape');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('<w:r>text</w:r>', Elements[0].Inner);
end;

procedure TXmlFindElementsTests.FindElements_OpenTag_ContainsAttributes;
begin
  var Elements := TXml.FindElements('<w:hyperlink r:id="rId7">text</w:hyperlink>', 'w:hyperlink');

  Assert.AreEqual(1, Integer(Length(Elements)));
  Assert.AreEqual('<w:hyperlink r:id="rId7">', Elements[0].OpenTag);
end;

procedure TXmlFindElementsTests.FindElements_UnknownElement_ReturnsNothing;
begin
  var Elements := TXml.FindElements('<w:p><w:r>A</w:r></w:p>', 'w:tbl');

  Assert.AreEqual(0, Integer(Length(Elements)));
end;

procedure TXmlFindElementsTests.RemoveElements_NestedElements_CutsThemOut;
begin
  const Xml = '<w:r>keep<w:p>drop</w:p>tail</w:r>';
  var Elements := TXml.FindElements(Xml, 'w:p');

  Assert.AreEqual('<w:r>keeptail</w:r>', TXml.RemoveElements(Xml, Elements));
end;

procedure TXmlFindElementsTests.RemoveElements_Nothing_ReturnsInput;
begin
  const Xml = '<w:r>text</w:r>';

  Assert.AreEqual(Xml, TXml.RemoveElements(Xml, TXml.FindElements(Xml, 'w:p')));
end;

procedure TXmlFindElementsTests.ReplaceElements_Replacement_SplicesItIn;
begin
  const Xml = '<w:p>a<w:pict>shape</w:pict>b</w:p>';
  var Elements := TXml.FindElements(Xml, 'w:pict');

  var Replaced := TXml.ReplaceElements(Xml, Elements,
    function(Element: TXmlElement): string
    begin
      Result := '[' + Element.Inner + ']';
    end);

  Assert.AreEqual('<w:p>a[shape]b</w:p>', Replaced);
end;

procedure TXmlFindElementsTests.FindElements_WhitespaceInCloseTag_FindsElement;
begin
  const Xml = '<row><c r="A1"><v>1</v></c ><c r="B1"><v>2</v></c></row>';

  var Elements := TXml.FindElements(Xml, 'c');

  Assert.AreEqual(2, Integer(Length(Elements)));
  Assert.AreEqual('<v>1</v>', Elements[0].Inner);
  Assert.AreEqual('<v>2</v>', Elements[1].Inner);
end;

{ TXmlAttributeTests }

procedure TXmlAttributeTests.AttributeValue_AnyShape_ReturnsValue(const StartTag, Expected: string);
begin
  const Actual = TXml.AttributeValue(StartTag, 'name');

  Assert.AreEqual(Expected, Actual);
end;

procedure TXmlAttributeTests.AttributeValue_LongerOrPrefixedName_IsNotMatched;
begin
  const Actual = TXml.AttributeValue('<sheet sheetname="Wrong" x:name="Wrong"/>', 'name');

  Assert.AreEqual('', Actual);
end;

procedure TXmlAttributeTests.AttributeValue_PrefixedName_MatchesWithPrefix;
begin
  const StartTag = '<p:sldId id="256" r:id="rId2"/>';

  Assert.AreEqual('256', TXml.AttributeValue(StartTag, 'id'));
  Assert.AreEqual('rId2', TXml.AttributeValue(StartTag, 'r:id'));
end;

procedure TXmlAttributeTests.TryGetAttribute_EmptyValue_ReturnsTrue;
begin
  var Value: string;
  const Found = TXml.TryGetAttribute('<c r="A1" t=""/>', 't', Value);

  Assert.IsTrue(Found);
  Assert.AreEqual('', Value);
end;

procedure TXmlAttributeTests.TryGetAttribute_MissingAttribute_ReturnsFalse;
begin
  var Value: string;
  const Found = TXml.TryGetAttribute('<c r="A1"/>', 't', Value);

  Assert.IsFalse(Found);
end;

procedure TXmlAttributeTests.PrefixOf_Declaration_ReturnsPrefix(const Xml, Expected: string);
begin
  const Actual = TXml.PrefixOf(Xml, 'urn:main');

  Assert.AreEqual(Expected, Actual);
end;

procedure TXmlAttributeTests.PrefixOf_DefaultOrUndeclared_ReturnsEmpty(const Xml: string);
begin
  const Actual = TXml.PrefixOf(Xml, 'urn:main');

  Assert.AreEqual('', Actual);
end;

procedure TXmlAttributeTests.BooleanOf_Absent_ReturnsDefault;
begin
  Assert.IsTrue(TXml.BooleanOf('', True));
  Assert.IsFalse(TXml.BooleanOf('', False));
end;

procedure TXmlAttributeTests.BooleanOf_Value_ReturnsBoolean(const Value: string; const Default, Expected: Boolean);
begin
  const Actual = TXml.BooleanOf(Value, Default);

  Assert.AreEqual(Expected, Actual);
end;

initialization
  TDUnitX.RegisterTestFixture(TXmlFindElementsTests);

end.
