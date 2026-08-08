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

initialization
  TDUnitX.RegisterTestFixture(TXmlFindElementsTests);

end.
