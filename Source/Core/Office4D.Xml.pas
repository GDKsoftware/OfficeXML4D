unit Office4D.Xml;

interface

uses
  System.SysUtils;

type
  /// <summary>
  /// One element located by <see cref="TXml.FindElements"/>. Positions are
  /// 1-based indexes into the scanned XML: StartPos points at the '&lt;' of the
  /// opening tag, EndPos at the first character after the closing tag.
  /// </summary>
  TXmlElement = record
  public
    StartPos: Integer;
    EndPos: Integer;
    OpenTag: string;
    Inner: string;

    /// <summary>
    /// The value of the attribute Name on this element, see
    /// <see cref="TXml.AttributeValue"/>.
    /// </summary>
    function Attribute(const Name: string): string;
  end;

  /// <summary>
  /// Shared XML text helpers used by the Word, Excel and PowerPoint modules.
  /// Escape is applied when writing element/attribute text, Unescape when
  /// reading it back, so text round-trips through the five predefined entities.
  /// </summary>
  TXml = record
  private
    class function ScanTagEnd(const Xml: string; TagStart: Integer): Integer; static;
    class function IsTagNameEnd(const Value: Char): Boolean; static;
    class function CloseTagEnd(const Xml, ClosePrefix: string; const Cursor: Integer): Integer; static;
    class function SelectAlternateContentBranch(const AlternateContentXml: string): string; static;
  public
    class function Escape(const Value: string): string; static;
    class function Unescape(const Value: string): string; static;

    /// <summary>
    /// Returns the value of the attribute Name in StartTag with its entities
    /// decoded, or an empty string when the tag has no such attribute. The
    /// order of the attributes, the quote character and whitespace around '='
    /// make no difference. Name is matched as written, prefix included.
    /// </summary>
    class function AttributeValue(const StartTag, Name: string): string; static;

    /// <summary>
    /// Returns True and the decoded value when StartTag carries the attribute
    /// Name, so an attribute that is present but empty can be told apart from
    /// one that is missing.
    /// </summary>
    class function TryGetAttribute(const StartTag, Name: string; out Value: string): Boolean; static;

    /// <summary>
    /// Returns the prefix, colon included, that Xml binds to NamespaceUri, or
    /// an empty string when NamespaceUri is the default namespace or is not
    /// declared. Writers are free to choose a prefix, so a reader prepends this
    /// to the local names it looks for.
    /// </summary>
    class function PrefixOf(const Xml, NamespaceUri: string): string; static;

    /// <summary>
    /// Reads an xsd:boolean: '1' and 'true' are True, '0' and 'false' are
    /// False, and anything else, an absent attribute included, is Default.
    /// </summary>
    class function BooleanOf(const Value: string; const Default: Boolean): Boolean; static;

    /// <summary>
    /// Returns True and the first ElementName element of Xml, as
    /// <see cref="FindElements"/> would list it first.
    /// </summary>
    class function TryFindElement(const Xml, ElementName: string; out Element: TXmlElement): Boolean; static;

    /// <summary>
    /// Returns every ElementName element at the outermost nesting level of Xml,
    /// in document order. Elements of the same name nested inside a result are
    /// part of its Inner and are never returned separately, which lets callers
    /// deal with constructs such as a w:p inside a text box explicitly.
    /// Matching is exact and delimiter-aware, so scanning for 'w:r' never hits
    /// 'w:rPr'. Self-closing elements are returned with an empty Inner.
    /// </summary>
    class function FindElements(const Xml, ElementName: string): TArray<TXmlElement>; static;

    /// <summary>
    /// Returns Xml with each of the given elements, opening and closing tag
    /// included, replaced by what GetReplacement yields for it. A nil
    /// GetReplacement drops the elements. The elements must belong to Xml and
    /// be ordered by StartPos, which is what FindElements returns.
    /// </summary>
    class function ReplaceElements(const Xml: string; const Elements: TArray<TXmlElement>;
      const GetReplacement: TFunc<TXmlElement, string>): string; static;

    /// <summary>
    /// Returns Xml with the given elements cut out, opening and closing tag
    /// included.
    /// </summary>
    class function RemoveElements(const Xml: string; const Elements: TArray<TXmlElement>): string; static;

    /// <summary>
    /// Replaces every mc:AlternateContent element by the branch a consumer is
    /// meant to read: the first mc:Choice, or mc:Fallback when no choice is
    /// offered. Both branches describe the same content, so reading them both
    /// yields every shape and its text twice. Each pass unwraps one nesting
    /// level and always shortens the XML, so the loop terminates.
    /// </summary>
    class function ReduceAlternateContent(const Xml: string): string; static;
  end;

implementation

uses
  System.Generics.Collections,
  System.RegularExpressions;

type
  /// <summary>
  /// Walks the attributes of one start tag in document order, so a value that
  /// happens to contain name="..." text is never mistaken for an attribute.
  /// </summary>
  TAttributeScanner = record
  private
    FTag: string;
    FCursor: Integer;

    function IsAtEnd: Boolean;
    function IsWhitespace: Boolean;
    procedure SkipWhitespace;
    function ReadName: string;
    function TryReadQuotedValue(out RawValue: string): Boolean;
  public
    constructor Create(const StartTag: string);

    function TryReadNext(out Name, RawValue: string): Boolean;
  end;

const
  ElementAlternateContent = 'mc:AlternateContent';
  ElementChoice = 'mc:Choice';
  ElementFallback = 'mc:Fallback';

{ TAttributeScanner }

constructor TAttributeScanner.Create(const StartTag: string);
begin
  FTag := StartTag;
  FCursor := 2;

  while not IsAtEnd and not IsWhitespace and not CharInSet(FTag[FCursor], ['/', '>']) do
    Inc(FCursor);
end;

function TAttributeScanner.TryReadNext(out Name, RawValue: string): Boolean;
begin
  SkipWhitespace;
  Name := ReadName;
  if Name.IsEmpty then
    Exit(False);

  SkipWhitespace;
  const HasEqualsSign = (not IsAtEnd and (FTag[FCursor] = '='));
  if not HasEqualsSign then
    Exit(False);

  Inc(FCursor);
  SkipWhitespace;
  Result := TryReadQuotedValue(RawValue);
end;

function TAttributeScanner.IsAtEnd: Boolean;
begin
  Result := FCursor > Length(FTag);
end;

function TAttributeScanner.IsWhitespace: Boolean;
begin
  Result := CharInSet(FTag[FCursor], [' ', #9, #10, #13]);
end;

procedure TAttributeScanner.SkipWhitespace;
begin
  while not IsAtEnd and IsWhitespace do
    Inc(FCursor);
end;

function TAttributeScanner.ReadName: string;
begin
  const NameStart = FCursor;
  while not IsAtEnd and not IsWhitespace and not CharInSet(FTag[FCursor], ['=', '/', '>']) do
    Inc(FCursor);

  Result := Copy(FTag, NameStart, FCursor - NameStart);
end;

function TAttributeScanner.TryReadQuotedValue(out RawValue: string): Boolean;
begin
  RawValue := '';
  if IsAtEnd or not CharInSet(FTag[FCursor], ['"', '''']) then
    Exit(False);

  const Quote = FTag[FCursor];
  const ValueStart = FCursor + 1;
  const ClosingQuote = Pos(Quote, FTag, ValueStart);
  const IsUnterminated = (ClosingQuote = 0);
  if IsUnterminated then
    Exit(False);

  RawValue := Copy(FTag, ValueStart, ClosingQuote - ValueStart);
  FCursor := ClosingQuote + 1;
  Result := True;
end;

{ TXmlElement }

function TXmlElement.Attribute(const Name: string): string;
begin
  Result := TXml.AttributeValue(OpenTag, Name);
end;

{ TXml }

class function TXml.Escape(const Value: string): string;
begin
  Result := Value;
  // & must be replaced first so the entities added below are not re-escaped.
  Result := StringReplace(Result, '&', '&amp;', [rfReplaceAll]);
  Result := StringReplace(Result, '<', '&lt;', [rfReplaceAll]);
  Result := StringReplace(Result, '>', '&gt;', [rfReplaceAll]);
  Result := StringReplace(Result, '"', '&quot;', [rfReplaceAll]);
  Result := StringReplace(Result, '''', '&apos;', [rfReplaceAll]);
end;

class function TXml.Unescape(const Value: string): string;
begin
  Result := Value;
  Result := StringReplace(Result, '&lt;', '<', [rfReplaceAll]);
  Result := StringReplace(Result, '&gt;', '>', [rfReplaceAll]);
  Result := StringReplace(Result, '&quot;', '"', [rfReplaceAll]);
  Result := StringReplace(Result, '&apos;', '''', [rfReplaceAll]);
  // &amp; must be decoded last so escaped entities like &amp;lt; survive intact.
  Result := StringReplace(Result, '&amp;', '&', [rfReplaceAll]);
end;

class function TXml.AttributeValue(const StartTag, Name: string): string;
begin
  if not TryGetAttribute(StartTag, Name, Result) then
    Result := '';
end;

class function TXml.TryGetAttribute(const StartTag, Name: string; out Value: string): Boolean;
begin
  Value := '';

  var Scanner := TAttributeScanner.Create(StartTag);
  var AttributeName: string;
  var RawValue: string;
  while Scanner.TryReadNext(AttributeName, RawValue) do
  begin
    const IsRequestedAttribute = (AttributeName = Name);
    if IsRequestedAttribute then
    begin
      Value := Unescape(RawValue);
      Exit(True);
    end;
  end;

  Result := False;
end;

class function TXml.PrefixOf(const Xml, NamespaceUri: string): string;
begin
  const EscapedUri = TRegEx.Escape(NamespaceUri);
  const Pattern = Format('\sxmlns(?::([^\s=/>]+))?\s*=\s*(["''])%s\2', [EscapedUri]);
  const Declaration = TRegEx.Match(Xml, Pattern);
  if not Declaration.Success then
    Exit('');

  const Prefix = Declaration.Groups[1].Value;
  if Prefix.IsEmpty then
    Exit('');

  Result := Prefix + ':';
end;

class function TXml.BooleanOf(const Value: string; const Default: Boolean): Boolean;
begin
  if (Value = '1') or (Value = 'true') then
    Exit(True);

  if (Value = '0') or (Value = 'false') then
    Exit(False);

  Result := Default;
end;

class function TXml.TryFindElement(const Xml, ElementName: string; out Element: TXmlElement): Boolean;
begin
  const Elements = FindElements(Xml, ElementName);
  Result := (Length(Elements) > 0);
  if Result then
    Element := Elements[0];
end;

class function TXml.IsTagNameEnd(const Value: Char): Boolean;
begin
  Result := CharInSet(Value, [' ', #9, #10, #13, '>', '/']);
end;

class function TXml.ScanTagEnd(const Xml: string; TagStart: Integer): Integer;
begin
  var Quote := #0;

  for var I := TagStart to Length(Xml) do
  begin
    var Current := Xml[I];

    if Quote <> #0 then
    begin
      if Current = Quote then
        Quote := #0;
    end
    else if CharInSet(Current, ['"', '''']) then
      Quote := Current
    else if Current = '>' then
      Exit(I);
  end;

  Result := 0;
end;

class function TXml.CloseTagEnd(const Xml, ClosePrefix: string; const Cursor: Integer): Integer;
begin
  if Copy(Xml, Cursor, Length(ClosePrefix)) <> ClosePrefix then
    Exit(0);

  var TagEnd := Cursor + Length(ClosePrefix);
  while (TagEnd <= Length(Xml)) and CharInSet(Xml[TagEnd], [' ', #9, #10, #13]) do
    Inc(TagEnd);

  const IsCloseTag = (TagEnd <= Length(Xml)) and (Xml[TagEnd] = '>');
  if IsCloseTag then
    Result := TagEnd
  else
    Result := 0;
end;

class function TXml.FindElements(const Xml, ElementName: string): TArray<TXmlElement>;
begin
  var OpenPrefix := '<' + ElementName;
  var ClosePrefix := '</' + ElementName;

  var Elements := TList<TXmlElement>.Create;
  try
    var Depth := 0;
    var ElementStart := 0;
    var ContentStart := 0;
    var Cursor := 1;

    while Cursor <= Length(Xml) do
    begin
      if Xml[Cursor] <> '<' then
      begin
        Inc(Cursor);
        Continue;
      end;

      var CloseEnd := 0;
      if Depth > 0 then
        CloseEnd := CloseTagEnd(Xml, ClosePrefix, Cursor);

      if CloseEnd > 0 then
      begin
        Dec(Depth);

        if Depth = 0 then
        begin
          var Element: TXmlElement;
          Element.StartPos := ElementStart;
          Element.EndPos := CloseEnd + 1;
          Element.OpenTag := Copy(Xml, ElementStart, ContentStart - ElementStart);
          Element.Inner := Copy(Xml, ContentStart, Cursor - ContentStart);
          Elements.Add(Element);
        end;

        Cursor := CloseEnd + 1;
        Continue;
      end;

      var IsOpenTag := (Copy(Xml, Cursor, Length(OpenPrefix)) = OpenPrefix) and
        (Cursor + Length(OpenPrefix) <= Length(Xml)) and
        IsTagNameEnd(Xml[Cursor + Length(OpenPrefix)]);

      if not IsOpenTag then
      begin
        Inc(Cursor);
        Continue;
      end;

      var TagEnd := ScanTagEnd(Xml, Cursor);
      if TagEnd = 0 then
        Break;

      if Xml[TagEnd - 1] = '/' then
      begin
        if Depth = 0 then
        begin
          var Element: TXmlElement;
          Element.StartPos := Cursor;
          Element.EndPos := TagEnd + 1;
          Element.OpenTag := Copy(Xml, Cursor, TagEnd - Cursor + 1);
          Element.Inner := '';
          Elements.Add(Element);
        end;
      end
      else
      begin
        if Depth = 0 then
        begin
          ElementStart := Cursor;
          ContentStart := TagEnd + 1;
        end;
        Inc(Depth);
      end;

      Cursor := TagEnd + 1;
    end;

    Result := Elements.ToArray;
  finally
    Elements.Free;
  end;
end;

class function TXml.ReplaceElements(const Xml: string; const Elements: TArray<TXmlElement>;
  const GetReplacement: TFunc<TXmlElement, string>): string;
begin
  if Length(Elements) = 0 then
    Exit(Xml);

  var Builder := TStringBuilder.Create;
  try
    var Cursor := 1;

    for var Element in Elements do
    begin
      Builder.Append(Copy(Xml, Cursor, Element.StartPos - Cursor));

      if Assigned(GetReplacement) then
        Builder.Append(GetReplacement(Element));

      Cursor := Element.EndPos;
    end;

    Builder.Append(Copy(Xml, Cursor, Length(Xml) - Cursor + 1));

    Result := Builder.ToString;
  finally
    Builder.Free;
  end;
end;

class function TXml.RemoveElements(const Xml: string; const Elements: TArray<TXmlElement>): string;
begin
  Result := ReplaceElements(Xml, Elements, nil);
end;

class function TXml.SelectAlternateContentBranch(const AlternateContentXml: string): string;
begin
  var Choices := FindElements(AlternateContentXml, ElementChoice);
  if Length(Choices) > 0 then
    Exit(Choices[0].Inner);

  var Fallbacks := FindElements(AlternateContentXml, ElementFallback);
  if Length(Fallbacks) > 0 then
    Exit(Fallbacks[0].Inner);

  Result := '';
end;

class function TXml.ReduceAlternateContent(const Xml: string): string;
begin
  Result := Xml;

  var Elements := FindElements(Result, ElementAlternateContent);
  while Length(Elements) > 0 do
  begin
    Result := ReplaceElements(Result, Elements,
      function(Element: TXmlElement): string
      begin
        Result := TXml.SelectAlternateContentBranch(Element.Inner);
      end);

    Elements := FindElements(Result, ElementAlternateContent);
  end;
end;

end.
