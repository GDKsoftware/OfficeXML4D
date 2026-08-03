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
    class function SelectAlternateContentBranch(const AlternateContentXml: string): string; static;
  public
    class function Escape(const Value: string): string; static;
    class function Unescape(const Value: string): string; static;

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
  System.Generics.Collections;

const
  ElementAlternateContent = 'mc:AlternateContent';
  ElementChoice = 'mc:Choice';
  ElementFallback = 'mc:Fallback';

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

class function TXml.FindElements(const Xml, ElementName: string): TArray<TXmlElement>;
begin
  var OpenPrefix := '<' + ElementName;
  var ClosePrefix := '</' + ElementName + '>';

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

      if (Depth > 0) and (Copy(Xml, Cursor, Length(ClosePrefix)) = ClosePrefix) then
      begin
        Dec(Depth);
        Inc(Cursor, Length(ClosePrefix));

        if Depth = 0 then
        begin
          var Element: TXmlElement;
          Element.StartPos := ElementStart;
          Element.EndPos := Cursor;
          Element.OpenTag := Copy(Xml, ElementStart, ContentStart - ElementStart);
          Element.Inner := Copy(Xml, ContentStart, Cursor - Length(ClosePrefix) - ContentStart);
          Elements.Add(Element);
        end;

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
