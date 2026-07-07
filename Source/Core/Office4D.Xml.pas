unit Office4D.Xml;

interface

type
  /// <summary>
  /// Shared XML text helpers used by the Word, Excel and PowerPoint modules.
  /// Escape is applied when writing element/attribute text, Unescape when
  /// reading it back, so text round-trips through the five predefined entities.
  /// </summary>
  TXml = record
  public
    class function Escape(const Value: string): string; static;
    class function Unescape(const Value: string): string; static;
  end;

implementation

uses
  System.SysUtils;

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

end.
