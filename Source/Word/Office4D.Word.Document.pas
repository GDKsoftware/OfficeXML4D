unit Office4D.Word.Document;

interface

uses
  System.Classes,
  System.Zip,
  Office4D.Word,
  Office4D.Metadata,
  Office4D.Package,
  Office4D.Word.Model;

type
  /// <summary>
  /// A Word document as its users see it: the content, plus loading and saving.
  /// The content itself lives in TWordDocumentContent; reading a package and
  /// writing one are handed to TWordDocumentReader and TWordDocumentWriter.
  /// </summary>
  TWordDocument = class(TInterfacedObject, IWordDocument)
  private
    FContent: TWordDocumentContent;
    FPackage: TOXMLPackage;

    procedure LoadFromPackage;
    procedure WriteToZip(const Zip: TZipFile);
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;

    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IWordParagraph;
    function AddParagraph: IWordParagraph;

    function GetTableCount: Integer;
    function GetTable(Index: Integer): IWordTable;
    function AddTable(const Rows, Cols: Integer): IWordTable;

    function GetPageOrientation: TPageOrientation;
    procedure SetPageOrientation(const Value: TPageOrientation);
    function GetPageMargins: TPageMargins;
    procedure SetPageMargins(const Value: TPageMargins);
    function GetHeader: IWordHeaderFooter;
    function GetFooter: IWordHeaderFooter;
  end;

implementation

uses
  System.SysUtils,
  Office4D.Word.Reader,
  Office4D.Word.Writer;

{ TWordDocument }

constructor TWordDocument.Create;
begin
  inherited Create;
  FContent := TWordDocumentContent.Create;
end;

destructor TWordDocument.Destroy;
begin
  FreeAndNil(FPackage);
  FContent.Free;
  inherited;
end;

procedure TWordDocument.LoadFromFile(const FileName: string);
begin
  FreeAndNil(FPackage);
  FContent.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(FileName);

  LoadFromPackage;
end;

procedure TWordDocument.LoadFromStream(const Stream: TStream);
begin
  FreeAndNil(FPackage);
  FContent.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(Stream);

  LoadFromPackage;
end;

procedure TWordDocument.LoadFromPackage;
begin
  var Reader := TWordDocumentReader.Create(FPackage, FContent);
  try
    Reader.Read;
  finally
    Reader.Free;
  end;
end;

procedure TWordDocument.SaveToFile(const FileName: string);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(FileName, zmWrite);
    WriteToZip(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TWordDocument.SaveToStream(const Stream: TStream);
begin
  var Zip := TZipFile.Create;
  try
    Zip.Open(Stream, zmWrite);
    WriteToZip(Zip);
    Zip.Close;
  finally
    Zip.Free;
  end;
end;

procedure TWordDocument.WriteToZip(const Zip: TZipFile);
begin
  var Writer := TWordDocumentWriter.Create(FContent);
  try
    Writer.WriteParts(Zip);
  finally
    Writer.Free;
  end;
end;

function TWordDocument.GetText: string;
begin
  var Lines := TStringList.Create;
  try
    for var Para in FContent.Paragraphs do
      Lines.Add(Para.Text);
    Result := Lines.Text;
  finally
    Lines.Free;
  end;
end;

function TWordDocument.GetMetadata: TDocumentMetadata;
begin
  Result := FContent.Metadata;
end;

function TWordDocument.GetParagraphCount: Integer;
begin
  Result := FContent.Paragraphs.Count;
end;

function TWordDocument.GetParagraph(Index: Integer): IWordParagraph;
begin
  Result := FContent.Paragraphs[Index];
end;

function TWordDocument.AddParagraph: IWordParagraph;
begin
  var Para := TWordParagraph.Create;
  FContent.Paragraphs.Add(Para);
  Result := Para;
end;

function TWordDocument.GetTableCount: Integer;
begin
  Result := FContent.Tables.Count;
end;

function TWordDocument.GetTable(Index: Integer): IWordTable;
begin
  Result := FContent.Tables[Index];
end;

function TWordDocument.AddTable(const Rows, Cols: Integer): IWordTable;
begin
  var Table := TWordTable.Create(Rows, Cols);
  FContent.Tables.Add(Table);
  Result := Table;
end;

function TWordDocument.GetPageOrientation: TPageOrientation;
begin
  Result := FContent.PageOrientation;
end;

procedure TWordDocument.SetPageOrientation(const Value: TPageOrientation);
begin
  FContent.PageOrientation := Value;
end;

function TWordDocument.GetPageMargins: TPageMargins;
begin
  Result := FContent.PageMargins;
end;

procedure TWordDocument.SetPageMargins(const Value: TPageMargins);
begin
  FContent.PageMargins := Value;
end;

function TWordDocument.GetHeader: IWordHeaderFooter;
begin
  Result := FContent.Header;
end;

function TWordDocument.GetFooter: IWordHeaderFooter;
begin
  Result := FContent.Footer;
end;

end.
