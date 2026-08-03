unit Office4D.PowerPoint.Presentation;

interface

uses
  System.Classes,
  System.Zip,
  Office4D.PowerPoint,
  Office4D.Metadata,
  Office4D.Package,
  Office4D.PowerPoint.Model;

type
  /// <summary>
  /// A presentation as its users see it: the slides, plus loading and saving.
  /// The slides themselves live in TPowerPointContent; reading a package and
  /// writing one are handed to TPowerPointReader and TPowerPointWriter.
  /// </summary>
  TPowerPointPresentation = class(TInterfacedObject, IPowerPointPresentation)
  private
    FContent: TPowerPointContent;
    FPackage: TOXMLPackage;

    procedure ParsePackage;
    procedure WriteToZip(const Zip: TZipFile);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;
    function GetSlideCount: Integer;
    function GetSlide(Index: Integer): IPowerPointSlide;
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function AddSlide: IPowerPointSlide; overload;
    function AddSlide(const Title: string): IPowerPointSlide; overload;
  end;

implementation

uses
  System.SysUtils,
  Office4D.PowerPoint.Reader,
  Office4D.PowerPoint.Writer;

{ TPowerPointPresentation }

constructor TPowerPointPresentation.Create;
begin
  inherited Create;
  FContent := TPowerPointContent.Create;
  FPackage := nil;
end;

destructor TPowerPointPresentation.Destroy;
begin
  FreeAndNil(FPackage);
  FContent.Free;
  inherited;
end;

function TPowerPointPresentation.AddSlide: IPowerPointSlide;
begin
  Result := FContent.AddSlide;
end;

function TPowerPointPresentation.AddSlide(const Title: string): IPowerPointSlide;
begin
  Result := FContent.AddSlide;
  Result.Title := Title;
end;

function TPowerPointPresentation.GetSlideCount: Integer;
begin
  Result := FContent.Slides.Count;
end;

function TPowerPointPresentation.GetSlide(Index: Integer): IPowerPointSlide;
begin
  Result := FContent.Slides[Index];
end;

function TPowerPointPresentation.GetText: string;
begin
  Result := '';
  for var Slide in FContent.Slides do
  begin
    if Result <> '' then
      Result := Result + sLineBreak;
    Result := Result + Slide.Text;
  end;
end;

function TPowerPointPresentation.GetMetadata: TDocumentMetadata;
begin
  Result := FContent.Metadata;
end;

procedure TPowerPointPresentation.LoadFromFile(const FileName: string);
begin
  FreeAndNil(FPackage);
  FContent.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(FileName);
  ParsePackage;
end;

procedure TPowerPointPresentation.LoadFromStream(const Stream: TStream);
begin
  FreeAndNil(FPackage);
  FContent.Clear;

  FPackage := TOXMLPackage.Create;
  FPackage.Open(Stream);
  ParsePackage;
end;

procedure TPowerPointPresentation.ParsePackage;
begin
  var Reader := TPowerPointReader.Create(FPackage, FContent);
  try
    Reader.ParsePackage;
  finally
    Reader.Free;
  end;
end;

procedure TPowerPointPresentation.SaveToFile(const FileName: string);
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

procedure TPowerPointPresentation.SaveToStream(const Stream: TStream);
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

procedure TPowerPointPresentation.WriteToZip(const Zip: TZipFile);
begin
  var Writer := TPowerPointWriter.Create(FContent);
  try
    Writer.AddPartsToZip(Zip);
  finally
    Writer.Free;
  end;
end;

end.
