unit Office4D.Excel.Workbook;

interface

uses
  System.Classes,
  System.Zip,
  Office4D.Excel,
  Office4D.Metadata,
  Office4D.Package,
  Office4D.Excel.Model;

type
  /// <summary>
  /// A workbook as its users see it: the sheets, plus loading and saving. The
  /// sheets themselves live in TExcelWorkbookContent; reading a package and
  /// writing one are handed to TExcelWorkbookReader and TExcelWorkbookWriter.
  /// </summary>
  TExcelWorkbook = class(TInterfacedObject, IExcelWorkbook)
  private
    FContent: TExcelWorkbookContent;

    procedure ReadPackage(const Package: TOXMLPackage);
    procedure WriteToZip(const Zip: TZipFile);
  public
    constructor Create;
    destructor Destroy; override;

    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetSheetCount: Integer;
    function GetSheet(Index: Integer): IExcelSheet;
    function GetSheetByName(const Name: string): IExcelSheet;
    function GetMetadata: TDocumentMetadata;

    function AddSheet(const Name: string): IExcelSheet;
    function SheetByName(const Name: string): IExcelSheet;

    procedure RemoveSheet(Index: Integer);
    procedure RemoveSheetByName(const Name: string);
  end;

implementation

uses
  System.SysUtils,
  Office4D.Errors,
  Office4D.Excel.Reader,
  Office4D.Excel.Writer;

{ TExcelWorkbook }

constructor TExcelWorkbook.Create;
begin
  inherited;
  FContent := TExcelWorkbookContent.Create;
end;

destructor TExcelWorkbook.Destroy;
begin
  FContent.Free;
  inherited;
end;

procedure TExcelWorkbook.LoadFromFile(const FileName: string);
begin
  if not FileExists(FileName) then
    raise EPackageNotFound.Create('File not found: ' + FileName);

  var Package := TOXMLPackage.Create;
  try
    Package.Open(FileName);
    ReadPackage(Package);
    Package.Close;
  finally
    Package.Free;
  end;
end;

procedure TExcelWorkbook.LoadFromStream(const Stream: TStream);
begin
  var Package := TOXMLPackage.Create;
  try
    Package.Open(Stream);
    ReadPackage(Package);
    Package.Close;
  finally
    Package.Free;
  end;
end;

procedure TExcelWorkbook.ReadPackage(const Package: TOXMLPackage);
begin
  var Reader := TExcelWorkbookReader.Create(FContent);
  try
    Reader.Read(Package);
  finally
    Reader.Free;
  end;
end;

procedure TExcelWorkbook.SaveToFile(const FileName: string);
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

procedure TExcelWorkbook.SaveToStream(const Stream: TStream);
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

procedure TExcelWorkbook.WriteToZip(const Zip: TZipFile);
begin
  var Writer := TExcelWorkbookWriter.Create(FContent);
  try
    Writer.WriteParts(Zip);
  finally
    Writer.Free;
  end;
end;

function TExcelWorkbook.GetSheetCount: Integer;
begin
  Result := FContent.Sheets.Count;
end;

function TExcelWorkbook.GetSheet(Index: Integer): IExcelSheet;
begin
  Result := FContent.Sheets[Index];
end;

function TExcelWorkbook.GetSheetByName(const Name: string): IExcelSheet;
begin
  Result := FContent.SheetByName(Name);
end;

function TExcelWorkbook.GetMetadata: TDocumentMetadata;
begin
  Result := FContent.Metadata;
end;

function TExcelWorkbook.AddSheet(const Name: string): IExcelSheet;
begin
  Result := FContent.AddSheet(Name);
end;

function TExcelWorkbook.SheetByName(const Name: string): IExcelSheet;
begin
  Result := FContent.SheetByName(Name);
end;

procedure TExcelWorkbook.RemoveSheet(Index: Integer);
begin
  FContent.RemoveSheet(Index);
end;

procedure TExcelWorkbook.RemoveSheetByName(const Name: string);
begin
  FContent.RemoveSheetByName(Name);
end;

end.
