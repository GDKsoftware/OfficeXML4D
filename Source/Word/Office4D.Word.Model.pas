unit Office4D.Word.Model;

interface

uses
  System.SysUtils,
  System.Generics.Collections,
  Office4D.Word,
  Office4D.Metadata;

type
  TWordRun = class(TInterfacedObject, IWordRun)
  private
    FText: string;
    FBold: Boolean;
    FItalic: Boolean;
    FUnderline: Boolean;
    FHyperlink: string;
    FFontName: string;
    FFontSize: Integer;
    FFontColor: string;
    FImage: TWordImage;
    FHasImage: Boolean;
  public
    function GetText: string;
    procedure SetText(const Value: string);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetHyperlink: string;
    procedure SetHyperlink(const Value: string);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Integer;
    procedure SetFontSize(const Value: Integer);
    function GetFontColor: string;
    procedure SetFontColor(const Value: string);
    function GetImage: TWordImage;
    function HasImage: Boolean;
    procedure AddImage(const AData: TBytes; const AExtension: string; AWidthPx, AHeightPx: Integer);
  end;

  TWordParagraph = class(TInterfacedObject, IWordParagraph)
  private
    FRuns: TList<IWordRun>;
    FListStyle: TListStyle;
    FAlignment: TParagraphAlignment;
    FLineSpacing: TLineSpacing;
    FIndent: TParagraphIndent;
  public
    constructor Create;
    destructor Destroy; override;

    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IWordRun;
    function AddRun(const Text: string): IWordRun;
    procedure AddLineBreak;
    procedure AddTab;
    procedure AddPageBreak;
    function GetListStyle: TListStyle;
    procedure SetListStyle(const Value: TListStyle);
    function GetAlignment: TParagraphAlignment;
    procedure SetAlignment(const Value: TParagraphAlignment);
    function GetLineSpacing: TLineSpacing;
    procedure SetLineSpacing(const Value: TLineSpacing);
    function GetIndent: TParagraphIndent;
    procedure SetIndent(const Value: TParagraphIndent);
  end;

  TWordTableCell = class(TInterfacedObject, IWordTableCell)
  private
    FText: string;
    FShading: string;
    FWidth: Integer;
  public
    function GetText: string;
    procedure SetText(const Value: string);
    function GetShading: string;
    procedure SetShading(const Value: string);
    function GetWidth: Integer;
    procedure SetWidth(const Value: Integer);
  end;

  TWordTable = class(TInterfacedObject, IWordTable)
  private
    FRowCount: Integer;
    FColCount: Integer;
    FCells: TList<TList<IWordTableCell>>;
    FBorders: TTableBorders;
    FColumnWidths: TArray<Integer>;
  public
    constructor Create(const Rows, Cols: Integer);
    destructor Destroy; override;

    function GetRowCount: Integer;
    function GetColCount: Integer;
    function GetCell(Row, Col: Integer): IWordTableCell;
    function GetBorders: TTableBorders;
    procedure SetBorders(const Value: TTableBorders);
    procedure SetColumnWidths(const Widths: array of Integer);

    /// Not part of IWordTable: only a writer needs to read the widths back.
    function GetColumnWidths: TArray<Integer>;
  end;

  TWordHeaderFooter = class(TInterfacedObject, IWordHeaderFooter)
  private
    FText: string;
  public
    function GetText: string;
    procedure SetText(const Value: string);
  end;

  /// <summary>
  /// Everything a Word document holds: what a reader fills and a writer emits.
  /// Keeping it in one object lets the reader and the writer work on a document
  /// without knowing TWordDocument, or each other.
  /// </summary>
  TWordDocumentContent = class
  private
    FParagraphs: TList<IWordParagraph>;
    FTables: TList<IWordTable>;
    FMetadata: TDocumentMetadata;
    FHeader: IWordHeaderFooter;
    FFooter: IWordHeaderFooter;
    FPageOrientation: TPageOrientation;
    FPageMargins: TPageMargins;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;

    property Paragraphs: TList<IWordParagraph> read FParagraphs;
    property Tables: TList<IWordTable> read FTables;
    property Metadata: TDocumentMetadata read FMetadata write FMetadata;
    property Header: IWordHeaderFooter read FHeader;
    property Footer: IWordHeaderFooter read FFooter;
    property PageOrientation: TPageOrientation read FPageOrientation write FPageOrientation;
    property PageMargins: TPageMargins read FPageMargins write FPageMargins;
  end;

implementation

const
  EmuPerPixel = 9525;

{ TWordRun }

function TWordRun.GetText: string;
begin
  Result := FText;
end;

procedure TWordRun.SetText(const Value: string);
begin
  FText := Value;
end;

function TWordRun.GetBold: Boolean;
begin
  Result := FBold;
end;

procedure TWordRun.SetBold(const Value: Boolean);
begin
  FBold := Value;
end;

function TWordRun.GetItalic: Boolean;
begin
  Result := FItalic;
end;

procedure TWordRun.SetItalic(const Value: Boolean);
begin
  FItalic := Value;
end;

function TWordRun.GetUnderline: Boolean;
begin
  Result := FUnderline;
end;

procedure TWordRun.SetUnderline(const Value: Boolean);
begin
  FUnderline := Value;
end;

function TWordRun.GetHyperlink: string;
begin
  Result := FHyperlink;
end;

procedure TWordRun.SetHyperlink(const Value: string);
begin
  FHyperlink := Value;
end;

function TWordRun.GetFontName: string;
begin
  Result := FFontName;
end;

procedure TWordRun.SetFontName(const Value: string);
begin
  FFontName := Value;
end;

function TWordRun.GetFontSize: Integer;
begin
  Result := FFontSize;
end;

procedure TWordRun.SetFontSize(const Value: Integer);
begin
  FFontSize := Value;
end;

function TWordRun.GetFontColor: string;
begin
  Result := FFontColor;
end;

procedure TWordRun.SetFontColor(const Value: string);
begin
  FFontColor := Value;
end;

function TWordRun.GetImage: TWordImage;
begin
  Result := FImage;
end;

function TWordRun.HasImage: Boolean;
begin
  Result := FHasImage;
end;

procedure TWordRun.AddImage(const AData: TBytes; const AExtension: string; AWidthPx, AHeightPx: Integer);
begin
  FImage := TWordImage.Create(AData, AExtension, AWidthPx * EmuPerPixel, AHeightPx * EmuPerPixel);
  FHasImage := True;
  FText := '';
end;

{ TWordParagraph }

constructor TWordParagraph.Create;
begin
  inherited Create;
  FRuns := TList<IWordRun>.Create;
end;

destructor TWordParagraph.Destroy;
begin
  FRuns.Free;
  inherited;
end;

function TWordParagraph.GetText: string;
begin
  Result := '';
  for var Run in FRuns do
    Result := Result + Run.Text;
end;

function TWordParagraph.GetRunCount: Integer;
begin
  Result := FRuns.Count;
end;

function TWordParagraph.GetRun(Index: Integer): IWordRun;
begin
  Result := FRuns[Index];
end;

function TWordParagraph.AddRun(const Text: string): IWordRun;
begin
  var Run := TWordRun.Create;
  Run.FText := Text;
  FRuns.Add(Run);
  Result := Run;
end;

procedure TWordParagraph.AddLineBreak;
begin
  var Run := TWordRun.Create;
  Run.FText := sLineBreak;
  FRuns.Add(Run);
end;

procedure TWordParagraph.AddTab;
begin
  var Run := TWordRun.Create;
  Run.FText := #9;
  FRuns.Add(Run);
end;

function TWordParagraph.GetListStyle: TListStyle;
begin
  Result := FListStyle;
end;

procedure TWordParagraph.SetListStyle(const Value: TListStyle);
begin
  FListStyle := Value;
end;

function TWordParagraph.GetAlignment: TParagraphAlignment;
begin
  Result := FAlignment;
end;

procedure TWordParagraph.SetAlignment(const Value: TParagraphAlignment);
begin
  FAlignment := Value;
end;

procedure TWordParagraph.AddPageBreak;
begin
  var Run := TWordRun.Create;
  Run.FText := #12;
  FRuns.Add(Run);
end;

function TWordParagraph.GetLineSpacing: TLineSpacing;
begin
  Result := FLineSpacing;
end;

procedure TWordParagraph.SetLineSpacing(const Value: TLineSpacing);
begin
  FLineSpacing := Value;
end;

function TWordParagraph.GetIndent: TParagraphIndent;
begin
  Result := FIndent;
end;

procedure TWordParagraph.SetIndent(const Value: TParagraphIndent);
begin
  FIndent := Value;
end;

{ TWordTableCell }

function TWordTableCell.GetText: string;
begin
  Result := FText;
end;

procedure TWordTableCell.SetText(const Value: string);
begin
  FText := Value;
end;

function TWordTableCell.GetShading: string;
begin
  Result := FShading;
end;

procedure TWordTableCell.SetShading(const Value: string);
begin
  FShading := Value;
end;

function TWordTableCell.GetWidth: Integer;
begin
  Result := FWidth;
end;

procedure TWordTableCell.SetWidth(const Value: Integer);
begin
  FWidth := Value;
end;

{ TWordTable }

constructor TWordTable.Create(const Rows, Cols: Integer);
begin
  inherited Create;
  FRowCount := Rows;
  FColCount := Cols;
  FCells := TList<TList<IWordTableCell>>.Create;

  for var RowIndex := 0 to Rows - 1 do
  begin
    var Row := TList<IWordTableCell>.Create;
    for var ColIndex := 0 to Cols - 1 do
    begin
      Row.Add(TWordTableCell.Create);
    end;
    FCells.Add(Row);
  end;
end;

destructor TWordTable.Destroy;
begin
  for var Row in FCells do
  begin
    Row.Free;
  end;
  FCells.Free;
  inherited;
end;

function TWordTable.GetRowCount: Integer;
begin
  Result := FRowCount;
end;

function TWordTable.GetColCount: Integer;
begin
  Result := FColCount;
end;

function TWordTable.GetCell(Row, Col: Integer): IWordTableCell;
begin
  Result := FCells[Row][Col];
end;

function TWordTable.GetBorders: TTableBorders;
begin
  Result := FBorders;
end;

procedure TWordTable.SetBorders(const Value: TTableBorders);
begin
  FBorders := Value;
end;

procedure TWordTable.SetColumnWidths(const Widths: array of Integer);
begin
  SetLength(FColumnWidths, Length(Widths));
  for var I := 0 to High(Widths) do
    FColumnWidths[I] := Widths[I];
end;

function TWordTable.GetColumnWidths: TArray<Integer>;
begin
  Result := FColumnWidths;
end;

{ TWordHeaderFooter }

function TWordHeaderFooter.GetText: string;
begin
  Result := FText;
end;

procedure TWordHeaderFooter.SetText(const Value: string);
begin
  FText := Value;
end;

{ TWordDocumentContent }

constructor TWordDocumentContent.Create;
begin
  inherited Create;
  FParagraphs := TList<IWordParagraph>.Create;
  FTables := TList<IWordTable>.Create;
  FMetadata.Clear;
  FHeader := TWordHeaderFooter.Create;
  FFooter := TWordHeaderFooter.Create;
end;

destructor TWordDocumentContent.Destroy;
begin
  FTables.Free;
  FParagraphs.Free;
  inherited;
end;

procedure TWordDocumentContent.Clear;
begin
  FParagraphs.Clear;
  FTables.Clear;
  FMetadata.Clear;
  FHeader.Text := '';
  FFooter.Text := '';
end;

end.
