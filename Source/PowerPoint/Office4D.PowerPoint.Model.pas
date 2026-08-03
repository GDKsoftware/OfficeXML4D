unit Office4D.PowerPoint.Model;

interface

uses
  System.SysUtils,
  System.Generics.Collections,
  Office4D.PowerPoint,
  Office4D.Metadata;

type
  TPowerPointRun = class(TInterfacedObject, IPowerPointRun)
  private
    FText: string;
    FBold: Boolean;
    FItalic: Boolean;
    FUnderline: Boolean;
    FFontName: string;
    FFontSize: Integer;
    FFontColor: string;

    function GetText: string;
    procedure SetText(const Value: string);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Integer;
    procedure SetFontSize(const Value: Integer);
    function GetFontColor: string;
    procedure SetFontColor(const Value: string);
  end;

  TPowerPointParagraph = class(TInterfacedObject, IPowerPointParagraph)
  private
    FRuns: TList<IPowerPointRun>;
    FBullet: Boolean;
    FIndentLevel: Integer;

    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IPowerPointRun;
    function GetBullet: Boolean;
    procedure SetBullet(const Value: Boolean);
    function GetIndentLevel: Integer;
    procedure SetIndentLevel(const Value: Integer);
  public
    constructor Create;
    destructor Destroy; override;

    function AddRun(const Text: string): IPowerPointRun;
  end;

  TPowerPointSlide = class(TInterfacedObject, IPowerPointSlide)
  private
    FTitle: string;
    FParagraphs: TList<IPowerPointParagraph>;

    function GetTitle: string;
    procedure SetTitle(const Value: string);
    function GetText: string;
    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IPowerPointParagraph;
  public
    constructor Create;
    destructor Destroy; override;

    function AddParagraph: IPowerPointParagraph; overload;
    function AddParagraph(const Text: string): IPowerPointParagraph; overload;
  end;

  /// <summary>
  /// Everything a presentation holds: what a reader fills and a writer emits.
  /// Keeping it in one object lets the reader and the writer work on a
  /// presentation without knowing TPowerPointPresentation, or each other.
  /// </summary>
  TPowerPointContent = class
  private
    FSlides: TList<IPowerPointSlide>;
    FMetadata: TDocumentMetadata;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;
    function AddSlide: IPowerPointSlide;

    property Slides: TList<IPowerPointSlide> read FSlides;
    property Metadata: TDocumentMetadata read FMetadata write FMetadata;
  end;

implementation

{ TPowerPointRun }

function TPowerPointRun.GetText: string;
begin
  Result := FText;
end;

procedure TPowerPointRun.SetText(const Value: string);
begin
  FText := Value;
end;

function TPowerPointRun.GetBold: Boolean;
begin
  Result := FBold;
end;

procedure TPowerPointRun.SetBold(const Value: Boolean);
begin
  FBold := Value;
end;

function TPowerPointRun.GetItalic: Boolean;
begin
  Result := FItalic;
end;

procedure TPowerPointRun.SetItalic(const Value: Boolean);
begin
  FItalic := Value;
end;

function TPowerPointRun.GetUnderline: Boolean;
begin
  Result := FUnderline;
end;

procedure TPowerPointRun.SetUnderline(const Value: Boolean);
begin
  FUnderline := Value;
end;

function TPowerPointRun.GetFontName: string;
begin
  Result := FFontName;
end;

procedure TPowerPointRun.SetFontName(const Value: string);
begin
  FFontName := Value;
end;

function TPowerPointRun.GetFontSize: Integer;
begin
  Result := FFontSize;
end;

procedure TPowerPointRun.SetFontSize(const Value: Integer);
begin
  FFontSize := Value;
end;

function TPowerPointRun.GetFontColor: string;
begin
  Result := FFontColor;
end;

procedure TPowerPointRun.SetFontColor(const Value: string);
begin
  FFontColor := Value;
end;

{ TPowerPointParagraph }

constructor TPowerPointParagraph.Create;
begin
  inherited Create;
  FRuns := TList<IPowerPointRun>.Create;
end;

destructor TPowerPointParagraph.Destroy;
begin
  FRuns.Free;
  inherited;
end;

function TPowerPointParagraph.GetText: string;
begin
  Result := '';
  for var Run in FRuns do
    Result := Result + Run.Text;
end;

function TPowerPointParagraph.GetRunCount: Integer;
begin
  Result := FRuns.Count;
end;

function TPowerPointParagraph.GetRun(Index: Integer): IPowerPointRun;
begin
  Result := FRuns[Index];
end;

function TPowerPointParagraph.AddRun(const Text: string): IPowerPointRun;
begin
  Result := TPowerPointRun.Create;
  Result.Text := Text;
  FRuns.Add(Result);
end;

function TPowerPointParagraph.GetBullet: Boolean;
begin
  Result := FBullet;
end;

procedure TPowerPointParagraph.SetBullet(const Value: Boolean);
begin
  FBullet := Value;
end;

function TPowerPointParagraph.GetIndentLevel: Integer;
begin
  Result := FIndentLevel;
end;

procedure TPowerPointParagraph.SetIndentLevel(const Value: Integer);
begin
  FIndentLevel := Value;
end;

{ TPowerPointSlide }

constructor TPowerPointSlide.Create;
begin
  inherited Create;
  FParagraphs := TList<IPowerPointParagraph>.Create;
end;

destructor TPowerPointSlide.Destroy;
begin
  FParagraphs.Free;
  inherited;
end;

function TPowerPointSlide.GetTitle: string;
begin
  Result := FTitle;
end;

procedure TPowerPointSlide.SetTitle(const Value: string);
begin
  FTitle := Value;
end;

function TPowerPointSlide.GetText: string;
begin
  Result := FTitle;
  for var Paragraph in FParagraphs do
  begin
    if Result <> '' then
      Result := Result + sLineBreak;
    Result := Result + Paragraph.Text;
  end;
end;

function TPowerPointSlide.GetParagraphCount: Integer;
begin
  Result := FParagraphs.Count;
end;

function TPowerPointSlide.GetParagraph(Index: Integer): IPowerPointParagraph;
begin
  Result := FParagraphs[Index];
end;

function TPowerPointSlide.AddParagraph: IPowerPointParagraph;
begin
  Result := TPowerPointParagraph.Create;
  FParagraphs.Add(Result);
end;

function TPowerPointSlide.AddParagraph(const Text: string): IPowerPointParagraph;
begin
  Result := AddParagraph;
  Result.AddRun(Text);
end;


{ TPowerPointContent }

constructor TPowerPointContent.Create;
begin
  inherited Create;
  FSlides := TList<IPowerPointSlide>.Create;
  FMetadata.Clear;
end;

destructor TPowerPointContent.Destroy;
begin
  FSlides.Free;
  inherited;
end;

procedure TPowerPointContent.Clear;
begin
  FSlides.Clear;
  FMetadata.Clear;
end;

function TPowerPointContent.AddSlide: IPowerPointSlide;
begin
  Result := TPowerPointSlide.Create;
  FSlides.Add(Result);
end;

end.
