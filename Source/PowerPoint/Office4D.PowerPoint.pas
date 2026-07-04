unit Office4D.PowerPoint;

interface

uses
  System.Classes,
  Office4D.Metadata;

type
  IPowerPointRun = interface
    ['{7E31A6C2-9B44-4D18-8A5F-2C6E90D1B3A7}']
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

    property Text: string read GetText write SetText;
    property Bold: Boolean read GetBold write SetBold;
    property Italic: Boolean read GetItalic write SetItalic;
    property Underline: Boolean read GetUnderline write SetUnderline;
    property FontName: string read GetFontName write SetFontName;
    property FontSize: Integer read GetFontSize write SetFontSize;
    property FontColor: string read GetFontColor write SetFontColor;
  end;

  IPowerPointParagraph = interface
    ['{7E31A6C2-9B44-4D18-8A5F-2C6E90D1B3A8}']
    function GetText: string;
    function GetRunCount: Integer;
    function GetRun(Index: Integer): IPowerPointRun;
    function AddRun(const Text: string): IPowerPointRun;
    function GetBullet: Boolean;
    procedure SetBullet(const Value: Boolean);
    function GetIndentLevel: Integer;
    procedure SetIndentLevel(const Value: Integer);

    property Text: string read GetText;
    property RunCount: Integer read GetRunCount;
    property Runs[Index: Integer]: IPowerPointRun read GetRun;
    property Bullet: Boolean read GetBullet write SetBullet;
    property IndentLevel: Integer read GetIndentLevel write SetIndentLevel;
  end;

  IPowerPointSlide = interface
    ['{7E31A6C2-9B44-4D18-8A5F-2C6E90D1B3A9}']
    function GetTitle: string;
    procedure SetTitle(const Value: string);
    function GetText: string;
    function GetParagraphCount: Integer;
    function GetParagraph(Index: Integer): IPowerPointParagraph;
    function AddParagraph: IPowerPointParagraph; overload;
    function AddParagraph(const Text: string): IPowerPointParagraph; overload;

    property Title: string read GetTitle write SetTitle;
    property Text: string read GetText;
    property ParagraphCount: Integer read GetParagraphCount;
    property Paragraphs[Index: Integer]: IPowerPointParagraph read GetParagraph;
  end;

  IPowerPointPresentation = interface
    ['{7E31A6C2-9B44-4D18-8A5F-2C6E90D1B3AA}']
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetText: string;
    function GetMetadata: TDocumentMetadata;

    function GetSlideCount: Integer;
    function GetSlide(Index: Integer): IPowerPointSlide;
    function AddSlide: IPowerPointSlide; overload;
    function AddSlide(const Title: string): IPowerPointSlide; overload;

    property Text: string read GetText;
    property Metadata: TDocumentMetadata read GetMetadata;
    property SlideCount: Integer read GetSlideCount;
    property Slides[Index: Integer]: IPowerPointSlide read GetSlide;
  end;

  TPowerPointPresentationFactory = class
  public
    class function CreatePresentation: IPowerPointPresentation;
  end;

implementation

uses
  Office4D.PowerPoint.Presentation;

class function TPowerPointPresentationFactory.CreatePresentation: IPowerPointPresentation;
begin
  Result := TPowerPointPresentation.Create;
end;

end.
