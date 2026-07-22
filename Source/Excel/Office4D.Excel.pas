unit Office4D.Excel;

interface

uses
  System.Classes,
  System.Generics.Collections,
  System.SysUtils,
  Office4D.Metadata;

type
  {$SCOPEDENUMS ON}
  TExcelBorderStyle = (None, Thin, Medium, Thick, Dashed, Dotted, Double);
  TExcelHAlign = (None, Left, Center, Right, Justify);
  TExcelVAlign = (None, Top, Center, Bottom);
  TExcelBorderSide = (Top, Right, Bottom, Left);
  TExcelBorderSides = set of TExcelBorderSide;
  // VeryHidden sheets can only be made visible again through code (or the VBA editor),
  // not through the Excel UI.
  TExcelSheetVisibility = (Visible, Hidden, VeryHidden);
  TExcelFontStyle = (Bold, Italic, Underline, Strikeout);
  TExcelFontStyles = set of TExcelFontStyle;
  {$SCOPEDENUMS OFF}

  IExcelCell = interface;
  IExcelSheet = interface;
  IExcelWorkbook = interface;

  IExcelCell = interface
    ['{B8E7F3A1-4D2C-4E8F-9A1B-3C5D7E9F0A2B}']
    function GetAsString: string;
    procedure SetAsString(const Value: string);
    function GetAsFloat: Double;
    procedure SetAsFloat(const Value: Double);
    function GetAsBoolean: Boolean;
    procedure SetAsBoolean(const Value: Boolean);
    function GetAsDateTime: TDateTime;
    procedure SetAsDateTime(const Value: TDateTime);
    function GetFormula: string;
    function GetHasFormula: Boolean;
    procedure SetFormula(const Formula: string; const CalculatedValue: Double);
    function GetBold: Boolean;
    procedure SetBold(const Value: Boolean);
    function GetItalic: Boolean;
    procedure SetItalic(const Value: Boolean);
    function GetUnderline: Boolean;
    procedure SetUnderline(const Value: Boolean);
    function GetStrikeout: Boolean;
    procedure SetStrikeout(const Value: Boolean);
    function GetFontName: string;
    procedure SetFontName(const Value: string);
    function GetFontSize: Double;
    procedure SetFontSize(const Value: Double);
    function GetBackgroundColor: Cardinal;
    procedure SetBackgroundColor(const Value: Cardinal);
    function GetNumberFormat: string;
    procedure SetNumberFormat(const Value: string);
    function GetBorderStyle(ASides: TExcelBorderSides): TExcelBorderStyle;
    procedure SetBorderStyle(ASides: TExcelBorderSides; const Value: TExcelBorderStyle);
    function GetBorderColor(ASides: TExcelBorderSides): Cardinal;
    procedure SetBorderColor(ASides: TExcelBorderSides; const Value: Cardinal);
    function GetHAlign: TExcelHAlign;
    procedure SetHAlign(const Value: TExcelHAlign);
    function GetVAlign: TExcelVAlign;
    procedure SetVAlign(const Value: TExcelVAlign);
    function GetWrapText: Boolean;
    procedure SetWrapText(const Value: Boolean);
    function GetFontColor: Cardinal;
    procedure SetFontColor(const Value: Cardinal);
    function GetFontStyle: TExcelFontStyles;
    procedure SetFontStyle(const Value: TExcelFontStyles);

    property AsString: string read GetAsString write SetAsString;
    property AsFloat: Double read GetAsFloat write SetAsFloat;
    property AsBoolean: Boolean read GetAsBoolean write SetAsBoolean;
    property AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;
    property Formula: string read GetFormula;
    property HasFormula: Boolean read GetHasFormula;
    property Bold: Boolean read GetBold write SetBold;
    property Italic: Boolean read GetItalic write SetItalic;
    property Underline: Boolean read GetUnderline write SetUnderline;
    property Strikeout: Boolean read GetStrikeout write SetStrikeout;
    property FontName: string read GetFontName write SetFontName;
    property FontSize: Double read GetFontSize write SetFontSize;
    property BackgroundColor: Cardinal read GetBackgroundColor write SetBackgroundColor;
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    property BorderStyle[ASides: TExcelBorderSides]: TExcelBorderStyle read GetBorderStyle write SetBorderStyle;
    property BorderColor[ASides: TExcelBorderSides]: Cardinal read GetBorderColor write SetBorderColor;
    property HAlign: TExcelHAlign read GetHAlign write SetHAlign;
    property VAlign: TExcelVAlign read GetVAlign write SetVAlign;
    property WrapText: Boolean read GetWrapText write SetWrapText;
    property FontColor: Cardinal read GetFontColor write SetFontColor;
    property FontStyle: TExcelFontStyles read GetFontStyle write SetFontStyle;
  end;

  IExcelSheet = interface
    ['{C9F8A4B2-5E3D-4F9A-AB2C-4D6E8F0A1B3C}']
    function GetName: string;
    function GetCell(const Address: string): IExcelCell;

    procedure SetColumnWidth(const Column: string; const Width: Double);
    function GetColumnWidth(const Column: string): Double;

    procedure SetRowHeight(const Row: Integer; const Height: Double);
    function GetRowHeight(const Row: Integer): Double;

    procedure MergeCells(const Range: string);
    function GetMergedRanges: TArray<string>;

    function GetVisibility: TExcelSheetVisibility;
    procedure SetVisibility(const Value: TExcelSheetVisibility);

    procedure FreezePanes(const TopLeftCell: string);
    procedure UnfreezePanes;
    function GetFrozenRows: Integer;
    function GetFrozenColumns: Integer;

    function GetNote(const Address: string): string;
    procedure SetNote(const Address: string; const Value: string);

    procedure ClearColumn(const Column: string);
    procedure ClearRow(const Row: Integer);
    procedure DeleteColumn(const Column: string);
    procedure DeleteRow(const Row: Integer);

    function GetCells: TDictionary<string, IExcelCell>;

    property Name: string read GetName;
    property Cell[const Address: string]: IExcelCell read GetCell;
    property Cells: TDictionary<string, IExcelCell> read GetCells;
    property Visibility: TExcelSheetVisibility read GetVisibility write SetVisibility;
    property FrozenRows: Integer read GetFrozenRows;
    property FrozenColumns: Integer read GetFrozenColumns;
    property Note[const Address: string]: string read GetNote write SetNote;
  end;

  IExcelWorkbook = interface
    ['{D0A9B5C3-6F4E-5AAB-BC3D-5E7F9A1B2C4D}']
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(const Stream: TStream);
    procedure SaveToFile(const FileName: string);
    procedure SaveToStream(const Stream: TStream);

    function GetSheetCount: Integer;
    function GetSheet(Index: Integer): IExcelSheet;
    function GetSheetByName(const Name: string): IExcelSheet;
    function GetMetadata: TDocumentMetadata;

    function AddSheet(const Name: string): IExcelSheet;
    procedure RemoveSheet(Index: Integer);
    procedure RemoveSheetByName(const Name: string);

    property SheetCount: Integer read GetSheetCount;
    property Sheets[Index: Integer]: IExcelSheet read GetSheet;
    function SheetByName(const Name: string): IExcelSheet;
    property Metadata: TDocumentMetadata read GetMetadata;
  end;

  TExcelWorkbookFactory = class
  public
    class function Create: IExcelWorkbook;
  end;

const
  AllBorderSides: TExcelBorderSides = [TExcelBorderSide.Top, TExcelBorderSide.Right, TExcelBorderSide.Bottom, TExcelBorderSide.Left];

implementation

uses
  Office4D.Excel.Workbook;

{ TExcelWorkbookFactory }

class function TExcelWorkbookFactory.Create: IExcelWorkbook;
begin
  Result := TExcelWorkbook.Create;
end;

end.
