unit Office4D.Relationships;

interface

uses
  System.Classes,
  System.SysUtils,
  System.Generics.Collections;

type
  TRelationship = record
    Id: string;
    RelType: string;
    Target: string;
  end;

  TRelationshipList = TList<TRelationship>;

  TRelationships = class
  private
    FRelationships: TRelationshipList;

    class function ParentFolderOf(const PartName: string): string; static;
    class function NormalizeSegments(const PartName: string): string; static;
  public
    /// <summary>
    /// Returns the package part a relationship Target points at, seen from the
    /// part SourcePart whose relationships hold it. Writers state targets
    /// relative to that part ('worksheets/sheet1.xml', '../comments1.xml') or
    /// absolute from the package root ('/xl/worksheets/sheet1.xml'); both
    /// resolve to a part name without a leading slash.
    /// </summary>
    class function ResolveTarget(const SourcePart, Target: string): string; static;

    /// <summary>
    /// Returns the name of the part that holds the relationships of PartName,
    /// for example 'xl/worksheets/_rels/sheet1.xml.rels' for
    /// 'xl/worksheets/sheet1.xml'.
    /// </summary>
    class function RelationshipsPartOf(const PartName: string): string; static;

    constructor Create;
    destructor Destroy; override;

    procedure LoadFromXml(const XmlContent: string);
    function TryGetById(const Id: string; out Relationship: TRelationship): Boolean;
    function GetById(const Id: string): TRelationship;
    function GetByType(const RelType: string): TRelationship;
    function GetTargetByType(const RelType: string): string;
    function HasType(const RelType: string): Boolean;
    function Count: Integer;

    property Relationships: TRelationshipList read FRelationships;
  end;

implementation

uses
  Office4D.Errors,
  Office4D.Types,
  Office4D.Xml;

const
  ElementRelationship = 'Relationship';
  PartSeparator = '/';
  RelationshipsFolder = '_rels/';
  RelationshipsExtension = '.rels';
  ParentSegment = '..';
  CurrentSegment = '.';

{ TRelationships }

class function TRelationships.ResolveTarget(const SourcePart, Target: string): string;
begin
  const IsAbsolute = Target.StartsWith(PartSeparator);
  if IsAbsolute then
  begin
    const PathFromRoot = Target.Substring(1);
    Result := NormalizeSegments(PathFromRoot);
    Exit;
  end;

  const SourceFolder = ParentFolderOf(SourcePart);
  Result := NormalizeSegments(SourceFolder + Target);
end;

class function TRelationships.RelationshipsPartOf(const PartName: string): string;
begin
  const FileNameStart = PartName.LastIndexOf(PartSeparator) + 1;
  const FileName = PartName.Substring(FileNameStart);
  const Folder = ParentFolderOf(PartName);
  Result := Folder + RelationshipsFolder + FileName + RelationshipsExtension;
end;

class function TRelationships.ParentFolderOf(const PartName: string): string;
begin
  const FolderLength = PartName.LastIndexOf(PartSeparator) + 1;
  Result := PartName.Substring(0, FolderLength);
end;

class function TRelationships.NormalizeSegments(const PartName: string): string;
begin
  var Segments := TList<string>.Create;
  try
    for var Segment in PartName.Split([PartSeparator]) do
    begin
      const IsParent = (Segment = ParentSegment);
      const IsNamed = ((Segment <> CurrentSegment) and not Segment.IsEmpty);
      const HasSegments = (Segments.Count > 0);

      if IsParent then
      begin
        if HasSegments then
          Segments.Delete(Segments.Count - 1);
      end
      else if IsNamed then
      begin
        Segments.Add(Segment);
      end;
    end;

    Result := string.Join(PartSeparator, Segments.ToArray);
  finally
    Segments.Free;
  end;
end;

constructor TRelationships.Create;
begin
  inherited Create;
  FRelationships := TRelationshipList.Create;
end;

destructor TRelationships.Destroy;
begin
  FRelationships.Free;
  inherited;
end;

procedure TRelationships.LoadFromXml(const XmlContent: string);
begin
  FRelationships.Clear;

  const Prefix = TXml.PrefixOf(XmlContent, NsPackageRelationships);
  for var Element in TXml.FindElements(XmlContent, Prefix + ElementRelationship) do
  begin
    var Rel: TRelationship;
    Rel.Id := Element.Attribute('Id');
    Rel.RelType := Element.Attribute('Type');
    Rel.Target := Element.Attribute('Target');
    FRelationships.Add(Rel);
  end;
end;

function TRelationships.TryGetById(const Id: string; out Relationship: TRelationship): Boolean;
begin
  for var Rel in FRelationships do
  begin
    if SameText(Rel.Id, Id) then
    begin
      Relationship := Rel;
      Exit(True);
    end;
  end;

  Result := False;
end;

function TRelationships.GetById(const Id: string): TRelationship;
begin
  if TryGetById(Id, Result) then
    Exit;

  raise EOfficeDocumentException.CreateFmt('Relationship not found: %s', [Id]);
end;

function TRelationships.GetByType(const RelType: string): TRelationship;
begin
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(Rel);
  end;

  raise EOfficeDocumentException.CreateFmt('Relationship type not found: %s', [RelType]);
end;

function TRelationships.GetTargetByType(const RelType: string): string;
begin
  Result := '';
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(Rel.Target);
  end;
end;

function TRelationships.HasType(const RelType: string): Boolean;
begin
  Result := False;
  for var Rel in FRelationships do
  begin
    if SameText(Rel.RelType, RelType) then
      Exit(True);
  end;
end;

function TRelationships.Count: Integer;
begin
  Result := FRelationships.Count;
end;

end.
