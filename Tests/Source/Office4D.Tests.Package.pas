unit Office4D.Tests.Package;

interface

uses
  System.Classes,
  System.SysUtils,
  System.IOUtils,
  DUnitX.TestFramework,
  Office4D.Tests.Samples,
  Office4D.Package,
  Office4D.Errors;

type
  [TestFixture]
  TPackageTests = class(TOffice4DTests)
  private
    FPackage: TOXMLPackage;
  public
    [Setup]
    procedure Setup;

    [TearDown]
    procedure TearDown;

    [Test]
    procedure Open_ValidDocx_OpensSuccessfully;

    [Test]
    procedure Open_ValidXlsx_OpensSuccessfully;

    [Test]
    procedure Open_NonExistentFile_RaisesException;

    [Test]
    procedure Open_InvalidZipFile_RaisesException;

    [Test]
    procedure PartExists_ContentTypes_ReturnsTrue;

    [Test]
    procedure PartExists_RootRels_ReturnsTrue;

    [Test]
    procedure PartExists_NonExistentPart_ReturnsFalse;

    [Test]
    procedure GetPartStream_ContentTypes_ReturnsStream;

    [Test]
    procedure GetPartStream_NonExistentPart_RaisesException;

    [Test]
    procedure GetPartContent_ContentTypes_ReturnsXml;

    [Test]
    procedure Close_OpenPackage_ReleasesResources;
  end;

implementation

{ TPackageTests }

procedure TPackageTests.Setup;
begin
  FPackage := TOXMLPackage.Create;
end;

procedure TPackageTests.TearDown;
begin
  FPackage.Free;
end;

procedure TPackageTests.Open_ValidDocx_OpensSuccessfully;
begin
  FPackage.Open(GetWordSamplePath);

  Assert.IsTrue(FPackage.IsOpen);
end;

procedure TPackageTests.Open_ValidXlsx_OpensSuccessfully;
begin
  FPackage.Open(GetExcelSamplePath);

  Assert.IsTrue(FPackage.IsOpen);
end;

procedure TPackageTests.Open_NonExistentFile_RaisesException;
begin
  Assert.WillRaise(
    procedure
    begin
      FPackage.Open('C:\nonexistent\file.docx');
    end,
    EPackageNotFound
  );
end;

procedure TPackageTests.Open_InvalidZipFile_RaisesException;
begin
  var TempFile := TPath.Combine(TPath.GetTempPath, 'invalid_test.docx');
  try
    TFile.WriteAllText(TempFile, 'This is not a valid ZIP file');

    Assert.WillRaise(
      procedure
      begin
        FPackage.Open(TempFile);
      end,
      EPackageInvalid
    );
  finally
    if TFile.Exists(TempFile) then
      TFile.Delete(TempFile);
  end;
end;

procedure TPackageTests.PartExists_ContentTypes_ReturnsTrue;
begin
  FPackage.Open(GetWordSamplePath);

  Assert.IsTrue(FPackage.PartExists('[Content_Types].xml'));
end;

procedure TPackageTests.PartExists_RootRels_ReturnsTrue;
begin
  FPackage.Open(GetWordSamplePath);

  Assert.IsTrue(FPackage.PartExists('_rels/.rels'));
end;

procedure TPackageTests.PartExists_NonExistentPart_ReturnsFalse;
begin
  FPackage.Open(GetWordSamplePath);

  Assert.IsFalse(FPackage.PartExists('nonexistent/part.xml'));
end;

procedure TPackageTests.GetPartStream_ContentTypes_ReturnsStream;
begin
  FPackage.Open(GetWordSamplePath);

  var Stream := FPackage.GetPartStream('[Content_Types].xml');
  try
    Assert.IsNotNull(Stream);
    Assert.IsTrue(Stream.Size > 0);
  finally
    Stream.Free;
  end;
end;

procedure TPackageTests.GetPartStream_NonExistentPart_RaisesException;
begin
  FPackage.Open(GetWordSamplePath);

  Assert.WillRaise(
    procedure
    begin
      FPackage.GetPartStream('nonexistent/part.xml').Free;
    end,
    EPartNotFound
  );
end;

procedure TPackageTests.GetPartContent_ContentTypes_ReturnsXml;
begin
  FPackage.Open(GetWordSamplePath);

  var Content := FPackage.GetPartContent('[Content_Types].xml');

  Assert.IsNotEmpty(Content);
  Assert.Contains(Content, '<?xml');
  Assert.Contains(Content, 'Types');
end;

procedure TPackageTests.Close_OpenPackage_ReleasesResources;
begin
  FPackage.Open(GetWordSamplePath);
  Assert.IsTrue(FPackage.IsOpen);

  FPackage.Close;

  Assert.IsFalse(FPackage.IsOpen);
end;

initialization
  TDUnitX.RegisterTestFixture(TPackageTests);

end.
