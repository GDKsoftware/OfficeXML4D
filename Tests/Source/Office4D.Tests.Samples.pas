unit Office4D.Tests.Samples;

interface

Uses
  System.SysUtils,
  System.IOUtils;

Type
  TOffice4DTests = Class
  protected
    function GetSamplesPath: string;
    function GetWordSamplePath: string;
    function GetExcelSamplePath: string;
  end;

implementation

function TOffice4DTests.GetSamplesPath: string;
begin
  Result := TPath.GetFullPath(TPath.Combine(ExtractFilePath(ParamStr(0)), '..\Samples'));
end;

function TOffice4DTests.GetWordSamplePath: string;
begin
  Result := TPath.Combine(GetSamplesPath, 'Word\simple_word.docx');
end;

function TOffice4DTests.GetExcelSamplePath: string;
begin
  Result := TPath.Combine(GetSamplesPath, 'Excel\excel.xlsx');
end;

end.
