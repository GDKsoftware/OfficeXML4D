program Office4D.Tests;

{$APPTYPE CONSOLE}

{$STRONGLINKTYPES ON}

uses
  System.SysUtils,
  DUnitX.CommandLine.Options,
  DUnitX.Loggers.Console,
  DUnitX.Loggers.Xml.JUnit,
  DUnitX.Loggers.Xml.NUnit,
  DUnitX.TestFramework,
  Office4D.Types in '..\..\Source\Core\Office4D.Types.pas',
  Office4D.Errors in '..\..\Source\Core\Office4D.Errors.pas',
  Office4D.Package in '..\..\Source\Core\Office4D.Package.pas',
  Office4D.Relationships in '..\..\Source\Core\Office4D.Relationships.pas',
  Office4D.Xml in '..\..\Source\Core\Office4D.Xml.pas',
  Office4D.Metadata in '..\..\Source\Common\Office4D.Metadata.pas',
  Office4D.Word in '..\..\Source\Word\Office4D.Word.pas',
  Office4D.Word.Document in '..\..\Source\Word\Office4D.Word.Document.pas',
  Office4D.Excel in '..\..\Source\Excel\Office4D.Excel.pas',
  Office4D.Excel.Workbook in '..\..\Source\Excel\Office4D.Excel.Workbook.pas',
  Office4D.PowerPoint in '..\..\Source\PowerPoint\Office4D.PowerPoint.pas',
  Office4D.PowerPoint.Presentation in '..\..\Source\PowerPoint\Office4D.PowerPoint.Presentation.pas',
  Office4D.Tests.Samples in 'Office4D.Tests.Samples.pas',
  Office4D.Tests.Package in 'Office4D.Tests.Package.pas',
  Office4D.Tests.Relationships in 'Office4D.Tests.Relationships.pas',
  Office4D.Tests.Metadata in 'Office4D.Tests.Metadata.pas',
  Office4D.Tests.Word in 'Office4D.Tests.Word.pas',
  Office4D.Tests.Word.Write in 'Office4D.Tests.Word.Write.pas',
  Office4D.Tests.Excel in 'Office4D.Tests.Excel.pas',
  Office4D.Tests.Excel.Write in 'Office4D.Tests.Excel.Write.pas',
  Office4D.Tests.PowerPoint in 'Office4D.Tests.PowerPoint.pas',
  Office4D.Tests.Excel.BorderSides in 'Office4D.Tests.Excel.BorderSides.pas',
  Office4D.Tests.Excel.FontStyle in 'Office4D.Tests.Excel.FontStyle.pas',
  Office4D.Tests.Excel.FreezePanes in 'Office4D.Tests.Excel.FreezePanes.pas',
  Office4D.Tests.Excel.Notes in 'Office4D.Tests.Excel.Notes.pas',
  Office4D.Tests.Excel.RemoveSheet in 'Office4D.Tests.Excel.RemoveSheet.pas',
  Office4D.Tests.Excel.ClearColumnRow in 'Office4D.Tests.Excel.ClearColumnRow.pas',
  Office4D.Tests.Excel.DeleteColumnRow in 'Office4D.Tests.Excel.DeleteColumnRow.pas';

var
  Runner: ITestRunner;
  Results: IRunResults;
  Logger: ITestLogger;
  XmlLogger: ITestLogger;
  UseJUnit: Boolean = False;

begin
  ReportMemoryLeaksOnShutdown := True;

  try
    TOptionsRegistry.RegisterOption<Boolean>('JUnit','','Write JUnit XML output instead of NUnit',
      procedure(Value: Boolean)
      begin
        UseJUnit := Value;
      end).HasValue := False;

    TDUnitX.CheckCommandLine;

    Runner := TDUnitX.CreateRunner;
    Runner.UseRTTI := True;

    Logger := TDUnitXConsoleLogger.Create(True);
    Runner.AddLogger(Logger);

    if UseJUnit then
      XmlLogger := TDUnitXXMLJUnitFileLogger.Create(TDUnitX.Options.XMLOutputFile)
    else
      XmlLogger := TDUnitXXMLNUnitFileLogger.Create(TDUnitX.Options.XMLOutputFile);
    Runner.AddLogger(XmlLogger);
    Runner.FailsOnNoAsserts := False;

    Results := Runner.Execute;

    if not Results.AllPassed then
      System.ExitCode := EXIT_ERRORS;

    {$IFNDEF CI}
    if TDUnitX.Options.ExitBehavior = TDUnitXExitBehavior.Pause then
    begin
      System.Write('Done.. press <Enter> key to quit.');
      System.Readln;
    end;
    {$ENDIF}
  except
    on E: Exception do
      System.Writeln(E.ClassName, ': ', E.Message);
  end;
end.
