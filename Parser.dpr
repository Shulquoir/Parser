program Parser;

uses
  Vcl.Forms,
  ufMain in 'ufMain.pas' {MainForm},
  FileOpener in 'FileOpener.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
