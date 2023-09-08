program Parser;

uses
  Vcl.Forms,
  ufMain in 'src\view\ufMain.pas' {MainForm},
  Parsing in 'src\logic\Parsing.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
