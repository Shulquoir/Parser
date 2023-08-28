unit FileOpener;

interface

uses
  Dialogs, Classes, SysUtils, StdCtrls;

type
  TFileOpener = class
  public
    class function ChooseFile(var filePath: string; const FileExtension: string): Boolean;
    class function OpenFile(Sender: TObject; const Key: Char): Boolean;
  end;

implementation

class function TFileOpener.ChooseFile(var filePath: string; const FileExtension: string): Boolean;
var
  odOpenFileDialog: TOpenDialog;

begin
  Result := False;

  odOpenFileDialog := TOpenDialog.Create(nil);
  try
    odOpenFileDialog.Filter := FileExtension;
    if odOpenFileDialog.Execute then
    begin
      filePath := odOpenFileDialog.FileName;
      ShowMessage('Файл DFN был выбран.');
      Result := True;
    end;
  finally
    odOpenFileDialog.Free;
  end;
end;

class function TFileOpener.OpenFile(Sender: TObject; const Key: Char): Boolean;
const
  ENTER_KEY_CODE = #13;

begin
  Result := False;
  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(TEdit(Sender).Text) = false then
    begin
      ShowMessage('Файл не существует: ' + TEdit(Sender).Text);
      Exit;
    end;

  ShowMessage('Файл был открыттот: ' + TEdit(Sender).Text);
  Result := True;
end;

end.
