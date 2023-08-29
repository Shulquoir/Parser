unit Controller;

interface

uses
  Winapi.Windows, Dialogs, Classes, SysUtils, StdCtrls, ShellAPI, Model;

type
  TController = class
  public
    class function ChooseFile(var filePath: string; const FileExtension: string): Boolean;
    class function OpenFileByName(Sender: TObject; const Key: Char): Boolean;
    class function CreateTSV(const OpenedDFNFilePath: string; var SavedTSVFilePath: string): Boolean;
    class function RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
    class procedure OpenExcel(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
    class procedure OpenCalc(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
  end;

implementation

class function TController.ChooseFile(var filePath: string; const FileExtension: string): Boolean;
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
      MessageDlg('���� DFN ��� ������.', mtInformation, [mbOK], 0);
      Result := True;
    end;
  finally
    odOpenFileDialog.Free;
  end;
end;

//------------------------------------------------------------------------------

class function TController.OpenFileByName(Sender: TObject; const Key: Char): Boolean;
const
  ENTER_KEY_CODE = #13;

begin
  Result := False;

  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(TEdit(Sender).Text) = false then
    begin
      MessageDlg('���� �� ����������: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
      Exit;
    end;

  MessageDlg('���� DFN ��� ������: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
  Result := True;
end;

//------------------------------------------------------------------------------

class function TController.CreateTSV(const OpenedDFNFilePath: string; var SavedTSVFilePath: string): Boolean;
var
   sdSaveTSV: TSaveDialog;
begin
  Result := false; // ��������� ��������, �� ���� �� ��� ����������

  if FileExists(OpenedDFNFilePath) = false then
  begin
    MessageDlg('������� �������� DFN ����', mtInformation, [mbOK], 0);
    Exit;
  end;

  sdSaveTSV := TSaveDialog.Create(nil);
  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // �������� �����������
                                  // ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
  if sdSaveTSV.Execute = false then  // �������� �� ���� ��������� ������ ���������� ����� � ���������� ���
    Exit;

  SavedTSVFilePath := TModel.ConvertDFNToTSV(OpenedDFNFilePath, sdSaveTSV);
  Result := true; // ��������� ��������, �� ���� ��� ����������
  MessageDlg('���� ��� ������� �������������� � �������� � ������� TSV.' + SavedTSVFilePath,
    mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

class function TController.RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
begin
   if (OpenedTSVFilePath = '') and (OpenedDFNToModFilePath = '') then
  begin
    MessageDlg('�������� TSV � DFN ����� � ������������ �������!', mtInformation, [mbOK], 0); // ��������� ����������, ���� �� ����
    Exit;                                                    // ������� ������ �����, ��� ������ �� �����
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // �������� ����������� ���� �������� �����
  begin
    MessageDlg('����������, �������� TSV � DFN ����� � ������������ �������.', mtInformation, [mbOK], 0);
    Exit; // ���������� ���������, � ���, ���� ����� �������� ����� �� ����������
  end;

  TModel.ModifyDFN(OpenedTSVFilePath, OpenedDFNToModFilePath);
  MessageDlg('���� DFN ��� �����������.', mtInformation, [mbOK], 0);
end;


class procedure TController.OpenExcel(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
var
  ExcelPath: string;
  odOpenExcel: TOpenDialog;

begin
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe'; // ����������� ���� �� Excel

  if IsFileCreate = false then // ��������, �� ��� ��������� TSV ����
  begin
    MessageDlg('������� �������� TSV ����', mtInformation, [mbOK], 0); // ��������� �����������, ���� �� ���� �������� TSV ����
    Exit;
  end;

  if FileExists(ExcelPath) then // ���� ���� �������� Excel
  begin
    ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW); // ³�������
    Exit;
  end;                                                              // ���������� ����� � ������� Excel

  ShowMessage('Excel �� ��������. ������� ���� Excel ����� �����.'); // ���� �� ���� �������� Excel, �� ����� ������ ��������� �������� Excel
  odOpenExcel := TOpenDialog.Create(nil);

  if odOpenExcel.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ���
    ShellExecute(0, 'open', PChar(odOpenExcel.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
    // ³������� ���������� ����� � ������� Excel � ����������� �������� ����� ��� �������� Excel
end;


class procedure TController.OpenCalc(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
var
  CalcPath: string;
  odOpenCalc: TOpenDialog;

begin
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe'; // ����������� ���� �� Calc

  if IsFileCreate then // // ��������, �� ��� ��������� TSV ����
  begin
    MessageDlg('������� �������� TSV ����', mtInformation, [mbOK], 0); // ��������� �����������, ���� �� ���� �������� TSV ����
    Exit;
  end;

  if FileExists(CalcPath) then // // ���� ���� �������� Calc
  begin
    ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // ³�������
  end;                                                               // ���������� ����� � ������� Calc

  ShowMessage('Calc �� ��������. ������� ���� Calc ����� �����.'); // ���� �� ���� �������� Calc, �� ����� ������ ��������� �������� Calc
  odOpenCalc := TOpenDialog.Create(nil);

  if odOpenCalc.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ���
    ShellExecute(0, 'open', PChar(odOpenCalc.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
end;   // ³������� ���������� ����� � ������� Calc � ����������� �������� ����� ��� �������� Calc

end.
