unit Controller;

interface

uses
  Winapi.Windows, Vcl.Dialogs, Classes, SysUtils, Vcl.StdCtrls, ShellAPI, Model;

type
  TController = class
  private
    const DFN_NAME: string = 'DFN';
    const TSV_NAME: string = 'TSV';
  public
    class function ChooseFile(var odOpenFileDialog: TOpenDialog; const FileExtension: string): string;
    class function OpenFileByName(Sender: TObject; const Key: Char): Boolean;
    class function CreateTSV(const OpenedDFNFilePath: string; var SavedTSVFilePath: string): Boolean;
    class function RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
    class function OpenByProgram(const SavedTSVFilePath: string; const IsFileCreate: Boolean; const ProgramName: string): Boolean;
  end;

implementation

/// <summary>
/// ����� ���� � �������� �����������.
/// </summary>
/// <param name="odOpenFileDialog">ĳ������� ���� �������� �����.</param>
/// <param name="FileExtension">����� � �������� ����� ��� ���������� ����.</param>
/// <returns>���� �� �������� ����� ��� �������� �����, ���� ���� �� ��� �������.</returns>
class function TController.ChooseFile(var odOpenFileDialog: TOpenDialog; const FileExtension: string): string;
var
  TypeOpenFile: string;

begin
  Result := '';
  TypeOpenFile := DFN_NAME;

  if Pos(FileExtension, 'TSV Files|*.tsv') = 1 then
    TypeOpenFile := TSV_NAME;

  odOpenFileDialog.Filter := FileExtension;
  if odOpenFileDialog.Execute then
  begin
    Result := odOpenFileDialog.FileName;
    MessageDlg('���� ' + TypeOpenFile + ' ��� ������.', mtInformation, [mbOK], 0);
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// ³������ ���� �� ������, ���� ��������� ������� ������.
/// </summary>
/// <param name="Sender">��'���, ���� ��������� ����, � ������ ������� TEdit.</param>
/// <param name="Key">������, ��� ���� ���������.</param>
/// <returns>������� True, ���� ���� ���� ������ �������; � ������ ������� - False.</returns>
class function TController.OpenFileByName(Sender: TObject; const Key: Char): Boolean;
const
  ENTER_KEY_CODE = #13;

var
  TypeOpenFile: string;

begin
  Result := False;

  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(TEdit(Sender).Text) = false then
    begin
      MessageDlg('���� �� ����������: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
      Exit;
    end;

  TypeOpenFile := DFN_NAME;

  if Pos(TEdit(Sender).Text, '.tsv') > 0 then
    TypeOpenFile := TSV_NAME;

  MessageDlg('���� ' + TypeOpenFile + ' ��� ������: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
  Result := True;
end;

//------------------------------------------------------------------------------

/// <summary>
/// ������� ���� � ������ TSV � ����� ��������� DFN ����� �� ������ ���� �� ������� ������.
/// </summary>
/// <param name="OpenedDFNFilePath">���� �� ��������� DFN �����.</param>
/// <param name="SavedTSVFilePath">�����, � ��� ���� ���������� ���� �� ����������� TSV �����.</param>
/// <returns>������� True, ���� ���� ������ ��� ��������� �� ����������; � ������ ������� - False.</returns>
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
  sdSaveTSV.Filter := 'TSV Files|*.tsv';
  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // �������� �����������
                                  // ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
  if sdSaveTSV.Execute = false then  // ��������� �� ���������� �� ���� ��������� ������ ���������� �����
    Exit;                                                                            // � ���������� ����

  SavedTSVFilePath := TModel.ConvertDFNToTSV(OpenedDFNFilePath, sdSaveTSV); //��������� ������������ TSV ����
  Result := true; // ��������� ��������, �� ���� ��� ����������
  MessageDlg('���� ��� ������� �������������� � �������� � ������� TSV.', mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

/// <summary>
/// ���������� ���� DFN � ������������� ����� � ����� TSV.
/// </summary>
/// <param name="OpenedTSVFilePath">���� �� ��������� TSV �����.</param>
/// <param name="OpenedDFNToModFilePath">���� �� ��������� DFN ����� ��� �����������.</param>
/// <returns>True, ���� ���� DFN ��� ������ �������������, ������ False.</returns>
class function TController.RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
begin
  Result := false;

   if (OpenedTSVFilePath = '') and (OpenedDFNToModFilePath = '') then
  begin
    MessageDlg('�������� TSV � DFN ����� � ������������ �������!', mtInformation, [mbOK], 0); // ���������
    Exit;                         // ����������, ���� �� ���� ������� ������ �����, ��� ������ �� �����
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // �������� ����������� ���� �������� �����
  begin
    MessageDlg('����������, �������� TSV � DFN ����� � ������������ �������.', mtInformation, [mbOK], 0);
    Exit; // ���������� ���������, � ���, ���� ����� �������� ����� �� ����������
  end;

  TModel.ModifyDFN(OpenedTSVFilePath, OpenedDFNToModFilePath);
  Result := true;
  MessageDlg('���� DFN ��� �����������.', mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

/// <summary>
/// ³������ TSV ���� � ��������� ������� ��� �������� ������� �������� ��� ��������.
/// </summary>
/// <param name="SavedTSVFilePath">���� �� ����������� TSV �����.</param>
/// <param name="IsFileCreate">���������, �� ����� �� ��������� ���������� TSV �����.</param>
/// <param name="ProgramName">����� ��������, � ��� ������� ������� ��������� TSV ����.</param>
/// <returns>������� True, ���� TSV ���� ��� ������ ��������; � ������ ������� - False.</returns>
class function TController.OpenByProgram(const SavedTSVFilePath: string;
  const IsFileCreate: Boolean; const ProgramName: string): Boolean;
var
  ExcelPath, CalcPath, ProgramPath: string;
  odOpenProgram: TOpenDialog;

begin
  Result := false;
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe'; // ����������� ���� �� Excel
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe'; // ����������� ���� �� Calc
  ProgramPath := ExcelPath; // ������������ �� ������������� ����������� ���� �� �������� Excel

  if (IsFileCreate = false) or (FileExists(SavedTSVFilePath) = false) then // ��������, �� ��� ��������� TSV ����
  begin                                                                                        // �� �� ���� ����
    MessageDlg('������� �������� TSV ����', mtInformation, [mbOK], 0); // ��������� �����������, ���� �� ����
    Exit;                                                                                 // �������� TSV ����
  end;

  if Pos(ProgramName, 'Calc') > 0 then  // ���� � ���������� �������� ���� ������� ����� �������� Calc
    ProgramPath := CalcPath; // ��� ������������ ����������� ���� �� �������� Calc

  if FileExists(ProgramPath) then // ���� ���� �������� ����������� ���� ������� ��������
  begin
    ShellExecute(0, 'open', PChar(ProgramPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW); // ³�������
    Result := true;                                                           // ���������� ����� � �������
    Exit;
  end;

  MessageDlg(ProgramName + ' �� �������. �������� ���� ' + ProgramName + ' ����� ���������� ����.',
    mtInformation, [mbOK], 0);
                                                                            // ���� �� ���� �������� ��������,
  odOpenProgram := TOpenDialog.Create(nil);     // �� ����� ������ ��������� �������� ����� �������� ����
  try
    if odOpenProgram.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
      begin
        ShellExecute(0, 'open', PChar(odOpenProgram.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
        Result := true;          // ³������� ���������� ����� � ������� � ����������� �������� ����� ��� ��������
      end;

  finally
    FreeAndNil(odOpenProgram);
  end;
end;

end.
