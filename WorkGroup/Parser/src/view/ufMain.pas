unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs,
  Vcl.Buttons, ShellAPI, Vcl.ExtCtrls, Parsing;

type
  TMainForm = class(TForm)
    ePathDFN: TEdit;
    bCreateTSV: TButton;
    bOpenExcel: TButton;
    bOpenCalc: TButton;
    odOpenDFN: TOpenDialog;
    odOpenTSV: TOpenDialog;
    odOpenDFNToMod: TOpenDialog;
    spChooseDFN: TSpeedButton;
    lChoiceDFNForConversion: TLabel;
    sdSaveTSV: TSaveDialog;
    odOpenExcel: TOpenDialog;
    odOpenCalc: TOpenDialog;
    pCreateTSV: TPanel;
    lChoiceTSVForRewriting: TLabel;
    lChoiceDFNToMod: TLabel;
    spChooseDFNToMod: TSpeedButton;
    spChooseTSV: TSpeedButton;
    bWriteDFN: TButton;
    ePathDFNToMod: TEdit;
    ePathTSV: TEdit;
    pRewriteDFN: TPanel;
    lPanelTitleCreateTSV: TLabel;
    lPanelTitleRewriteDFN: TLabel;
    procedure ePathDFNKeyPress(Sender: TObject; var Key: Char);
    procedure spChooseDFNClick(Sender: TObject);
    procedure spChooseTSVClick(Sender: TObject);
    procedure spChooseDFNToModClick(Sender: TObject);
    procedure bCreateTSVClick(Sender: TObject);
    procedure bOpenExcelClick(Sender: TObject);
    procedure bOpenCalcClick(Sender: TObject);
    procedure ePathTSVKeyPress(Sender: TObject; var Key: Char);
    procedure ePathDFNToModKeyPress(Sender: TObject; var Key: Char);
    procedure bWriteDFNClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ApplyFilePath(const filePath : String; var GlobalFilePath : String; const fileType : String);
    function OpenByProgram(const programName: string): Boolean;
  private
    { Private declarations }
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFNToModFilePath: string; // �����, �� ������ �������� ���� �������� �����
    SavedTSVFilePath: string; // ����� ��� ��������� ����� ���������� TSV ����� ��� ��������� ������� ���� � ��������� ���������
    IsFileCreate: Boolean; // ����� ��� ��������, �� ��� ���������� TSV ����, ��� �������� ���� � ��������� ���������
    const ENTER_KEY_CODE = #13;
    const STR_OPEN_FILE_INFORMATION = '���� "%s" ��� ������';
    const STR_DFN = 'DFN';
    const STR_TSV = 'TSV';
    const STR_FILE_NOT_EXIST = '���� �� ����������: ';
    const STR_CHOOSE_DFN_FILE = '������� �������� "DFN" ����';
    const STR_TSV_EXTENSION = '.tsv';
    const STR_SUCCESSFUL_CREATION_OF_TSV_FILE = '���� ��� ������� �������������� � �������� � ������� TSV';
    const STR_FIRST_OPEN_TSV_AND_DFN_FILES = '������ �������� TSV � DFN ����� � ������������ �������!';
    const STR_OPEN_FILES_WITH_MATCHING_NAME = '����������, �������� TSV � DFN ����� � ������������ �������.';
    const SRT_SUCCESSFUL_REWRITE_OF_DFN_FILE = '���� DFN ��� �����������.';
    const STR_DEFAULT_EXCEL_PATH = 'C:\Program Files\Microsoft Office\root\Office16\excel.exe';
    const STR_DEFAULT_CALC_PATH = 'C:\Program Files\LibreOffice\program\soffice.exe';
    const STR_TSV_FILE_NOT_CREATED = '������� �������� TSV ����';
    const STR_OPEN = 'open';
    const STR_EXCEL = 'Excel';
    const STR_CALC = 'Calc';
    const STR_FIND_THE_APPLICATION = '%s �� �������. �������� ����������� ���� ��������� %s ����� ���������� ����';
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.bCreateTSVClick(Sender: TObject);
var
  OutputFileContent: TStringList;

begin
  if FileExists(OpenedDFNFilePath) = false then
  begin
    MessageDlg(STR_CHOOSE_DFN_FILE, mtInformation, [mbOK], 0);
    Exit;
  end;

  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), STR_TSV_EXTENSION); // ��������
                  // ����������� ��'� ����� ��� ���������, ����� ��'� ��������� ����� �� ������ �������
  if sdSaveTSV.Execute = false then
    Exit;

  try
    OutputFileContent := TStringList.Create;
    IsFileCreate := TParsing.ConvertDFNToTSV(OpenedDFNFilePath, OutputFileContent); // ��������� TSV ����

    if Pos(STR_TSV_EXTENSION, sdSaveTSV.FileName) = 0 then // ��������, �� � � ����� ���������� �����
    begin                                                                             // ���������� .tsv
      OutputFileContent.SaveToFile(sdSaveTSV.FileName + STR_TSV_EXTENSION); //��������� ����� � ������������
      SavedTSVFilePath := sdSaveTSV.FileName + STR_TSV_EXTENSION;                          // ���������� .tsv
      Exit;
    end;

    OutputFileContent.SaveToFile(sdSaveTSV.FileName); //��������� ����� ��� ����������� ���������� .tsv
    SavedTSVFilePath := sdSaveTSV.FileName;

  finally
    FreeAndNil(OutputFileContent);
    MessageDlg(STR_SUCCESSFUL_CREATION_OF_TSV_FILE, mtInformation, [mbOK], 0);
  end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bWriteDFNClick(Sender: TObject);
begin
  if (OpenedTSVFilePath = '') or (OpenedDFNToModFilePath = '') then
  begin
    MessageDlg(STR_FIRST_OPEN_TSV_AND_DFN_FILES, mtInformation, [mbOK], 0); // ��������� ����������, ���� �� ����
    Exit;                                                           // ������� ������ �����, ��� ������ �� �����
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // �������� ����������� ���� �������� �����
  begin
    MessageDlg(STR_OPEN_FILES_WITH_MATCHING_NAME, mtInformation, [mbOK], 0);
    Exit; // ���������� ���������, � ���, ���� ����� �������� ����� �� ����������
  end;

  TParsing.ModifyDFN(OpenedTSVFilePath, OpenedDFNToModFilePath); // ��������� ��������� DFN �����
  MessageDlg(SRT_SUCCESSFUL_REWRITE_OF_DFN_FILE, mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNClick(Sender: TObject);
begin
  if odOpenDFN.Execute = false then
    Exit;

  ePathDFN.Text := odOpenDFN.FileName;
  ApplyFilePath(ePathDFN.Text, OpenedDFNFilePath, STR_DFN);
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(ePathDFN.Text) = false then
  begin
    MessageDlg(STR_FILE_NOT_EXIST + ePathDFN.Text, mtInformation, [mbOK], 0);
    Exit;
  end;

  ApplyFilePath(ePathDFN.Text, OpenedDFNFilePath, STR_DFN);
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseTSVClick(Sender: TObject);
begin
  if odOpenTSV.Execute = false then
    Exit;

  ePathTSV.Text := odOpenTSV.FileName;
  ApplyFilePath(ePathTSV.Text, OpenedTSVFilePath, STR_TSV);
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathTSVKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(ePathTSV.Text) = false then
  begin
    MessageDlg(STR_FILE_NOT_EXIST + ePathTSV.Text, mtInformation, [mbOK], 0);
    Exit;
  end;

  ApplyFilePath(ePathTSV.Text, OpenedTSVFilePath, STR_TSV);
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNToModClick(Sender: TObject);
begin
  if odOpenDFNToMod.Execute = false then
    Exit;

  ePathDFNToMod.Text := odOpenDFNToMod.FileName;
  ApplyFilePath(ePathDFNToMod.Text, OpenedDFNToModFilePath, STR_DFN);
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNToModKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(ePathDFNToMod.Text) = false then
  begin
    MessageDlg(STR_FILE_NOT_EXIST + ePathDFNToMod.Text, mtInformation, [mbOK], 0);
    Exit;
  end;

  ApplyFilePath(ePathDFNToMod.Text, OpenedDFNToModFilePath, STR_DFN);
end;

//------------------------------------------------------------------------------

/// <summary>
/// ��������� ���� �� ����� �� ������ ���� � ���������� ������.
/// </summary>
/// <param name="filePath">���� �� �����, ���� ������� �����������.</param>
/// <param name="globalFilePath">��������� �����, � ��� ���� ��������� ���� �� �����.</param>
/// <param name="fileType">��� ����� (���������, "TSV" ��� "DFN").</param>
procedure TMainForm.ApplyFilePath(const filePath : String; var globalFilePath : String; const fileType : String);
begin
   globalFilePath := filePath;
   MessageDlg(Format(STR_OPEN_FILE_INFORMATION, [fileType]), mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenExcelClick(Sender: TObject);
begin
  if OpenByProgram(STR_EXCEL) then
    Exit;

  // ���� �� ���� �������� ��������, �� ����� ������ ��������� �������� ����� �������� ����
  if odOpenExcel.Execute then
    ShellExecute(0, STR_OPEN, PChar(odOpenExcel.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenCalcClick(Sender: TObject);
begin
  if OpenByProgram(STR_CALC) then
    Exit;

  // ���� �� ���� �������� ��������, �� ����� ������ ��������� �������� ����� �������� ����
  if odOpenCalc.Execute then
    ShellExecute(0, STR_OPEN, PChar(odOpenCalc.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
end;

//------------------------------------------------------------------------------

/// <summary>
/// ³������ ���� � �������� ������� �� �� ������ �� ������ �� �����.
/// </summary>
/// <param name="programName">����� ��������, � ��� ������� ������� ���� (���������, "Excel" ��� "Calc").</param>
/// <returns>True, ���� ���� ��� ������ �������; False, ���� ������� ����� ������� ��������� ����������� ����.</returns>
function TMainForm.OpenByProgram(const programName: String): Boolean;
var
  programPath: String;

begin
  Result := true; // ��� ������� ����������� ��������, ������, �� ��������� �� ���������� ��������� ���������
  if (IsFileCreate = false) and (FileExists(SavedTSVFilePath) = false) then // ��������, �� ��� ��������� TSV ����
  begin                                                                                        // �� �� ���� ����
    MessageDlg(STR_TSV_FILE_NOT_CREATED, mtInformation, [mbOK], 0); // �����������, ���� �� ���� �������� TSV ����
    Exit;
  end;

  programPath := STR_DEFAULT_EXCEL_PATH; // ������������ ����������� ���� �� �������� Excel

  if Pos(programName, STR_CALC) > 0 then  // ���� � ���������� �������� ���� ������� ����� �������� Calc
    programPath := STR_DEFAULT_CALC_PATH; // ��� ������������ ����������� ���� �� �������� Calc

  if FileExists(programPath) then // ���� ���� �������� ����������� ���� ������� ��������
  begin
    ShellExecute(0, STR_OPEN, PChar(programPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
    Exit;
  end;

  Result := false; // ��� ������� ����������� ��������, ���� ��������, �� ��������� �� ���� ��������� ������
                                                                         // ����������� ���� ������� ��������
  MessageDlg(Format(STR_FIND_THE_APPLICATION, [ProgramName, ProgramName]), mtInformation, [mbOK], 0);
end;                                 // �����������, ���� �� ���� �������� ����������� ���� ��������

//------------------------------------------------------------------------------

procedure TMainForm.FormShow(Sender: TObject);
begin
  ePathDFN.SetFocus; // ������������ ������ �� ����� ���� ����� �����, ��� �������� �������
end;

//------------------------------------------------------------------------------

procedure TMainForm.FormCreate(Sender: TObject);
begin
  IsFileCreate := false; // ��� ��������� ����� ������������ ���������� ������, �� ������� �� ��������
end;                                                            // �� ��� ��������� TSV ����, �������� false

end.
