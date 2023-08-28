unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs,
  Vcl.Buttons, ShellAPI, Vcl.ExtCtrls, FileOpener;

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
    sdSaveDFNToMod: TSaveDialog;
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
    procedure FormDestroy(Sender: TObject);
    procedure bCreateTSVClick(Sender: TObject);
    procedure bOpenExcelClick(Sender: TObject);
    procedure bOpenCalcClick(Sender: TObject);
    procedure ePathTSVKeyPress(Sender: TObject; var Key: Char);
    procedure ePathDFNToModKeyPress(Sender: TObject; var Key: Char);
    procedure bWriteDFNClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    OriginalDFNFileContent, OriginalTSVFileContent, OriginalDFNToModFileContent: TStringList; // �����, �� ������ �������� ������� ��������� ����� �� ������� ����������� ��� ��������� �����
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFNToModFilePath: String; // �����, �� ������ �������� ���� �������� �����
    SavedTSVFilePath: String; // ����� ��� ��������� ����� ����������� TSV ����� ��� ��������� ������� ���� � ��������� ���������
    IsFileCreate: boolean; // ����� ��� ��������, �� ��� ���������� TSV ����, ��� �������� ���� � ��������� ���������
    ShowWarningMessage: string;
    function ContainsCyrillicCharacters(const input: string): Boolean;
    function IsCharCyrillic(c: Char): Boolean;
    function ExtractField(const fields: TArray<string>; const fieldName: string): string;
  public
    { Public declarations }

  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.bCreateTSVClick(Sender: TObject);
var
  OutputFileContent: TStringList; // ����� ��� ��������� ������������� �������� �����

begin
  if FileExists(OpenedDFNFilePath) = false then
  begin
    ShowMessage('������� �������� DFN ����');
    Exit;
  end;

  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // �������� �����������
                                  // ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
  if sdSaveTSV.Execute = false then  // �������� �� ���� ��������� ������ ���������� ����� � ���������� ����
    Exit;

  try
    OriginalDFNFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
    OriginalDFNFileContent.LoadFromFile(OpenedDFNFilePath); // �������� ������� � �����
    OutputFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
    OutputFileContent.Add('ID' + #9 + '��������' + #9 + '�������'); // ������ ������ ����� � �����,
                                                     // ��� ������� �� ������������ ������� �����
    for var i := 0 to OriginalDFNFileContent.Count - 1 do // ���������� ����� � ����
    begin
      var line := OriginalDFNFileContent[i]; // ��������� ����� ��� ���� ������ ��� �����

      if ContainsCyrillicCharacters(line) and  // ���������� �� ��������� �������� ������� �� ��������� ��������
      (Pos('ID', line) > 0) and (Pos('Orig', line) > 0) and (Pos('Curr', line) > 0) then // �����
      begin     // ���������� �������� �� ������� � ����� ������ ����� ���� 'ID', 'Orig', 'Curr'
        var fields := line.Split([#4]); // ��������� ����� �� ����, ���� ��������� �������� EOT
          var id := ExtractField(fields, 'ID'); // �� ��������� �������� �������, � ��� �������� ���� ��
          var orig := ExtractField(fields, 'Orig'); // ������������� 'ID'/'Orig'/'Curr', �������� ����
          var curr := ExtractField(fields, 'Curr'); // ���������� ���� ��� ��������� �����, ���� ��� ��������
        OutputFileContent.Add(id + #9 + orig + #9 + curr);
      end;
    end;

    if Pos('.tsv', sdSaveTSV.FileName) > 0 then // ��������, �� � � ����� ���������� ����� ���������� tsv
      OutputFileContent.SaveToFile(sdSaveTSV.FileName) // ����� ������������� �������� � ��������� ����
    else
      OutputFileContent.SaveToFile(sdSaveTSV.FileName + '.tsv'); // ����� ������������� ��������
                                                    // � ��������� ���� � ������������ �������
    SavedTSVFilePath := sdSaveTSV.FileName; // �������� ���� ���������� ����� � ��������� �����,
                                // ��������� ��� ���������� �������� ����� � ��������� ���������
    IsFileCreate := true; // ϳ�����������, �� TSV ���� ��� ����������, ��� ���������� ������������ � ��������� ���������
    ShowMessage('���� ��� ������� �������������� � �������� � ������� TSV.' + sdSaveTSV.FileName);
  finally
    FreeAndNil(OriginalDFNFileContent); // ��������� ���� ��ﳿ ��������� �����
    FreeAndNil(OutputFileContent); // ��������� ���� � ������������ ���������
  end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bWriteDFNClick(Sender: TObject);
var
  tsvIDField, dfnIDField: string;

begin
  if (OpenedTSVFilePath = '') and (OpenedDFNToModFilePath = '') then
  begin
    ShowMessage('�������� TSV � DFN ����� � ������������ �������!'); // ��������� ����������, ���� �� ����
    Exit;                                                    // ������� ������ �����, ��� ������ �� �����
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // �������� ����������� ���� �������� �����
  begin
    ShowMessage('����������, �������� TSV � DFN ����� � ������������ �������.');
    Exit; // ���������� ���������, � ���, ���� ����� �������� ����� �� ����������
  end;

  try
    OriginalTSVFileContent := TStringList.Create;
    OriginalDFNToModFileContent := TStringList.Create;
    OriginalTSVFileContent.LoadFromFile(OpenedTSVFilePath); // �������� ������� � �����
    OriginalDFNToModFileContent.LoadFromFile(OpenedDFNToModFilePath); // �������� ������� � �����

    for var TSVLine := 0 to OriginalTSVFileContent.Count - 1 do // ���������� ����� � ���� TSV
    begin
      tsvIDField := OriginalTSVFileContent[TSVLine].Split([#9])[0]; // �������� � ����� ����� ���� � ����� TSV �����
      for var DFNLine := 0 to OriginalDFNToModFileContent.Count - 1 do // ���������� ����� � ���� DFN
      begin
        dfnIDField := ExtractField(OriginalDFNToModFileContent[DFNLine].Split([#4]), 'ID'); // �������� ���� "ID" DFN �����
        if Pos(tsvIDField, dfnIDField) = 1 then // ��������, �� ���� ���� ID � DFN ���� ��������� ������� ���� TSV �����
        begin // ���� ���� �������� ���������, �� ������ ���� "Curr"
          var OriginalDFNLine := OriginalDFNToModFileContent[DFNLine]; // ���� ������� ����� DFN ����� ���
                                                                      // �������� ����� � ����� ���� Curr
          for var CurrField := 0 to Length(OriginalDFNToModFileContent[DFNLine].Split([#4])) - 1 do // ����������
          begin  // ���� ����� � ����� ���� �������� ��������� ���� "ID" ��� ����������� ������ ���� � �����
                  // �������� ����� ���� "Curr", ����������� ���� ��� ����� ��� ���������� ���������� ����
            if Pos('Curr:', OriginalDFNToModFileContent[DFNLine].Split([#4])[CurrField]) >= 1 then // ������������
            begin                                                               // �������� ���� �� ���� 'Curr:'
              var DFNCurr := OriginalDFNToModFileContent[DFNLine].Split([#4])[CurrField]; // �������� ���� Curr DFN �����
              var TSVCurr := OriginalTSVFileContent[TSVLine].Split([#9])[2]; // �������� ���� Curr TSV �����
              OriginalDFNLine := OriginalDFNLine.Replace(DFNCurr, 'Curr:' + QuotedStr(TSVCurr)); // � ����� � ����������
              // ����� "ID" �������� ���� ���� "Curr", ���������� ����� ����� ������� CurrField, �� 3 ���� � ����� TSV
              OriginalDFNToModFileContent[DFNLine] := OriginalDFNLine; // ����� ������������ ����� ������������ ��ﳺ� �����
              Break; // ����� � �����, ���� �������� �������� ����
            end;
          end;
          Break; // ����� � �����, ���� �������� ��������� ���� ID, ���������� �� ���������� ����� � ���� TSV
        end;
      end;
    end;
    OriginalDFNToModFileContent.SaveToFile(OpenedDFNToModFilePath);
    // �������� ���� � ���� DFN, ������������� ��������� ����� � ���������� ������ ��������� �����,
    // ��������� � �������, ���� ���� ��� �������� �� ��������� ���������� ����
  finally
    FreeAndNil(OriginalTSVFileContent);
    FreeAndNil(OriginalDFNToModFileContent);
  end;
  ShowMessage('���� DFN ��� �����������.');
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNClick(Sender: TObject);
begin
  if TFileOpener.ChooseFile(OpenedDFNFilePath, 'DFN Files|*.dfn') then
    ePathDFN.Text := OpenedDFNFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNKeyPress(Sender: TObject; var Key: Char);
begin
  if TFileOpener.OpenFile(Sender, Key) then
    OpenedDFNFilePath := ePathDFN.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseTSVClick(Sender: TObject);
begin
  if TFileOpener.ChooseFile(OpenedTSVFilePath, 'TSV Files|*.tsv') then
    ePathTSV.Text := OpenedTSVFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathTSVKeyPress(Sender: TObject; var Key: Char);

begin
  if TFileOpener.OpenFile(Sender, Key) then
    OpenedTSVFilePath := ePathTSV.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNToModClick(Sender: TObject);
begin
  if TFileOpener.ChooseFile(OpenedDFNToModFilePath, 'DFN Files|*.dfn') then
    ePathDFNToMod.Text := OpenedDFNToModFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNToModKeyPress(Sender: TObject; var Key: Char);
begin
  if TFileOpener.OpenFile(Sender, Key) then
    OpenedDFNToModFilePath := ePathDFNToMod.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenExcelClick(Sender: TObject);
var
  ExcelPath: string;

begin
  ShowWarningMessage := '';
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe'; // ����������� ���� �� Excel
  if IsFileCreate then // ��������, �� ��� ��������� TSV ����
  begin
    if FileExists(ExcelPath) then // ���� ���� �������� Excel
      ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // ³�������
    else                                                               // ���������� ����� � ������� Excel
    begin // ���� �� ���� �������� Excel, �� ����� ������ ��������� �������� Excel
      ShowMessage('Excel �� ��������. ������� ���� Excel ����� �����.');
      if odOpenExcel.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
      begin
        ShellExecute(0, 'open', PChar(odOpenExcel.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;        // ³������� ���������� ����� � ������� Excel � ����������� �������� ����� ��� �������� Excel
    end;
  end
  else
    begin
      ShowMessage('������� �������� TSV ����'); // ��������� �����������, ���� �� ���� �������� TSV ����
      ShowWarningMessage := '�������� ������� TSV ����'; // �������� ����������� � ��������� ����� ��� �����
    end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenCalcClick(Sender: TObject);
var
  CalcPath: string;

begin
  ShowWarningMessage := '';
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe'; // ����������� ���� �� Calc
  if IsFileCreate then // // ��������, �� ��� ��������� TSV ����
  begin
    if FileExists(CalcPath) then // // ���� ���� �������� Calc
      ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // ³�������
    else                                                               // ���������� ����� � ������� Calc
    begin // ���� �� ���� �������� Calc, �� ����� ������ ��������� �������� Calc
      ShowMessage('Calc �� ��������. ������� ���� Calc ����� �����.');
      if odOpenCalc.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
      begin
        ShellExecute(0, 'open', PChar(odOpenCalc.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;          // ³������� ���������� ����� � ������� Calc � ����������� �������� ����� ��� �������� Calc
    end;
  end
  else
    begin
      ShowMessage('������� �������� TSV ����'); // ��������� �����������, ���� �� ���� �������� TSV ����
      ShowWarningMessage := '�������� ������� TSV ����'; // �������� ����������� � ��������� ����� ��� �����
    end;
end;

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

//------------------------------------------------------------------------------

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  OriginalTSVFileContent.Free;
  OriginalDFNToModFileContent.Free;
end;

//------------------------------------------------------------------------------

function TMainForm.ContainsCyrillicCharacters(const input: string): Boolean;
var i: Integer;                   // ������� ��� �������� �� ������ ����� ��������
begin
  Result := False; // �������� � ��������� �������� false, ���� �� ���� �������� ��������
  for i := 1 to Length(input) do // ���������� ����� � ���������� �����
  begin
    if IsCharCyrillic(input[i]) then // ������������� �� ���� ������� ��� �������� �� � ����� ���������
    begin
      Result := True; // ���� ���� �������� ����������� ������� ��������, ����������� � ��������� IsCharCyrillic
      Exit; // ��������� ��������� ������� � ����������� true
    end;
  end;
end;

//------------------------------------------------------------------------------

function TMainForm.IsCharCyrillic(c: Char): Boolean; // ������� ��� �������� �� � ����� ���������
begin
  Result := (c >= '�') and (c <= '�') or (c >= '�') and (c <= '�') or (c = '�') or (c = '�'); // ������� true
end;                                                 // ���� ����� "�" ������� � ������� ���������� �������

//------------------------------------------------------------------------------

function TMainForm.ExtractField(const fields: TArray<string>; const fieldName: string): string;
// ������� ��� ����������� ����� ���� �� ���� ������
var fieldValue: string;

begin
  for var field in fields do // ���������� ���� � �����
  begin
    if Pos(fieldName, field) > 0 then // ��������, �� �������� � ��� ������ ����� ����
    begin
      fieldValue := (Copy(field, Pos(fieldName, field) + Length(fieldName) + 1, MaxInt));  // ���� ��������, �� ������� ���� ����,
                   // ����� ������� ���� ����, ��� ����� ���� �� ������� ':', �� ���� ���� �� ���� ����
      if (Length(fieldValue) >= 2) and (fieldValue[1] = '''') and (fieldValue[Length(fieldValue)] = '''') then
      // �������� �� ������ ���� ���� � ���� �� �� ������� �������� �����
        Result := Copy(fieldValue, 2, Length(fieldValue) - 2) // ���� ������, �� �������
      else
        Result := fieldValue; // ���� ��, �� ������ �� ���������
      Break; // �������� � ����� ��� ����������� �������� ����� ����
    end
    else
      Result := ''; // ���� �� ���� �������� ��'� ����
  end;
end;

end.
