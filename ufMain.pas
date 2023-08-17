unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs, Vcl.Buttons, ShellAPI;

type
  TForm1 = class(TForm)
    ePath1: TEdit;
    bCreateTSV: TButton;
    ePath2: TEdit;
    ePath3: TEdit;
    bOpenExcel: TButton;
    bOpenCalc: TButton;
    bWriteDFN: TButton;
    OpenDialog1: TOpenDialog;
    OpenDialog2: TOpenDialog;
    OpenDialog3: TOpenDialog;
    spChooseDFN: TSpeedButton;
    Label1: TLabel;
    spChooseTSV: TSpeedButton;
    Label2: TLabel;
    spChooseDFN2: TSpeedButton;
    Label3: TLabel;
    SaveDialog1: TSaveDialog;
    OpenDialog4: TOpenDialog;
    OpenDialog5: TOpenDialog;
    procedure ePath1KeyPress(Sender: TObject; var Key: Char);
    procedure spChooseDFNClick(Sender: TObject);
    procedure spChooseTSVClick(Sender: TObject);
    procedure spChooseDFN2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bCreateTSVClick(Sender: TObject);
    procedure bOpenExcelClick(Sender: TObject);
    procedure bOpenCalcClick(Sender: TObject);
  private
    { Private declarations }
    function ContainsCyrillicCharacters(const input: string): Boolean;
    function IsCharCyrillic(c: Char): Boolean;
    function ExtractFieldValue(const field: string): string;
    function ExtractField(const fields: TArray<string>; const fieldName: string): string;
  public
    { Public declarations }

  end;

var
  Form1: TForm1;
  OriginalFileContent: TStringList; // ��������� ��������� �����, ��� ���� �������� ������� ��������� ����� �� ������� ����������� ��� ��������� �����
  OpenedFilePath: String; // ��������� ��������� �����, ��� ���� �������� ���� ��������� �����
  SavedFilePath: String; // ��������� ��������� ����� ��� ��������� ����� ����������� tsv ����� ��� ��������� ������� ���� � ��������� ���������
  IsFileCreate: boolean = false;

implementation

{$R *.dfm}

procedure TForm1.bCreateTSVClick(Sender: TObject);

var OutputFileContent: TStringList; // ��������� ����� ��� ��������� ������������� �������� �����

begin
  if OriginalFileContent <> nil then // �������� �� ��� �������� ����
  begin
      SaveDialog1.FileName := ChangeFileExt(ExtractFileName(OpenedFilePath), '.tsv'); // �������� ����������� ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
      if SaveDialog1.Execute then  // �������� �� ���� ��������� ������ ���������� ����� � ���������� ����
      begin
        IsFileCreate := true;
        OutputFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        try
          OutputFileContent.Add('ID' + #9 + '��������' + #9 + '�������'); // ������ ����� ������ � �����, ��� ������� �� ������������ ������� �����
          for var i := 0 to OriginalFileContent.Count - 1 do
        begin
          var line := OriginalFileContent[i]; //
          if ContainsCyrillicCharacters(line) and
          (Pos('ID', line) > 0) and (Pos('Orig', line) > 0) and (Pos('Curr', line) > 0) then //
          begin
            var fields := line.Split([#4]);
            if Length(fields) >= 3 then
            begin
              var id := ExtractFieldValue(ExtractField(fields, 'ID'));
              var orig := ExtractFieldValue(ExtractField(fields, 'Orig'));
              var curr := ExtractFieldValue(ExtractField(fields, 'Curr'));
              OutputFileContent.Add(id + #9 + orig + #9 + curr);
            end;
          end;
        end;
            if Pos('.tsv', SaveDialog1.FileName) > 0 then // �������� �� � � ����� ���������� ����� ���������� tsv
              OutputFileContent.SaveToFile(SaveDialog1.FileName) // ����� ������������� �������� � ��������� ����
            else
              OutputFileContent.SaveToFile(SaveDialog1.FileName + '.tsv'); // ����� ������������� �������� � ��������� ���� � ������������ �������
            SavedFilePath := SaveDialog1.FileName; // �������� ���� ���������� �����
            ShowMessage('���� ��� ������ �������������� �� ���������� � ������ TSV.' + SaveDialog1.FileName);
        finally
          OutputFileContent.Free;
        end;
      end;
  end
  else
    ShowMessage('�������� �������� ���� DFN.');
end;

procedure TForm1.bOpenCalcClick(Sender: TObject);
var
  CalcPath: string;
begin
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe';
  if IsFileCreate then
  begin
    if FileExists(CalcPath) then
    begin
      if IsFileCreate then
      begin
        ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedFilePath+'"'), nil, SW_SHOW);
      end;
    end
    else
    begin
      ShowMessage('Calc �� ��������. ������� ���� Calc ����� �����.');
      if OpenDialog5.Execute then
      begin
        ShellExecute(0, 'open', PChar(OpenDialog5.FileName), PChar('"'+SavedFilePath+'"'), nil, SW_SHOW);
      end;
    end;
  end
  else
    ShowMessage('�������� ������� TSV ����');
end;

procedure TForm1.bOpenExcelClick(Sender: TObject);
var
  ExcelPath: string;
begin
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe';
  if IsFileCreate then
  begin
    if FileExists(ExcelPath) then
    begin
      if IsFileCreate then
      begin
        ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedFilePath+'"'), nil, SW_SHOW);
      end;
    end
    else
    begin
      ShowMessage('Excel �� ��������. ������� ���� Excel ����� �����.');
      if OpenDialog4.Execute then
      begin
        ShellExecute(0, 'open', PChar(OpenDialog4.FileName), PChar('"'+SavedFilePath+'"'), nil, SW_SHOW);
      end;
    end;
  end
  else
    ShowMessage('�������� ������� TSV ����');
end;

function TForm1.ContainsCyrillicCharacters(const input: string): Boolean; // ������� ��� �������� �� � ����� ���������
var i: Integer;
begin
  Result := False;
  for i := 1 to Length(input) do
  begin
    if IsCharCyrillic(input[i]) then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

procedure TForm1.ePath1KeyPress(Sender: TObject; var Key: Char);

var FilePath: string;

begin
  if (Key = #13) then // ���� ���� ��������� ������ #13 - ��� ������ "Enter"
  begin
    FilePath := ePath1.Text; // �������� ����, ���� ��� �������� � �������� ���� "edit"
    if FileExists(FilePath) then // �������� �� ���� ����
    begin
      try
        if OriginalFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
          OriginalFileContent.Free; // ��������� ���������� ����, ���� ���

        OriginalFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        OriginalFileContent.LoadFromFile(FilePath); // �������� ������� � ��������� �����
        OpenedFilePath := FilePath; // �������� ���� ��������� ����� � ��������� �����
        ShowMessage('���� �������: ' + FilePath);
      except
        on E: Exception do
          ShowMessage('������� ��� ������� �����: ' + E.Message);
      end;
    end
    else
    begin
    ShowMessage('���� �� ����: ' + FilePath); // ���� ��� �������� �� ��������� ���� �� �����
    end;
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  OriginalFileContent.Free; // ��������� ���� ������������ �������� � ����� ��� ������� ����������.
end;

procedure TForm1.spChooseDFN2Click(Sender: TObject);
begin
  if OpenDialog3.Execute then
  ePath3.Text := OpenDialog3.FileName;
end;

procedure TForm1.spChooseDFNClick(Sender: TObject);
begin
  if OpenDialog1.Execute then // �������� �� ���� ��������� ������ �������� �������� �����
  begin
    try
      if OriginalFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalFileContent.Free; // ��������� ���������� ����, ���� ���

      ePath1.Text := OpenDialog1.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalFileContent.LoadFromFile(OpenDialog1.FileName); // �������� ������� � ��������� �����
      OpenedFilePath := OpenDialog1.FileName; // �������� ���� ��������� ����� � ��������� �����
      ShowMessage('���� DFN ��� ��������.');
    except
      on E: Exception do
          ShowMessage('������� ��� ������� �����: ' + E.Message);
    end;
  end;
end;

procedure TForm1.spChooseTSVClick(Sender: TObject);
begin
  if OpenDialog2.Execute then
  ePath2.Text := OpenDialog2.FileName;
end;

function TForm1.IsCharCyrillic(c: Char): Boolean; // ������� ��� �������� �� � ����� ���������
begin
  Result := (c >= '�') and (c <= '�') or (c >= '�') and (c <= '�') or (c = '�') or (c = '�');
end;

function TForm1.ExtractFieldValue(const field: string): string;
var
  colonPos: Integer;
  fieldValue: string;
begin
  colonPos := Pos(':', field);
  if colonPos > 0 then
  begin
    fieldValue := Trim(Copy(field, colonPos + 1, MaxInt));
    if (Length(fieldValue) >= 2) and (fieldValue[1] = '''') and (fieldValue[Length(fieldValue)] = '''') then
      Result := Copy(fieldValue, 2, Length(fieldValue) - 2)
    else
      Result := fieldValue;
  end
  else
    Result := '';
end;

function TForm1.ExtractField(const fields: TArray<string>; const fieldName: string): string;
begin
  for var field in fields do
  begin
    if Pos(fieldName, field) > 0 then
      Exit(Trim(Copy(field, Length(fieldName) + 1, MaxInt)));
  end;
  Result := '';
end;

end.
