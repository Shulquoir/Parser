unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs, Vcl.Buttons, ShellAPI;

type
  TForm1 = class(TForm)
    ePathDFN: TEdit;
    bCreateTSV: TButton;
    ePathTSV: TEdit;
    ePathDFN2: TEdit;
    bOpenExcel: TButton;
    bOpenCalc: TButton;
    bWriteDFN: TButton;
    odOpenDFN: TOpenDialog;
    odOpenTSV: TOpenDialog;
    odOpenDFN2: TOpenDialog;
    spChooseDFN: TSpeedButton;
    Label1: TLabel;
    spChooseTSV: TSpeedButton;
    Label2: TLabel;
    spChooseDFN2: TSpeedButton;
    Label3: TLabel;
    sdSaveTSV: TSaveDialog;
    OpenDialog4: TOpenDialog;
    OpenDialog5: TOpenDialog;
    sdSaveDFN: TSaveDialog;
    procedure ePathDFNKeyPress(Sender: TObject; var Key: Char);
    procedure spChooseDFNClick(Sender: TObject);
    procedure spChooseTSVClick(Sender: TObject);
    procedure spChooseDFN2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bCreateTSVClick(Sender: TObject);
    procedure bOpenExcelClick(Sender: TObject);
    procedure bOpenCalcClick(Sender: TObject);
    procedure ePathTSVKeyPress(Sender: TObject; var Key: Char);
    procedure ePathDFN2KeyPress(Sender: TObject; var Key: Char);
    procedure bWriteDFNClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    OriginalDFNFileContent, OriginalTSVFileContent, OriginalDFN2FileContent: TStringList; // ��������� ��������� �����, ��� ���� �������� ������� ��������� ����� �� ������� ����������� ��� ��������� �����
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFN2FilePath: String; // ��������� ��������� �����, �� ������ �������� ���� �������� �����
    SavedTSVFilePath: String; // ��������� ��������� ����� ��� ��������� ����� ����������� TSV ����� ��� ��������� ������� ���� � ��������� ���������
    IsFileCreate: boolean; // ����� ��� ��������, �� ��� ���������� TSV ����, ��� �������� ���� � ��������� ���������
    ShowWarningMessage: string;
    function ContainsCyrillicCharacters(const input: string): Boolean;
    function IsCharCyrillic(c: Char): Boolean;
    function ExtractField(const fields: TArray<string>; const fieldName: string): string;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.bCreateTSVClick(Sender: TObject);

var OutputFileContent: TStringList; // ��������� ����� ��� ��������� ������������� �������� �����

begin
  if OriginalDFNFileContent <> nil then // �������� �� ��� �������� ����
  begin
    sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // �������� �����������
    // ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
    if sdSaveTSV.Execute then  // �������� �� ���� ��������� ������ ���������� ����� � ���������� ����
    begin
      try
        IsFileCreate := true;
        OutputFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        try
          OutputFileContent.Add('ID' + #9 + '��������' + #9 + '�������'); // ������ ����� ������ � �����,
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
                OutputFileContent.Add(id + #9 + orig + #9 + curr); //
            end;
          end;
          if Pos('.tsv', sdSaveTSV.FileName) > 0 then // ��������, �� � � ����� ���������� ����� ���������� tsv
            OutputFileContent.SaveToFile(sdSaveTSV.FileName) // ����� ������������� �������� � ��������� ����
          else
            OutputFileContent.SaveToFile(sdSaveTSV.FileName + '.tsv'); // ����� ������������� ��������
            // � ��������� ���� � ������������ �������

          SavedTSVFilePath := sdSaveTSV.FileName; // �������� ���� ���������� ����� � ��������� �����,
          // ��������� ��� ���������� �������� ����� � ��������� ���������
          ShowMessage('���� ��� ������� �������������� � �������� � ������� TSV.' + sdSaveTSV.FileName);
        finally
          OutputFileContent.Free; // ��������� ���� ����� � ������������ ���������
        end;
      except
         on E: Exception do
           ShowMessage('������ ��� �������� �����: ' + E.Message);
      end;
    end;
  end
  else
    ShowMessage('������� �������� ���� DFN.');
end;

procedure TForm1.bWriteDFNClick(Sender: TObject);

var FieldA, FieldB: string;

begin
   if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
     ChangeFileExt(ExtractFileName(OpenedDFN2FilePath), '') then // �������� ����������� ���� �������� �����
   begin
     ShowMessage('����������, �������� TSV � DFN ����� � ������������ �������.');
     Exit; // ���������� ���������, � ���, ���� ����� �������� ����� �� ����������
   end;

   if (OriginalTSVFileContent <> nil) and (OriginalDFN2FileContent <> nil) then // �������� �� ���� ������ �����
   begin
     for var TSVLine := 0 to OriginalTSVFileContent.Count - 1 do // ���������� ����� � ���� TSV
     begin
         FieldA := OriginalTSVFileContent[TSVLine].Split([#9])[0]; // �������� � ����� ����� ���� � ����� TSV �����
         for var DFNLine := 0 to OriginalDFN2FileContent.Count - 1 do // ���������� ����� � ���� DFN
         begin
           FieldB := OriginalDFN2FileContent[DFNLine].Split([#4])[0] // �������� � ����� ���� "ID" DFN �����;
             .Substring(Pos('ID:', OriginalDFN2FileContent[DFNLine].Split([#4])[0]) + 2);
           if Pos(FieldA, FieldB) = 1 then // ��������, �� ���� ���� ID � DFN ���� ��������� ������� ���� TSV �����
           begin // ���� ���� �������� ���������, �� ������ ���� "Curr"
             for var CurrField := 0 to Length(OriginalDFN2FileContent[DFNLine].Split([#4])) - 1 do // ����������
             begin  // ���� ����� � ����� ���� �������� ��������� ���� "ID" ��� ����������� ������ ���� � �����
                    // �������� ����� ���� "Curr", ����������� ���� ��� ����� ��� ���������� ���������� ����
               if Pos('Curr:', OriginalDFN2FileContent[DFNLine].Split([#4])[CurrField]) >= 1 then // ������������
               begin                                                           // �������� ���� �� ���� 'Curr:'
                 OriginalDFN2FileContent[DFNLine] := OriginalDFN2FileContent[DFNLine] // � ����� � ����������
                   .Replace(OriginalDFN2FileContent[DFNLine].Split([#4])[CurrField], // ����� "ID" �������� ����
                   'Curr:' + QuotedStr(OriginalTSVFileContent[TSVLine].Split([#9])[2])); // ���� "Curr", ����������
                                                           // ����� ����� ������� CurrField, �� 3 ���� � ����� TSV
                 Break; // ����� � �����, ���� �������� �������� ����
               end;
             end;
           Break; // ����� � �����, ���� �������� ��������� ���� ID, ���������� �� ���������� ����� � ���� TSV
           end;
         end;
     end;
       OriginalDFN2FileContent.SaveToFile(OpenedDFN2FilePath);
       // �������� ���� � ���� DFN, ������������� ��������� ����� � ���������� ������ ��������� �����,
       // ��������� � �������, ���� ���� ��� �������� �� ��������� ���������� ����
       ShowMessage('���� DFN ��� �����������.');
   end
   else
     ShowMessage('�������� TSV � DFN ����� � ������������ �������!'); // ��������� ����������, ���� �� ����
                                                              // ������� ������ �����, ��� ������ �� �����
end;

procedure TForm1.spChooseDFNClick(Sender: TObject);
begin
  if odOpenDFN.Execute then // �������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
  begin
    try
      if OriginalDFNFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalDFNFileContent.Free; // ��������� ���������� ����, ���� ���

      ePathDFN.Text := odOpenDFN.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalDFNFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalDFNFileContent.LoadFromFile(odOpenDFN.FileName); // �������� ������� � ��������� �����
      OpenedDFNFilePath := odOpenDFN.FileName; // �������� ���� ��������� ����� � ��������� �����
      ShowMessage('���� DFN ��� ������.');
    except
      on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
    end;
  end;
end;

procedure TForm1.ePathDFNKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then // ���� ���� ��������� ������ #13 - ��� ������ "Enter"
  begin
    if FileExists(ePathDFN.Text) then // �������� �� ���� ����
    begin
      try
        if OriginalDFNFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
          OriginalDFNFileContent.Free; // ��������� ���������� ����, ���� ���

        OriginalDFNFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        OriginalDFNFileContent.LoadFromFile(ePathDFN.Text); // �������� ������� � ��������� �����
        OpenedDFNFilePath := ePathDFN.Text; // �������� ���� ��������� ����� � ��������� �����
        ShowMessage('���� ��� ������: ' + ePathDFN.Text);
      except
        on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
      end;
    end
    else
    begin
    ShowMessage('���� �� ����������: ' + ePathDFN.Text); // ���� ��� �������� �� ��������� ���� �� �����
    end;
  end;
end;

procedure TForm1.spChooseTSVClick(Sender: TObject);
begin
  if odOpenTSV.Execute then // �������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
  begin
    try
      if OriginalTSVFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalTSVFileContent.Free; // ��������� ���������� ����, ���� ���

      ePathTSV.Text := odOpenTSV.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalTSVFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalTSVFileContent.LoadFromFile(odOpenTSV.FileName); // �������� ������� � ��������� �����
      OpenedTSVFilePath := odOpenTSV.FileName; // �������� ���� ��������� ����� � ��������� �����
      ShowMessage('���� TSV ��� ������.');
    except
      on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
    end;
  end;
end;

procedure TForm1.ePathTSVKeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then // ���� ���� ��������� ������ #13 - ��� ������ "Enter"
  begin
    if FileExists(ePathTSV.Text) then // �������� �� ���� ����
    begin
      try
        if OriginalTSVFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
          OriginalTSVFileContent.Free; // ��������� ���������� ����, ���� ���

        OriginalTSVFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        OriginalTSVFileContent.LoadFromFile(ePathTSV.Text); // �������� ������� � ��������� �����
        OpenedTSVFilePath := ePathTSV.Text; // �������� ���� ��������� ����� � ��������� �����
        ShowMessage('���� TSV ��� ������: ');
      except
        on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
      end;
    end
    else
    begin
    ShowMessage('���� �� ����������: ' + ePathTSV.Text); // ���� ��� �������� �� ��������� ���� �� �����
    end;
  end;
end;

procedure TForm1.spChooseDFN2Click(Sender: TObject);
begin
  if odOpenDFN2.Execute then // �������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
  begin
    try
      if OriginalDFN2FileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalDFN2FileContent.Free; // ��������� ���������� ����, ���� ���

      ePathDFN2.Text := odOpenDFN2.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalDFN2FileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalDFN2FileContent.LoadFromFile(odOpenDFN2.FileName); // �������� ������� � ��������� �����
      OpenedDFN2FilePath := odOpenDFN2.FileName; // �������� ���� ��������� ����� � ��������� �����
      ShowMessage('���� DFN ��� ������.');
    except
      on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
    end;
  end;
end;

procedure TForm1.ePathDFN2KeyPress(Sender: TObject; var Key: Char);
begin
  if (Key = #13) then // ���� ���� ��������� ������ #13 - ��� ������ "Enter"
  begin
    if FileExists(ePathDFN2.Text) then // �������� �� ���� ����
    begin
      try
        if OriginalDFN2FileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
          OriginalDFN2FileContent.Free; // ��������� ���������� ����, ���� ���

        OriginalDFN2FileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        OriginalDFN2FileContent.LoadFromFile(ePathDFN2.Text); // �������� ������� � ��������� �����
        OpenedDFN2FilePath := ePathDFN2.Text; // �������� ���� ��������� ����� � ��������� �����
        ShowMessage('���� ��� ������: ' + ePathDFN2.Text);
      except
        on E: Exception do
          ShowMessage('������ ��� �������� �����: ' + E.Message);
      end;
    end
    else
    begin
    ShowMessage('���� �� ����������: ' + ePathDFN2.Text); // ���� ��� �������� �� ��������� ���� �� �����
    end;
  end;
end;

procedure TForm1.bOpenExcelClick(Sender: TObject);
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
      if OpenDialog4.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
      begin
        ShellExecute(0, 'open', PChar(OpenDialog4.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;        // ³������� ���������� ����� � ������� Excel � ����������� �������� ����� ��� �������� Excel
    end;
  end
  else
    begin
    ShowMessage('�������� ������� TSV ����'); // ��������� �����������, ���� �� ���� �������� TSV ����
    ShowWarningMessage := '�������� ������� TSV ����'; // �������� ����������� � ��������� ����� ��� �����
    end;
end;

procedure TForm1.bOpenCalcClick(Sender: TObject);
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
      if OpenDialog5.Execute then //�������� �� ���� ��������� ������ �������� �������� ����� � ���������� ����
      begin
        ShellExecute(0, 'open', PChar(OpenDialog5.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;          // ³������� ���������� ����� � ������� Calc � ����������� �������� ����� ��� �������� Calc
    end;
  end
  else
    begin
    ShowMessage('�������� ������� TSV ����'); // ��������� �����������, ���� �� ���� �������� TSV ����
    ShowWarningMessage := '�������� ������� TSV ����'; // �������� ����������� � ��������� ����� ��� �����
    end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  IsFileCreate := false; // ��� ��������� ����� ������������ ���������� ������, �� ������� �� ��������
end;                                                            // �� ��� ��������� TSV ���� �������� false

procedure TForm1.FormDestroy(Sender: TObject);
begin
  OriginalDFNFileContent.Free; // ��������� ���� ������������ �������� � DFN ����� ��� ������� ����������
  OriginalTSVFileContent.Free; // ��������� ���� ������������ �������� � TSV ����� ��� ������� ����������
  OriginalDFN2FileContent.Free; // ��������� ���� ������������ �������� � DFN ����� ��� ������� ����������
end;

function TForm1.ContainsCyrillicCharacters(const input: string): Boolean;
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

function TForm1.IsCharCyrillic(c: Char): Boolean; // ������� ��� �������� �� � ����� ���������
begin
  Result := (c >= '�') and (c <= '�') or (c >= '�') and (c <= '�') or (c = '�') or (c = '�'); // ������� true
end;                                                    // ���� ����� "�" ������� � ������� ���������� ����

function TForm1.ExtractField(const fields: TArray<string>; const fieldName: string): string;
// ������� ��� ����������� ����� ���� �� ���� ������
var fieldValue: string;

begin
  for var field in fields do // ���������� ���� � �����
  begin
    if Pos(fieldName, field) > 0 then // ��������, �� �������� � ��� ������ ����� ����
    begin
      fieldValue := (Copy(field, Length(fieldName) + 2, MaxInt));  // ���� ��������, �� ������� ���� ����,
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
