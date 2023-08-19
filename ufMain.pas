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
  OriginalDFNFileContent, OriginalTSVFileContent, OriginalDFN2FileContent: TStringList; // ��������� ��������� �����, ��� ���� �������� ������� ��������� ����� �� ������� ����������� ��� ��������� �����
  OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFN2FilePath: String; // ��������� ��������� �����, �� ������ �������� ���� �������� �����
  SavedTSVFilePath: String; // ��������� ��������� ����� ��� ��������� ����� ����������� TSV ����� ��� ��������� ������� ���� � ��������� ���������
  IsFileCreate: boolean = false; // ����� ��� ��������, �� ��� ���������� TSV ����, ��� �������� ���� � ��������� ���������

implementation

{$R *.dfm}

procedure TForm1.bCreateTSVClick(Sender: TObject);

var OutputFileContent: TStringList; // ��������� ����� ��� ��������� ������������� �������� �����

begin
  if OriginalDFNFileContent <> nil then // �������� �� ��� �������� ����
  begin
    sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // �������� ����������� ��'� ����� ��� ���������, ����� ��'� ��������� ����� � ������ �������
    if sdSaveTSV.Execute then  // �������� �� ���� ��������� ������ ���������� ����� � ���������� ����
    begin
      IsFileCreate := true;
      OutputFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      try
        OutputFileContent.Add('ID' + #9 + '��������' + #9 + '�������'); // ������ ����� ������ � �����, ��� ������� �� ������������ ������� �����
        for var i := 0 to OriginalDFNFileContent.Count - 1 do
        begin
          var line := OriginalDFNFileContent[i]; //
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
        if Pos('.tsv', sdSaveTSV.FileName) > 0 then // �������� �� � � ����� ���������� ����� ���������� tsv
          OutputFileContent.SaveToFile(sdSaveTSV.FileName) // ����� ������������� �������� � ��������� ����
        else
          OutputFileContent.SaveToFile(sdSaveTSV.FileName + '.tsv'); // ����� ������������� �������� � ��������� ���� � ������������ �������

        SavedTSVFilePath := sdSaveTSV.FileName; // �������� ���� ���������� �����
        ShowMessage('���� ��� ������� �������������� � �������� � ������� TSV.' + sdSaveTSV.FileName);
      finally
        OutputFileContent.Free;
      end;
    end;
  end
  else
    ShowMessage('������� �������� ���� DFN.');
end;

procedure TForm1.bWriteDFNClick(Sender: TObject);

var FieldA, FieldB: string;
begin
   if ChangeFileExt(ExtractFileName(odOpenTSV.FileName), '') <> ChangeFileExt(ExtractFileName(odOpenDFN2.FileName), '') then
   begin
     ShowMessage('����������, �������� TSV � DFN ����� � ������������ �������.');
     Exit;
   end;

   if (OriginalTSVFileContent <> nil) and (OriginalDFN2FileContent <> nil) then
   begin
     try
       for var TSVLine := 0 to OriginalTSVFileContent.Count - 1 do
       begin
         FieldA := OriginalTSVFileContent[TSVLine].Split([#9])[0];
         for var DFNLine := 0 to OriginalDFN2FileContent.Count - 1 do
         begin
           FieldB := OriginalDFN2FileContent[DFNLine].Split([#4])[0]
             .Substring(Pos('ID:', OriginalDFN2FileContent[DFNLine].Split([#4])[0]) + 2); // ����������� �������� ���� "ID:";
           if Pos(FieldA, FieldB) = 1 then // ����������, �� ���� ���������� � ����������� ����
           begin
             for var CurrField := 0 to Length(OriginalDFN2FileContent[DFNLine].Split([#4])) - 1 do
             begin
               if Pos('Curr:', OriginalDFN2FileContent[DFNLine].Split([#4])[CurrField]) >= 1 then
               begin
                 OriginalDFN2FileContent[DFNLine] := OriginalDFN2FileContent[DFNLine]
                   .Replace(OriginalDFN2FileContent[DFNLine].Split([#4])[CurrField],
                   'Curr:' + QuotedStr(OriginalTSVFileContent[TSVLine].Split([#9])[2]));
                 Break; // ����� � �����, ���� �������� �������� ����
               end;
             end;
           Break;
           end;
         end;
       end;

       // ���������� ���������� ����� DFN
        OriginalDFN2FileContent.SaveToFile(odOpenDFN2.FileName);
        ShowMessage('���� DFN ��� �����������.');
     finally

     end;
   end
   else
   ShowMessage('�������� TSV � DFN �����!');
end;

procedure TForm1.spChooseDFNClick(Sender: TObject);
begin
  if odOpenDFN.Execute then // �������� �� ���� ��������� ������ �������� �������� �����
  begin
    try
      if OriginalDFNFileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalDFNFileContent.Free; // ��������� ���������� ����, ���� ���

      ePathDFN.Text := odOpenDFN.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalDFNFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalDFNFileContent.LoadFromFile(odOpenDFN.FileName); // �������� ������� � ��������� �����
      OpenedDFNFilePath := odOpenDFN.FileName; // �������� ���� ��������� ����� � ��������� �����
      //ShowMessage('���� DFN ��� ������.');
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
        //ShowMessage('���� ��� ������: ' + ePathDFN.Text);
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
  if odOpenTSV.Execute then // �������� �� ���� ��������� ������ �������� �������� �����
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
        //ShowMessage('���� ��� ������: ' + ePathTSV.Text);
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
  if odOpenDFN2.Execute then // �������� �� ���� ��������� ������ �������� �������� �����
  begin
    try
      if OriginalDFN2FileContent <> nil then // �������� �� ��� ��� �������� ���� �� ����� �������
        OriginalDFN2FileContent.Free; // ��������� ���������� ����, ���� ���

      ePathDFN2.Text := odOpenDFN2.FileName; // �������� ���� �� ����� � �������� ���� "edit"
      OriginalDFN2FileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
      OriginalDFN2FileContent.LoadFromFile(odOpenDFN2.FileName); // �������� ������� � ��������� �����
      OpenedDFN2FilePath := odOpenDFN2.FileName; // �������� ���� ��������� ����� � ��������� �����
      //ShowMessage('���� DFN ��� ������.');
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
        //ShowMessage('���� ��� ������: ' + ePathDFN2.Text);
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
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe';
  if IsFileCreate then
  begin
    if FileExists(ExcelPath) then
    begin
      if IsFileCreate then
      begin
        ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;
    end
    else
    begin
      ShowMessage('Excel �� ��������. ������� ���� Excel ����� �����.');
      if OpenDialog4.Execute then
      begin
        ShellExecute(0, 'open', PChar(OpenDialog4.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;
    end;
  end
  else
    ShowMessage('�������� ������� TSV ����');
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
        ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;
    end
    else
    begin
      ShowMessage('Calc �� ��������. ������� ���� Calc ����� �����.');
      if OpenDialog5.Execute then
      begin
        ShellExecute(0, 'open', PChar(OpenDialog5.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;
    end;
  end
  else
    ShowMessage('�������� ������� TSV ����');
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  OriginalDFNFileContent.Free; // ��������� ���� ������������ �������� � ����� ��� ������� ����������.
  OriginalTSVFileContent.Free;
  OriginalDFN2FileContent.Free;
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
