unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs, Vcl.Buttons;

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
    procedure ePath1KeyPress(Sender: TObject; var Key: Char);
    procedure spChooseDFNClick(Sender: TObject);
    procedure spChooseTSVClick(Sender: TObject);
    procedure spChooseDFN2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bCreateTSVClick(Sender: TObject);
  private
    { Private declarations }
    function ContainsCyrillicCharacters(const input: string): Boolean;
    function IsCharCyrillic(c: Char): Boolean;
  public
    { Public declarations }

  end;

var
  Form1: TForm1;
  OriginalFileContent: TStringList; // ��������� ��������� �����, ��� ���� �������� ������� ��������� ����� �� ������� ����������� ��� ��������� �����
  OpenedFilePath: string; // ��������� ��������� �����, ��� ���� �������� ���� ��������� �����

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
        OutputFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
        try
          OutputFileContent.Add('ID' + #9 + '��������' + #9 + '�������'); // ������ ����� ������ � �����, ��� ������� �� ������������ ������� �����
          for var i := 0 to OriginalFileContent.Count - 1 do
        begin
          var line := OriginalFileContent[i]; //
          if ContainsCyrillicCharacters(line) then //
            OutputFileContent.Add(line); //
        end;
            if Pos('.tsv', SaveDialog1.FileName) > 0 then // �������� �� � � ����� ���������� ����� ���������� tsv
              OutputFileContent.SaveToFile(SaveDialog1.FileName) // ����� ������������� �������� � ��������� ����
            else
              OutputFileContent.SaveToFile(SaveDialog1.FileName + '.tsv'); // ����� ������������� �������� � ��������� ���� � ������������ �������
            ShowMessage('���� ��� ������ �������������� �� ���������� � ������ TSV.' + SaveDialog1.FileName);
        finally
          OutputFileContent.Free;
        end;
      end;
  end
  else
    ShowMessage('�������� �������� ���� DFN.');
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

end.
