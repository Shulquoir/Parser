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
  OriginalFileContent: TStringList; // Створюємо глобальну змінну, яка буде зберігати контент відкритого файлу до моменту переобрання або зберігання файлу
  OpenedFilePath: string; // Створюємо глобальну змінну, яка буде зберігати шлях відкритого файлу

implementation

{$R *.dfm}

procedure TForm1.bCreateTSVClick(Sender: TObject);

var OutputFileContent: TStringList; // Створюємо змінну для зберігання форматованого контенту файлу

begin
  if OriginalFileContent <> nil then // Перевірка чи був відкритий файл
  begin
      SaveDialog1.FileName := ChangeFileExt(ExtractFileName(OpenedFilePath), '.tsv'); // Отримуємо пропоноване ім'я файлу для зберігання, через ім'я відкритого файлу з заміною формата
      if SaveDialog1.Execute then  // Перевірка чи була натиснута кнопка збереження файлу в діалоговому вікні
      begin
        OutputFileContent := TStringList.Create; // Створюємо об'єкт типу TStringList
        try
          OutputFileContent.Add('ID' + #9 + 'Значения' + #9 + 'Перевод'); // Додаємо першу строку в змінну, яка відповідає за форматований контент файлу
          for var i := 0 to OriginalFileContent.Count - 1 do
        begin
          var line := OriginalFileContent[i]; //
          if ContainsCyrillicCharacters(line) then //
            OutputFileContent.Add(line); //
        end;
            if Pos('.tsv', SaveDialog1.FileName) > 0 then // Перевірка чи є в імені зберігаємого файлу розширення tsv
              OutputFileContent.SaveToFile(SaveDialog1.FileName) // Запис форматованого контенту в зберігаємий файл
            else
              OutputFileContent.SaveToFile(SaveDialog1.FileName + '.tsv'); // Запис форматованого контенту в зберігаємий файл з дописуванням формату
            ShowMessage('Файл був успішно сконвертований та збережений у форматі TSV.' + SaveDialog1.FileName);
        finally
          OutputFileContent.Free;
        end;
      end;
  end
  else
    ShowMessage('Спочатку відкрийте файл DFN.');
end;

function TForm1.ContainsCyrillicCharacters(const input: string): Boolean; // Функція для перевірки чи є слова кирилицею
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

function TForm1.IsCharCyrillic(c: Char): Boolean; // Функція для перевірки чи є літера кирилицею
begin
  Result := (c >= 'А') and (c <= 'Я') or (c >= 'а') and (c <= 'я') or (c = 'Ё') or (c = 'ё');
end;

procedure TForm1.ePath1KeyPress(Sender: TObject; var Key: Char);

var FilePath: string;

begin
  if (Key = #13) then // Якщо була натиснута клавіша #13 - код клавіші "Enter"
  begin
    FilePath := ePath1.Text; // Записуємо шлях, який був вписаний в відповідне поле "edit"
    if FileExists(FilePath) then // Перевірка чи існує файл
    begin
      try
        if OriginalFileContent <> nil then // Перевірка чи був вже відкритий файл до цього моменту
          OriginalFileContent.Free; // Звільнюємо попередній вміст, якщо був

        OriginalFileContent := TStringList.Create; // Створюємо об'єкт типу TStringList
        OriginalFileContent.LoadFromFile(FilePath); // Записуємо контент з відкритого файлу
        OpenedFilePath := FilePath; // Записуємо шлях відкритого файлу в глобальну змінну
        ShowMessage('Файл відкрито: ' + FilePath);
      except
        on E: Exception do
          ShowMessage('Помилка при відкритті файлу: ' + E.Message);
      end;
    end
    else
    begin
    ShowMessage('Файл не існує: ' + FilePath); // Якщо був введений не коректний шлях до файлу
    end;
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  OriginalFileContent.Free; // Звільняємо вміст скопійованого контенту з файлу при закритті застосунку.
end;

procedure TForm1.spChooseDFN2Click(Sender: TObject);
begin
  if OpenDialog3.Execute then
  ePath3.Text := OpenDialog3.FileName;
end;

procedure TForm1.spChooseDFNClick(Sender: TObject);
begin
  if OpenDialog1.Execute then // Перевірка чи була натиснута кнопка відкриття обраного файлу
  begin
    try
      if OriginalFileContent <> nil then // Перевірка чи був вже відкритий файл до цього моменту
        OriginalFileContent.Free; // Звільнюємо попередній вміст, якщо був

      ePath1.Text := OpenDialog1.FileName; // Записуємо шлях до файла в відповідне поле "edit"
      OriginalFileContent := TStringList.Create; // Створюємо об'єкт типу TStringList
      OriginalFileContent.LoadFromFile(OpenDialog1.FileName); // Записуємо контент з відкритого файлу
      OpenedFilePath := OpenDialog1.FileName; // Записуємо шлях відкритого файлу в глобальну змінну
      ShowMessage('Файл DFN був відкритий.');
    except
      on E: Exception do
          ShowMessage('Помилка при відкритті файлу: ' + E.Message);
    end;
  end;
end;

procedure TForm1.spChooseTSVClick(Sender: TObject);
begin
  if OpenDialog2.Execute then
  ePath2.Text := OpenDialog2.FileName;
end;

end.
