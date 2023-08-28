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
    OriginalDFNFileContent, OriginalTSVFileContent, OriginalDFNToModFileContent: TStringList; // Змінні, які будуть зберігати контент відкритого файлу до моменту переобрання або зберігання файлу
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFNToModFilePath: String; // Змінні, які будуть зберігати шлях відкритих файлів
    SavedTSVFilePath: String; // Змінна для зберігання шляху збереженого TSV файлу для можливості відкрити його в сторонніх програмах
    IsFileCreate: boolean; // Змінна для перевірки, чи був збережений TSV файл, для відкриття його в сторонніх програмах
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
  OutputFileContent: TStringList; // Змінна для зберігання форматованого контенту файлу

begin
  if FileExists(OpenedDFNFilePath) = false then
  begin
    ShowMessage('Сначала откройте DFN файл');
    Exit;
  end;

  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // Отримуємо пропоноване
                                  // ім'я файлу для зберігання, через ім'я відкритого файлу з заміною формата
  if sdSaveTSV.Execute = false then  // Перевірка чи була натиснута кнопка збереження файлу в діалоговому вікні
    Exit;

  try
    OriginalDFNFileContent := TStringList.Create; // Створюємо об'єкт типу TStringList
    OriginalDFNFileContent.LoadFromFile(OpenedDFNFilePath); // Записуємо контент з файлу
    OutputFileContent := TStringList.Create; // Створюємо об'єкт типу TStringList
    OutputFileContent.Add('ID' + #9 + 'Значения' + #9 + 'Перевод'); // Додаємо перший рядок в змінну,
                                                     // яка відповідає за форматований контент файлу
    for var i := 0 to OriginalDFNFileContent.Count - 1 do // Перебираємо рядки в файлі
    begin
      var line := OriginalDFNFileContent[i]; // Оголошуємо змінну яка буде містити наш радок

      if ContainsCyrillicCharacters(line) and  // Перевіряємо за допомогою створеної функції на наявність кирилиці
      (Pos('ID', line) > 0) and (Pos('Orig', line) > 0) and (Pos('Curr', line) > 0) then // Також
      begin     // відбувається перевірка на наявість в рядку одразу трьох полів 'ID', 'Orig', 'Curr'
        var fields := line.Split([#4]); // Розділяємо рядок на поля, який розділений символом EOT
          var id := ExtractField(fields, 'ID'); // За допомогою створеної функції, в яку передаємо поле та
          var orig := ExtractField(fields, 'Orig'); // ідентифікатор 'ID'/'Orig'/'Curr', отримуємо зміст
          var curr := ExtractField(fields, 'Curr'); // відповідного поля без одинарних лапок, якщо такі присутні
        OutputFileContent.Add(id + #9 + orig + #9 + curr);
      end;
    end;

    if Pos('.tsv', sdSaveTSV.FileName) > 0 then // Перевірка, чи є в імені зберігаємого файлу розширення tsv
      OutputFileContent.SaveToFile(sdSaveTSV.FileName) // Запис форматованого контенту в зберігаємий файл
    else
      OutputFileContent.SaveToFile(sdSaveTSV.FileName + '.tsv'); // Запис форматованого контенту
                                                    // в зберігаємий файл з дописуванням формату
    SavedTSVFilePath := sdSaveTSV.FileName; // Записуємо шлях створеного файлу в глобальну змінну,
                                // необхідно для подальшого відкриття файлу в сторонніх програмах
    IsFileCreate := true; // Підтвердження, що TSV файл був збережений, для подальшого використання в сторонніх програмах
    ShowMessage('Файл был успешно сконвертирован и сохранен в формате TSV.' + sdSaveTSV.FileName);
  finally
    FreeAndNil(OriginalDFNFileContent); // Звільняємо зміст копії відкритого файлу
    FreeAndNil(OutputFileContent); // Звільняємо зміст з форматованим контентом
  end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bWriteDFNClick(Sender: TObject);
var
  tsvIDField, dfnIDField: string;

begin
  if (OpenedTSVFilePath = '') and (OpenedDFNToModFilePath = '') then
  begin
    ShowMessage('Откройте TSV и DFN файлы с совпадающими именами!'); // Виведення повідомленн, якщо не було
    Exit;                                                    // відкрито обидва файла, або одного із файлів
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // Перевірка ідентичності імен відкритих файлів
  begin
    ShowMessage('Пожалуйста, откройте TSV и DFN файлы с совпадающими именами.');
    Exit; // Завершення процедури, в разі, якшо імена відкритих файлів не співпадають
  end;

  try
    OriginalTSVFileContent := TStringList.Create;
    OriginalDFNToModFileContent := TStringList.Create;
    OriginalTSVFileContent.LoadFromFile(OpenedTSVFilePath); // Записуємо контент з файлу
    OriginalDFNToModFileContent.LoadFromFile(OpenedDFNToModFilePath); // Записуємо контент з файлу

    for var TSVLine := 0 to OriginalTSVFileContent.Count - 1 do // Перебираємо рядки в файлі TSV
    begin
      tsvIDField := OriginalTSVFileContent[TSVLine].Split([#9])[0]; // Передаємо в змінну перше поле з рядка TSV файла
      for var DFNLine := 0 to OriginalDFNToModFileContent.Count - 1 do // Перебираємо рядки в файлі DFN
      begin
        dfnIDField := ExtractField(OriginalDFNToModFileContent[DFNLine].Split([#4]), 'ID'); // Записуємо поле "ID" DFN файла
        if Pos(tsvIDField, dfnIDField) = 1 then // Перевірка, чи зміст поля ID в DFN файлі ідентичне першому полю TSV файла
        begin // Якщо було знайдено співпадіння, то шукаємо поле "Curr"
          var OriginalDFNLine := OriginalDFNToModFileContent[DFNLine]; // Копія повного рядка DFN файлу для
                                                                      // подальшої заміни в ньому поля Curr
          for var CurrField := 0 to Length(OriginalDFNToModFileContent[DFNLine].Split([#4])) - 1 do // Перебираємо
          begin  // поля рядка в якому було знайдено співпадіння поля "ID" для знаходження номера поля в якому
                  // присутня назва поля "Curr", використаємо потім цей номер для перезапису відповідного поля
            if Pos('Curr:', OriginalDFNToModFileContent[DFNLine].Split([#4])[CurrField]) >= 1 then // Безпосередня
            begin                                                               // перевірка поля на вміст 'Curr:'
              var DFNCurr := OriginalDFNToModFileContent[DFNLine].Split([#4])[CurrField]; // Записуємо поле Curr DFN файлу
              var TSVCurr := OriginalTSVFileContent[TSVLine].Split([#9])[2]; // Записуємо поле Curr TSV файлу
              OriginalDFNLine := OriginalDFNLine.Replace(DFNCurr, 'Curr:' + QuotedStr(TSVCurr)); // В рядку з ідентичним
              // полем "ID" замінюємо зміст поля "Curr", порядковий номер якого дорівнює CurrField, на 3 поле з файлу TSV
              OriginalDFNToModFileContent[DFNLine] := OriginalDFNLine; // Заміна оригінального рядка форматованою копією рядка
              Break; // Вийти з циклу, якщо знайдено відповідне поле
            end;
          end;
          Break; // Вийти з циклу, якщо знайдено ідентичне поле ID, переходимо до наступного рядка в файлі TSV
        end;
      end;
    end;
    OriginalDFNToModFileContent.SaveToFile(OpenedDFNToModFilePath);
    // Зберігаємо зміни у файлі DFN, використовуємо глобальну змінну зі збереженим шляхом відкритого файлу,
    // необхідно у випадку, якщо файл був відкритий за допомогою текстового поля
  finally
    FreeAndNil(OriginalTSVFileContent);
    FreeAndNil(OriginalDFNToModFileContent);
  end;
  ShowMessage('Файл DFN был перезаписан.');
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
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe'; // Стандартний шлях до Excel
  if IsFileCreate then // Перевірка, чи був створений TSV файл
  begin
    if FileExists(ExcelPath) then // Якщо було знайдено Excel
      ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // Відкриття
    else                                                               // створеного файлу в програмі Excel
    begin // Якщо не було знайдено Excel, дає змогу обрати самостійно програму Excel
      ShowMessage('Excel не знайдено. Виберіть файл Excel через діалог.');
      if odOpenExcel.Execute then //Перевірка чи була натиснута кнопка відкриття обраного файлу в діалоговому вікні
      begin
        ShellExecute(0, 'open', PChar(odOpenExcel.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;        // Відкриття створеного файлу в програмі Excel з урахуванням обраного шляху для програми Excel
    end;
  end
  else
    begin
      ShowMessage('Сначала создайте TSV файл'); // Виведення повідомлення, якщо не було створено TSV файл
      ShowWarningMessage := 'Спочатку створіть TSV файл'; // Записуємо повідомлення в глобальну змінну для тестів
    end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenCalcClick(Sender: TObject);
var
  CalcPath: string;

begin
  ShowWarningMessage := '';
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe'; // Стандартний шлях до Calc
  if IsFileCreate then // // Перевірка, чи був створений TSV файл
  begin
    if FileExists(CalcPath) then // // Якщо було знайдено Calc
      ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // Відкриття
    else                                                               // створеного файлу в програмі Calc
    begin // Якщо не було знайдено Calc, дає змогу обрати самостійно програму Calc
      ShowMessage('Calc не знайдено. Виберіть файл Calc через діалог.');
      if odOpenCalc.Execute then //Перевірка чи була натиснута кнопка відкриття обраного файлу в діалоговому вікні
      begin
        ShellExecute(0, 'open', PChar(odOpenCalc.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
      end;          // Відкриття створеного файлу в програмі Calc з урахуванням обраного шляху для програми Calc
    end;
  end
  else
    begin
      ShowMessage('Сначала создайте TSV файл'); // Виведення повідомлення, якщо не було створено TSV файл
      ShowWarningMessage := 'Спочатку створіть TSV файл'; // Записуємо повідомлення в глобальну змінну для тестів
    end;
end;

//------------------------------------------------------------------------------

procedure TMainForm.FormShow(Sender: TObject);
begin
  ePathDFN.SetFocus; // Встановлюємо курсор на перше поле вводу шляху, для швидкого доступу
end;


//------------------------------------------------------------------------------

procedure TMainForm.FormCreate(Sender: TObject);
begin
  IsFileCreate := false; // При створенні форми привласнюємо глобальній змінній, що відповідає за перевірку
end;                                                            // чи був створений TSV файл, значення false

//------------------------------------------------------------------------------

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  OriginalTSVFileContent.Free;
  OriginalDFNToModFileContent.Free;
end;

//------------------------------------------------------------------------------

function TMainForm.ContainsCyrillicCharacters(const input: string): Boolean;
var i: Integer;                   // Функція для перевірки чи містить рядок кирилицю
begin
  Result := False; // Записуємо в результат фунцкції false, якщо не буде знайдено кирилицю
  for i := 1 to Length(input) do // Перебираємо літери в переданому рядку
  begin
    if IsCharCyrillic(input[i]) then // Використовуємо ще одну функцію для перевірки чи є літера кирилицею
    begin
      Result := True; // Якщо було знайдено відповідність символу кирилиці, зазначеному в фунцкціїї IsCharCyrillic
      Exit; // Завершуємо виконання функції з результатом true
    end;
  end;
end;

//------------------------------------------------------------------------------

function TMainForm.IsCharCyrillic(c: Char): Boolean; // Функція для перевірки чи є літера кирилицею
begin
  Result := (c >= 'А') and (c <= 'Я') or (c >= 'а') and (c <= 'я') or (c = 'Ё') or (c = 'ё'); // Повертає true
end;                                                 // якщо змінна "с" входить в діапазон зазначених символів

//------------------------------------------------------------------------------

function TMainForm.ExtractField(const fields: TArray<string>; const fieldName: string): string;
// Функція для знаходження змісту поля за його назвою
var fieldValue: string;

begin
  for var field in fields do // Перебираємо поля в рядку
  begin
    if Pos(fieldName, field) > 0 then // Перевірка, чи міститься в полі задана назва поля
    begin
      fieldValue := (Copy(field, Pos(fieldName, field) + Length(fieldName) + 1, MaxInt));  // Якщо знайдено, то повертає зміст поля,
                   // тобто повертає зміст поля, без імені поля та символу ':', та весь зміст до кінця поля
      if (Length(fieldValue) >= 2) and (fieldValue[1] = '''') and (fieldValue[Length(fieldValue)] = '''') then
      // Перевірка чи містить зміст поля в кінці та на початку одинарні лапки
        Result := Copy(fieldValue, 2, Length(fieldValue) - 2) // Якщо містить, то прибирає
      else
        Result := fieldValue; // Якщо ні, то нічого не змінюється
      Break; // Виходимо з циклу при знаходженні відповідної назви поля
    end
    else
      Result := ''; // Якщо не було знайдено ім'я поля
  end;
end;

end.
