unit Model;

interface

uses
  SysUtils, Classes, Vcl.StdCtrls, Vcl.Dialogs;

type
  TModel = class
  public
    class function ConvertDFNToTSV(const OpenedDFNFilePath: string; var sdSaveTSV: TSaveDialog): string;
    class function ModifyDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
    class function ContainsCyrillicCharacters(const input: string): Boolean;
    class function IsCharCyrillic(c: Char): Boolean;
    class function ExtractField(const fields: TArray<string>; const fieldName: string): string;
  end;

implementation

/// <summary>
/// Конвертує вміст DFN файлу в формат TSV та зберігає його з обраним ім'ям та розширенням.
/// </summary>
/// <param name="OpenedDFNFilePath">Шлях до відкритого DFN файлу.</param>
/// <param name="sdSaveTSV">Діалогове вікно для збереження TSV файлу.</param>
/// <returns>Шлях до збереженого TSV файлу.</returns>
class function TModel.ConvertDFNToTSV(const OpenedDFNFilePath: string; var sdSaveTSV: TSaveDialog): string;
var
  OriginalDFNFileContent, OutputFileContent: TStringList;

begin
  Result := '';

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

    if Pos('.tsv', sdSaveTSV.FileName) = 0 then // Перевірка, чи є в імені зберігаємого файлу розширення tsv
    begin
      OutputFileContent.SaveToFile(sdSaveTSV.FileName + '.tsv'); // Запис форматованого контенту в зберігаємий файл
      Result := sdSaveTSV.FileName + '.tsv';                                       // з додаванням розширення файла
      Exit;
    end;

    OutputFileContent.SaveToFile(sdSaveTSV.FileName); // Запис форматованого контенту в зберігаємий файл
    Result := sdSaveTSV.FileName;

  finally
    FreeAndNil(OriginalDFNFileContent); // Звільняємо зміст копії відкритого файлу
    FreeAndNil(OutputFileContent); // Звільняємо зміст з форматованим контентом
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// Модифікує вміст DFN файлу на основі відкритого TSV файлу та зберігає зміни у відкритому DFN файлі.
/// </summary>
/// <param name="OpenedTSVFilePath">Шлях до відкритого TSV файлу.</param>
/// <param name="OpenedDFNToModFilePath">Шлях до відкритого DFN файлу для модифікації.</param>
/// <returns>True, якщо зміни в DFN файлі були успішно збережені.</returns>
class function TModel.ModifyDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
var
  OriginalTSVFileContent, OriginalDFNToModFileContent: TStringList;
  tsvIDField, dfnIDField: string;

begin
  Result := false;

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
          var OriginalDFNLine := OriginalDFNToModFileContent[DFNLine]; // Копіюємо рядок DFN файлу для подальшої
                                                                                     // заміни в ньому поля Curr
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

    OriginalDFNToModFileContent.SaveToFile(OpenedDFNToModFilePath); // Зберігаємо зміни у файлі DFN
    Result := true;

  finally
    FreeAndNil(OriginalTSVFileContent);
    FreeAndNil(OriginalDFNToModFileContent);
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// Перевіряє, чи містить рядок кириличні символи.
/// </summary>
/// <param name="input">Рядок для перевірки.</param>
/// <returns>True, якщо рядок містить кириличні символи, інакше - False.</returns>
class function TModel.ContainsCyrillicCharacters(const input: string): Boolean;
var
  i: Integer;

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

/// <summary>
/// Перевіряє, чи є символ кириличним.
/// </summary>
/// <param name="c">Символ для перевірки.</param>
/// <returns>True, якщо символ є кириличним, інакше - False.</returns>
class function TModel.IsCharCyrillic(c: Char): Boolean;
begin
  Result := (c >= 'А') and (c <= 'Я') or (c >= 'а') and (c <= 'я') or (c = 'Ё') or (c = 'ё');
end;

//------------------------------------------------------------------------------

/// <summary>
/// Знаходить зміст поля за його назвою.
/// </summary>
/// <param name="fields">Масив полів, серед яких здійснюється пошук.</param>
/// <param name="fieldName">Назва поля, яке необхідно знайти.</param>
/// <returns>Зміст поля, якщо знайдено, або порожній рядок, якщо не знайдено.</returns>
class function TModel.ExtractField(const fields: TArray<string>; const fieldName: string): string;
var
  fieldValue: string;

begin
  for var field in fields do // Перебираємо поля в рядку
  begin
    if Pos(fieldName, field) > 0 then // Перевірка, чи міститься в полі задана назва поля
    begin
      fieldValue := (Copy(field, Pos(fieldName, field) + Length(fieldName) + 1, MaxInt));  // Якщо знайдено,
                      // то повертає зміст поля - без імені поля та символу ':' та весь зміст до кінця поля
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
