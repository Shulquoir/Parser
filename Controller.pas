unit Controller;

interface

uses
  Winapi.Windows, Dialogs, Classes, SysUtils, StdCtrls, ShellAPI, Model;

type
  TController = class
  public
    class function ChooseFile(var filePath: string; const FileExtension: string): Boolean;
    class function OpenFileByName(Sender: TObject; const Key: Char): Boolean;
    class function CreateTSV(const OpenedDFNFilePath: string; var SavedTSVFilePath: string): Boolean;
    class function RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
    class procedure OpenExcel(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
    class procedure OpenCalc(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
  end;

implementation

class function TController.ChooseFile(var filePath: string; const FileExtension: string): Boolean;
var
  odOpenFileDialog: TOpenDialog;

begin
  Result := False;

  odOpenFileDialog := TOpenDialog.Create(nil);
  try
    odOpenFileDialog.Filter := FileExtension;
    if odOpenFileDialog.Execute then
    begin
      filePath := odOpenFileDialog.FileName;
      MessageDlg('Файл DFN был выбран.', mtInformation, [mbOK], 0);
      Result := True;
    end;
  finally
    odOpenFileDialog.Free;
  end;
end;

//------------------------------------------------------------------------------

class function TController.OpenFileByName(Sender: TObject; const Key: Char): Boolean;
const
  ENTER_KEY_CODE = #13;

begin
  Result := False;

  if (Key <> ENTER_KEY_CODE) then
    Exit;

  if FileExists(TEdit(Sender).Text) = false then
    begin
      MessageDlg('Файл не существует: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
      Exit;
    end;

  MessageDlg('Файл DFN был выбран: ' + TEdit(Sender).Text, mtInformation, [mbOK], 0);
  Result := True;
end;

//------------------------------------------------------------------------------

class function TController.CreateTSV(const OpenedDFNFilePath: string; var SavedTSVFilePath: string): Boolean;
var
   sdSaveTSV: TSaveDialog;
begin
  Result := false; // Повертаємо значення, що файл не був збережений

  if FileExists(OpenedDFNFilePath) = false then
  begin
    MessageDlg('Сначала выберите DFN файл', mtInformation, [mbOK], 0);
    Exit;
  end;

  sdSaveTSV := TSaveDialog.Create(nil);
  sdSaveTSV.FileName := ChangeFileExt(ExtractFileName(OpenedDFNFilePath), '.tsv'); // Отримуємо пропоноване
                                  // ім'я файлу для зберігання, через ім'я відкритого файлу з заміною формата
  if sdSaveTSV.Execute = false then  // Перевірка чи була натиснута кнопка збереження файлу в діалоговому вікні
    Exit;

  SavedTSVFilePath := TModel.ConvertDFNToTSV(OpenedDFNFilePath, sdSaveTSV);
  Result := true; // Повертаємо значення, що файл був збережений
  MessageDlg('Файл был успешно сконвертирован и сохранен в формате TSV.' + SavedTSVFilePath,
    mtInformation, [mbOK], 0);
end;

//------------------------------------------------------------------------------

class function TController.RewriteDFN(const OpenedTSVFilePath: string; const OpenedDFNToModFilePath: string): Boolean;
begin
   if (OpenedTSVFilePath = '') and (OpenedDFNToModFilePath = '') then
  begin
    MessageDlg('Откройте TSV и DFN файлы с совпадающими именами!', mtInformation, [mbOK], 0); // Виведення повідомленн, якщо не було
    Exit;                                                    // відкрито обидва файла, або одного із файлів
  end;

  if ChangeFileExt(ExtractFileName(OpenedTSVFilePath), '') <>
    ChangeFileExt(ExtractFileName(OpenedDFNToModFilePath), '') then // Перевірка ідентичності імен відкритих файлів
  begin
    MessageDlg('Пожалуйста, откройте TSV и DFN файлы с совпадающими именами.', mtInformation, [mbOK], 0);
    Exit; // Завершення процедури, в разі, якшо імена відкритих файлів не співпадають
  end;

  TModel.ModifyDFN(OpenedTSVFilePath, OpenedDFNToModFilePath);
  MessageDlg('Файл DFN был перезаписан.', mtInformation, [mbOK], 0);
end;


class procedure TController.OpenExcel(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
var
  ExcelPath: string;
  odOpenExcel: TOpenDialog;

begin
  ExcelPath := 'C:\Program Files\Microsoft Office\root\Office16\excel.exe'; // Стандартний шлях до Excel

  if IsFileCreate = false then // Перевірка, чи був створений TSV файл
  begin
    MessageDlg('Сначала создайте TSV файл', mtInformation, [mbOK], 0); // Виведення повідомлення, якщо не було створено TSV файл
    Exit;
  end;

  if FileExists(ExcelPath) then // Якщо було знайдено Excel
  begin
    ShellExecute(0, 'open', PChar(ExcelPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW); // Відкриття
    Exit;
  end;                                                              // створеного файлу в програмі Excel

  ShowMessage('Excel не знайдено. Виберіть файл Excel через діалог.'); // Якщо не було знайдено Excel, дає змогу обрати самостійно програму Excel
  odOpenExcel := TOpenDialog.Create(nil);

  if odOpenExcel.Execute then //Перевірка чи була натиснута кнопка відкриття обраного файлу в діалоговому вікні
    ShellExecute(0, 'open', PChar(odOpenExcel.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
    // Відкриття створеного файлу в програмі Excel з урахуванням обраного шляху для програми Excel
end;


class procedure TController.OpenCalc(const SavedTSVFilePath: string; const IsFileCreate: Boolean);
var
  CalcPath: string;
  odOpenCalc: TOpenDialog;

begin
  CalcPath := 'C:\Program Files\LibreOffice\program\soffice.exe'; // Стандартний шлях до Calc

  if IsFileCreate then // // Перевірка, чи був створений TSV файл
  begin
    MessageDlg('Сначала создайте TSV файл', mtInformation, [mbOK], 0); // Виведення повідомлення, якщо не було створено TSV файл
    Exit;
  end;

  if FileExists(CalcPath) then // // Якщо було знайдено Calc
  begin
    ShellExecute(0, 'open', PChar(CalcPath), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW) // Відкриття
  end;                                                               // створеного файлу в програмі Calc

  ShowMessage('Calc не знайдено. Виберіть файл Calc через діалог.'); // Якщо не було знайдено Calc, дає змогу обрати самостійно програму Calc
  odOpenCalc := TOpenDialog.Create(nil);

  if odOpenCalc.Execute then //Перевірка чи була натиснута кнопка відкриття обраного файлу в діалоговому вікні
    ShellExecute(0, 'open', PChar(odOpenCalc.FileName), PChar('"'+SavedTSVFilePath+'"'), nil, SW_SHOW);
end;   // Відкриття створеного файлу в програмі Calc з урахуванням обраного шляху для програми Calc

end.
