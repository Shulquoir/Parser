unit Unit1;

interface

uses
  DUnitX.TestFramework, ufMain, Vcl.Dialogs, Winapi.Windows, Vcl.Buttons, System.SysUtils,
  Vcl.Graphics, System.Variants, ShellAPI, Winapi.Messages, Vcl.StdCtrls, Vcl.Controls,
  Vcl.Forms, Vcl.ExtDlgs, System.Classes;

type
  [TestFixture]
  TParserTests = class
    FForm: TForm1;
    FExcelPath: string;
    FIsFileCreate: Boolean;
    FSavedTSVFilePath: string;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure TestCreateTSVButtonClick;
    [Test]
    procedure TestWriteDFNButtonClick;
    [Test]
    procedure TestOpenDFNClick;
    [Test]
    procedure TestePathDFNKeyPress_FileExists;
    [Test]
    procedure TestePathDFNKeyPress_FileNotExists;
    [Test]
    procedure TestOpenTSVClick;
    [Test]
    procedure TestePathTSVKeyPress_FileExists;
    [Test]
    procedure TestePathTSVKeyPress_FileNotExists;
    [Test]
    procedure TestOpenDFN2Click;
    [Test]
    procedure TestePathDFN2KeyPress_FileExists;
    [Test]
    procedure TestePathDFN2KeyPress_FileNotExists;
    [Test]
    procedure TestOpenExcel_NoFileCreated;
    [Test]
    procedure TestOpenCalc_NoFileCreated;
    [Test]
    procedure TestFormCreate;
    [Test]
    procedure TestContainsCyrillicCharacters_True;
    [Test]
    procedure TestContainsCyrillicCharacters_False;
    [Test]
    procedure TestIsCharCyrillic_True;
    [Test]
    procedure TestIsCharCyrillic_False;
    [Test]
    procedure TestExtractField_FoundWithQuotes;
    [Test]
    procedure TestExtractField_FoundWithoutQuotes;
    [Test]
    procedure TestExtractField_NotFound;
    [Test]
    procedure TestFormDestroy;
  end;

implementation

procedure TParserTests.Setup;
begin
  Application.Initialize; // Ініціалізуємо додаток
  FForm := TForm1.Create(Application); // Створюємо екземпляр форми
  FForm.IsFileCreate := FIsFileCreate;
  FForm.SavedTSVFilePath := FSavedTSVFilePath;
end;

procedure TParserTests.TearDown;
begin
  FForm.Free; // Звільняємо екземпляр форми
end;

procedure TParserTests.TestCreateTSVButtonClick;
var
  OutputFilePath: string;
  ExpectedOutput: TStringList;
begin
  // Встановлюємо потрібні дані та налаштування перед тестом
  FForm.OpenedDFNFilePath := 'D:\Programs\Embarcadero\ТЗ\UnitTest.dfn';
  FForm.OriginalDFNFileContent := TStringList.Create;
  FForm.OriginalDFNFileContent.Add('ID:000' + #4 + 'Orig:111' + #4 + 'Curr:222');
  FForm.OriginalDFNFileContent.Add('ID:1' + #4 + 'Orig:Hello' + #4 + 'Curr:Привіт');

  // Викликаємо процедуру, яку хочемо протестувати
  FForm.bCreateTSVClick(nil);

  // Очікувані результати
  OutputFilePath := 'D:\Programs\Embarcadero\ТЗ\Unit1.tsv';
  ExpectedOutput := TStringList.Create;
  ExpectedOutput.Add('ID' + #9 + 'Значения' + #9 + 'Перевод');
  ExpectedOutput.Add('1' + #9 + 'Hello' + #9 + 'Привіт');

  // Завантажуємо створений файл та порівнюємо з очікуваним результатом
  var ActualOutput := TStringList.Create;
  ActualOutput.LoadFromFile(FForm.sdSaveTSV.FileName);
  Assert.AreEqual(ExpectedOutput.Text, ActualOutput.Text);

  // Прибираємо за собою
  ActualOutput.Free;
  ExpectedOutput.Free;
  DeleteFile(FForm.sdSaveTSV.FileName);
end;

procedure TParserTests.TestWriteDFNButtonClick;
begin
  // Встановлюємо потрібні дані та налаштування перед тестом
  FForm.OpenedTSVFilePath := 'D:\Programs\Embarcadero\ТЗ\16.tsv';
  FForm.OpenedDFN2FilePath := 'D:\Programs\Embarcadero\ТЗ\16.dfn';

  FForm.OriginalTSVFileContent := TStringList.Create;
  FForm.OriginalTSVFileContent.Add('456' + #9 + 'Значения' + #9 + 'Перевод');
  FForm.OriginalTSVFileContent.Add('001' + #9 + 'Строка' + #9 + 'Рядок');

  FForm.OriginalDFN2FileContent := TStringList.Create;
  FForm.OriginalDFN2FileContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Значения''');
  FForm.OriginalDFN2FileContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:');
  FForm.OriginalDFN2FileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

  // Викликаємо процедуру, яку хочемо протестувати
  FForm.bWriteDFNClick(nil);

  // Очікувані результати
  var ExpectedDFNContent := TStringList.Create;
  ExpectedDFNContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Перевод''');
  ExpectedDFNContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:''Рядок''');
  ExpectedDFNContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

  // Завантажуємо створений файл та порівнюємо з очікуваним результатом
  var ActualDFNContent := TStringList.Create;
  ActualDFNContent.LoadFromFile(FForm.OpenedDFN2FilePath);
  Assert.AreEqual(ExpectedDFNContent.Text, ActualDFNContent.Text);

  // Прибираємо за собою
  ActualDFNContent.Free;
  ExpectedDFNContent.Free;
  DeleteFile(FForm.OpenedDFN2FilePath);
end;

procedure TParserTests.TestOpenDFNClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // Встановлюємо потрібні дані та налаштування перед тестом
  TestFileName := 'D:\Programs\Embarcadero\ТЗ\OpenFile.dfn';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // Викликаємо процедуру, яку хочемо протестувати
  FForm.odOpenDFN.FileName := TestFileName; // Симулюємо натискання кнопки та передаємо ім'я тестового файлу
  FForm.spChooseDFNClick(nil);

  // Перевірка, чи поля форми мають правильні значення після виконання процедури
  Assert.AreEqual(TestFileName, FForm.ePathDFN.Text);
  Assert.IsNotNull(FForm.OriginalDFNFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalDFNFileContent.Text);

  // Прибираємо за собою
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathDFNKeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\ТЗ\Unit1.dfn';
  FForm.ePathDFN.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent не є nil
  Assert.IsNotNull(FForm.OriginalDFNFileContent);

  // Перевірка, чи OpenedDFNFilePath містить правильний шлях
  Assert.AreEqual(TestFilePath, FForm.OpenedDFNFilePath);
end;

procedure TParserTests.TestePathDFNKeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFN.Text := 'D:\Programs\Embarcadero\ТЗ\Unit112.dfn';
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent є nil
  Assert.IsNull(FForm.OriginalDFNFileContent);
end;

procedure TParserTests.TestOpenTSVClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // Встановлюємо потрібні дані та налаштування перед тестом
  TestFileName := 'D:\Programs\Embarcadero\ТЗ\OpenFile.tvs';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // Викликаємо процедуру, яку хочемо протестувати
  FForm.odOpenTSV.FileName := TestFileName; // Симулюємо натискання кнопки та передаємо ім'я тестового файлу
  FForm.spChooseTSVClick(nil);

  // Перевірка, чи поля форми мають правильні значення після виконання процедури
  Assert.AreEqual(TestFileName, FForm.ePathTSV.Text);
  Assert.IsNotNull(FForm.OriginalTSVFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalTSVFileContent.Text);

  // Прибираємо за собою
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathTSVKeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\ТЗ\Unit1.tsv';
  FForm.ePathDFN.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent не є nil
  Assert.IsNotNull(FForm.OriginalDFNFileContent);

  // Перевірка, чи OpenedDFNFilePath містить правильний шлях
  Assert.AreEqual(TestFilePath, FForm.OpenedDFNFilePath);
end;

procedure TParserTests.TestePathTSVKeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFN.Text := 'D:\Programs\Embarcadero\ТЗ\Unit112.tsv';
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent є nil
  Assert.IsNull(FForm.OriginalDFNFileContent);
end;

procedure TParserTests.TestOpenDFN2Click;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // Встановлюємо потрібні дані та налаштування перед тестом
  TestFileName := 'D:\Programs\Embarcadero\ТЗ\OpenFile2.dfn';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // Викликаємо процедуру, яку хочемо протестувати
  FForm.odOpenDFN2.FileName := TestFileName; // Симулюємо натискання кнопки та передаємо ім'я тестового файлу
  FForm.spChooseDFN2Click(nil);

  // Перевірка, чи поля форми мають правильні значення після виконання процедури
  Assert.AreEqual(TestFileName, FForm.ePathDFN2.Text);
  Assert.IsNotNull(FForm.OriginalDFN2FileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalDFN2FileContent.Text);

  // Прибираємо за собою
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathDFN2KeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\ТЗ\Unit1.dfn';
  FForm.ePathDFN2.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFN2KeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent не є nil
  Assert.IsNotNull(FForm.OriginalDFN2FileContent);

  // Перевірка, чи OpenedDFNFilePath містить правильний шлях
  Assert.AreEqual(TestFilePath, FForm.OpenedDFN2FilePath);
end;

procedure TParserTests.TestePathDFN2KeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFN2.Text := 'D:\Programs\Embarcadero\ТЗ\Unit112.tsv';
  Key := #13;

  FForm.ePathDFN2KeyPress(nil, Key);

  // Перевірка, чи OriginalDFNFileContent є nil
  Assert.IsNull(FForm.OriginalDFN2FileContent);
end;

procedure TParserTests.TestOpenExcel_NoFileCreated;
var
  Msg: string;
begin
  FIsFileCreate := False; // Встановлюємо, що файл не був створений
  FSavedTSVFilePath := ''; // Порожній шлях до TSV файлу

  FForm.bOpenExcelClick(nil); // Виклик процедури для тестування

  // Очікуване повідомлення про необхідність створення TSV файлу
  Msg := 'Спочатку створіть TSV файл';
  Assert.AreEqual(Msg, FForm.ShowWarningMessage);
end;

procedure TParserTests.TestOpenCalc_NoFileCreated;
var
  Msg: string;
begin
  FIsFileCreate := False; // Встановлюємо, що файл не був створений
  FSavedTSVFilePath := ''; // Порожній шлях до TSV файлу

  FForm.bOpenCalcClick(nil); // Виклик процедури для тестування

  // Очікуване повідомлення про необхідність створення TSV файлу
  Msg := 'Спочатку створіть TSV файл';
  Assert.AreEqual(Msg, FForm.ShowWarningMessage);
end;

procedure TParserTests.TestFormCreate;
begin
  // Викликаємо процедуру FormCreate
  FForm.FormCreate(nil);

  // Перевіряємо, чи змінна IsFileCreate має значення false після виконання процедури
  Assert.IsFalse(FForm.IsFileCreate);
end;

procedure TParserTests.TestContainsCyrillicCharacters_True;
begin
  // Передаємо рядок, який містить кириличні символи
  Assert.IsTrue(FForm.ContainsCyrillicCharacters('Привіт, світе!'));
end;

procedure TParserTests.TestContainsCyrillicCharacters_False;
begin
  // Передаємо рядок, який не містить кириличних символів
  Assert.IsFalse(FForm.ContainsCyrillicCharacters('Hello, world!'));
end;

procedure TParserTests.TestIsCharCyrillic_True;
begin
  // Передаємо символи, які є кириличними
  Assert.IsTrue(FForm.IsCharCyrillic('А'));
  Assert.IsTrue(FForm.IsCharCyrillic('а'));
  Assert.IsTrue(FForm.IsCharCyrillic('Я'));
  Assert.IsTrue(FForm.IsCharCyrillic('я'));
  Assert.IsTrue(FForm.IsCharCyrillic('Ё'));
  Assert.IsTrue(FForm.IsCharCyrillic('ё'));
end;

procedure TParserTests.TestIsCharCyrillic_False;
begin
  // Передаємо символи, які не є кириличними
  Assert.IsFalse(FForm.IsCharCyrillic('A'));
  Assert.IsFalse(FForm.IsCharCyrillic('B'));
  Assert.IsFalse(FForm.IsCharCyrillic('1'));
  Assert.IsFalse(FForm.IsCharCyrillic(' '));
  Assert.IsFalse(FForm.IsCharCyrillic('.'));
end;

procedure TParserTests.TestExtractField_FoundWithQuotes;
var
  Fields: TArray<string>;
  FieldName: string;
  FieldValue: string;
begin
  // Вхідні дані
  Fields := ['ID:001', 'Name:''John''', 'Age:25'];
  FieldName := 'Name';

  // Викликаємо функцію
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('John', FieldValue);
end;

procedure TParserTests.TestExtractField_FoundWithoutQuotes;
var
  Fields: TArray<string>;
  FieldName: string;
  FieldValue: string;
begin
  // Вхідні дані
  Fields := ['ID:001', 'Name:John', 'Age:25'];
  FieldName := 'Name';

  // Викликаємо функцію
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('John', FieldValue);
end;

procedure TParserTests.TestExtractField_NotFound;
var
  Fields: TArray<string>;
  FieldName: string;
  FieldValue: string;
begin
  // Вхідні дані
  Fields := ['ID:001', 'Name:''John''', 'Age:25'];
  FieldName := 'Email';

  // Викликаємо функцію
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('', FieldValue);
end;

procedure TParserTests.TestFormDestroy;
begin
  // Викликаємо процедуру FormDestroy
  FForm.FormDestroy(nil);

  // Перевіряємо, чи об'єкти звільнились
  Assert.IsNull(FForm.OriginalDFNFileContent);
  Assert.IsNull(FForm.OriginalTSVFileContent);
  Assert.IsNull(FForm.OriginalDFN2FileContent);
end;

initialization
  TDUnitX.RegisterTestFixture(TParserTests);
end.
