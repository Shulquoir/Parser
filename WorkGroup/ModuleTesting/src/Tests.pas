unit Tests;

interface

uses
  DUnitX.TestFramework, Classes, Dialogs;

type
  [TestFixture]
  TParserTests = class
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure TestConvertDFNToTSV;
    [Test]
    procedure TestRewriteDFN;
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
  end;

implementation

uses
  SysUtils, Parsing;

procedure TParserTests.Setup;
begin
end;

procedure TParserTests.TearDown;
begin
end;

//------------------------------------------------------------------------------
//    Тестування функції створення TSV файлу на основі відкритого DFN файлу
//------------------------------------------------------------------------------
procedure TParserTests.TestConvertDFNToTSV;
var
  ProjectPath, OpenedDFNFilePath: String;
  TestDFNFileContent, OutputFileContent: TStringList;
  FileCreate: TFileStream;

begin
  // Проводимо налаштування перед тестом
  ProjectPath := GetCurrentDir; // Отримуємо шлях до папки з exe файлом проекту, тобто Parser\WorkGroup\bin
  OpenedDFNFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnit.dfn'; // Шлях до файлу, що буде створений

  try
    FileCreate := TFileStream.Create(OpenedDFNFilePath, fmCreate); // Створюємо файл в кореневій папці проекту bin
  finally                                                                                 // з іменем TestUnit.dfn
    FreeAndNil(FileCreate);
  end;

  // Встановлюємо тестові значення та записуємо в створений DFN файл
  TestDFNFileContent := TStringList.Create;
  TestDFNFileContent.Add('ID:000' + #4 + 'Orig:111' + #4 + 'Curr:222');
  TestDFNFileContent.Add('ID:1:50' + #4 + 'Orig:line' + #4 + 'Curr:');
  TestDFNFileContent.Add('ID:1' + #4 + 'Orig:Hello' + #4 + 'Curr:Привіт');
  TestDFNFileContent.SaveToFile(OpenedDFNFilePath);

  //Створюємо TStringList з очікуваним результатом Викликаємо процедуру, яку хочемо протестувати
  OutputFileContent := TStringList.Create;
  TParsing.ConvertDFNToTSV(OpenedDFNFilePath, OutputFileContent);

  // Створюємо TStringList з очікуваним результатом
  var ExpectedOutput := TStringList.Create;
  ExpectedOutput.Add('ID' + #9 + 'Значения' + #9 + 'Перевод');
  ExpectedOutput.Add('1' + #9 + 'Hello' + #9 + 'Привіт');

  // Порівнюємо результат з очікуваним результатом
  Assert.AreEqual(ExpectedOutput.Text, OutputFileContent.Text);

  // Прибираємо за собою
  FreeAndNil(ExpectedOutput);
  FreeAndNil(TestDFNFileContent);
  FreeAndNil(OutputFileContent);
  DeleteFile(OpenedDFNFilePath);
end;

//------------------------------------------------------------------------------
//            Тестування функції перезапису відкритого DFN файлу
//------------------------------------------------------------------------------
procedure TParserTests.TestRewriteDFN;
var
  ProjectPath, OpenedTSVFilePath, OpenedDFNToModFilePath: String;
  TestTSVFileContent, TestDFNToModFileContent: TStringList;
  TSVFileCreate, DFNFileCreate, DifTSVFileCreate : TFileStream;

begin
  // Проводимо налаштування перед тестом
  ProjectPath := GetCurrentDir; // Отримуємо шлях до папки з exe файлом проекту, тобто Parser\WorkGroup\bin
  OpenedTSVFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnit1.tsv'; // Шлях до файлу TSV, що буде створений
  OpenedDFNToMoDFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnit1.dfn'; // Шлях до файлу DFN, що буде створений

  try
    TSVFileCreate := TFileStream.Create(OpenedTSVFilePath, fmCreate); // Створюємо файл TestUnit1.tsv в папці bin
    DFNFileCreate := TFileStream.Create(OpenedDFNToModFilePath, fmCreate); // Створюємо файл TestUnit1.dfn в папці bin
  finally
    FreeAndNil(TSVFileCreate);
    FreeAndNil(DFNFileCreate);
  end;

  // Встановлюємо тестові значення та записуємо в створений TSV файл
  TestTSVFileContent := TStringList.Create;
  TestTSVFileContent.Add('456' + #9 + 'Значения' + #9 + 'Перевод');
  TestTSVFileContent.Add('001' + #9 + 'Строка' + #9 + 'Рядок');
  TestTSVFileContent.SaveToFile(OpenedTSVFilePath);

  // Встановлюємо тестові значення та записуємо в створений DFN файл
  TestDFNToModFileContent := TStringList.Create;
  TestDFNToModFileContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Значения''');
  TestDFNToModFileContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:');
  TestDFNToModFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');
  TestDFNToModFileContent.SaveToFile(OpenedDFNToMoDFilePath);

  // Викликаємо процедуру, яку хочемо протестувати
  TParsing.ModifyDFN(OpenedTSVFilePath, OpenedDFNToMoDFilePath);

  // Очікувані результати
  var ExpectedDFNContent := TStringList.Create;
  ExpectedDFNContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Перевод''');
  ExpectedDFNContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:''Рядок''');
  ExpectedDFNContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

  // Завантажуємо створений файл та порівнюємо з очікуваним результатом
  var ActualDFNContent := TStringList.Create;
  ActualDFNContent.LoadFromFile(OpenedDFNToModFilePath);
  Assert.AreEqual(ExpectedDFNContent.Text, ActualDFNContent.Text);

  // Прибираємо за собою
  FreeAndNil(ExpectedDFNContent);
  FreeAndNil(ActualDFNContent);
  FreeAndNil(TestTSVFileContent);
  FreeAndNil(TestDFNToModFileContent);
  DeleteFile(OpenedTSVFilePath);
  DeleteFile(OpenedDFNToModFilePath);
end;

//------------------------------------------------------------------------------
//      Тестування функцій для обробки рядків, полів та символів файлу
//------------------------------------------------------------------------------
procedure TParserTests.TestContainsCyrillicCharacters_True;
begin
  // Передаємо рядок, який містить кириличні символи
  Assert.IsTrue(TParsing.ContainsCyrillicCharacters('Привіт, світе!'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestContainsCyrillicCharacters_False;
begin
  // Передаємо рядок, який не містить кириличних символів
  Assert.IsFalse(TParsing.ContainsCyrillicCharacters('Hello, world!'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestIsCharCyrillic_True;
begin
  // Передаємо символи, які є кириличними
  Assert.IsTrue(TParsing.IsCharCyrillic('А'));
  Assert.IsTrue(TParsing.IsCharCyrillic('а'));
  Assert.IsTrue(TParsing.IsCharCyrillic('Я'));
  Assert.IsTrue(TParsing.IsCharCyrillic('я'));
  Assert.IsTrue(TParsing.IsCharCyrillic('Ё'));
  Assert.IsTrue(TParsing.IsCharCyrillic('ё'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestIsCharCyrillic_False;
begin
  // Передаємо символи, які не є кириличними
  Assert.IsFalse(TParsing.IsCharCyrillic('A'));
  Assert.IsFalse(TParsing.IsCharCyrillic('B'));
  Assert.IsFalse(TParsing.IsCharCyrillic('1'));
  Assert.IsFalse(TParsing.IsCharCyrillic(' '));
  Assert.IsFalse(TParsing.IsCharCyrillic('.'));
end;

//------------------------------------------------------------------------------

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
  FieldValue := TParsing.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('John', FieldValue);
end;

//------------------------------------------------------------------------------

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
  FieldValue := TParsing.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('John', FieldValue);
end;

//------------------------------------------------------------------------------

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
  FieldValue := TParsing.ExtractField(Fields, FieldName);

  // Перевірка
  Assert.AreEqual('', FieldValue);
end;

//-------------------------------------TODO-------------------------------------
//------------------------------------------------------------------------------
// Тестування функціоналу відкриття DFN файлу, створення форматованого TSV файлу
//           та відкриття створеного TSV файлу в сторонніх програмах
//------------------------------------------------------------------------------
{procedure TParserTests.TestCreateTSV;
const
  ENTER_KEY_CODE = #13;

var
  Key: Char;
  ProjectPath, FilePath, OutputFilePath: string;
  FileCreate: TFileStream;
  TestFileContent, ExpectedOutput: TStringList;

begin
  ProjectPath := GetCurrentDir; // Отримуємо шлях до папки з exe файлом проекту, тобто Parser\Win32\Debug

//------------Перевірка, якщо був введений шлях до неіснуючого файлу------------

  FilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'nonexistentfile.dfn'; // Шлях до неіснуючого файлу
  FForm.ePathDFN.Text := FilePath; // Імітуємо введення шляху до не існуючого файлу в елементі edit
  Assert.IsFalse(FForm.ePathDFNKeyPress(FForm.ePathDFN, #13),
    'OpenFileByName має повернути значення False для неіснуючого файлу.');

//-------------Перевірка, якщо був введений шлях до існуючого файлу-------------

  FilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnit.dfn'; // Шлях до файлу, що буде створений
  try
    FileCreate := TFileStream.Create(FilePath, fmCreate); // Створюємо файл в кореневій папці проекту Debug
  finally                                                                          // з іменем TestUnit.dfn
    FreeAndNil(FileCreate);
  end;

  FForm.ePathDFN.Text := FilePath; // Імітуємо введення шляху до створеного файлу в елементі edit
  Assert.IsTrue(FForm.ePathDFNKeyPress(FForm.ePathDFN, ENTER_KEY_CODE),
    'OpenFileByName має повернути значення True для існуючого файлу.');

//------------------------------Відкриття DFN файлу-----------------------------

  Key := ENTER_KEY_CODE;
  FForm.ePathDFNKeyPress(FForm.ePathDFN, Key);

//----------------Записуємо в створений DFN файл тестові значення---------------

  TestFileContent := TStringList.Create;
  TestFileContent.Add('ID:000' + #4 + 'Orig:''111''' + #4 + 'Curr:''222''');
  TestFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Привет''');
  TestFileContent.SaveToFile(FilePath);

//------------------------------Створення TSV файлу-----------------------------

  TController.CreateTSV(FilePath, OutputFilePath); // Створюємо файл таким чином, щоб отримати шлях
                                                                                //збереженого файлу
//----------------Створення TStringList з очікуваним результатом----------------

  ExpectedOutput := TStringList.Create;
  ExpectedOutput.Add('ID' + #9 + 'Значения' + #9 + 'Перевод');
  ExpectedOutput.Add('1' + #9 + 'Hello' + #9 + 'Привет');

//------Завантажуємо створений файл та порівнюємо з очікуваним результатом------

  var ActualOutput := TStringList.Create;
  ActualOutput.LoadFromFile(OutputFilePath);
  Assert.AreEqual(ExpectedOutput.Text, ActualOutput.Text);

// Перевірка відкриття TSV файлу в сторонній програмі, якщо не був створений TSV файл

  Assert.IsFalse(TController.OpenByProgram(OutputFilePath, False, 'Excel'),
    'ShellExecute не має бути викликано.');

// Перевірка відкриття TSV файлу в сторонній програмі, якщо був створений TSV файл

  Assert.IsTrue(TController.OpenByProgram(OutputFilePath, True, 'Excel'),
    'ShellExecute має бути викликано.');

// Перевірка відкриття TSV файлу в сторонній програмі, якщо був переданий неіснуючий шлях файлу

  Assert.IsFalse(TController.OpenByProgram('nonexistentfile.tsv', True, 'Calc'),
    'ShellExecute не має бути викликано.');

//------------------------------Прибираємо за собою-----------------------------

  FreeAndNil(ActualOutput);
  FreeAndNil(ExpectedOutput);
  DeleteFile(FilePath); // Видаляємо створений DFN файл
  DeleteFile(OutputFilePath); // Видаляємо створений TSV файл
end;

//------------------------------------------------------------------------------
// Тестування функціоналу відкриття TSV та DFN файлів, створення редагованого
//                       DFN файлу на основі TSV файлу
//------------------------------------------------------------------------------
procedure TParserTests.RewriteDFN;
var
  ProjectPath, TSVFilePath, DifTSVFilePath, DFNFilePath: string;
  TSVFileCreate, DFNFileCreate, DifTSVFileCreate : TFileStream;
  TestTSVFileContent, TestDFNFileContent, ExpectedOutput: TStringList;

begin
  ProjectPath := GetCurrentDir; // Отримуємо шлях до папки з exe файлом проекту, тобто Parser\Win32\Debug

//----------------------------Створюємо тестові файли---------------------------

  TSVFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithSameName.tsv'; // Шлях до файлу, що буде створений
  DFNFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithSameName.dfn'; // Шлях до файлу, що буде створений
  DifTSVFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithDifName.tsv'; // Шлях до файлу, що буде створений

  try
    TSVFileCreate := TFileStream.Create(TSVFilePath, fmCreate); // Створюємо файл в кореневій папці проекту Debug
    DFNFileCreate := TFileStream.Create(DFNFilePath, fmCreate); // Створюємо файл в кореневій папці проекту Debug
    DifTSVFileCreate := TFileStream.Create(DifTSVFilePath, fmCreate); // Створюємо файл в кореневій папці проекту Debug
  finally
    FreeAndNil(TSVFileCreate);
    FreeAndNil(DFNFileCreate);
    FreeAndNil(DifTSVFileCreate);
  end;

//---------------------Передаємо різні імена відкритих файлів-------------------

  FForm.odOpenTSV.FileName := DifTSVFilePath;
  FForm.spChooseTSVClick(nil);
  FForm.odOpenDFNToMod.FileName := DFNFilePath;
  FForm.spChooseDFNToModClick(nil);

//Перевіряємо, що відповідні імена були записані у відповідні поля edit та файл DFN не перезаписався

  Assert.AreEqual(DifTSVFilePath, FForm.ePathTSV.Text);
  Assert.AreEqual(DFNFilePath, FForm.ePathDFNToMod.Text);
  Assert.IsFalse(TController.RewriteDFN(FForm.ePathTSV.Text, FForm.ePathDFNToMod.Text),
   'RewriteDFN має повернути значення Fasle для неспівпадаючих імен файлів.');

//-------------------Передаємо однакові імена відкритих файлів------------------

  FForm.odOpenTSV.FileName := TSVFilePath;
  FForm.spChooseTSVClick(nil);
  FForm.odOpenDFNToMod.FileName := DFNFilePath;
  FForm.spChooseDFNToModClick(nil);

//------Перевірка, що відповідні імена були записані у відповідні поля edit-----

  Assert.AreEqual(TSVFilePath, FForm.ePathTSV.Text);
  Assert.AreEqual(DFNFilePath, FForm.ePathDFNToMod.Text);

//----------------Записуємо в створений TSV файл тестові значення---------------

  TestTSVFileContent := TStringList.Create;
  TestTSVFileContent.Add('456' + #9 + 'Значения' + #9 + 'Перевод');
  TestTSVFileContent.Add('001' + #9 + 'Строка' + #9 + 'Рядок');
  TestTSVFileContent.SaveToFile(TSVFilePath);

//----------------Записуємо в створений DFN файл тестові значення---------------

  TestDFNFileContent := TStringList.Create;
  TestDFNFileContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Значения''');
  TestDFNFileContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:');
  TestDFNFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');
  TestDFNFileContent.SaveToFile(DFNFilePath);

//------------------------------Перезапис DFN файлу-----------------------------

  FForm.bWriteDFNClick(nil);

//----------------Створення TStringList з очікуваним результатом----------------

  var ExpectedDFNContent := TStringList.Create;
  ExpectedDFNContent.Add('ID:456' + #4 + 'Orig:''Значения''' + #4 + 'Curr:''Перевод''');
  ExpectedDFNContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:''Рядок''');
  ExpectedDFNContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

//------Завантажуємо створений файл та порівнюємо з очікуваним результатом------

  var ActualDFNContent := TStringList.Create;
  ActualDFNContent.LoadFromFile(DFNFilePath);
  Assert.AreEqual(ExpectedDFNContent.Text, ActualDFNContent.Text);

//------------------------------Прибираємо за собою-----------------------------

  FreeAndNil(TestTSVFileContent);
  FreeAndNil(TestDFNFileContent);
  FreeAndNil(ActualDFNContent);
  FreeAndNil(ExpectedDFNContent);
  DeleteFile(TSVFilePath); // Видаляємо створений DFN файл
  DeleteFile(DFNFilePath); // Видаляємо створений TSV файл
  DeleteFile(DifTSVFilePath); // Видаляємо створений TSV файл
end;}
//------------------------------------------------------------------------------

initialization
  TDUnitX.RegisterTestFixture(TParserTests);

end.
