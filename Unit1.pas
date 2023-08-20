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
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    procedure TestCreateTSVButtonClick;
    procedure TestOpenDFNClick;
    procedure TestFormCreate;
  end;

implementation

procedure TParserTests.Setup;
begin
  FForm := TForm1.Create(Application);
end;

procedure TParserTests.TearDown;
begin
  FForm.Free;
end;

procedure TParserTests.TestCreateTSVButtonClick;
var
  OutputFilePath: string;
  ExpectedOutput: TStringList;
begin
  // Встановіть потрібні дані та налаштування перед тестом
  FForm.OpenedDFNFilePath := 'D:\Programs\Embarcadero\ТЗ\Unit1.dfn';
  FForm.OriginalDFNFileContent := TStringList.Create;
  FForm.OriginalDFNFileContent.Add('ID' + #4 + 'Orig' + #4 + 'Curr');
  FForm.OriginalDFNFileContent.Add('1' + #4 + 'Hello' + #4 + 'Привіт');

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

procedure TParserTests.TestOpenDFNClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // Встановіть потрібні дані та налаштування перед тестом
  TestFileName := 'D:\Programs\Embarcadero\ТЗ\Unit1.dfn';
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

procedure TParserTests.TestFormCreate;
begin
  // Викликаємо процедуру FormCreate
  FForm.FormCreate(nil);

  // Перевіряємо, чи змінна IsFileCreate має значення false після виконання процедури
  Assert.IsFalse(FForm.IsFileCreate);
end;

initialization
  TDUnitX.RegisterTestFixture(TParserTests);
end.
