unit Unit1;

interface

uses
  DUnitX.TestFramework, ufMain, Vcl.Dialogs, Winapi.Windows, Vcl.Buttons, System.SysUtils,
  Vcl.Graphics, System.Variants, ShellAPI, Winapi.Messages, Vcl.StdCtrls, Vcl.Controls,
  Vcl.Forms, Vcl.ExtDlgs, System.Classes;

type
  [TestFixture]
  TParserTests = class
    FForm: TMainForm;
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
  Application.Initialize; // ����������� �������
  FForm := TMainForm.Create(Application); // ��������� ��������� �����
  FForm.IsFileCreate := FIsFileCreate;
  FForm.SavedTSVFilePath := FSavedTSVFilePath;
end;

procedure TParserTests.TearDown;
begin
  FForm.Free; // ��������� ��������� �����
end;

procedure TParserTests.TestCreateTSVButtonClick;
var
  OutputFilePath: string;
  ExpectedOutput: TStringList;
begin
  // ������������ ������� ���� �� ������������ ����� ������
  FForm.OpenedDFNFilePath := 'D:\Programs\Embarcadero\��\UnitTest.dfn';
  FForm.OriginalDFNFileContent := TStringList.Create;
  FForm.OriginalDFNFileContent.Add('ID:000' + #4 + 'Orig:111' + #4 + 'Curr:222');
  FForm.OriginalDFNFileContent.Add('ID:1' + #4 + 'Orig:Hello' + #4 + 'Curr:�����');

  // ��������� ���������, ��� ������ ������������
  FForm.bCreateTSVClick(nil);

  // ��������� ����������
  OutputFilePath := 'D:\Programs\Embarcadero\��\Unit1.tsv';
  ExpectedOutput := TStringList.Create;
  ExpectedOutput.Add('ID' + #9 + '��������' + #9 + '�������');
  ExpectedOutput.Add('1' + #9 + 'Hello' + #9 + '�����');

  // ����������� ��������� ���� �� ��������� � ���������� �����������
  var ActualOutput := TStringList.Create;
  ActualOutput.LoadFromFile(FForm.sdSaveTSV.FileName);
  Assert.AreEqual(ExpectedOutput.Text, ActualOutput.Text);

  // ��������� �� �����
  ActualOutput.Free;
  ExpectedOutput.Free;
  DeleteFile(FForm.sdSaveTSV.FileName);
end;

procedure TParserTests.TestWriteDFNButtonClick;
begin
  // ������������ ������� ���� �� ������������ ����� ������
  FForm.OpenedTSVFilePath := 'D:\Programs\Embarcadero\��\16.tsv';
  FForm.OpenedDFNToModFilePath := 'D:\Programs\Embarcadero\��\16.dfn';

  FForm.OriginalTSVFileContent := TStringList.Create;
  FForm.OriginalTSVFileContent.Add('456' + #9 + '��������' + #9 + '�������');
  FForm.OriginalTSVFileContent.Add('001' + #9 + '������' + #9 + '�����');

  FForm.OriginalDFNToModFileContent := TStringList.Create;
  FForm.OriginalDFNToModFileContent.Add('ID:456' + #4 + 'Orig:''��������''' + #4 + 'Curr:''��������''');
  FForm.OriginalDFNToModFileContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:');
  FForm.OriginalDFNToModFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

  // ��������� ���������, ��� ������ ������������
  FForm.bWriteDFNClick(nil);

  // ��������� ����������
  var ExpectedDFNContent := TStringList.Create;
  ExpectedDFNContent.Add('ID:456' + #4 + 'Orig:''��������''' + #4 + 'Curr:''�������''');
  ExpectedDFNContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:''�����''');
  ExpectedDFNContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

  // ����������� ��������� ���� �� ��������� � ���������� �����������
  var ActualDFNContent := TStringList.Create;
  ActualDFNContent.LoadFromFile(FForm.OpenedDFNToModFilePath);
  Assert.AreEqual(ExpectedDFNContent.Text, ActualDFNContent.Text);

  // ��������� �� �����
  ActualDFNContent.Free;
  ExpectedDFNContent.Free;
  DeleteFile(FForm.OpenedDFNToModFilePath);
end;

procedure TParserTests.TestOpenDFNClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // ������������ ������� ���� �� ������������ ����� ������
  TestFileName := 'D:\Programs\Embarcadero\��\OpenFile.dfn';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // ��������� ���������, ��� ������ ������������
  FForm.odOpenDFN.FileName := TestFileName; // ��������� ���������� ������ �� �������� ��'� ��������� �����
  FForm.spChooseDFNClick(nil);

  // ��������, �� ���� ����� ����� ��������� �������� ���� ��������� ���������
  Assert.AreEqual(TestFileName, FForm.ePathDFN.Text);
  Assert.IsNotNull(FForm.OriginalDFNFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalDFNFileContent.Text);

  // ��������� �� �����
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathDFNKeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\��\Unit1.dfn';
  FForm.ePathDFN.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent �� � nil
  Assert.IsNotNull(FForm.OriginalDFNFileContent);

  // ��������, �� OpenedDFNFilePath ������ ���������� ����
  Assert.AreEqual(TestFilePath, FForm.OpenedDFNFilePath);
end;

procedure TParserTests.TestePathDFNKeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFN.Text := 'D:\Programs\Embarcadero\��\Unit112.dfn';
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent � nil
  Assert.IsNull(FForm.OriginalDFNFileContent);
end;

procedure TParserTests.TestOpenTSVClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // ������������ ������� ���� �� ������������ ����� ������
  TestFileName := 'D:\Programs\Embarcadero\��\OpenFile.tvs';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // ��������� ���������, ��� ������ ������������
  FForm.odOpenTSV.FileName := TestFileName; // ��������� ���������� ������ �� �������� ��'� ��������� �����
  FForm.spChooseTSVClick(nil);

  // ��������, �� ���� ����� ����� ��������� �������� ���� ��������� ���������
  Assert.AreEqual(TestFileName, FForm.ePathTSV.Text);
  Assert.IsNotNull(FForm.OriginalTSVFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalTSVFileContent.Text);

  // ��������� �� �����
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathTSVKeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\��\Unit1.tsv';
  FForm.ePathDFN.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent �� � nil
  Assert.IsNotNull(FForm.OriginalDFNFileContent);

  // ��������, �� OpenedDFNFilePath ������ ���������� ����
  Assert.AreEqual(TestFilePath, FForm.OpenedDFNFilePath);
end;

procedure TParserTests.TestePathTSVKeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFN.Text := 'D:\Programs\Embarcadero\��\Unit112.tsv';
  Key := #13;

  FForm.ePathDFNKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent � nil
  Assert.IsNull(FForm.OriginalDFNFileContent);
end;

procedure TParserTests.TestOpenDFN2Click;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // ������������ ������� ���� �� ������������ ����� ������
  TestFileName := 'D:\Programs\Embarcadero\��\OpenFile2.dfn';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // ��������� ���������, ��� ������ ������������
  FForm.odOpenDFNToMod.FileName := TestFileName; // ��������� ���������� ������ �� �������� ��'� ��������� �����
  FForm.spChooseDFNToModClick(nil);

  // ��������, �� ���� ����� ����� ��������� �������� ���� ��������� ���������
  Assert.AreEqual(TestFileName, FForm.ePathDFNToMod.Text);
  Assert.IsNotNull(FForm.OriginalDFNToModFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalDFNToModFileContent.Text);

  // ��������� �� �����
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestePathDFN2KeyPress_FileExists;
var
  TestFilePath: string;
  Key: Char;
begin
  TestFilePath := 'D:\Programs\Embarcadero\��\Unit1.dfn';
  FForm.ePathDFNToMod.Text := TestFilePath;
  Key := #13;

  FForm.ePathDFNToModKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent �� � nil
  Assert.IsNotNull(FForm.OriginalDFNToModFileContent);

  // ��������, �� OpenedDFNFilePath ������ ���������� ����
  Assert.AreEqual(TestFilePath, FForm.OpenedDFNToModFilePath);
end;

procedure TParserTests.TestePathDFN2KeyPress_FileNotExists;
var
  Key: Char;
begin
  FForm.ePathDFNToMod.Text := 'D:\Programs\Embarcadero\��\Unit112.tsv';
  Key := #13;

  FForm.ePathDFNToModKeyPress(nil, Key);

  // ��������, �� OriginalDFNFileContent � nil
  Assert.IsNull(FForm.OriginalDFNToModFileContent);
end;

procedure TParserTests.TestOpenExcel_NoFileCreated;
var
  Msg: string;
begin
  FIsFileCreate := False; // ������������, �� ���� �� ��� ���������
  FSavedTSVFilePath := ''; // �������� ���� �� TSV �����

  FForm.bOpenExcelClick(nil); // ������ ��������� ��� ����������

  // ��������� ����������� ��� ������������ ��������� TSV �����
  Msg := '�������� ������� TSV ����';
  Assert.AreEqual(Msg, FForm.ShowWarningMessage);
end;

procedure TParserTests.TestOpenCalc_NoFileCreated;
var
  Msg: string;
begin
  FIsFileCreate := False; // ������������, �� ���� �� ��� ���������
  FSavedTSVFilePath := ''; // �������� ���� �� TSV �����

  FForm.bOpenCalcClick(nil); // ������ ��������� ��� ����������

  // ��������� ����������� ��� ������������ ��������� TSV �����
  Msg := '�������� ������� TSV ����';
  Assert.AreEqual(Msg, FForm.ShowWarningMessage);
end;

procedure TParserTests.TestFormCreate;
begin
  // ��������� ��������� FormCreate
  FForm.FormCreate(nil);

  // ����������, �� ����� IsFileCreate �� �������� false ���� ��������� ���������
  Assert.IsFalse(FForm.IsFileCreate);
end;

procedure TParserTests.TestContainsCyrillicCharacters_True;
begin
  // �������� �����, ���� ������ ��������� �������
  Assert.IsTrue(FForm.ContainsCyrillicCharacters('�����, ����!'));
end;

procedure TParserTests.TestContainsCyrillicCharacters_False;
begin
  // �������� �����, ���� �� ������ ���������� �������
  Assert.IsFalse(FForm.ContainsCyrillicCharacters('Hello, world!'));
end;

procedure TParserTests.TestIsCharCyrillic_True;
begin
  // �������� �������, �� � �����������
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
  Assert.IsTrue(FForm.IsCharCyrillic('�'));
end;

procedure TParserTests.TestIsCharCyrillic_False;
begin
  // �������� �������, �� �� � �����������
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
  // ������ ����
  Fields := ['ID:001', 'Name:''John''', 'Age:25'];
  FieldName := 'Name';

  // ��������� �������
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('John', FieldValue);
end;

procedure TParserTests.TestExtractField_FoundWithoutQuotes;
var
  Fields: TArray<string>;
  FieldName: string;
  FieldValue: string;
begin
  // ������ ����
  Fields := ['ID:001', 'Name:John', 'Age:25'];
  FieldName := 'Name';

  // ��������� �������
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('John', FieldValue);
end;

procedure TParserTests.TestExtractField_NotFound;
var
  Fields: TArray<string>;
  FieldName: string;
  FieldValue: string;
begin
  // ������ ����
  Fields := ['ID:001', 'Name:''John''', 'Age:25'];
  FieldName := 'Email';

  // ��������� �������
  FieldValue := FForm.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('', FieldValue);
end;

procedure TParserTests.TestFormDestroy;
begin
  // ��������� ��������� FormDestroy
  FForm.FormDestroy(nil);

  // ����������, �� ��'���� ����������
  Assert.IsNull(FForm.OriginalDFNFileContent);
  Assert.IsNull(FForm.OriginalTSVFileContent);
  Assert.IsNull(FForm.OriginalDFNToModFileContent);
end;

initialization
  TDUnitX.RegisterTestFixture(TParserTests);
end.
