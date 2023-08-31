unit Tests;

interface

uses
  DUnitX.TestFramework, Vcl.Dialogs, Vcl.Forms, Classes, SysUtils, ShellAPI, ufMain, Controller, Model;

type
  [TestFixture]
  TParserTests = class
    FForm: TMainForm;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;
    [Test]
    procedure TestCreateTSV;
    [Test]
    procedure RewriteDFN;
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

procedure TParserTests.Setup;
begin
  FForm := TMainForm.Create(nil);
end;

procedure TParserTests.TearDown;
begin
  FreeAndNil(FForm);
end;

//------------------------------------------------------------------------------
// ���������� ����������� �������� DFN �����, ��������� ������������� TSV �����
//           �� �������� ���������� TSV ����� � ��������� ���������
//------------------------------------------------------------------------------
procedure TParserTests.TestCreateTSV;
const
  ENTER_KEY_CODE = #13;

var
  Key: Char;
  ProjectPath, FilePath, OutputFilePath: string;
  FileCreate: TFileStream;
  TestFileContent, ExpectedOutput: TStringList;

begin
  ProjectPath := GetCurrentDir; // �������� ���� �� ����� � exe ������ �������, ����� Parser\Win32\Debug

//------------��������, ���� ��� �������� ���� �� ���������� �����------------

  FilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'nonexistentfile.dfn'; // ���� �� ���������� �����
  FForm.ePathDFN.Text := FilePath; // ������ �������� ����� �� �� ��������� ����� � ������� edit
  Assert.IsFalse(TController.OpenFileByName(FForm.ePathDFN, ENTER_KEY_CODE),
    'OpenFileByName �� ��������� �������� False ��� ���������� �����.');

//-------------��������, ���� ��� �������� ���� �� ��������� �����-------------

  FilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnit.dfn'; // ���� �� �����, �� ���� ���������
  try
    FileCreate := TFileStream.Create(FilePath, fmCreate); // ��������� ���� � �������� ����� ������� Debug
  finally                                                                          // � ������ TestUnit.dfn
    FreeAndNil(FileCreate);
  end;

  FForm.ePathDFN.Text := FilePath; // ������ �������� ����� �� ���������� ����� � ������� edit
  Assert.IsTrue(TController.OpenFileByName(FForm.ePathDFN, ENTER_KEY_CODE),
    'OpenFileByName �� ��������� �������� True ��� ��������� �����.');

//------------------------------³������� DFN �����-----------------------------

  Key := ENTER_KEY_CODE;
  FForm.ePathDFNKeyPress(FForm.ePathDFN, Key);

//----------------�������� � ��������� DFN ���� ������ ��������---------------

  TestFileContent := TStringList.Create;
  TestFileContent.Add('ID:000' + #4 + 'Orig:''111''' + #4 + 'Curr:''222''');
  TestFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''������''');
  TestFileContent.SaveToFile(FilePath);

//------------------------------��������� TSV �����-----------------------------

  TController.CreateTSV(FilePath, OutputFilePath); // ��������� ���� ����� �����, ��� �������� ����
                                                                                //����������� �����
//----------------��������� TStringList � ���������� �����������----------------

  ExpectedOutput := TStringList.Create;
  ExpectedOutput.Add('ID' + #9 + '��������' + #9 + '�������');
  ExpectedOutput.Add('1' + #9 + 'Hello' + #9 + '������');

//------����������� ��������� ���� �� ��������� � ���������� �����������------

  var ActualOutput := TStringList.Create;
  ActualOutput.LoadFromFile(OutputFilePath);
  Assert.AreEqual(ExpectedOutput.Text, ActualOutput.Text);

// �������� �������� TSV ����� � ��������� �������, ���� �� ��� ��������� TSV ����

  Assert.IsFalse(TController.OpenByProgram(OutputFilePath, False, 'Excel'),
    'ShellExecute �� �� ���� ���������.');

// �������� �������� TSV ����� � ��������� �������, ���� ��� ��������� TSV ����

  Assert.IsTrue(TController.OpenByProgram(OutputFilePath, True, 'Excel'),
    'ShellExecute �� ���� ���������.');

// �������� �������� TSV ����� � ��������� �������, ���� ��� ��������� ��������� ���� �����

  Assert.IsFalse(TController.OpenByProgram('nonexistentfile.tsv', True, 'Calc'),
    'ShellExecute �� �� ���� ���������.');

//------------------------------��������� �� �����-----------------------------

  FreeAndNil(ActualOutput);
  FreeAndNil(ExpectedOutput);
  DeleteFile(FilePath); // ��������� ��������� DFN ����
  DeleteFile(OutputFilePath); // ��������� ��������� TSV ����
end;

//------------------------------------------------------------------------------
// ���������� ����������� �������� TSV �� DFN �����, ��������� ������������
//                       DFN ����� �� ����� TSV �����
//------------------------------------------------------------------------------
procedure TParserTests.RewriteDFN;
var
  ProjectPath, TSVFilePath, DifTSVFilePath, DFNFilePath: string;
  TSVFileCreate, DFNFileCreate, DifTSVFileCreate : TFileStream;
  TestTSVFileContent, TestDFNFileContent, ExpectedOutput: TStringList;

begin
  ProjectPath := GetCurrentDir; // �������� ���� �� ����� � exe ������ �������, ����� Parser\Win32\Debug

//----------------------------��������� ������ �����---------------------------

  TSVFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithSameName.tsv'; // ���� �� �����, �� ���� ���������
  DFNFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithSameName.dfn'; // ���� �� �����, �� ���� ���������
  DifTSVFilePath := IncludeTrailingPathDelimiter(ProjectPath) + 'TestUnitWithDifName.tsv'; // ���� �� �����, �� ���� ���������

  try
    TSVFileCreate := TFileStream.Create(TSVFilePath, fmCreate); // ��������� ���� � �������� ����� ������� Debug
    DFNFileCreate := TFileStream.Create(DFNFilePath, fmCreate); // ��������� ���� � �������� ����� ������� Debug
    DifTSVFileCreate := TFileStream.Create(DifTSVFilePath, fmCreate); // ��������� ���� � �������� ����� ������� Debug
  finally
    FreeAndNil(TSVFileCreate);
    FreeAndNil(DFNFileCreate);
    FreeAndNil(DifTSVFileCreate);
  end;

//---------------------�������� ���� ����� �������� �����-------------------

  FForm.odOpenTSV.FileName := DifTSVFilePath;
  FForm.spChooseTSVClick(nil);
  FForm.odOpenDFNToMod.FileName := DFNFilePath;
  FForm.spChooseDFNToModClick(nil);

//����������, �� �������� ����� ���� �������� � �������� ���� edit �� ���� DFN �� �������������

  Assert.AreEqual(DifTSVFilePath, FForm.ePathTSV.Text);
  Assert.AreEqual(DFNFilePath, FForm.ePathDFNToMod.Text);
  Assert.IsFalse(TController.RewriteDFN(FForm.ePathTSV.Text, FForm.ePathDFNToMod.Text),
   'RewriteDFN �� ��������� �������� Fasle ��� ������������� ���� �����.');

//-------------------�������� ������� ����� �������� �����------------------

  FForm.odOpenTSV.FileName := TSVFilePath;
  FForm.spChooseTSVClick(nil);
  FForm.odOpenDFNToMod.FileName := DFNFilePath;
  FForm.spChooseDFNToModClick(nil);

//------��������, �� �������� ����� ���� �������� � �������� ���� edit-----

  Assert.AreEqual(TSVFilePath, FForm.ePathTSV.Text);
  Assert.AreEqual(DFNFilePath, FForm.ePathDFNToMod.Text);

//----------------�������� � ��������� TSV ���� ������ ��������---------------

  TestTSVFileContent := TStringList.Create;
  TestTSVFileContent.Add('456' + #9 + '��������' + #9 + '�������');
  TestTSVFileContent.Add('001' + #9 + '������' + #9 + '�����');
  TestTSVFileContent.SaveToFile(TSVFilePath);

//----------------�������� � ��������� DFN ���� ������ ��������---------------

  TestDFNFileContent := TStringList.Create;
  TestDFNFileContent.Add('ID:456' + #4 + 'Orig:''��������''' + #4 + 'Curr:''��������''');
  TestDFNFileContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:');
  TestDFNFileContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');
  TestDFNFileContent.SaveToFile(DFNFilePath);

//------------------------------��������� DFN �����-----------------------------

  FForm.bWriteDFNClick(nil);

//----------------��������� TStringList � ���������� �����������----------------

  var ExpectedDFNContent := TStringList.Create;
  ExpectedDFNContent.Add('ID:456' + #4 + 'Orig:''��������''' + #4 + 'Curr:''�������''');
  ExpectedDFNContent.Add('ID:001' + #4 + 'Orig:' + #4 + 'Curr:''�����''');
  ExpectedDFNContent.Add('ID:1' + #4 + 'Orig:''Hello''' + #4 + 'Curr:''Hello''');

//------����������� ��������� ���� �� ��������� � ���������� �����������------

  var ActualDFNContent := TStringList.Create;
  ActualDFNContent.LoadFromFile(DFNFilePath);
  Assert.AreEqual(ExpectedDFNContent.Text, ActualDFNContent.Text);

//------------------------------��������� �� �����-----------------------------

  FreeAndNil(TestTSVFileContent);
  FreeAndNil(TestDFNFileContent);
  FreeAndNil(ActualDFNContent);
  FreeAndNil(ExpectedDFNContent);
  DeleteFile(TSVFilePath); // ��������� ��������� DFN ����
  DeleteFile(DFNFilePath); // ��������� ��������� TSV ����
  DeleteFile(DifTSVFilePath); // ��������� ��������� TSV ����
end;

//------------------------------------------------------------------------------
//      ���������� ������� ��� ������� �����, ���� �� ������� �����
//------------------------------------------------------------------------------
procedure TParserTests.TestContainsCyrillicCharacters_True;
begin
  // �������� �����, ���� ������ ��������� �������
  Assert.IsTrue(TModel.ContainsCyrillicCharacters('�����, ����!'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestContainsCyrillicCharacters_False;
begin
  // �������� �����, ���� �� ������ ���������� �������
  Assert.IsFalse(TModel.ContainsCyrillicCharacters('Hello, world!'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestIsCharCyrillic_True;
begin
  // �������� �������, �� � �����������
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
  Assert.IsTrue(TModel.IsCharCyrillic('�'));
end;

//------------------------------------------------------------------------------

procedure TParserTests.TestIsCharCyrillic_False;
begin
  // �������� �������, �� �� � �����������
  Assert.IsFalse(TModel.IsCharCyrillic('A'));
  Assert.IsFalse(TModel.IsCharCyrillic('B'));
  Assert.IsFalse(TModel.IsCharCyrillic('1'));
  Assert.IsFalse(TModel.IsCharCyrillic(' '));
  Assert.IsFalse(TModel.IsCharCyrillic('.'));
end;

//------------------------------------------------------------------------------

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
  FieldValue := TModel.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('John', FieldValue);
end;

//------------------------------------------------------------------------------

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
  FieldValue := TModel.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('John', FieldValue);
end;

//------------------------------------------------------------------------------

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
  FieldValue := TModel.ExtractField(Fields, FieldName);

  // ��������
  Assert.AreEqual('', FieldValue);
end;

initialization
  TDUnitX.RegisterTestFixture(TParserTests);

end.