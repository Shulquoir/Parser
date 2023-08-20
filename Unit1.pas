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
  // ��������� ������ ��� �� ������������ ����� ������
  FForm.OpenedDFNFilePath := 'D:\Programs\Embarcadero\��\Unit1.dfn';
  FForm.OriginalDFNFileContent := TStringList.Create;
  FForm.OriginalDFNFileContent.Add('ID' + #4 + 'Orig' + #4 + 'Curr');
  FForm.OriginalDFNFileContent.Add('1' + #4 + 'Hello' + #4 + '�����');

  // ��������� ���������, ��� ������ ������������
  FForm.bCreateTSVClick(nil);

  // �������� ����������
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

procedure TParserTests.TestOpenDFNClick;
var
  TestFileName: string;
  TestFileContent: TStringList;
begin
  // ��������� ������ ��� �� ������������ ����� ������
  TestFileName := 'D:\Programs\Embarcadero\��\Unit1.dfn';
  TestFileContent := TStringList.Create;
  TestFileContent.Add('Test content line 1');
  TestFileContent.Add('Test content line 2');
  TestFileContent.SaveToFile(TestFileName);

  // ��������� ���������, ��� ������ ������������
  FForm.odOpenDFN.FileName := TestFileName; // ��������� ���������� ������ �� �������� ��'� ��������� �����
  FForm.spChooseDFNClick(nil);

  // ��������, �� ���� ����� ����� �������� �������� ���� ��������� ���������
  Assert.AreEqual(TestFileName, FForm.ePathDFN.Text);
  Assert.IsNotNull(FForm.OriginalDFNFileContent);
  Assert.AreEqual(TestFileContent.Text, FForm.OriginalDFNFileContent.Text);

  // ��������� �� �����
  TestFileContent.Free;
  DeleteFile(TestFileName);
end;

procedure TParserTests.TestFormCreate;
begin
  // ��������� ��������� FormCreate
  FForm.FormCreate(nil);

  // ����������, �� ����� IsFileCreate �� �������� false ���� ��������� ���������
  Assert.IsFalse(FForm.IsFileCreate);
end;

initialization
  TDUnitX.RegisterTestFixture(TParserTests);
end.
