unit ufMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtDlgs,
  Vcl.Buttons, ShellAPI, Vcl.ExtCtrls, Controller;

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
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFNToModFilePath: string; // Змінні, які будуть зберігати шлях відкритих файлів
    SavedTSVFilePath: string; // Змінна для зберігання шляху збереженого TSV файлу для можливості відкрити його в сторонніх програмах
    IsFileCreate: Boolean; // Змінна для перевірки, чи був збережений TSV файл, для відкриття його в сторонніх програмах
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.bCreateTSVClick(Sender: TObject);
begin
  IsFileCreate := TController.CreateTSV(OpenedDFNFilePath, SavedTSVFilePath); // Створюємо TSV файл за допомогою
end;                                             // функції, яка повертає значення true, якщо файл був створений

//------------------------------------------------------------------------------

procedure TMainForm.bWriteDFNClick(Sender: TObject);
begin
  TController.RewriteDFN(OpenedTSVFilePath, OpenedDFNToModFilePath);
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNClick(Sender: TObject);
begin
  OpenedDFNFilePath := TController.ChooseFile(odOpenDFN, 'DFN Files|*.dfn');
  if OpenedDFNFilePath <> '' then
    ePathDFN.Text := OpenedDFNFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNKeyPress(Sender: TObject; var Key: Char);
begin
  if TController.OpenFileByName(Sender, Key) then
    OpenedDFNFilePath := ePathDFN.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseTSVClick(Sender: TObject);
begin
  OpenedTSVFilePath := TController.ChooseFile(odOpenTSV, 'TSV Files|*.tsv');
  if OpenedTSVFilePath <> '' then
    ePathTSV.Text := OpenedTSVFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathTSVKeyPress(Sender: TObject; var Key: Char);
begin
  if TController.OpenFileByName(Sender, Key) then
    OpenedTSVFilePath := ePathTSV.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNToModClick(Sender: TObject);
begin
  OpenedDFNToModFilePath := TController.ChooseFile(odOpenDFNToMod, 'DFN Files|*.dfn');
  if OpenedDFNToModFilePath <> '' then
    ePathDFNToMod.Text := OpenedDFNToModFilePath;
end;

//------------------------------------------------------------------------------

procedure TMainForm.ePathDFNToModKeyPress(Sender: TObject; var Key: Char);
begin
  if TController.OpenFileByName(Sender, Key) then
    OpenedDFNToModFilePath := ePathDFNToMod.Text;
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenExcelClick(Sender: TObject);
begin
  TController.OpenByProgram(SavedTSVFilePath, IsFileCreate, 'Excel');
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenCalcClick(Sender: TObject);
begin
   TController.OpenByProgram(SavedTSVFilePath, IsFileCreate, 'Calc');
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

end.
