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
    OpenedDFNFilePath, OpenedTSVFilePath, OpenedDFNToModFilePath: String; // Змінні, які будуть зберігати шлях відкритих файлів
    SavedTSVFilePath: String; // Змінна для зберігання шляху збереженого TSV файлу для можливості відкрити його в сторонніх програмах
    IsFileCreate: boolean; // Змінна для перевірки, чи був збережений TSV файл, для відкриття його в сторонніх програмах
  public
    { Public declarations }

  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.bCreateTSVClick(Sender: TObject);
begin
  ShowMessage('1' + SavedTSVFilePath);
  IsFileCreate := TController.CreateTSV(OpenedDFNFilePath, SavedTSVFilePath); // Створюємо TSV файл за допомогою
                                                // функції, яка повертає значення true, якщо файл був створений
  ShowMessage('2' + SavedTSVFilePath);
 end;

//------------------------------------------------------------------------------

procedure TMainForm.bWriteDFNClick(Sender: TObject);
begin
  TController.RewriteDFN(OpenedTSVFilePath, OpenedDFNToModFilePath);
end;

//------------------------------------------------------------------------------

procedure TMainForm.spChooseDFNClick(Sender: TObject);
begin
  if TController.ChooseFile(OpenedDFNFilePath, 'DFN Files|*.dfn') then
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
  if TController.ChooseFile(OpenedTSVFilePath, 'TSV Files|*.tsv') then
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
  if TController.ChooseFile(OpenedDFNToModFilePath, 'DFN Files|*.dfn') then
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
  TController.OpenExcel(SavedTSVFilePath, IsFileCreate);
end;

//------------------------------------------------------------------------------

procedure TMainForm.bOpenCalcClick(Sender: TObject);
begin
   TController.OpenCalc(SavedTSVFilePath, IsFileCreate);
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

end.
