unit SalePlanWFWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, StdCtrls, ExtCtrls, CommUtils, IniFiles,
  SOPReaderUnit, ComObj, DateUtils, ProjYearWin;

type
  TfrmSalePlanWF = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    Memo1: TMemo;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    leCurrWeekSP: TLabeledEdit;
    leLastWeekWF: TLabeledEdit;
    btnLastWeekWF: TButton;
    btnCurrWeekSP: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure btnCurrWeekSPClick(Sender: TObject);
    procedure btnLastWeekWFClick(Sender: TObject);
  private
    { Private declarations }
    procedure DoWFWithLassWeek(const sfile_save:string; aSOPReader: TSOPReader);
    procedure DoWF(const sfile_save:string; aSOPReader: TSOPReader);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmSalePlanWF.ShowForm;
var
  frmSalePlan: TfrmSalePlanWF;
begin
  frmSalePlan := TfrmSalePlanWF.Create(nil);
  frmSalePlan.ShowModal;
  frmSalePlan.Free;
end;

procedure TfrmSalePlanWF.FormCreate(Sender: TObject);
var
  ini: TIniFile; 
begin
  ini := TIniFile.Create(AppIni);
  try
    leCurrWeekSP.Text := ini.ReadString(self.ClassName, leCurrWeekSP.Name, '');
    leLastWeekWF.Text := ini.ReadString(self.ClassName, leLastWeekWF.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSalePlanWF.FormDestroy(Sender: TObject);
var
  ini: TIniFile; 
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leCurrWeekSP.Name, leCurrWeekSP.Text);
    ini.WriteString(self.ClassName, leLastWeekWF.Name, leLastWeekWF.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSalePlanWF.ToolButton2Click(Sender: TObject);
begin
  TfrmProjYear.ShowForm;
end;

procedure TfrmSalePlanWF.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSalePlanWF.btnCurrWeekSPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leCurrWeekSP.Text := sfile;
end;

procedure TfrmSalePlanWF.btnLastWeekWFClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leLastWeekWF.Text := sfile;
end;
  
procedure TfrmSalePlanWF.btnSave2Click(Sender: TObject);
var
  sfile_save: string;
  bLastWeekWF: Boolean;
  ee: Integer;
  aSOPReader: TSOPReader;
begin
  bLastWeekWF := FileExists(leLastWeekWF.Text);

  if not bLastWeekWF then
  begin
    ee := MessageBox(Handle, '����Waterfall�����ڣ��Ƿ������', '��ʾ', MB_YESNO);
    if ee <> IDYES then Exit;
  end;

  if not ExcelSaveDialog(sfile_save) then Exit;

  aSOPReader := TSOPReader.Create(nil, leCurrWeekSP.Text);
  try
    if bLastWeekWF then
    begin
      DoWFWithLassWeek(sfile_save, aSOPReader);
    end
    else
    begin
      DoWF(sfile_save, aSOPReader);
    end;
  finally
    aSOPReader.Free;
  end;
 
  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmSalePlanWF.DoWFWithLassWeek(const sfile_save:string; aSOPReader: TSOPReader);
begin

end;
    
procedure TfrmSalePlanWF.DoWF(const sfile_save:string; aSOPReader: TSOPReader);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  slMonth: TStringList;
begin
  // ��ʼ���� Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '�����ʾ', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;
  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  slMonth := TStringList.Create;
  try

    ExcelApp.Sheets[1].Activate;
    ExcelApp.Sheets[1].Name := '�����ܶԱ�';

    irow := 1;

    ExcelApp.Cells[irow, 1].Value := '��Ŀ';
    MergeCells(ExcelApp, irow, 1, irow + 1, 1);

    ExcelApp.Cells[irow, 1].Value := '��Ŀ';                                                                   
    ExcelApp.Cells[irow, 1].Value := '����/����';
    ExcelApp.Cells[irow, 1].Value := '���ۼƻ�';


    aSOPReader.GetMonthList(slMonth);

    ExcelApp.Sheets[1].Activate;

    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

  finally
    slMonth.Free;
    
    WorkBook.Close;
    ExcelApp.Quit;
  end; 


end;  

end.
