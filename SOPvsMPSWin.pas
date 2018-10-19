unit SOPvsMPSWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DailyPlanVsActReader, ComCtrls, ToolWin, ImgList,
  ExtCtrls, CommUtils, IniFiles, ComObj, SOPReaderUnit, ProjYearWin,
  DateUtils;

type
  TfrmSOPvsMPS = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    ToolButton7: TToolButton;
    btnSave: TToolButton;
    ToolButton10: TToolButton;
    btnExit: TToolButton;
    lbSOP: TLabeledEdit;
    lbMPS: TLabeledEdit;
    btnSOP: TButton;
    btnMPS: TButton;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Memo1: TMemo;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSOPClick(Sender: TObject);
    procedure btnMPSClick(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
  private
    { Private declarations } 
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmSOPvsMPS.ShowForm;
var
  frmScheActException: TfrmSOPvsMPS;
begin
  frmScheActException := TfrmSOPvsMPS.Create(nil);
  try
    frmScheActException.ShowModal;
  finally
    frmScheActException.Free;
  end;
end;

procedure TfrmSOPvsMPS.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    lbSOP.Text := ini.ReadString(self.ClassName, lbSOP.Name, '');
    lbMPS.Text := ini.ReadString(self.ClassName, lbMPS.Name, '');
    DateTimePicker1.DateTime := ini.ReadDateTime(self.ClassName,
      DateTimePicker1.Name, EncodeDate(YearOf(Now) , MonthOf(Now), DayOf(Now)));
  finally
    ini.Free;
  end;
end;

procedure TfrmSOPvsMPS.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, lbSOP.Name, lbSOP.Text);
    ini.WriteString(self.ClassName, lbMPS.Name, lbMPS.Text);
    ini.WriteDateTime(self.ClassName, DateTimePicker1.Name, DateTimePicker1.DateTime);
  finally
    ini.Free;
  end;
end;

procedure TfrmSOPvsMPS.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSOPvsMPS.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  lbSOP.Text := sfile;
end;

procedure TfrmSOPvsMPS.btnMPSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  lbMPS.Text := sfile;
end;
    
procedure TfrmSOPvsMPS.btnSaveClick(Sender: TObject);
var                 
  ExcelApp, WorkBook: Variant; 
  sfile_save: string;
  aSOPReader_SOP: TSOPReader;
  aSOPReader_MPS: TSOPReader;
  slyear: TStringList;
  iProj: Integer;
  aSOPProj_SOP: TSOPProj;  
  aSOPProj_MPS: TSOPProj;
  iline: Integer;
  aSOPLine_SOP: TSOPLine;
  aSOPLine_MPS: TSOPLine;
  irow: Integer;
  icol: Integer;
  icolMax: Integer;
  idate: Integer;
  aSOPCol_SOP: TSOPCol;
  aSOPCol_MPS: TSOPCol;
  bTitleWritten: Boolean;
  f: string;
begin

  if not ExcelSaveDialog(sfile_save) then Exit;
  slyear := TfrmProjYear.GetProjYears;
  aSOPReader_SOP := TSOPReader.Create(slyear, lbSOP.Text);
  aSOPReader_MPS := TSOPReader.Create(slyear, lbMPS.Text);

  bTitleWritten := False; 
  aSOPLine_MPS := nil;
  aSOPCol_MPS := nil;
  icolMax := 0;
  
  try

    // 开始保存 Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
        Exit;
      end;
    end;

    WorkBook := ExcelApp.WorkBooks.Add;
    try
      ExcelApp.Sheets[1].Name := 'S&OP vs MPS';
    
      ExcelApp.Cells[1, 1].Value := '项目';
      ExcelApp.Cells[1, 2].Value := '物料编码';
      ExcelApp.Cells[1, 3].Value := '颜色';
      ExcelApp.Cells[1, 4].Value := '容量';
      ExcelApp.Cells[1, 5].Value := '制式';
      ExcelApp.Cells[1, 6].Value := '计划';

      ExcelApp.Columns[1].ColumnWidth := 6;
      ExcelApp.Columns[2].ColumnWidth := 16;
      ExcelApp.Columns[3].ColumnWidth := 6;
      ExcelApp.Columns[4].ColumnWidth := 8;
      ExcelApp.Columns[5].ColumnWidth := 13;
      ExcelApp.Columns[6].ColumnWidth := 12;

      MergeCells(ExcelApp, 1, 1, 2, 1);
      MergeCells(ExcelApp, 1, 2, 2, 2);
      MergeCells(ExcelApp, 1, 3, 2, 3);
      MergeCells(ExcelApp, 1, 4, 2, 4);
      MergeCells(ExcelApp, 1, 5, 2, 5);
      MergeCells(ExcelApp, 1, 6, 2, 6);

      irow := 3;
      for iProj := 0 to aSOPReader_SOP.FProjs.Count - 1 do
      begin
        aSOPProj_SOP := TSOPProj(aSOPReader_SOP.FProjs.Objects[iProj]);
        aSOPProj_MPS := aSOPReader_MPS.GetProj(aSOPProj_SOP.FName);
        if aSOPProj_MPS = nil then
        begin
          Memo1.Lines.Add('mps of proj ' + aSOPProj_SOP.FName + ' not found.' );
          //Continue;
        end;

        for iline := 0 to aSOPProj_SOP.LineCount - 1 do
        begin
          aSOPLine_SOP := aSOPProj_SOP.Lines[iline];
          if aSOPProj_MPS <> nil then
          begin
            aSOPLine_MPS := aSOPProj_MPS.GetLine(aSOPLine_SOP.sVer,
              aSOPLine_SOP.sNumber, aSOPLine_SOP.sColor, aSOPLine_SOP.sCap);
          end;
                   
          ExcelApp.Cells[irow,     1].Value := aSOPProj_SOP.FName;
          ExcelApp.Cells[irow + 1, 1].Value := aSOPProj_SOP.FName;
          ExcelApp.Cells[irow + 2, 1].Value := aSOPProj_SOP.FName;
          ExcelApp.Cells[irow,     2].Value := aSOPLine_SOP.sNumber;
          ExcelApp.Cells[irow + 1, 2].Value := aSOPLine_SOP.sNumber;
          ExcelApp.Cells[irow + 2, 2].Value := aSOPLine_SOP.sNumber;
          ExcelApp.Cells[irow,     3].Value := aSOPLine_SOP.sColor;
          ExcelApp.Cells[irow + 1, 3].Value := aSOPLine_SOP.sColor;
          ExcelApp.Cells[irow + 2, 3].Value := aSOPLine_SOP.sColor;
          ExcelApp.Cells[irow,     4].Value := aSOPLine_SOP.sCap;
          ExcelApp.Cells[irow + 1, 4].Value := aSOPLine_SOP.sCap;
          ExcelApp.Cells[irow + 2, 4].Value := aSOPLine_SOP.sCap;
          ExcelApp.Cells[irow,     5].Value := aSOPLine_SOP.sVer;
          ExcelApp.Cells[irow + 1, 5].Value := aSOPLine_SOP.sVer;
          ExcelApp.Cells[irow + 2, 5].Value := aSOPLine_SOP.sVer;
          ExcelApp.Cells[irow,     6].Value := 'S&OP';
          ExcelApp.Cells[irow + 1, 6].Value := 'MPS';
          ExcelApp.Cells[irow + 2, 6].Value := 'S&OP vs MPS';

          icol := 7;
          for idate := 0 to aSOPLine_SOP.DateCount - 1 do
          begin
            aSOPCol_SOP := aSOPLine_SOP.Dates[idate];
            if aSOPCol_SOP.dt2 < DateTimePicker1.DateTime then
            begin
              Memo1.Lines.Add(aSOPCol_SOP.sDate + ' 小于设定日期，不显示');
              Continue;
            end;

            if aSOPLine_MPS <> nil then
            begin
              aSOPCol_MPS := aSOPLine_MPS.GetCol(aSOPCol_SOP.sDate);
            end;           

            if not bTitleWritten then
            begin
              ExcelApp.Cells[1, icol].Value := aSOPCol_SOP.sWeek;
              ExcelApp.Cells[2, icol].Value := aSOPCol_SOP.sDate;  
              ExcelApp.Columns[icol].ColumnWidth := 10;

              if icolMax < icol then
              begin
                icolMax := icol;
              end;
            end;
            
            ExcelApp.Cells[irow, icol].Value := aSOPCol_SOP.iQty;
            if aSOPCol_MPS <> nil then
            begin
              ExcelApp.Cells[irow + 1, icol].Value := aSOPCol_MPS.iQty;
            end;
            f := GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
            if icol > 7 then
            begin
              f := GetRef(icol - 1) + IntToStr(irow + 2) + '+' + f;
            end;
            ExcelApp.Cells[irow + 2, icol].Value := '=' + f;
            
            icol := icol + 1;

          end;
          bTitleWritten := True;
          
          irow := irow + 3;
        end;
      end;
                              
      AddColor(ExcelApp, 1, 1, 2, icolMax, $DBDCF2);
      AddBorder(ExcelApp, 1, 1, irow + 3 - 1, icolMax);


      WorkBook.SaveAs(sfile_save);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;

  finally
    slyear.Free;
    aSOPReader_SOP.Free;
    aSOPReader_MPS.Free;
  end;
 
  MessageBox(Handle, '完成', '提示', 0);

end;

procedure TfrmSOPvsMPS.ToolButton1Click(Sender: TObject);
begin
  TfrmProjYear.ShowForm;
end;

end.
