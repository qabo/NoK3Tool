unit ScheActExceptionWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DailyPlanVsActReader, ComCtrls, ToolWin, ImgList,
  ExtCtrls, CommUtils, IniFiles, ComObj;

type
  TfrmScheActException = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    ToolButton7: TToolButton;
    btnManage: TToolButton;
    ToolButton10: TToolButton;
    btnExit: TToolButton;
    mmoFiles: TMemo;
    tbAdd: TToolButton;
    procedure btnManageClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure tbAddClick(Sender: TObject);
  private
    { Private declarations }
    procedure GetNumberMTD(aDailyPlanVsActReader: TDailyPlanVsActReader; const snumbe: string;
      dt: TDateTime; var dPlan, dAct: Double; var sReason, sAction: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

class procedure TfrmScheActException.ShowForm;
var
  frmScheActException: TfrmScheActException;
begin
  frmScheActException := TfrmScheActException.Create(nil);
  try
    frmScheActException.ShowModal;
  finally
    frmScheActException.Free;
  end;
end;

procedure TfrmScheActException.GetNumberMTD(aDailyPlanVsActReader: TDailyPlanVsActReader;
  const snumbe: string; dt: TDateTime; var dPlan, dAct: Double;
  var sReason, sAction: string);
var
  isheet: Integer;
  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
  iline: Integer;
  aDPVALine: TDPVALine;
  idate: Integer;
  aDatePlanActPtr: PDatePlanAct; 
  sl: TStringList;
begin
  sl := TStringList.Create;
  try
    for isheet := 0 to aDailyPlanVsActReader.Count - 1 do
    begin
      aDailyPlanVsAcsSheet := TDailyPlanVsAcsSheet(aDailyPlanVsActReader.Items[isheet]);
      for iline := 0 to aDailyPlanVsAcsSheet.FList.Count - 1 do
      begin
        aDPVALine := TDPVALine(aDailyPlanVsAcsSheet.FList[iline]);
        if aDPVALine.snumber <> snumbe then Continue;

        for idate := 0 to aDPVALine.FList.Count - 1 do
        begin
          aDatePlanActPtr := PDatePlanAct(aDPVALine.FList[idate]);
          if aDatePlanActPtr^.dt < dt then
          begin
            dPlan := dPlan + aDatePlanActPtr^.dQty;
            dAct := dAct + aDatePlanActPtr^.dQtyAct;
            if aDatePlanActPtr^.dQty > aDatePlanActPtr^.dQtyAct then
            begin
              if aDatePlanActPtr^.scomment <> '' then
              begin
                sl.Text := aDatePlanActPtr^.scomment;
                if sl.Count > 1 then
                begin
                  sReason := sReason + sl.Names[1] + '; ';
                  sAction := sAction + sl.ValueFromIndex[1] + '; ';
                end;
              end;
            end;
          end
          else Break; 
        end; 
        Break;
      end;
    end;
  finally
    sl.Free;
  end;
end;

procedure TfrmScheActException.btnManageClick(Sender: TObject);
var                 
  ExcelApp, WorkBook: Variant;
  sfile: string;
  sfile_save: string;
  i: Integer;
  aDailyPlanVsActReader: TDailyPlanVsActReader;
  isheet: Integer;
  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
  iline: Integer;
  aDPVALine: TDPVALine;
  slnumber: TStringList;
  irow: Integer;
  inumber: Integer;
  dPlan, dAct: Double;
  dt: TDateTime;
  irow1: Integer;
  sReason, sAction: string;
begin

  if not ExcelSaveDialog(sfile_save) then Exit;

  dt := Now;

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
    ExcelApp.Cells[1, 1].Value := 'ODM各量产项目生产异常分析_跟进_补货计划汇总';
    MergeCells(ExcelApp, 1, 1, 1, 10);

    ExcelApp.Cells[3, 1].Value := 'Update: ' + FormatDateTime('yyyy-MM-dd', dt);
    MergeCells(ExcelApp, 3, 1, 3, 3);
    
    ExcelApp.Cells[4, 1].Value := '项目';
    ExcelApp.Cells[4, 2].Value := 'P/N';
    ExcelApp.Cells[4, 3].Value := '描述';
    ExcelApp.Cells[4, 4].Value := 'MTD 计划';
    ExcelApp.Cells[4, 5].Value := 'MTD 产出';
    ExcelApp.Cells[4, 6].Value := 'MTD 差异';
    ExcelApp.Cells[4, 7].Value := '原因比例';
    ExcelApp.Cells[4, 8].Value := '原因';
    ExcelApp.Cells[4, 9].Value := '补货计划';
    ExcelApp.Cells[4, 10].Value := '备注';


    ExcelApp.Columns[1].ColumnWidth := 8.5;
    ExcelApp.Columns[2].ColumnWidth := 16;
    ExcelApp.Columns[3].ColumnWidth := 16;
    ExcelApp.Columns[4].ColumnWidth := 8.5;
    ExcelApp.Columns[5].ColumnWidth := 8.5;
    ExcelApp.Columns[6].ColumnWidth := 8.5;
    ExcelApp.Columns[7].ColumnWidth := 9;
    ExcelApp.Columns[8].ColumnWidth := 19;
    ExcelApp.Columns[9].ColumnWidth := 15;
    ExcelApp.Columns[10].ColumnWidth := 15;

    irow := 5;

    slnumber := TStringList.Create;
    try
      for i := 0 to mmoFiles.Lines.Count - 1 do
      begin
        irow1 := irow;
        
        sfile := mmoFiles.Lines[i];
        aDailyPlanVsActReader := TDailyPlanVsActReader.Create(sfile);
        try
          for isheet := 0 to aDailyPlanVsActReader.Count - 1 do
          begin
            aDailyPlanVsAcsSheet := TDailyPlanVsAcsSheet(aDailyPlanVsActReader.Items[isheet]);
            for iline := 0 to aDailyPlanVsAcsSheet.FList.Count - 1 do
            begin
              aDPVALine := TDPVALine(aDailyPlanVsAcsSheet.FList[iline]);
              if slnumber.IndexOf(aDPVALine.snumber) < 0 then
              begin
                slnumber.Add( aDPVALine.snumber  );
              end;
            end;
          end;

          ExcelApp.Cells[irow, 1].Value := ChangeFileExt(ExtractFileName(sfile), ''); 
          for inumber := 0 to slnumber.Count - 1 do
          begin
            dPlan := 0;
            dAct := 0;
            sReason := '';
            sAction := '';
            GetNumberMTD(aDailyPlanVsActReader, slnumber[inumber], dt, dPlan,
              dAct, sReason, sAction);
            ExcelApp.Cells[irow, 2].Value := slnumber[inumber];
            ExcelApp.Cells[irow, 4].Value :=  dPlan;
            ExcelApp.Cells[irow, 5].Value :=  dAct;
            ExcelApp.Cells[irow, 6].Value :=  '=' + GetRef(5) + IntToStr(irow)
              + '-' + GetRef(4) + IntToStr(irow);
            if sReason <> '' then
            begin
              ExcelApp.Cells[irow, 8].Value :=  sReason;
              ExcelApp.Cells[irow, 9].Value :=  sAction;
            end;  


            irow := irow + 1;
          end;
          MergeCells(ExcelApp, irow, 2, irow, 3);
          ExcelApp.Cells[irow, 2].Value := 'Total';

          ExcelApp.Cells[irow, 4].Value :=  '=SUM(' + GetRef(4) + IntToStr(irow1) + ':' + GetRef(4) + IntToStr(irow - 1) + ')';
          ExcelApp.Cells[irow, 5].Value :=  '=SUM(' + GetRef(5) + IntToStr(irow1) + ':' + GetRef(5) + IntToStr(irow - 1) + ')';
          ExcelApp.Cells[irow, 6].Value :=  '=SUM(' + GetRef(6) + IntToStr(irow1) + ':' + GetRef(6) + IntToStr(irow - 1) + ')';

          AddColor(ExcelApp, irow, 2, irow, 10, $E8DEB7);

          MergeCells(ExcelApp, irow1, 1, irow, 1);
          
          irow := irow + 1;
        finally
          aDailyPlanVsActReader.Free;
        end;
      end;
    finally
      slnumber.Free;
    end;

    AddBorder(ExcelApp, 4, 1, irow - 1, 10);

    
    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit; 
  end;

  MessageBox(Handle, '完成', '提示', 0);

end;

procedure TfrmScheActException.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    mmoFiles.Text := StringReplace( ini.ReadString(self.ClassName, mmoFiles.Name, ''), ';', #13#10, [rfReplaceAll] ) ;
  finally
    ini.Free;
  end;
end;

procedure TfrmScheActException.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, mmoFiles.Name, StringReplace(mmoFiles.Text, #13#10, ';', [rfReplaceAll]));
  finally
    ini.Free;
  end;
end;

procedure TfrmScheActException.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmScheActException.tbAddClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  mmoFiles.Lines.Add(StringReplace(sfile, ';', #13#10, [rfReplaceAll]));
end;

end.
