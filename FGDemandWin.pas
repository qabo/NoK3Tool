unit FGDemandWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, SOPReaderUnit, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin,
  CommUtils, DB, ADODB, ProjYearWin, SelectFGDemandWin, DBGridEhGrouping,
  ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh, GridsEh, DBAxisGridsEh, IniFiles,
  DBGridEh, Provider, DBClient, FGDemandConfigWin, ComObj, ExcelConsts,
  FGDemandManageWin;

type
  TfrmFGDemand = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    ADOConnection1: TADOConnection;
    tbImport: TToolButton;
    ADOQuery1: TADOQuery;
    Memo1: TMemo;
    btnProjYear: TToolButton;
    ToolButton2: TToolButton;
    tbCompare: TToolButton;
    ToolButton3: TToolButton;
    mmosql: TMemo;
    DBGridEh1: TDBGridEh;
    ClientDataSet1: TClientDataSet;
    DataSetProvider1: TDataSetProvider;
    DataSource1: TDataSource;
    ToolButton1: TToolButton;
    ToolButton4: TToolButton;
    ToolButton6: TToolButton;
    ToolButton8: TToolButton;
    GroupBox1: TGroupBox;
    btnManage: TToolButton;
    ToolButton10: TToolButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbImportClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnProjYearClick(Sender: TObject);
    procedure tbCompareClick(Sender: TObject);
    procedure DBGridEh1GetCellParams(Sender: TObject; Column: TColumnEh;
      AFont: TFont; var Background: TColor; State: TGridDrawState);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure btnManageClick(Sender: TObject);
  private         
    slFGPlans: TStringList;
    FFGDemandConfig: TFGDemandConfig;
    { Private declarations }
    function GetWeekAndDateStr(const sproj, sdate: string; slname: TStringList;
      var sweek, sdatestr: string): Boolean;
    function FGPlanExists(const splanname: string): Boolean;
//    procedure DeleteFGPlan(const splanname: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

const
  CSConnStr = 'Provider=SQLOLEDB.1;Password=Pmc010161;Persist Security Info=True;User ID=sa;Initial Catalog=sop;Data Source=.';

class procedure TfrmFGDemand.ShowForm;
var
  frmFGDemand: TfrmFGDemand;
begin
  frmFGDemand := TfrmFGDemand.Create(nil);
  try
    frmFGDemand.ShowModal;
  finally
    frmFGDemand.Free;
  end;
end;

procedure TfrmFGDemand.FormCreate(Sender: TObject); 
begin
  ADOConnection1.Connected := False;
  ADOConnection1.ConnectionString := 'Provider=SQLOLEDB.1;Password=' + gpwd + ';Persist Security Info=True;User ID=' + guser + ';Initial Catalog=sop;Data Source=' + gserver;

  FFGDemandConfig := TFGDemandConfig.Create;
  FFGDemandConfig.LoadConfig(Self.ClassName, AppIni);
  ADOConnection1.Connected := True;

  slFGPlans := TStringList.Create;
end;

procedure TfrmFGDemand.FormDestroy(Sender: TObject);
begin
  FFGDemandConfig.SaveConfig(Self.ClassName, AppIni);
  FFGDemandConfig.Free;
  ADOConnection1.Close;
  
  slFGPlans.Free;
end;

function TfrmFGDemand.FGPlanExists(const splanname: string): Boolean;
begin
  Result := False;
  try
    ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(' select * from fgdemand where fname=''' + splanname + ''' ');
    ADOQuery1.Open;
    Result := not ADOQuery1.IsEmpty;
    ADOQuery1.Close;
  except
    Memo1.Lines.Add('异常： FGPlanExists ' + splanname);
  end;
end;

//procedure TfrmFGDemand.DeleteFGPlan(const splanname: string);
//begin
//  try
//    ADOQuery1.Close;
//    ADOQuery1.SQL.Clear;
//    ADOQuery1.SQL.Add(' delete from fgdemand where fname=''' + splanname + ''' ');
//    ADOQuery1.ExecSQL;
//  except
//    Memo1.Lines.Add('异常： DeleteFGPlan ' + splanname);
//  end;
//end;

procedure TfrmFGDemand.tbImportClick(Sender: TObject);
var
  aSOPReader: TSOPReader;
  splanname: string;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iLine: Integer;
  aSOPLine: TSOPLine;
  iCol: Integer;
  aSOPCol: TSOPCol;
  slProjYear: TStringList;
  syear: string;
  dt0, dt: TDateTime;
  sdt: string;

  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;


  if MessageBox(Handle, '确定导入销售计划？', '提示', MB_YESNO) <> IDYES then Exit;

  splanname := ExtractFileName(sfile);
  splanname := ChangeFileExt(splanname, '');

  Memo1.Lines.Add('销售计划名称：' + splanname);

  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;

  if FGPlanExists(splanname) then
  begin
    if MessageBox(Handle, '销售计划已存在，是否替换？', '提示', MB_YESNO) <> IDYES then Exit;

    ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(' declare @fid int     ');
    ADOQuery1.SQL.Add(' select @fid=fid from fgdemand where fname=''' + splanname + '''  ');
    ADOQuery1.SQL.Add(' delete from fgdemand where fname=''' + splanname + ''' ');
    ADOQuery1.SQL.Add(' delete from fgdemand_entry where fid=@fid');   
    ADOQuery1.ExecSQL;
    //DeleteFGPlan(splanname);
  end;

  Memo1.Lines.Add('读取销售计划');             
  slProjYear := TfrmProjYear.GetProjYears;

  aSOPReader := TSOPReader.Create(slProjYear, sfile);
  try
    Memo1.Lines.Add('销售计划写入数据库 开始......');

    ADOQuery1.Close;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add(' begin tran ');
    ADOQuery1.SQL.Add(' insert into fgdemand(fname) values(''' + splanname + ''') ');
    ADOQuery1.SQL.Add(' declare @fid int  ');
    ADOQuery1.SQL.Add(' select @fid = fid from fgdemand where fname=''' + splanname + ''' ');

    for iproj := 0 to aSOPReader.FProjs.Count - 1 do
    begin
      aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[iproj]);
      for iLine := 0 to aSOPProj.FList.Count - 1 do
      begin
        syear := slProjYear.Values[aSOPProj.FName];
        dt0 := 0;
        aSOPLine := TSOPLine(aSOPProj.FList.Objects[iLine]);
        for iCol := 0 to aSOPLine.FList.Count - 1 do
        begin
          aSOPCol := TSOPCol(aSOPLine.FList.Objects[iCol]);
          sdt := aSOPCol.sDate;
          sdt := Copy(sdt, 1, Pos('-', sdt) - 1);
          sdt := syear + '-' + StringReplace(sdt, '/', '-', [rfReplaceAll]);
          dt := myStrToDateTime(sdt);            
          if dt0 = 0 then
          begin
            dt0 := dt;
          end
          else
          begin
            if dt0 > dt then
            begin
              syear := IntToStr(StrToInt(syear) + 1);   
              sdt := aSOPCol.sDate;
              sdt := Copy(sdt, 1, Pos('-', sdt) - 1);
              sdt := syear + '-' + StringReplace(sdt, '/', '-', [rfReplaceAll]);
              dt := myStrToDateTime(sdt);
            end;  
            dt0 := dt;
          end;

          if aSOPLine.sNumber = '' then
          begin
            aSOPLine.sNumber := aSOPLine.sVer + aSOPLine.sColor + aSOPLine.sCap;
          end;
          ADOQuery1.SQL.Add(' insert into fgdemand_entry(fid, fproj, fver, fnumber, fcolor, fcap, fweek, fdatestr, fdate, fqty) ');
          ADOQuery1.SQL.Add(' values( @fid, ''' + aSOPProj.FName + ''', ''' + aSOPLine.sVer + ''', ''' + aSOPLine.sNumber + ''', ''' + aSOPLine.sColor + ''', ''' + aSOPLine.sCap + ''', ''' + aSOPCol.sWeek + ''', ''' + aSOPCol.sDate + ''', ''' + FormatDateTime('yyyy-MM-dd', dt) + ''', ' + FloatToStr(aSOPCol.iQty) + ' ) ');

        end;
      end;
    end;
                        
    ADOQuery1.SQL.Add(' commit tran ');
    ADOQuery1.ExecSQL;

    Memo1.Lines.Add('销售计划写入数据库 完成......');

    MessageBox(Handle, '引入销售计划成功', '提示', 0);
  finally
    aSOPReader.Free;
    slProjYear.Free;
  end;
end;

procedure TfrmFGDemand.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmFGDemand.btnProjYearClick(Sender: TObject);
begin
  TfrmProjYear.ShowForm;
end;

function TfrmFGDemand.GetWeekAndDateStr(const sproj, sdate: string; slname: TStringList;
  var sweek, sdatestr: string): Boolean;
var
  sname: string;
  i: Integer;
begin
  sname := '';
  for i := 0 to slname.Count - 1 do
  begin
    sname := sname + ',''' + slname[i] + '''';
  end;
  
  Result := False;
  try
    ADOQuery1.Close;
    ADOQuery1.SQL.Clear;

    ADOQuery1.SQL.Add(' select top 1 t1.fweek, fdatestr from fgdemand_entry t1 ');
    ADOQuery1.SQL.Add(' inner join fgdemand t2 on t1.fid=t2.fid ');
    ADOQuery1.SQL.Add(' where t2.fname in (''xx''' + sname + ') and t1.fproj=''' +
      sproj + ''' and t1.fdate=''' + sdate + ''' ');
    Memo1.Lines.Add(ADOQuery1.SQL.Text);
    ADOQuery1.Open;
    if not ADOQuery1.IsEmpty then
    begin
      ADOQuery1.First;
      sweek := ADOQuery1.FieldByName('fweek').AsString;
      sdatestr := ADOQuery1.FieldByName('fdatestr').AsString;
      Result := True;
    end;
    ADOQuery1.Close;
  except
    on e: Exception do
    begin
      raise Exception.Create(e.Message);
    end;
  end;
end;

procedure TfrmFGDemand.tbCompareClick(Sender: TObject);
var
  sproj: string;
  i: Integer;
  sweek: string;
  sdatestr: string;
  sdate: string;
  dt: TDateTime;
begin
  ClientDataSet1.Close;
  DBGridEh1.Columns.Clear;
  
  if not  TfrmSelectFGDemand.GetFGPlans(ADOConnection1, sproj, dt, slFGPlans) then Exit;
  
  Memo1.Lines.Add(sproj);
  Memo1.Lines.Add(slFGPlans.Text);

  if slFGPlans.Count < 2 then
  begin
    MessageBox(Handle, '请最少选取两个计划进行对比', '提示', 0);
  end;

  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(' declare @fproj varchar(50) ');
  ADOQuery1.SQL.Add(' select @fproj=''' + sproj + ''' ');
  ADOQuery1.SQL.Add(' declare @dt datetime   ');
  ADOQuery1.SQL.Add(' select @dt = ''' + FormatDateTime('yyyy-MM-dd', dt) + '''     ');

  ADOQuery1.SQL.Add(' create table #tmp_weeks(fid int, fname varchar(250)) ');
  for i := 0 to slFGPlans.Count - 1 do
  begin
    ADOQuery1.SQL.Add(' insert into #tmp_weeks(fid, fname) VALUES(' + IntToStr(i + 1) + ', ''' + slFGPlans[i] + ''') ');
  end;
  ADOQuery1.SQL.Add(mmosql.Text);

  Memo1.Lines.Add(ADOQuery1.SQL.Text);

  ADOQuery1.Open;

  ClientDataSet1.Open;

  ADOQuery1.Close;

  DBGridEh1.Columns[0].Title.Caption := 'Week';
  DBGridEh1.Columns[1].Title.Caption := '版本';
  DBGridEh1.Columns[2].Title.Caption := '产品编码';
  DBGridEh1.Columns[3].Title.Caption := '颜色';
  DBGridEh1.Columns[4].Title.Caption := '容量';

  DBGridEh1.Columns[0].Width := 250;
  DBGridEh1.Columns[1].Width := 100;
  DBGridEh1.Columns[2].Width := 100;
  DBGridEh1.Columns[3].Width := 60;
  DBGridEh1.Columns[4].Width := 50;
  
  for i := 5 to DBGridEh1.Columns.Count - 1  do
  begin
    sdate := DBGridEh1.Columns[i].FieldName;
    sdate := Copy(sdate, 1, Pos('|', sdate) - 1);
    if GetWeekAndDateStr(sproj, sdate, slFGPlans, sweek, sdatestr) then
    begin
      DBGridEh1.Columns[i].Title.Caption := sweek + '|' + sdatestr + '|' +
        Copy(DBGridEh1.Columns[i].FieldName, Pos('|', DBGridEh1.Columns[i].FieldName) + 1, Length(DBGridEh1.Columns[i].FieldName));
    end; 
    DBGridEh1.Columns[i].Width := 60;

    if Pos('%', DBGridEh1.Columns[i].FieldName) > 0 then
    begin
      DBGridEh1.Columns[i].DisplayFormat := '#.#%';
    end;
  end;  
end;

procedure TfrmFGDemand.DBGridEh1GetCellParams(Sender: TObject;
  Column: TColumnEh; AFont: TFont; var Background: TColor;
  State: TGridDrawState);
var
  icol: Integer;
  dPer: Double;
begin     
  if Pos('qty', Column.Name) > 0 then
  begin
    icol := Column.Index + 2;
  end
  else if Pos('delta', Column.Name) > 0 then
  begin
    icol := Column.Index + 1;
  end 
  else if Pos('%', Column.Name) > 0 then
  begin
    icol := Column.Index;
  end
  else Exit;

  dPer := DBGridEh1.DataSource.DataSet.FieldByName(DBGridEh1.Columns[icol].FieldName).AsFloat;
  if dPer >= FFGDemandConfig.UpperLimit * 100 then
  begin
    Background := FFGDemandConfig.UpperBrush;
    AFont.Color := FFGDemandConfig.UpperFont;
  end
  else if dPer <= FFGDemandConfig.LowerLimit * 100 then
  begin
    Background := FFGDemandConfig.LowerBrush;
    AFont.Color := FFGDemandConfig.LowerFont;
  end;
end;

procedure TfrmFGDemand.ToolButton1Click(Sender: TObject);
begin
  if not TfrmFGDemandConfig.ShowForm(FFGDemandConfig) then Exit;
  DBGridEh1.Invalidate;
end;

procedure TfrmFGDemand.ToolButton8Click(Sender: TObject);
var
  sfile: string;                        
  ExcelApp, WorkBook: Variant;
  irow: Integer; 
  ifield: Integer;
  sl: TStringList;
  sname: string; 
  ic: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;


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

  sl := TStringList.Create;

  WorkBook := ExcelApp.WorkBooks.Add;
  try

    ExcelApp.Columns[1].ColumnWidth := 45;
    ExcelApp.Columns[2].ColumnWidth := 12;
    ExcelApp.Columns[3].ColumnWidth := 15;
    ExcelApp.Columns[4].ColumnWidth := 7;
    ExcelApp.Columns[5].ColumnWidth := 7;

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := 'Week'; 
    ExcelApp.Cells[irow, 2].Value := '版本';
    ExcelApp.Cells[irow, 3].Value := '产品编码';
    ExcelApp.Cells[irow, 4].Value := '颜色';
    ExcelApp.Cells[irow, 5].Value := '容量';

    MergeCells(ExcelApp, irow, 1, irow + 2, 1);
    MergeCells(ExcelApp, irow, 2, irow + 2, 2);
    MergeCells(ExcelApp, irow, 3, irow + 2, 3);
    MergeCells(ExcelApp, irow, 4, irow + 2, 4);
    MergeCells(ExcelApp, irow, 5, irow + 2, 5);

    for ifield := 5 to DBGridEh1.Columns.Count - 1 do
    begin
      sl.Text := StringReplace(DBGridEh1.Columns[ifield].Title.Caption, '|', #13#10, [rfReplaceAll]);

      if (sl[2] <> '(delta)') and (sl[2] <> '(%)') then
      begin
        ExcelApp.Cells[irow,     ifield + 1].Value := sl[0];
        ExcelApp.Cells[irow + 1, ifield + 1].Value := sl[1];
        MergeCells(ExcelApp, irow, ifield + 1, irow, ifield + 3);   
        MergeCells(ExcelApp, irow + 1, ifield + 1, irow + 1, ifield + 3);
      end;

      ExcelApp.Cells[irow + 2, ifield + 1].Value := sl[2];
    end;

    AddColor(ExcelApp, irow, 1, irow + 2, DBGridEh1.Columns.Count, $DBDCF2);   
    ExcelApp.Range[ ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow + 2, DBGridEh1.Columns.Count] ].HorizontalAlignment := xlCenter;
 
    ic := 0;
    
    irow := 4;
    ClientDataSet1.First;
    while not ClientDataSet1.Eof do
    begin
      sname := ClientDataSet1.FieldByName('fname').AsString;
      
      ExcelApp.Cells[irow, 1].Value := sname;
      ExcelApp.Cells[irow, 2].Value := ClientDataSet1.FieldByName('fver').AsString;
      ExcelApp.Cells[irow, 3].Value := ClientDataSet1.FieldByName('fnumber').AsString;
      ExcelApp.Cells[irow, 4].Value := ClientDataSet1.FieldByName('fcolor').AsString;
      ExcelApp.Cells[irow, 5].Value := ClientDataSet1.FieldByName('fcap').AsString;

      for ifield := 5 to ClientDataSet1.Fields.Count - 1 do
      begin
        if Pos('qty', ClientDataSet1.Fields[ifield].FieldName) > 0 then
        begin
          ExcelApp.Cells[irow, ifield + 1].Value := ClientDataSet1.Fields[ifield].AsString; 
        end
        else if Pos('delta', ClientDataSet1.Fields[ifield].FieldName) > 0 then
        begin
          if ic = 0 then
          begin
            ExcelApp.Cells[irow, ifield + 1].Value := ClientDataSet1.Fields[ifield].AsString; 
          end
          else
          begin
            ExcelApp.Cells[irow, ifield + 1].Value := '=' + GetRef(ifield) + IntToStr(irow) + '-' + GetRef(ifield) + IntToStr(irow - 1);
          end;
        end 
        else if Pos('%', ClientDataSet1.Fields[ifield].FieldName) > 0 then
        begin
          if ic = 0 then
          begin
            ExcelApp.Cells[irow, ifield + 1].Value := ClientDataSet1.Fields[ifield].AsString;
          end
          else                        
          begin                                                                
            ExcelApp.Cells[irow, ifield + 1].Value := '=IF(' + GetRef(ifield - 1) + IntToStr(irow - 1) + '=0,IF(' + GetRef(ifield - 1) + IntToStr(irow) + '=0,0,1),' + GetRef(ifield) + IntToStr(irow) + '/' + GetRef(ifield - 1) + IntToStr(irow - 1) + ')';
            ExcelApp.Cells[irow, ifield + 1].NumberFormatLocal := '0.0%';

            ExcelApp.Range[ExcelApp.Cells[irow, ifield + 1], ExcelApp.Cells[irow, ifield + 1]].FormatConditions.Add(xlCellValue, xlLess, '=-0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow, ifield + 1], ExcelApp.Cells[irow, ifield + 1]].FormatConditions[1].Interior.Color := $0000FF;
            ExcelApp.Range[ExcelApp.Cells[irow, ifield + 1], ExcelApp.Cells[irow, ifield + 1]].FormatConditions.Add(xlCellValue, xlGreater, '=0.2', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow, ifield + 1], ExcelApp.Cells[irow, ifield + 1]].FormatConditions[2].Interior.Color := $0000FF;        
          end;
        end; 
      end;

      ic := ic + 1;
      if ic >= slFGPlans.Count then
      begin
        ic := 0;
      end;
      
      irow := irow + 1;
      ClientDataSet1.Next;
    end;

    AddBorder(ExcelApp, 1, 1, irow - 1, DBGridEh1.Columns.Count); 

    ExcelApp.Range[ ExcelApp.Cells[4, 6], ExcelApp.Cells[4, 6] ].Select;
    ExcelApp.ActiveWindow.FreezePanes := True;

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;

    sl.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);

end;

procedure TfrmFGDemand.btnManageClick(Sender: TObject);
begin
  TfrmFGDemandManage.ShowForm(ADOConnection1);
end;

end.
