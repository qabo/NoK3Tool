unit SalePlanWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, ComCtrls, ToolWin, StdCtrls, ExtCtrls, CommUtils, IniFiles,
  SOPReaderUnit, ComObj, DateUtils, ProjYearWin, SalePlanWFReader, ExcelConsts;

type
  TfrmSalePlan = class(TForm)
    ToolBar1: TToolBar;
    btnSave2: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    ImageList1: TImageList;
    mmoFiles: TMemo;
    Label1: TLabel;
    Button1: TButton;
    Memo1: TMemo;
    ToolButton1: TToolButton;
    leLastWeekWF: TLabeledEdit;
    btnLastWeekWF: TButton;
    procedure btnExitClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure btnLastWeekWFClick(Sender: TObject);
  private
    { Private declarations }
    procedure InsertDateCol(slDate: TStringList; aSOPCol: TSOPCol;
      const ssheet: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmSalePlan.ShowForm;
var
  frmSalePlan: TfrmSalePlan;
begin
  frmSalePlan := TfrmSalePlan.Create(nil);
  frmSalePlan.ShowModal;
  frmSalePlan.Free;
end;

procedure TfrmSalePlan.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := ini.ReadString(self.ClassName, mmoFiles.Name, '');
    mmoFiles.Text := Trim(StringReplace(s, '||', #13#10, [rfReplaceAll]));

    leLastWeekWF.Text := ini.ReadString(self.ClassName, leLastWeekWF.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSalePlan.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
  s: string;
begin
  ini := TIniFile.Create(AppIni);
  try
    s := StringReplace(Trim(mmoFiles.Text), #13#10, '||', [rfReplaceAll]);
    ini.WriteString(self.ClassName, mmoFiles.Name, s);

    ini.WriteString(self.ClassName, leLastWeekWF.Name, leLastWeekWF.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSalePlan.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSalePlan.Button1Click(Sender: TObject);
var
  sfile: string; 
begin
  if not ExcelOpenDialogs(sfile) then Exit;

  mmoFiles.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) ); 
end;
   
procedure TfrmSalePlan.btnLastWeekWFClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leLastWeekWF.Text := sfile;
end;

type
  TDateRecord = packed record
    dt1: TDateTime;
    dt2: TDateTime;
    sdate: string;
    sweek: string;
    snote: string;
  end;
  PDateRecord = ^TDateRecord;

procedure TfrmSalePlan.InsertDateCol(slDate: TStringList; aSOPCol: TSOPCol;
  const ssheet: string);
var
  i: Integer;
  s: string;
  idx: Integer;
  dt: TDateTime;
  aDateRecordPtr: PDateRecord;
begin
  s := FormatDateTime('yyyy-MM-dd', aSOPCol.dt1);
  idx := slDate.IndexOf(s);
  if idx < 0 then
  begin          
    aDateRecordPtr := New(PDateRecord);
    aDateRecordPtr^.dt1 := aSOPCol.dt1;
    aDateRecordPtr^.dt2 := aSOPCol.dt2;
    aDateRecordPtr^.sDate := aSOPCol.sDate;   
    aDateRecordPtr^.sweek := aSOPCol.sWeek;
    aDateRecordPtr^.snote := ssheet + aSOPCol.sDate;

//    if YearOf(aDateRecordPtr^.dt1) = 2019 then
//    asm
//      int 3
//
//    end;

    for i := 0 to slDate.Count - 1 do
    begin
      dt := myStrToDateTime(slDate[i]);
      if dt > aSOPCol.dt1 then
      begin
        Break;
      end;
    end;
    if (i >= 0) and (i < slDate.Count) then
    begin
      slDate.InsertObject(i, s, TObject(aDateRecordPtr));
    end
    else
    begin
      slDate.AddObject(s, TObject(aDateRecordPtr));
    end;
  end
  else
  begin
    aDateRecordPtr := PDateRecord(slDate.Objects[idx]);
    if aDateRecordPtr^.sdate <> aSOPCol.sDate then
    begin
      Memo1.Lines.Add('相同week的列日期不一致'#13#10 + aDateRecordPtr^.snote +
        #13#10 + ssheet + aSOPCol.sDate);
    end;
  end;
end;

function SumByDateVer(lstSOP: TList; const sproj, sver: string; dt1: TDateTime): Double;
var
  aSOPReader: TSOPReader;
  aProj: TSOPProj;
  iproj: Integer;
  ifile: Integer;
  iline: Integer;
  aSOPLine: TSOPLine;
  iDate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iproj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aProj := aSOPReader.Projs[iproj];
      if aProj.FName <> sproj then Continue;
      for iline := 0 to aProj.LineCount - 1 do
      begin
        aSOPLine := aProj.Lines[iline];
        if aSOPLine.sVer <> sver then Continue;

        for iDate := 0 to aSOPLine.DateCount - 1 do
        begin
          aSOPCol := aSOPLine.Dates[iDate];
          if aSOPCol.dt1 = dt1 then
          begin
            Result := Result + aSOPCol.iQty;
            Break;
          end;
        end;
      end;
    end;
  end;
end;
      
function SumByDateCap(lstSOP: TList; const sproj, scap: string; dt1: TDateTime): Double;
var       
  aSOPReader: TSOPReader;
  aProj: TSOPProj;
  iproj: Integer;
  ifile: Integer;
  iline: Integer;
  aSOPLine: TSOPLine;
  iDate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;           
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iproj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aProj := aSOPReader.Projs[iproj];
      if aProj.FName <> sproj then Continue;
      for iline := 0 to aProj.LineCount - 1 do
      begin
        aSOPLine := aProj.Lines[iline];
        if aSOPLine.sCap <> scap then Continue;

        for iDate := 0 to aSOPLine.DateCount - 1 do
        begin
          aSOPCol := aSOPLine.Dates[iDate];
          if aSOPCol.dt1 = dt1 then
          begin
            Result := Result + aSOPCol.iQty;
            Break;
          end;
        end;
      end;
    end;
  end;
end;

function SumByDateCol(lstSOP: TList; const sproj, scol: string; dt1: TDateTime): Double;
var                
  aSOPReader: TSOPReader;
  aProj: TSOPProj;
  iproj: Integer;
  ifile: Integer;
  iline: Integer;
  aSOPLine: TSOPLine;
  iDate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;                       
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iproj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aProj := aSOPReader.Projs[iproj];
      if aProj.FName <> sproj then Continue;
      for iline := 0 to aProj.LineCount - 1 do
      begin
        aSOPLine := aProj.Lines[iline];
        if aSOPLine.sColor <> scol then Continue;

        for iDate := 0 to aSOPLine.DateCount - 1 do
        begin
          aSOPCol := aSOPLine.Dates[iDate];
          if aSOPCol.dt1 = dt1 then
          begin
            Result := Result + aSOPCol.iQty;
            Break;
          end;
        end;
      end;   
    end;
  end;
end;
 
procedure TfrmSalePlan.btnSave2Click(Sender: TObject);
  function GetCol(sl: TStringList; const dt1: TDateTime): Integer;
  var
    i: Integer;
    sdt: string;
  begin
    Result := 0;
    sdt := FormatDateTime('yyyyMMdd', dt1);
    for i := 0 to sl.Count - 1 do
    begin
      if sl[i] = sdt then
      begin
        Result := TSOPCol(sl.Objects[i]).icol;
        Break;
      end;
    end;
  end;
var
  sfile: string;
  sfile_save: string;
  ifile: Integer;
  aSOPReader: TSOPReader;
  lstSOP: TList;
  iProj: Integer;
  jProj: Integer;
  aSOPProj: TSOPProj;
  aSOPProj0: TSOPProj;
  iLine: Integer;
  aSOPLine: TSOPLine;
  //slDate: TStringList;
  iDate: Integer;
  aSOPCol: TSOPCol;
  slver, slcap, slcol: TStringList;
  i: Integer;

  ExcelApp, WorkBook: Variant;
  irow: Integer;
  irow1: Integer;
  irow1_col: Integer;
  icol: Integer;     
  //aDateRecordPtr: PDateRecord;
  slSumRow: TStringList;
  s: string;
  dtMonth: TDateTime;
  icolMax: Integer;
  icol1: Integer;
  iSheet: Integer;
  slMonth: TStringList;
  idx: Integer;
  //slProj: TStringList;
  //slyear: TStringList;
  bLastWeekWF: Boolean;
  ee: Integer;

  aSalePlanWFReader: TSalePlanWFReader;
  iweek: Integer;
  aSalePlanWFWeek: TSalePlanWFWeek;
  aSalePlanWFProj: TSalePlanWFProj;
  aSalePlanWFRecordPtr: PSalePlanWFRecord;
  sweek_this: string;
  sweek_last: string;
  dt1: TDateTime;
  idxDate: Integer;

  slProjWeek: TStringList;
  slWeek: TStringList;
  slWeekAll: TStringList;
  sdt: string;
  imonth: Integer;
begin
  bLastWeekWF := FileExists(leLastWeekWF.Text);
  
  if not bLastWeekWF then
  begin
    ee := MessageBox(Handle, '上周Waterfall不存在，是否继续？', '提示', MB_YESNO);
    if ee <> IDYES then Exit;
  end;


  if not ExcelSaveDialog(sfile_save) then Exit;


  aSalePlanWFReader := TSalePlanWFReader.Create(leLastWeekWF.Text);

  mmoFiles.Lines.Add( StringReplace(sfile, ';', #13#10, [rfReplaceAll]) );

  lstSOP := TList.Create;
  //slDate := TStringList.Create;
  slSumRow := TStringList.Create;
  slMonth := TStringList.Create;
  //slProj := TStringList.Create;

  slProjWeek := TStringList.Create;
  slWeekAll := TStringList.Create;

  //slProj.Clear;
  for ifile := 0 to mmoFiles.Lines.Count - 1 do
  begin
    sfile := Trim(mmoFiles.Lines[ifile]);
    if sfile = '' then Continue;

    aSOPReader := TSOPReader.Create(nil, sfile);
    lstSOP.Add(aSOPReader);
  end;

  // 收集所有的列
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];

      if aSOPProj.FName = 'M1793' then
      begin
        Sleep(1);
      end;

      if aSOPProj.LineCount = 0 then Continue;

      idx := slProjWeek.IndexOf(aSOPProj.FName);
      if idx < 0 then
      begin
        slWeek := TStringList.Create;
        slProjWeek.AddObject(aSOPProj.FName, slWeek); 
      end
      else
      begin
        slWeek := TStringList(slProjWeek.Objects[idx]);
      end;

      for iDate := 0 to aSOPProj.Lines[0].DateCount - 1 do
      begin
        aSOPCol := aSOPProj.Lines[0].Dates[iDate];
        sdt := FormatDateTime('yyyyMMdd', aSOPCol.dt1);
        idx := slWeek.IndexOf(sdt);
        if idx < 0 then
        begin
          slWeek.AddObject( sdt, aSOPCol );
        end;

        idx := slWeekAll.IndexOf(sdt);
        if idx < 0 then
        begin
          slWeekAll.AddObject( sdt, aSOPCol );
        end;
      end;
              
    end;
  end;

  for iProj := 0 to slProjWeek.Count - 1 do
  begin
    slWeek := TStringList(slProjWeek.Objects[iProj]);
    slWeek.Sort;
  end;
  slWeekAll.Sort;

  (*
//   对齐日期 前面的列 /////////////////////////////////////////////////////////
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      if aSOPProj.LineCount = 0 then Continue;

      idx := slProj.IndexOf(aSOPProj.FName);
      if idx < 0 then
      begin
        slProj.AddObject(aSOPProj.FName, aSOPProj.Lines[0]);
      end
      else
      begin
        aSOPLine := TSOPLine(slProj.Objects[idx]);
        if aSOPLine.Dates[0].dt1 > aSOPProj.Lines[0].Dates[0].dt1 then
        begin
          slProj.Objects[idx] := aSOPProj.Lines[0];
        end;
      end;
    end;
  end;
  
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      if aSOPProj.LineCount = 0 then Continue;
      idx := slProj.IndexOf(aSOPProj.FName);
      if idx < 0 then Continue;

      aSOPLine := TSOPLine(slProj.Objects[idx]);
      idxDate := aSOPLine.DateIdx(aSOPProj.Lines[0].Dates[0].dt1);
      for iDate := idxDate - 1 downto 0 do
      begin
        for iLine := 0 to aSOPProj.LineCount - 1 do
        begin 
          aSOPProj.Lines[iLine].Insert(0, aSOPLine.Dates[iDate].sMonth,
            aSOPLine.Dates[iDate].sWeek,
            aSOPLine.Dates[iDate].sDate,
            aSOPLine.Dates[iDate].dt1,
            aSOPLine.Dates[iDate].dt2,
            Round(aSOPLine.Dates[iDate].iQty_sop),
            Round(aSOPLine.Dates[iDate].iQty_mps));
        end;
      end;
    end;
  end;      

//   对齐日期 后面的列 /////////////////////////////////////////////////////////
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      if aSOPProj.LineCount = 0 then Continue;

      idx := slProj.IndexOf(aSOPProj.FName);
      if idx < 0 then
      begin
        slProj.AddObject(aSOPProj.FName, aSOPProj.Lines[0]);
      end
      else
      begin
        aSOPLine := TSOPLine(slProj.Objects[idx]);
        if aSOPLine.Dates[aSOPLine.DateCount - 1].dt1 < aSOPProj.Lines[0].Dates[aSOPProj.Lines[0].DateCount - 1].dt1 then
        begin
          slProj.Objects[idx] := aSOPProj.Lines[0];
        end;
      end;
    end;
  end;
  
  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      if aSOPProj.LineCount = 0 then Continue;
      idx := slProj.IndexOf(aSOPProj.FName);
      if idx < 0 then Continue;

      aSOPLine := TSOPLine(slProj.Objects[idx]);
      idxDate := aSOPLine.DateIdx(aSOPProj.Lines[0].Dates[aSOPProj.Lines[0].DateCount - 1].dt1);
      for iDate := idxDate + 1 to aSOPLine.DateCount - 1 do
      begin
        for iLine := 0 to aSOPProj.LineCount - 1 do
        begin 
          aSOPProj.Lines[iLine].Add(aSOPLine.Dates[iDate].sMonth,
            aSOPLine.Dates[iDate].sWeek,
            aSOPLine.Dates[iDate].sDate,
            aSOPLine.Dates[iDate].dt1,
            aSOPLine.Dates[iDate].dt2,
            Round(aSOPLine.Dates[iDate].iQty_sop),
            Round(aSOPLine.Dates[iDate].iQty_mps));
        end;
      end;
    end;
  end;


  

  slProj.Clear;

  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      if aSOPProj.LineCount = 0 then Continue;
      if Pos('线上', aSOPProj.FName) > 0 then Continue;
      if Pos('线下', aSOPProj.FName) > 0 then Continue;

      idx := slProj.IndexOf(aSOPProj.FName);
      if idx < 0 then
      begin
        slProj.AddObject(aSOPProj.FName, aSOPProj);
      end
      else
      begin
        aSOPProj0 := TSOPProj(slProj.Objects[idx]);
        for iLine := aSOPProj.LineCount - 1 downto 0 do
        begin
          aSOPLine := aSOPProj.Lines[iLine];
          if aSOPProj0.GetLine(aSOPLine.sVer, aSOPLine.sNumber, aSOPLine.sColor, aSOPLine.sCap) = nil then
          begin
            if aSOPProj0.Lines[0].DateCount <> aSOPLine.DateCount then
            begin
              Sleep(15);
            end;
            aSOPProj0.FList.AddObject(aSOPLine.sNumber, aSOPLine);
            aSOPProj.FList.Delete(iLine);
          end;  
        end;
      end;
    end;
  end;

  for iProj := 0 to slProj.Count - 1 do
  begin
    aSOPProj := TSOPProj( slProj.Objects[iProj] );
    if aSOPProj.LineCount = 0 then Continue;
    if Pos('线上', aSOPProj.FName) > 0 then Continue;
    if Pos('线下', aSOPProj.FName) > 0 then Continue;

    aSOPLine := aSOPProj.Lines[0];
    for iDate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[iDate];
      InsertDateCol(slDate, aSOPCol, aSOPProj.FName);
    end;
  end;

  for iDate := 0 to slDate.Count - 1 do
  begin
    aDateRecordPtr := PDateRecord(slDate.Objects[iDate]);
    Memo1.Lines.Add( IntToStr(iDate) + ': ' + aDateRecordPtr^.sdate + '  ' + FormatDateTime('yyyy-MM-dd', aDateRecordPtr^.dt1) );
  end;
        
 *)
 
  // 开始保存 Excel
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := True;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;
  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  try
          
    ExcelApp.Sheets[1].Activate;   
    ExcelApp.Sheets[1].Name := '销售计划汇总';

    irow := 1;                               

    ExcelApp.Cells[irow, 1].Value := '项目';
    MergeCells(ExcelApp, irow, 1, irow + 1, 1);

    ExcelApp.Cells[irow, 2].Value := '分类';         
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);

    dtMonth := 0;
    icol := 3;

    for iDate := 0 to slWeekAll.Count - 1 do
    begin
      aSOPCol := TSOPCol(slWeekAll.Objects[iDate]);  

      if dtMonth = 0 then
      begin
        dtMonth := aSOPCol.dt1;
      end;

      if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
      begin
        ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);
        ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
        MergeCells(ExcelApp, irow, icol, irow + 1, icol);
        CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);

        dtMonth := aSOPCol.dt1;

        slMonth.Add(IntToStr(icol));

        icol := icol + 1;
      end;
      

      ExcelApp.Cells[irow, icol].Value := aSOPCol.sweek;
      ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.sdate;
      icol := icol + 1;
    end;

    ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);   
    ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
    MergeCells(ExcelApp, irow, icol, irow + 1, icol);
    CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
    slMonth.Add(IntToStr(icol));


    icolMax := icol;


    irow := irow + 2;
    
    slver := TStringList.Create;
    slcap := TStringList.Create;
    slcol := TStringList.Create;


    for iProj := 0 to slProjWeek.Count - 1 do
    begin   
      ExcelApp.Cells[irow, 1].Value := slProjWeek[iProj];
        
      for ifile := 0 to lstSOP.Count - 1 do
      begin
        aSOPReader := TSOPReader(lstSOP[ifile]);
        for jProj := 0 to aSOPReader.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader.Projs[jProj];
          if aSOPProj.FName <> slProjWeek[iProj] then Continue;

          if aSOPProj.LineCount = 0 then Continue;
          if Pos('线上', aSOPProj.FName) > 0 then Continue;
          if Pos('线下', aSOPProj.FName) > 0 then Continue;  

          for iLine := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iLine];
            if slver.IndexOf( aSOPLine.sVer ) < 0 then
            begin
              slver.Add( aSOPLine.sVer );
            end;  
            if slcap.IndexOf( aSOPLine.sCap ) < 0 then
            begin
              slcap.Add( aSOPLine.sCap );
            end;
            if slcol.IndexOf( aSOPLine.sColor ) < 0 then
            begin
              slcol.Add( aSOPLine.sColor );
            end;
          end;
        end;
      end;

      irow1 := irow;


      for i := 0 to slver.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slver[i];

        dtMonth := 0;
        icol := 3;
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );                   

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;   
            icol1 := icol; 
          end;
      

          ExcelApp.Cells[irow, icol].Value := SumByDateVer(lstSOP, slProjWeek[iProj], slver[i], aSOPCol.dt1);  
          icol := icol + 1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

        irow := irow + 1; 
      end;   

      for i := 0 to slcap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 2].Value := slcap[i];   

        dtMonth := 0;   
        icol := 3;      
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );                   

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;
            icol1 := icol;
          end;

          ExcelApp.Cells[irow, icol].Value := SumByDateCap(lstSOP, slProjWeek[iProj], slcap[i], aSOPCol.dt1);
          icol := icol + 1;
        end;            

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

        irow := irow + 1;
      end;     

      irow1_col := irow;
      for i := 0 to slcol.Count - 1 do
      begin            
        ExcelApp.Cells[irow, 2].Value := slcol[i];   

        dtMonth := 0;       
        icol := 3;        
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );     

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;
            icol1 := icol;
          end;

          ExcelApp.Cells[irow, icol].Value := SumByDateCol(lstSOP, slProjWeek[iProj], slcol[i], aSOPCol.dt1);   
          icol := icol + 1;
        end;               

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

        irow := irow + 1;
      end;

      ExcelApp.Cells[irow, 2].Value := '合计';
      CenterBoldCells(ExcelApp, irow, 2, irow, 2);

      dtMonth := 0;        
      icol := 3;    
      icol1 := icol;
      for idate := 0 to slWeekAll.Count - 1 do
      begin
        aSOPCol := TSOPCol( slWeekAll.Objects[idate] );

        if dtMonth = 0 then
        begin
          dtMonth := aSOPCol.dt1;
        end;

        if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
        begin
          ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
          dtMonth :=aSOPCol.dt1;
          icol := icol + 1;
          icol1 := icol;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + IntToStr(irow1_col) + ':' + GetRef(icol) + IntToStr(irow - 1) + ')';
        icol := icol + 1;
      end;             

      ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                                                       
      AddColor(ExcelApp, irow, 2, irow, icolMax, $EAEAEA);

      slSumRow.Add(IntToStr(irow));

      MergeCells(ExcelApp, irow1, 1, irow, 1);
      irow := irow + 1;
            
    end;

    ExcelApp.Cells[irow, 1].Value := '合计';
    MergeCells(ExcelApp, irow, 1, irow, 2);
    CenterBoldCells(ExcelApp, irow, 1, irow, 2);

    dtMonth := 0;
    icol := 3;         
    icol1 := icol;
    for idate := 0 to slWeekAll.Count - 1 do
    begin
      aSOPCol := TSOPCol( slWeekAll.Objects[idate] );

      if dtMonth = 0 then
      begin
        dtMonth := aSOPCol.dt1;
      end;

      if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
      begin
        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
        dtMonth := aSOPCol.dt1;
        icol := icol + 1;
        icol1 := icol;
      end;

      s := '=0';
      for i := 0 to slSumRow.Count - 1 do
      begin
        s := s + '+' + GetRef(icol) + slSumRow[i];
      end;
      ExcelApp.Cells[irow, icol].Value := s;     
      icol := icol + 1;
    end;
    
    ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                                                     
    AddColor(ExcelApp, 1, 1, 2, icolMax, $EAEAEA);

    AddColor(ExcelApp, irow, 1, irow, icolMax, $EAEAEA);

    for iDate := 0 to slMonth.Count - 1 do
    begin
      icol := StrToInt(slMonth[iDate]);
      AddColor(ExcelApp, 1, icol, irow, icol, $99E6FF);
    end;

    AddBorder(ExcelApp, 1, 1, irow, icolMax);

    slver.Free;
    slcap.Free;
    slcol.Free;

    /////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////


    slMonth.Clear;

               
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[1]);
    ExcelApp.Sheets[2].Activate;
    ExcelApp.Sheets[2].Name := '销售计划汇总（国内海外）';

    irow := 1;                               

    ExcelApp.Cells[irow, 1].Value := '项目';
    MergeCells(ExcelApp, irow, 1, irow + 1, 1);

    ExcelApp.Cells[irow, 2].Value := '分类';         
    MergeCells(ExcelApp, irow, 2, irow + 1, 2);

    dtMonth := 0;
    icol := 3;
    for idate := 0 to slWeekAll.Count - 1 do
    begin
      aSOPCol := TSOPCol( slWeekAll.Objects[idate] );

      if dtMonth = 0 then
      begin
        dtMonth := aSOPCol.dt1;
      end;

      if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
      begin
        ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);
        ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
        MergeCells(ExcelApp, irow, icol, irow + 1, icol);
        CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);

        dtMonth := aSOPCol.dt1;

        slMonth.Add(IntToStr(icol));

        icol := icol + 1;
      end;
      

//      ExcelApp.Cells[irow, icol].Value := aDateRecordPtr^.sweek;
//      ExcelApp.Cells[irow + 1, icol].Value := aDateRecordPtr^.sdate;
//      icol := icol + 1;
    end;

    ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);   
    ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
    MergeCells(ExcelApp, irow, icol, irow + 1, icol);
    CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
    slMonth.Add(IntToStr(icol));


    icolMax := icol;


    irow := irow + 2;
    
    slver := TStringList.Create; 

    for iProj := 0 to slProjWeek.Count - 1 do
    begin
//      aSOPProj := TSOPProj(slProj.Objects[iProj]);
//      if aSOPProj.LineCount = 0 then Continue;

      ExcelApp.Cells[irow, 1].Value := slProjWeek[iProj] ;

      for ifile := 0 to lstSOP.Count - 1 do
      begin
        aSOPReader := TSOPReader(lstSOP[ifile]);
        for jProj := 0 to aSOPReader.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader.Projs[jProj];
          if aSOPProj.FName <> slProjWeek[iProj] then Continue;

          if Pos('线上', aSOPProj.FName) > 0 then Continue;
          if Pos('线下', aSOPProj.FName) > 0 then Continue;

          for iLine := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iLine];
            if slver.IndexOf( aSOPLine.sVer ) < 0 then
            begin
              slver.Add( aSOPLine.sVer );
            end;   
          end;
        end;
      end;
      
      irow1 := irow;
 
      ExcelApp.Cells[irow, 2].Value := '国内';
      for i := 0 to slver.Count - 1 do
      begin
        if IsVerHW(slver[i]) then Continue;
        
        dtMonth := 0;
        icol := 3;
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );                  

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;   
            icol1 := icol; 
          end;
      

          ExcelApp.Cells[irow, icol].Value := ExcelApp.Cells[irow, icol].Value + SumByDateVer(lstSOP, slProjWeek[iProj], slver[i], aSOPCol.dt1);  
//          icol := icol + 1;
        end;

        //ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

        //irow := irow + 1; 
      end; 
      irow := irow + 1;

                                           
      ExcelApp.Cells[irow, 2].Value := '海外';
      for i := 0 to slver.Count - 1 do
      begin
        if not IsVerHW(slver[i]) then Continue;
        
        dtMonth := 0;
        icol := 3;
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );                       

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;   
            icol1 := icol; 
          end;
      

          ExcelApp.Cells[irow, icol].Value := ExcelApp.Cells[irow, icol].Value + SumByDateVer(lstSOP, slProjWeek[iProj], slver[i], aSOPCol.dt1);  
//          icol := icol + 1;
        end;

        //ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';

        //irow := irow + 1; 
      end; 
      irow := irow + 1;


      ExcelApp.Cells[irow, 2].Value := '合计';
      CenterBoldCells(ExcelApp, irow, 2, irow, 2);

      dtMonth := 0;        
      icol := 3;    
      icol1 := icol;
      for idate := 0 to slWeekAll.Count - 1 do
      begin
        aSOPCol := TSOPCol( slWeekAll.Objects[idate] );   

        if dtMonth = 0 then
        begin
          dtMonth := aSOPCol.dt1;
        end;

        if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
        begin
//          ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
          dtMonth :=aSOPCol.dt1;
          icol := icol + 1;
          icol1 := icol;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow - 1) + ')';
//        icol := icol + 1;
      end;

      //ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                                                       
      AddColor(ExcelApp, irow, 2, irow, icolMax, $EAEAEA);

      slSumRow.Add(IntToStr(irow));

      MergeCells(ExcelApp, irow1, 1, irow, 1);
      irow := irow + 1;

    end; 

    ExcelApp.Cells[irow, 1].Value := '合计';
    MergeCells(ExcelApp, irow, 1, irow, 2);
    CenterBoldCells(ExcelApp, irow, 1, irow, 2);

    dtMonth := 0;
    icol := 3;         
    icol1 := icol;
    for idate := 0 to slWeekAll.Count - 1 do
    begin
      aSOPCol := TSOPCol( slWeekAll.Objects[idate] );   

      if dtMonth = 0 then
      begin
        dtMonth := aSOPCol.dt1;
      end;

      if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
      begin
//        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
        dtMonth := aSOPCol.dt1;
        icol := icol + 1;
        icol1 := icol;
      end;

      s := '=0';
      for i := 0 to slSumRow.Count - 1 do
      begin
        s := s + '+' + GetRef(icol) + slSumRow[i];
      end;
      ExcelApp.Cells[irow, icol].Value := s;     
//      icol := icol + 1;
    end;
    
    //ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
                                                     
    AddColor(ExcelApp, 1, 1, 2, icolMax, $EAEAEA);

    AddColor(ExcelApp, irow, 1, irow, icolMax, $EAEAEA);

    for iDate := 0 to slMonth.Count - 1 do
    begin
      icol := StrToInt(slMonth[iDate]);
      AddColor(ExcelApp, 1, icol, irow, icol, $99E6FF);
    end;

    AddBorder(ExcelApp, 1, 1, irow, icolMax);

    slver.Free;


    /////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////

    slMonth.Clear;
               
    ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[2]);
    ExcelApp.Sheets[3].Activate;
    ExcelApp.Sheets[3].Name := 'Waterfall';

    irow := 1;
    
    ExcelApp.Cells[irow, 1].Value := '项目';
    ExcelApp.Cells[irow, 2].Value := '国内/海外';
    ExcelApp.Cells[irow, 3].Value := '销售计划';

    ExcelApp.Columns[1].ColumnWidth:= 8.38;
    ExcelApp.Columns[2].ColumnWidth:= 8.38;
    ExcelApp.Columns[3].ColumnWidth:= 18.25;
 
    dtMonth := 0;
    icol := 4;
    for idate := 0 to slWeekAll.Count - 1 do
    begin
      aSOPCol := TSOPCol( slWeekAll.Objects[idate] );   

      if dtMonth = 0 then
      begin
        dtMonth := aSOPCol.dt1;
      end;

      if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
      begin
        ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);
        ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';        
        ExcelApp.Columns[icol].ColumnWidth:= 8.38;
//        MergeCells(ExcelApp, irow, icol, irow + 1, icol);
        CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);

        dtMonth := aSOPCol.dt1;

        slMonth.Add(IntToStr(icol));

        icol := icol + 1;
      end;
      

//      ExcelApp.Cells[irow, icol].Value := aDateRecordPtr^.sweek;
//      ExcelApp.Cells[irow + 1, icol].Value := aDateRecordPtr^.sdate;
//      icol := icol + 1;
    end;

    ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);   
    ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';            
    ExcelApp.Columns[icol].ColumnWidth:= 8.38;
//    MergeCells(ExcelApp, irow, icol, irow + 1, icol);
    CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
    slMonth.Add(IntToStr(icol));

    icol := icol + 1;
    ExcelApp.Cells[irow, icol].Value := '合计';     
    ExcelApp.Columns[icol].ColumnWidth:= 10.75;
                     
    icol := icol + 1;
    ExcelApp.Cells[irow, icol].Value := '备注';    
    ExcelApp.Columns[icol].ColumnWidth:= 16.38;

    icolMax := icol;

    AddColor(ExcelApp, irow, 4, irow, icolMax, $C47244);

    sweek_this := '';
    if mmoFiles.Lines.Count > 0 then
    begin
      sweek_this := ChangeFileExt( ExtractFileName( mmoFiles.Lines[0] ), '' )
    end;

    sweek_last := '';
    if aSalePlanWFReader.Count > 0 then
    begin
      aSalePlanWFWeek := aSalePlanWFReader.Items[0];
      sweek_last := aSalePlanWFWeek.sweek1;
    end;

    irow := irow + 1;
    
    slver := TStringList.Create; 

    for iProj := 0 to slProjWeek.Count - 1 do
    begin
      ExcelApp.Cells[irow, 1].Value := slProjWeek[iproj];
              
      slver.Clear;

      for ifile := 0 to lstSOP.Count - 1 do
      begin
        aSOPReader := TSOPReader(lstSOP[ifile]);
        for jProj := 0 to aSOPReader.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader.Projs[jProj];
          if aSOPProj.FName <> slProjWeek[iProj] then Continue;

          if Pos('线上', aSOPProj.FName) > 0 then Continue;
          if Pos('线下', aSOPProj.FName) > 0 then Continue;

          for iLine := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iLine];
            if slver.IndexOf( aSOPLine.sVer ) < 0 then
            begin
              slver.Add( aSOPLine.sVer );
            end;
          end;
        end;
      end;
      
      irow1 := irow;
 
      ExcelApp.Cells[irow, 2].Value := '国内';
      MergeCells(ExcelApp, irow, 2, irow + 2, 2);

      ExcelApp.Cells[irow, 3].Value := '销售计划(' + sweek_this + ')';
      ExcelApp.Cells[irow + 1, 3].Value := sweek_last;
      ExcelApp.Cells[irow + 2, 3].Value := '差异值';
      for i := 0 to slver.Count - 1 do
      begin
        if IsVerHW(slver[i]) then Continue;
        
        dtMonth := 0;
        icol := 4;
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;   
            icol1 := icol; 
          end;
      

          ExcelApp.Cells[irow, icol].Value := aSalePlanWFReader.GetLastWeekQty(aSOPProj.FName, aSOPCol.dt1, True);
          ExcelApp.Cells[irow + 1, icol].Value := ExcelApp.Cells[irow + 1, icol].Value + SumByDateVer(lstSOP, slProjWeek[iproj], slver[i], aSOPCol.dt1);
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);  
//          icol := icol + 1;
        end;

        icol := icol + 1;   
        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')';

        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions[1].Font.Color := $0000FF;
        //irow := irow + 1; 
      end; 
      irow := irow + 3;

                                           
      ExcelApp.Cells[irow, 2].Value := '海外';
      MergeCells(ExcelApp, irow, 2, irow + 2, 2);   

      ExcelApp.Cells[irow, 3].Value := '销售计划(' + sweek_this + ')';
      ExcelApp.Cells[irow + 1, 3].Value := sweek_last;
      ExcelApp.Cells[irow + 2, 3].Value := '差异值';
      for i := 0 to slver.Count - 1 do
      begin
        if not IsVerHW(slver[i]) then Continue;
        
        dtMonth := 0;
        icol := 4;
        icol1 := icol;
        for idate := 0 to slWeekAll.Count - 1 do
        begin
          aSOPCol := TSOPCol( slWeekAll.Objects[idate] );              

          if dtMonth = 0 then
          begin
            dtMonth := aSOPCol.dt1;
          end;

          if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
            dtMonth := aSOPCol.dt1;
            icol := icol + 1;   
            icol1 := icol; 
          end;
       

          ExcelApp.Cells[irow, icol].Value := aSalePlanWFReader.GetLastWeekQty(aSOPProj.FName, aSOPCol.dt1, False);
          ExcelApp.Cells[irow + 1, icol].Value := ExcelApp.Cells[irow + 1, icol].Value + SumByDateVer(lstSOP, slProjWeek[iproj], slver[i], aSOPCol.dt1);
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);

//          icol := icol + 1;
        end;


        icol := icol + 1;   
        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')';

        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions[1].Font.Color := $0000FF;
        //irow := irow + 1; 
      end;
                                                   
      irow := irow + 3;
      
      MergeCells(ExcelApp, irow1, 1, irow - 1, 1);

    end; 
 
    AddBorder(ExcelApp, 1, 1, irow - 1, icol + 1);
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol + 1] ].Borders.Color := $F0B000;
    ExcelApp.Range[ ExcelApp.Cells[2, 4], ExcelApp.Cells[irow - 1, icol] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';


    slver.Free;

    //  写过去的week
    for iweek := 0 to aSalePlanWFReader.Count - 1 do
    begin
      irow := irow + 2;

      irow1 := irow;
      aSalePlanWFWeek := aSalePlanWFReader.Items[iweek];
 
      ExcelApp.Cells[irow, 1].Value := '项目';
      ExcelApp.Cells[irow, 2].Value := '国内/海外';
      ExcelApp.Cells[irow, 3].Value := '销售计划';

      ExcelApp.Columns[1].ColumnWidth:= 8.38;
      ExcelApp.Columns[2].ColumnWidth:= 8.38;
      ExcelApp.Columns[3].ColumnWidth:= 18.25;

      for iProj := 0 to aSalePlanWFWeek.Count - 1 do
      begin
        aSalePlanWFProj := aSalePlanWFWeek.Items[iProj];

        if iProj = 0 then
        begin
          icol := 4;
          for iDate := 0 to aSalePlanWFProj.Count - 1 do
          begin
            aSalePlanWFRecordPtr := aSalePlanWFProj.Items[iDate];
            
            ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', aSalePlanWFRecordPtr^.dt);
            ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';            
            ExcelApp.Columns[icol].ColumnWidth:= 8.38;
            CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
            icol := icol + 1;
          end;
          ExcelApp.Cells[irow, icol].Value := '合计';
          ExcelApp.Cells[irow, icol + 1].Value := '备注';

          AddColor(ExcelApp, irow, 4, irow, icol + 1, $C47244);
          
          irow := irow + 1;
        end;

        ExcelApp.Cells[irow, 1].Value := aSalePlanWFProj.sname;
        MergeCells(ExcelApp, irow, 1, irow + 5, 1);
        ExcelApp.Cells[irow, 2].Value := '国内';
        MergeCells(ExcelApp, irow, 2, irow + 2, 2);  
        ExcelApp.Cells[irow, 2].Value := '海外';
        MergeCells(ExcelApp, irow + 3, 2, irow + 5, 2);

                         
        ExcelApp.Cells[irow, 3].Value := aSalePlanWFWeek.sweek1;
        ExcelApp.Cells[irow + 1, 3].Value := aSalePlanWFWeek.sweek2;
        ExcelApp.Cells[irow + 2, 3].Value := '差异值';     
        ExcelApp.Cells[irow + 3, 3].Value := aSalePlanWFWeek.sweek1;
        ExcelApp.Cells[irow + 4, 3].Value := aSalePlanWFWeek.sweek2;
        ExcelApp.Cells[irow + 5, 3].Value := '差异值';

        icol := 4;
        for iDate := 0 to aSalePlanWFProj.Count - 1 do
        begin
          aSalePlanWFRecordPtr := aSalePlanWFProj.Items[iDate];

          ExcelApp.Cells[irow, icol].Value := aSalePlanWFRecordPtr^.dqty_ib1;
          ExcelApp.Cells[irow + 1, icol].Value := aSalePlanWFRecordPtr^.dqty_ib2;
          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
          ExcelApp.Cells[irow + 3, icol].Value := aSalePlanWFRecordPtr^.dqty_ob1;
          ExcelApp.Cells[irow + 4, icol].Value := aSalePlanWFRecordPtr^.dqty_ob2;
          ExcelApp.Cells[irow + 5, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 4);
          icol := icol + 1;
        end;

        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
        ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1) + ')';
        ExcelApp.Cells[irow + 2, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 2) + ':' + GetRef(icol - 1) + IntToStr(irow + 2) + ')';
        ExcelApp.Cells[irow + 3, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 3) + ':' + GetRef(icol - 1) + IntToStr(irow + 3) + ')';
        ExcelApp.Cells[irow + 4, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 4) + ':' + GetRef(icol - 1) + IntToStr(irow + 4) + ')';
        ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(4) + IntToStr(irow + 5) + ':' + GetRef(icol - 1) + IntToStr(irow + 5) + ')';

        ExcelApp.Cells[irow, icol + 1].Value := aSalePlanWFProj.snote_ib;
        MergeCells(ExcelApp, irow, icol + 1, irow + 2, icol + 1);
        ExcelApp.Cells[irow + 3, icol + 1].Value := aSalePlanWFProj.snote_ob;
        MergeCells(ExcelApp, irow + 3, icol + 1, irow + 5, icol + 1);

        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 2, 4], ExcelApp.Cells[irow + 2, icol] ].FormatConditions[1].Font.Color := $0000FF;
        ExcelApp.Range[ ExcelApp.Cells[irow + 5, 4], ExcelApp.Cells[irow + 5, icol] ].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
        ExcelApp.Range[ ExcelApp.Cells[irow + 5, 4], ExcelApp.Cells[irow + 5, icol] ].FormatConditions[1].Font.Color := $0000FF;

        irow := irow + 6;
      end;

      AddBorder(ExcelApp, irow1, 1, irow - 1, icol + 1);       
      ExcelApp.Range[ ExcelApp.Cells[irow1, 1], ExcelApp.Cells[irow - 1, icol + 1] ].Borders.Color := $F0B000;
      ExcelApp.Range[ ExcelApp.Cells[irow1 + 1, 4], ExcelApp.Cells[irow - 1, icol] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    end;
    


    /////////////////////////////////////////////////////////////
    /////////////////////////////////////////////////////////////
    // 项目明细 /////////////////////////////////////////////////

    iSheet := 3;

    for iProj := 0 to slProjWeek.Count - 1 do
    begin
      slWeek := TStringList( slProjWeek.Objects[iProj] );
      
      //aSOPProj := TSOPProj(slProjWeek.Objects[iProj]);
      
      ExcelApp.Sheets.Add(after:=ExcelApp.Sheets[iSheet]);
      iSheet := iSheet + 1;
      ExcelApp.Sheets[iSheet].Name := slProjWeek[iProj];

      
      slMonth.Clear;
      
      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '制式';
      ExcelApp.Cells[irow, 2].Value := '物料编码';
      ExcelApp.Cells[irow, 3].Value := '颜色';
      ExcelApp.Cells[irow, 4].Value := '容量';

      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
      MergeCells(ExcelApp, irow, 4, irow + 1, 4);

      dtMonth := 0;
      icol := 5;
      
      for iDate := 0 to slWeek.Count - 1 do
      begin
        aSOPCol := TSOPCol(slWeek.Objects[iDate]);

        if dtMonth = 0 then
        begin
          dtMonth := aSOPCol.dt1;
        end;

        if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
        begin
          ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);   
          ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
          MergeCells(ExcelApp, irow, icol, irow + 1, icol);
          CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
          dtMonth := aSOPCol.dt1;
          slMonth.Add(IntToStr(icol));
          icol := icol + 1;
        end;

        aSOPCol.icol := icol;

        ExcelApp.Cells[irow, icol].Value := aSOPCol.sWeek;
        ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.sDate;  
        icol := icol + 1;
      end;

      ExcelApp.Cells[irow, icol].Value := FormatDateTime('yyyy-MM-01', dtMonth);
      ExcelApp.Cells[irow, icol].NumberFormatLocal := 'm"月";@';
      MergeCells(ExcelApp, irow, icol, irow + 1, icol);
      CenterBoldCells(ExcelApp, irow, icol, irow + 1, icol);
      dtMonth := aSOPCol.dt1;
      slMonth.Add(IntToStr(icol));

      icolMax := icol;

      irow := irow + 2;
      for ifile := 0 to lstSOP.Count - 1 do
      begin
        aSOPReader := TSOPReader(lstSOP[ifile]);
        for jProj := 0 to aSOPReader.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader.Projs[jProj];
          if aSOPProj.FName <> slProjWeek[iProj] then Continue;

          for iLine := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iLine];

            /////////////////////////////////////////////////////////
            ExcelApp.Cells[irow, 1].Value := aSOPLine.sVer;
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sCap;
                  
            icol := 5;
            icol1 := icol;
            dtMonth := 0;
            for iDate := 0 to aSOPLine.DateCount - 1 do
            begin
              aSOPCol := aSOPLine.Dates[iDate];
              icol := GetCol(slWeek, aSOPCol.dt1);
              if icol = 0 then
              begin
                sleep(1);
              end;

//              if dtMonth = 0 then
//              begin
//                dtMonth := aSOPCol.dt1;
//              end;

//              if MonthOfTheYear(dtMonth) <> MonthOfTheYear(aSOPCol.dt1) then
//              begin
//                ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
//                dtMonth := aSOPCol.dt1;
//                icol := icol + 1;
//                icol1 := icol;
//              end;

              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;       
//              icol := icol + 1;
            end;
//            if not DoubleE(dtMonth, 0) then
//            begin
//              ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
//            end;

            icol1 := 0;
            for imonth := 0 to slMonth.Count - 1 do
            begin
              icol := StrToInt(slMonth[imonth]);
              if icol1 = 0 then
              begin
                icol1 := 5;
              end;
              ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow) + ')';
              icol1 := icol + 1;
            end;

            irow := irow + 1;
        
          end;
        end;
      end;
       
//      if not DoubleE(dtMonth, 0) then
//      begin
      ExcelApp.Cells[irow, 1].Value := '合计';
      MergeCells(ExcelApp, irow, 1, irow, 4);
      CenterBoldCells(ExcelApp, irow, 1, irow, 4);
      for icol := 5 to icolMax do
      begin
        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + '3:' + GetRef(icol) + IntToStr(irow - 1) + ')';
      end;
//      end;

                                                              
      AddColor(ExcelApp, 1, 1, 2, icolMax, $EAEAEA);
      AddColor(ExcelApp, irow, 1, irow, icolMax, $EAEAEA);
      
      for iDate := 0 to slMonth.Count - 1 do
      begin
        icol := StrToInt(slMonth[iDate]);
        AddColor(ExcelApp, 1, icol, irow, icol, $99E6FF);
      end;

      AddBorder(ExcelApp, 1, 1, irow, icolMax);
    end;
 

    ExcelApp.Sheets[3].Activate;

    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end; 


  for iproj := 0 to slProjWeek.Count - 1 do
  begin
    slWeek := TStringList(slProjWeek.Objects[iproj]);
    slWeek.Free;
  end;
  slProjWeek.Free;

  slWeekAll.Free;

  for ifile := 0 to lstSOP.Count - 1 do
  begin
    aSOPReader := TSOPReader(lstSOP[ifile]);
    aSOPReader.Free;
  end;
  lstSOP.Free;
//  slDate.Free;
  slSumRow.Free;
 
  slMonth.Free;
//  slProj.Free;

  aSalePlanWFReader.Free;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
