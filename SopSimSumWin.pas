unit SopSimSumWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, StdCtrls, ExtCtrls, CommUtils, ComObj,
  DateUtils, SOPReaderUnit, SOPSimReader, KeyICItemSupplyReader, MrpMPSReader,
  ProjYearWin, SBomReader, FGPriorityReader, DOSPlanReader, StockBalReader,
  SEOutReader, DailyPlanVsActReader, DOSReader, ExcelConsts, IniFiles;
 

type
  TfrmSopSimSum = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    btnSave_demand: TToolButton;
    ToolButton5: TToolButton;
    btnExit: TToolButton;
    ToolButton7: TToolButton;
    leDOSPlan: TLabeledEdit;
    leStockBal: TLabeledEdit;
    leSEOut: TLabeledEdit;
    leDailyPlanVsAct: TLabeledEdit;
    leSellPlan: TLabeledEdit;
    btnDOSPlan: TButton;
    btnStockBal: TButton;
    btnSEOut: TButton;
    btnDailyPlanVsAct: TButton;
    btnSellPlan: TButton;
    leDemand: TLabeledEdit;
    btnDemand: TButton;
    leSopSim: TLabeledEdit;
    btnSopSim: TButton;
    btnSave: TToolButton;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    procedure btnExitClick(Sender: TObject);
    procedure btnDOSPlanClick(Sender: TObject);
    procedure btnStockBalClick(Sender: TObject);
    procedure btnSEOutClick(Sender: TObject);
    procedure btnDailyPlanVsActClick(Sender: TObject);
    procedure btnSellPlanClick(Sender: TObject);
    procedure btnDemandClick(Sender: TObject);
    procedure btnSopSimClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSave_demandClick(Sender: TObject);
  private
    { Private declarations }
    procedure Log(const s: string); 
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation

{$R *.dfm}

class procedure TfrmSopSimSum.ShowForm;
var
  frmSopSim: TfrmSopSimSum;
begin
  frmSopSim := TfrmSopSimSum.Create(nil);
  frmSopSim.ShowModal;
  frmSopSim.Free;
end;

procedure TfrmSopSimSum.Log(const s: string);
begin
  Memo1.Lines.Add(s);
end;

procedure TfrmSopSimSum.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSopSimSum.btnDOSPlanClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDOSPlan.Text := sfile;
end;

procedure TfrmSopSimSum.btnStockBalClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leStockBal.Text := sfile;
end;

procedure TfrmSopSimSum.btnSEOutClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSEOut.Text := sfile;
end;

procedure TfrmSopSimSum.btnDailyPlanVsActClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDailyPlanVsAct.Text := sfile;
end;

procedure TfrmSopSimSum.btnSellPlanClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSellPlan.Text := sfile;
end;
  
procedure TfrmSopSimSum.btnDemandClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leDemand.Text := sfile;
end;

procedure TfrmSopSimSum.btnSopSimClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSopSim.Text := sfile;
end;
   
procedure SaveMPS_SOP(const sfile_save: string; aSOPReader: TSOPReader);
var
  ExcelApp, WorkBook: Variant;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  irow: Integer;
  icol: Integer;
  dt0: TDateTime;
  icol1: Integer;  
  icol2: Integer;
  irow1: Integer;
  sver0: string;
  slver: TStringList;
  slcap: TStringList;
  slcolor: TStringList;
  sver: string;
  idx: Integer;
  i: Integer;
  sl: TStringList;
  s: string;
begin
        
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  icol2 := 1;

  
  slver := TStringList.Create;
  slcap := TStringList.Create;
  slcolor := TStringList.Create;

  WorkBook := ExcelApp.WorkBooks.Add;

  try
    while ExcelApp.Sheets.Count < aSOPReader.FProjs.Count do
    begin
      ExcelApp.Sheets.Add;
    end;

    for iproj := 0 to aSOPReader.FProjs.Count - 1 do
    begin
      aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[iproj]);
      
      ExcelApp.Sheets[iproj + 1].Activate;
      ExcelApp.Sheets[iproj + 1].Name := aSOPProj.FName;    

      
      irow := 1;
      ExcelApp.Cells[irow, 1].Value := '项目';
      ExcelApp.Cells[irow, 2].Value := '整机/裸机';
      ExcelApp.Cells[irow, 3].Value := '包装';
      ExcelApp.Cells[irow, 4].Value := '标准制式';
      ExcelApp.Cells[irow, 5].Value := '制式';
      ExcelApp.Cells[irow, 6].Value := '物料编码';
      ExcelApp.Cells[irow, 7].Value := '颜色';
      ExcelApp.Cells[irow, 8].Value := '容量';

      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
      MergeCells(ExcelApp, irow, 4, irow + 1, 4);
      MergeCells(ExcelApp, irow, 5, irow + 1, 5);
      MergeCells(ExcelApp, irow, 6, irow + 1, 6);
      MergeCells(ExcelApp, irow, 7, irow + 1, 7);
      MergeCells(ExcelApp, irow, 8, irow + 1, 8);

      slver.Clear;
      slcap.Clear;
      slcolor.Clear;

      sver0 := '';
      irow := 3;
      irow1 := irow;
      for iline := 0 to aSOPProj.FList.Count - 1 do
      begin
        aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);

        ExcelApp.Cells[irow, 1].Value := aSOPLine.sProj;
        ExcelApp.Cells[irow, 2].Value := aSOPLine.sFG;
        ExcelApp.Cells[irow, 3].Value := aSOPLine.sPkg;
        ExcelApp.Cells[irow, 4].Value := aSOPLine.sStdVer;
        ExcelApp.Cells[irow, 6].Value := aSOPLine.sNumber;
        ExcelApp.Cells[irow, 7].Value := aSOPLine.sColor;
        ExcelApp.Cells[irow, 8].Value := aSOPLine.sCap;

        if sver0 = '' then
        begin
          ExcelApp.Cells[irow, 5].Value := aSOPLine.sVer;
        end
        else
        begin
          if sver0 <> aSOPLine.sVer then
          begin
            MergeCells(ExcelApp, irow1, 5, irow - 1, 5);
            ExcelApp.Cells[irow, 5].Value := aSOPLine.sVer;
            irow1 := irow;
          end;
        end;

        // 统计后面汇总行数据 /////////////////////////////////////////////////////
        sver := aSOPLine.sVer;
        idx := Pos('(', sver);
        if idx > 0 then
        begin
          sver := Copy(aSOPLine.sVer, 1, idx - 1);
        end;

        idx := slver.IndexOfName(sver);
        if idx >= 0 then
        begin
          slver.ValueFromIndex[idx] := slver.ValueFromIndex[idx] + ',' + IntToStr(irow);
        end
        else
        begin
          slver.Add(sver + '=' + IntToStr(irow));
        end;

        idx := slcap.IndexOfName(aSOPLine.sCap);
        if idx >= 0 then
        begin
          slcap.ValueFromIndex[idx] := slcap.ValueFromIndex[idx] + ',' + IntToStr(irow);
        end
        else
        begin
          slcap.Add(aSOPLine.sCap + '=' + IntToStr(irow));
        end;

        idx := slcolor.IndexOfName(aSOPLine.sColor);    
        if idx >= 0 then
        begin
          slcolor.ValueFromIndex[idx] := slcolor.ValueFromIndex[idx] + ',' + IntToStr(irow);
        end
        else
        begin
          slcolor.Add(aSOPLine.sColor + '=' + IntToStr(irow));
        end;
        
   
        dt0 := 0;     
        icol := 9;

        if iline = 0 then  // 第一个SKU编码，写列标题
        begin


          // 写标题列 /////////////////////////////////////////
          for idate := 0 to aSOPLine.FList.Count - 1 do
          begin
            aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);

            if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(dt0)) + '月';
              MergeCells(ExcelApp, 1, icol, 2, icol);
              icol := icol + 1;
            end;
               
            ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
            ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;   
            icol := icol + 1;

            // 最后一个日期
            if idate = aSOPLine.FList.Count - 1 then
            begin
              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(aSOPCol.dt1)) + '月';
              MergeCells(ExcelApp, 1, icol, 2, icol);

              icol2 := icol;
            end;

            dt0 := aSOPCol.dt1;
          end;
        end;
        
                         
        dt0 := 0;     
        icol := 9;
        icol1 := icol;

        for idate := 0 to aSOPLine.FList.Count - 1 do
        begin
          aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);   

          if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow);
            icol := icol + 1;
            icol1 := icol;
          end;
            
          ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty_ok;
          icol := icol + 1;

          // 最后一个日期
          if idate = aSOPLine.FList.Count - 1 then
          begin
            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow);
          end;
          
          dt0 := aSOPCol.dt1;

        end;

        if iline = aSOPProj.FList.Count - 1 then
        begin
          MergeCells(ExcelApp, irow1, 5, irow, 5);   
        end;

        sver0 := aSOPLine.sVer;
               
        irow := irow + 1;
      end;

      AddBorder(ExcelApp, 1, 1, irow - 1, icol2);
                  
      irow1 := irow + slver.Count + slcap.Count + slcolor.Count;
      
      ExcelApp.Cells[irow, 5].Value := aSOPProj.FName + ' Total';
      MergeCells(ExcelApp, irow, 5, irow1 - 1, 5);

      ExcelApp.Cells[irow1, 5].Value := aSOPProj.FName + ' TOTAL';
      MergeCells(ExcelApp, irow1, 5, irow1, 8);

      irow1 := irow;

      sl := TStringList.Create;

      for i := 0 to slver.Count - 1 do
      begin
        ExcelApp.Cells[irow, 6].Value := slver.Names[i];
        MergeCells(ExcelApp, irow, 6, irow, 8);

        sl.Text := StringReplace(slver.ValueFromIndex[i], ',', #13#10, [rfReplaceAll]);
        if sl.Count > 0 then
        begin
          for icol := 9 to icol2 do
          begin
            for idx := 0 to sl.Count - 1 do
            begin
              if idx = 0 then
              begin
                s := '=' + GetRef(icol) + sl[idx];
              end
              else
              begin
                s := s + '+' + GetRef(icol) + sl[idx];
              end;
            end;
            ExcelApp.Cells[irow, icol].Value := s;
          end;
        end;
        irow := irow + 1
      end;

      for i := 0 to slcap.Count - 1 do
      begin
        ExcelApp.Cells[irow, 6].Value := slcap.Names[i];
        MergeCells(ExcelApp, irow, 6, irow, 8);

        sl.Text := StringReplace(slcap.ValueFromIndex[i], ',', #13#10, [rfReplaceAll]);
        if sl.Count > 0 then
        begin
          for icol := 9 to icol2 do
          begin
            for idx := 0 to sl.Count - 1 do
            begin
              if idx = 0 then
              begin
                s := '=' + GetRef(icol) + sl[idx];
              end
              else
              begin
                s := s + '+' + GetRef(icol) + sl[idx];
              end;
            end;
            ExcelApp.Cells[irow, icol].Value := s;
          end;
        end;
        irow := irow + 1
      end;

      for i := 0 to slcolor.Count - 1 do
      begin
        ExcelApp.Cells[irow, 6].Value := slcolor.Names[i];
        MergeCells(ExcelApp, irow, 6, irow, 8);

        sl.Text := StringReplace(slcolor.ValueFromIndex[i], ',', #13#10, [rfReplaceAll]);
        if sl.Count > 0 then
        begin
          for icol := 9 to icol2 do
          begin
            for idx := 0 to sl.Count - 1 do
            begin
              if idx = 0 then
              begin
                s := '=' + GetRef(icol) + sl[idx];
              end
              else
              begin
                s := s + '+' + GetRef(icol) + sl[idx];
              end;
            end;
            ExcelApp.Cells[irow, icol].Value := s;
          end;
        end;
        irow := irow + 1
      end;

      for icol := 9 to icol2 do
      begin
        ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol) + IntToStr(3) + ':' + GetRef(icol) + IntToStr(irow1 - 1) + ')';
      end;

      AddBorder(ExcelApp, irow1, 5, irow, icol2);

      sl.Free;

    end;

    ExcelApp.Sheets[1].Activate;
      
    WorkBook.SaveAs(sfile_save);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  finally
    WorkBook.Close;
    ExcelApp.Quit;

    slver.Free;
    slcap.Free;
    slcolor.Free;
  end;
end;
// 
//procedure TfrmSopSimSum.SaveMPS(const sfile_save: string; aSOPReader: TSOPReader;
//  aSOPSimReader: TSOPSimReader; aMrpMPSReader: TMrpMPSReader);
//var
//  ExcelApp, WorkBook: Variant;
//  iproj: Integer;
//  aSOPProj: TSOPProj;
//  iline: Integer;
//  aSOPLine: TSOPLine;
//  idate: Integer;
//  aSOPCol: TSOPCol;
//  irow: Integer;
//  icol: Integer;
//  dt0: TDateTime;
//  icol1: Integer;  
//  icol2: Integer;
//
//  aSOPSimProj: TSOPSimProj;
//  aSOPSimLine: TSOPSimLine;
//begin
//        
//  try
//    ExcelApp := CreateOleObject('Excel.Application' );
//    ExcelApp.Visible := False;
//    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
//  except
//    on e: Exception do
//    begin
//      MessageBox(0, PChar(e.Message), '金蝶提示', 0);
//      Exit;
//    end;
//  end;
//
//  icol2 := 1;
//
//   
//  WorkBook := ExcelApp.WorkBooks.Add;
//
//  try
//    while ExcelApp.Sheets.Count < aSOPReader.FProjs.Count do
//    begin
//      ExcelApp.Sheets.Add;
//    end;
//
//    for iproj := 0 to aSOPReader.FProjs.Count - 1 do
//    begin
//      aSOPProj := TSOPProj(aSOPReader.FProjs.Objects[iproj]);
//
//      if aSOPSimReader <> nil then
//      begin
//        aSOPSimProj := aSOPSimReader.ProjByName[aSOPProj.FName];
//      end
//      else aSOPSimProj := nil;
//      
//      ExcelApp.Sheets[iproj + 1].Activate;
//      ExcelApp.Sheets[iproj + 1].Name := aSOPProj.FName;    
//
//      
//      irow := 1; 
//      ExcelApp.Cells[irow, 1].Value := '项目';
//      ExcelApp.Cells[irow, 2].Value := '物料编码';
//      ExcelApp.Cells[irow, 3].Value := '标准制式';
//      ExcelApp.Cells[irow, 4].Value := '颜色';
//      ExcelApp.Cells[irow, 5].Value := '容量';
//      ExcelApp.Cells[irow, 6].Value := '内容项';
//            
//
//      MergeCells(ExcelApp, irow, 1, irow + 1, 1);
//      MergeCells(ExcelApp, irow, 2, irow + 1, 2);
//      MergeCells(ExcelApp, irow, 3, irow + 1, 3);
//      MergeCells(ExcelApp, irow, 4, irow + 1, 4);
//      MergeCells(ExcelApp, irow, 5, irow + 1, 5);
//      MergeCells(ExcelApp, irow, 6, irow + 1, 6);
//      
//      ExcelApp.Columns[1].ColumnWidth := 8; 
//      ExcelApp.Columns[2].ColumnWidth := 13;
//      ExcelApp.Columns[3].ColumnWidth := 13;
//      ExcelApp.Columns[4].ColumnWidth := 8;
//      ExcelApp.Columns[5].ColumnWidth := 8;
//      ExcelApp.Columns[6].ColumnWidth := 23;
//
//      AddColor(ExcelApp, irow, 1, irow + 1, 6, $00CC99);
// 
//      irow := 3; 
//      for iline := 0 to aSOPProj.FList.Count - 1 do
//      begin
//        aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
//        if aSOPSimProj <> nil then
//        begin
//          aSOPSimLine := aSOPSimProj.GetLine(aSOPLine.sNumber);
//        end
//        else aSOPSimLine := nil;
//
//        ExcelApp.Cells[irow, 1].Value := aSOPLine.sProj;
//        ExcelApp.Cells[irow, 2].Value := aSOPLine.sNumber;
//        ExcelApp.Cells[irow, 3].Value := aSOPLine.sVer;
//        ExcelApp.Cells[irow, 4].Value := aSOPLine.sColor;
//        ExcelApp.Cells[irow, 5].Value := aSOPLine.sCap;
//        
//        ExcelApp.Cells[irow, 6].Value := '销售计划量';
//        ExcelApp.Cells[irow + 1, 6].Value := '可供应量';
//        ExcelApp.Cells[irow + 2, 6].Value := '销售计划与可供应量差异';
//        ExcelApp.Cells[irow + 3, 6].Value := '上周S&OP量';
//        ExcelApp.Cells[irow + 4, 6].Value := '上周S&OP量与可供应量差异';
//        ExcelApp.Cells[irow + 5, 6].Value := '上周MPS量';
//        ExcelApp.Cells[irow + 6, 6].Value := '上周MPS量与可供应量差异';
//
//                            
//        MergeCells(ExcelApp, irow, 1, irow + 6, 1);
//        MergeCells(ExcelApp, irow, 2, irow + 6, 2);
//        MergeCells(ExcelApp, irow, 3, irow + 6, 3);
//        MergeCells(ExcelApp, irow, 4, irow + 6, 4);
//        MergeCells(ExcelApp, irow, 5, irow + 6, 5); 
//
// 
//        dt0 := 0;     
//        icol := 7;
//
//        if iline = 0 then  // 第一个SKU编码，写列标题
//        begin
//
//
//          // 写标题列 /////////////////////////////////////////
//          for idate := 0 to aSOPLine.FList.Count - 1 do
//          begin
//            aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
//
//            if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
//            begin
//              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(dt0)) + '月';
//              MergeCells(ExcelApp, 1, icol, 2, icol);
//              AddColor(ExcelApp, 1, icol, 2, icol, $CCFFFF);
//              icol := icol + 1;
//            end;
//               
//            ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
//            ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;   
//            icol := icol + 1;
//
//            // 最后一个日期
//            if idate = aSOPLine.FList.Count - 1 then
//            begin
//              ExcelApp.Cells[1, icol].Value := IntToStr(MonthOf(aSOPCol.dt1)) + '月';
//              MergeCells(ExcelApp, 1, icol, 2, icol);        
//              AddColor(ExcelApp, 1, icol, 2, icol, $CCFFFF);
//
//              icol2 := icol;
//            end;
//
//            dt0 := aSOPCol.dt1;
//          end;
//        end;
//        
//                         
//        dt0 := 0;     
//        icol := 7;
//        icol1 := icol;
//
//        for idate := 0 to aSOPLine.FList.Count - 1 do
//        begin
//          aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);   
//
//          if (dt0 <> 0) and (MonthOf(dt0) <> MonthOf(aSOPCol.dt1)) then
//          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow); 
//            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1); 
//            ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//            ExcelApp.Cells[irow + 3, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 3) + ':' + GetRef(icol - 1) + IntToStr(irow + 3);
//            ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//            ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 5) + ':' + GetRef(icol - 1) + IntToStr(irow + 5);
//            ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
//            icol := icol + 1;
//            icol1 := icol;
//          end;
//            
//          ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty_ok;    
//          ExcelApp.Cells[irow + 1, icol].Value := aSOPCol.iQty;
//          ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//          if aSOPSimLine <> nil then
//          begin
//            ExcelApp.Cells[irow + 3, icol].Value := aSOPSimLine.GetQty(aSOPCol.dt1);
//          end;
//          ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//          if aMrpMPSReader <> nil then
//          begin
//            ExcelApp.Cells[irow + 5, icol].Value := aMrpMPSReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2);
//          end;
//          ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
//          icol := icol + 1;
//
//          // 最后一个日期
//          if idate = aSOPLine.FList.Count - 1 then
//          begin
//            ExcelApp.Cells[irow, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow) + ':' + GetRef(icol - 1) + IntToStr(irow);
//            ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 1) + ':' + GetRef(icol - 1) + IntToStr(irow + 1);
//            ExcelApp.Cells[irow + 2, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 1) + '-' + GetRef(icol) + IntToStr(irow);
//            ExcelApp.Cells[irow + 3, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 3) + ':' + GetRef(icol - 1) + IntToStr(irow + 3);
//            ExcelApp.Cells[irow + 4, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 1);
//            ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(icol1) + IntToStr(irow + 5) + ':' + GetRef(icol - 1) + IntToStr(irow + 5);
//            ExcelApp.Cells[irow + 6, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 5) + '-' + GetRef(icol) + IntToStr(irow + 1);
//          end;
//          
//          dt0 := aSOPCol.dt1;
//
//        end;
//
//        AddColor(ExcelApp, irow + 3, 6, irow + 4, icol2, $D6E4FC);     
//        AddColor(ExcelApp, irow + 5, 6, irow + 6, icol2, $99FFFF);
//
//        irow := irow + 7;
//      end;
//
//      AddBorder(ExcelApp, 1, 1, irow - 1, icol2);
//          
//      ExcelApp.Range[ ExcelApp.Cells[3, 7], ExcelApp.Cells[3, 7] ].Select;
//      ExcelApp.ActiveWindow.FreezePanes := True;
//
//    end;
//
//    ExcelApp.Sheets[1].Activate;
//      
//    WorkBook.SaveAs(sfile_save);
//    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
//
//  finally
//    WorkBook.Close;
//    ExcelApp.Quit; 
//  end;
//end;

procedure TfrmSopSimSum.btnSaveClick(Sender: TObject);
var
  aDOSPlanReader: TDOSPlanReader;
  aStockBalReader: TStockBalReader;
  aSEOutReader: TSEOutReader; 
  aDailyPlanVsActReader: TDailyPlanVsActReader;
  aSOPReader_sell: TSOPReader;
//  aSOPReader_demand: TSOPReader;    
  aSOPSimReader: TSOPSimReader;

  slProjYear: TStringList;

//  sfile: string;
  sfile_save: string;

  ExcelApp, WorkBook: Variant;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;

  aSOPSimProj: TSOPSimProj;
  aSOPSimLine: TSOPSimLine;

  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
  aDPVALine: TDPVALine;

  irow: Integer;
  icol: Integer;
  irow1_ver: Integer;
  irow1_proj: Integer;
  icol2: Integer;
  sver0: string;

  dSOPQtyPlan: Double;
  dSOPQtyAct: Double;
  
  slMonths: TStringList;
  smonth: string;
  idx_month: Integer;
  slmonth_cols: TStringList;

  slver: TStringList;
  slcolor: TStringList;
  slcap: TStringList;    
  slall: TStringList;

  slrow: TStringList;
  idx: Integer;
  ir: Integer;
  s: string;

  dqty: Double;

  days: Integer;
begin
  sfile_save := 'SOP计划汇总 ' + FormatDateTime('yyyyMMdd hhmmss', Now);
  if not ExcelSaveDialog(sfile_save) then Exit;

  Memo1.Lines.Add('open ' + leDOSPlan.Name);

  aDOSPlanReader := TDOSPlanReader.Create(leDOSPlan.Text);

  Memo1.Lines.Add('open ' + leStockBal.Name);
  aStockBalReader := TStockBalReader.Create(leStockBal.Text);

  Memo1.Lines.Add('open ' + leSEOut.Name);
  aSEOutReader := TSEOutReader.Create(leSEOut.Text);

  Memo1.Lines.Add('open ' + leDailyPlanVsAct.Name);
  aDailyPlanVsActReader := TDailyPlanVsActReader.Create(leDailyPlanVsAct.Text);

  slProjYear := TfrmProjYear.GetProjYears;

  Memo1.Lines.Add('open ' + leSellPlan.Name);
  aSOPReader_sell := TSOPReader.Create(slProjYear, leSellPlan.Text);

                                          
  Memo1.Lines.Add('open ' + leSopSim.Name); 
  aSOPSimReader := TSOPSimReader.Create(leSopSim.Text, slProjYear, Log);

  try


    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '金蝶提示', 0);
        Exit;
      end;
    end;


    WorkBook := ExcelApp.WorkBooks.Add;
 
    try
      while ExcelApp.Sheets.Count < aSOPReader_sell.FProjs.Count + 1 do
      begin
        ExcelApp.Sheets.Add;
      end;



      slMonths := TStringList.Create;

      slver := TStringList.Create;
      slcolor := TStringList.Create;
      slcap := TStringList.Create;
      slall := TStringList.Create;

      try
        ExcelApp.Sheets[1].Activate;
        ExcelApp.Sheets[1].Name := 'S&OP汇总';

        icol2 := 5;
        
        irow := 1;                                 
        ExcelApp.Cells[irow, 1].Value := 'OEM/ODM';
        ExcelApp.Cells[irow, 2].Value := '项目';
        ExcelApp.Cells[irow, 3].Value := '制式';
        ExcelApp.Cells[irow, 4].Value := '物料编码';
        ExcelApp.Cells[irow, 5].Value := '颜色';
        ExcelApp.Cells[irow, 6].Value := '容量';
        ExcelApp.Cells[irow, 7].Value := '计划项';
                                                     
        MergeCells(ExcelApp, irow, 1, irow + 1, 1);
        MergeCells(ExcelApp, irow, 2, irow + 1, 2);
        MergeCells(ExcelApp, irow, 3, irow + 1, 3);
        MergeCells(ExcelApp, irow, 4, irow + 1, 4);
        MergeCells(ExcelApp, irow, 5, irow + 1, 5);
        MergeCells(ExcelApp, irow, 6, irow + 1, 6);
        MergeCells(ExcelApp, irow, 7, irow + 1, 7);

        aSOPSimLine := nil;
        aDPVALine := nil;
        
        irow := 3;
        irow1_ver := irow;
        sver0 := '';

        // 所有项目写一个汇总的sheet //////////////////////////////////////////////////////////////////////////////
        for iproj := 0 to aSOPReader_sell.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader_sell.Projs[iproj];

          aSOPSimProj := aSOPSimReader.ProjByName[aSOPProj.FName];
          aDailyPlanVsAcsSheet := aDailyPlanVsActReader.SheetByName[aSOPProj.FName];

          for iline := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iline];
            if aSOPSimProj <> nil then
            begin
              aSOPSimLine := aSOPSimProj.GetLine(aSOPLine.sNumber);
            end;

            if aDailyPlanVsAcsSheet <> nil then
            begin
              aDPVALine := aDailyPlanVsAcsSheet.GetLine(aSOPLine.sNumber);
            end;

            ExcelApp.Cells[irow,     1].Value := '';
            ExcelApp.Cells[irow + 1, 1].Value := '';
            ExcelApp.Cells[irow + 2, 1].Value := '';
            ExcelApp.Cells[irow + 3, 1].Value := '';
            ExcelApp.Cells[irow + 4, 1].Value := '';
            ExcelApp.Cells[irow + 5, 1].Value := '';
            ExcelApp.Cells[irow + 6, 1].Value := '';
            ExcelApp.Cells[irow + 7, 1].Value := '';
            ExcelApp.Cells[irow + 8, 1].Value := '';
                      
            ExcelApp.Cells[irow,     2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 1, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 2, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 3, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 4, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 5, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 6, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 7, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 8, 2].Value := aSOPProj.FName;

          
            if (sver0 = '') or (sver0 <> aSOPLine.sVer) then
            begin                                       
              ExcelApp.Cells[irow, 3].Value := aSOPLine.sVer;
              if sver0 <> '' then
              begin
                MergeCells(ExcelApp, irow1_ver, 3, irow - 1, 3);
              end;
              sver0 := aSOPLine.sVer;
              irow1_ver := irow;
            end;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 5].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 6].Value := aSOPLine.sCap;

            MergeCells(ExcelApp, irow, 4, irow + 8, 4);
            MergeCells(ExcelApp, irow, 5, irow + 8, 5);
            MergeCells(ExcelApp, irow, 6, irow + 8, 6);

            ExcelApp.Cells[irow,     7].Value := '销售计划';
            ExcelApp.Cells[irow + 1, 7].Value := '实际出货';
            ExcelApp.Cells[irow + 2, 7].Value := 'DOS目标';
            ExcelApp.Cells[irow + 3, 7].Value := '期初库存';
            ExcelApp.Cells[irow + 4, 7].Value := '销售计划';
            ExcelApp.Cells[irow + 5, 7].Value := '供应能力';
            ExcelApp.Cells[irow + 6, 7].Value := 'S&OP供应计划';
            ExcelApp.Cells[irow + 7, 7].Value := 'S&OP实际产出';
            ExcelApp.Cells[irow + 8, 7].Value := '期末库存';

            for idate := 0 to aSOPLine.DateCount - 1 do
            begin
              aSOPCol := aSOPLine.Dates[idate];
              icol := idate + 8;

              days := (aSOPLine.DateCount - 1 - idate);
              if days > 4 then
              begin
                days := 4;
              end;

              if iline = 0 then
              begin
                ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
                ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;

                smonth := FormatDateTime('yyyy年MM月', aSOPCol.dt1);
                idx_month := slMonths.IndexOf(smonth);
                if idx_month < 0 then
                begin
                  slmonth_cols := TStringList.Create;
                  slMonths.AddObject(smonth, slmonth_cols);
                end
                else
                begin
                  slmonth_cols := TStringList(slMonths.Objects[idx_month]);
                end;
                slmonth_cols.Add(IntToStr(icol));
              end;
            
              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;    // '销售计划'
              dqty := aSEOutReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2);
              if dqty > 0 then
              begin
                ExcelApp.Cells[irow + 1, icol].Value := dqty; // '实际出货';
              end;

              dqty := aDOSPlanReader.GetDosPlan(aSOPProj.FName, aSOPLine.sNumber, aSOPCol.sDate);
              ExcelApp.Cells[irow + 2, icol].Value := dqty; //'DOS目标';
              if idate = 0 then
              begin
                ExcelApp.Cells[irow + 3, icol].Value := aStockBalReader.GetNumberBal(aSOPLine.sNumber);              // 期初库存
              end
              else
              begin
                ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(icol - 1) + IntToStr(irow + 8);
              end;

//              if idate = 0 then
//              begin
//                s := '=' + GetRef(icol) + IntToStr(irow) + '+SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + 4) + IntToStr(irow) + ')/28*' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 2);
//              end
//              else if idate = aSOPLine.DateCount - 1 then
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(0)/28-' + GetRef(icol) + IntToStr(irow + 2)
//                  + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) +',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,' + GetRef(icol) + IntToStr(irow + 3) +'*SUM(0)/28-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) +',-(' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')))';
//              end
//              else
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
//                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-' + GetRef(icol) + IntToStr(irow + 2) + '+IF(('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) + ',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,'
//                  + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) + ',-(' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')))';
//              end;

              if days = 0 then // 最后一个单元格，取差值，保证总量一致
              begin
                s := 'SUM(' + GetRef(8) + IntToStr(irow) + ':' + GetRef(icol) + IntToStr(irow) + ') - SUM(' + GetRef(8) + IntToStr(irow + 4) + ':' + GetRef(icol - 1) + IntToStr(irow + 4) + ')' + '-' + GetRef(8) + IntToStr(irow + 3);
                s := '=IF(' + s + '<0,0,' + s + ')';
              end
              else
              begin
                s := '=Round(IF(' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntToStr(irow + 3) + '<0,0,'
                  + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntTostr(irow + 3) + '), 0)';
              end;

              ExcelApp.Cells[irow + 4, icol].Value :=  s; //'S&OP要货计划';
              if aSOPSimLine <> nil then
              begin
                ExcelApp.Cells[irow + 5, icol].Value :=  aSOPSimLine.GetQtyAvail(aSOPCol.dt1); //'供应能力';
              end;
              if aDPVALine <> nil then
              begin
                aDPVALine.GetQty(aSOPCol.dt1, aSOPCol.dt2, dSOPQtyPlan, dSOPQtyAct);
                ExcelApp.Cells[irow + 6, icol].Value := '=MIN(' + GetRef(icol) + IntToStr(irow + 4) + ',' + GetRef(icol) + IntToStr(irow + 5) + ')'; //'S&OP供应计划';
                ExcelApp.Cells[irow + 7, icol].Value := dSOPQtyAct;  //'S&OP实际产出';
              end;
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '+' + GetRef(icol) + IntToStr(irow + 4) + '-' + GetRef(icol) + IntToStr(irow) // '期末库存';
            end;                                                                  // 实际库存                         //实际产出

            for idx_month := 0 to slMonths.Count - 1 do
            begin
              slmonth_cols := TStringList(slMonths.Objects[idx_month]);
              icol := aSOPLine.DateCount + idx_month + 8;
              ExcelApp.Cells[irow    , icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow    ) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow    ) + ')';
              ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 1) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 1) + ')';
              ExcelApp.Cells[irow + 2, icol].Value := '=AVERAGE(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 2) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 2) + ')';
              ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 3);  // 取月第一周的期初库存
              ExcelApp.Cells[irow + 4, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 4) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 4) + ')';
              ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 5) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 5) + ')';
              ExcelApp.Cells[irow + 6, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 6) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 6) + ')';
              ExcelApp.Cells[irow + 7, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 7) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 7) + ')';
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 8);  // 取月最后一周的期末库存
            end;

            if iline = 0 then
            begin 
              for idx_month := 0 to slMonths.Count - 1 do
              begin 
                icol := aSOPLine.DateCount + idx_month + 8;
                ExcelApp.Cells[1, icol].Value := slMonths[idx_month];
                MergeCells(ExcelApp, 1, icol, 2, icol);
              end;

              icol2 := aSOPLine.DateCount + slMonths.Count + 7;

            end;

          

            AddColor(ExcelApp, irow, 7, irow, icol2, $C4D9DD);
            AddColor(ExcelApp, irow, 7, irow + 1, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 2, 7, irow + 2, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 3, 7, irow + 3, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 4, 7, irow + 4, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 5, 7, irow + 5, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 6, 7, irow + 8, icol2, $DBDCF2);

            ExcelApp.Range[ExcelApp.Cells[irow + 3, 8], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 3, 8], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;

            ExcelApp.Range[ExcelApp.Cells[irow + 8, 8], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 8], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;


            idx := slver.IndexOfName(aSOPLine.sVer);
            if idx >= 0 then
            begin
              slver.ValueFromIndex[idx] := slver.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slver.Add(aSOPLine.sVer + '=' + IntToStr(irow));
            end;
                     

            idx := slcolor.IndexOfName(aSOPLine.sColor);
            if idx >= 0 then
            begin
              slcolor.ValueFromIndex[idx] := slcolor.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcolor.Add(aSOPLine.sColor + '=' + IntToStr(irow));
            end;

                
            idx := slcap.IndexOfName(aSOPLine.sCap);
            if idx >= 0 then
            begin
              slcap.ValueFromIndex[idx] := slcap.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcap.Add(aSOPLine.sCap + '=' + IntToStr(irow));
            end;

            slall.Add(IntToStr(irow));
            
            irow := irow + 9;
          end;
               
        end;
          
        if sver0 <> '' then
        begin
          MergeCells(ExcelApp, irow1_ver, 3, irow - 1, 3);
        end;

 

        AddColor(ExcelApp, 1, 1, 2, icol2 - slMonths.Count, $F1D9C5);  
        AddColor(ExcelApp, 1, icol2 - slMonths.Count + 1, 2, icol2, $D9E9FD);

        AddBorder(ExcelApp, 1, 1, irow - 1, icol2);

        ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Name := 'Arial';
        ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Size := 9;          

      finally
        slver.Free;
        slcolor.Free;
        slcap.Free;

        for idx_month := 0 to slMonths.Count - 1 do
        begin
          slmonth_cols := TStringList(slMonths.Objects[idx_month]);
          slmonth_cols.Free;
        end; 
        slMonths.Free;
      end;   


          
      Memo1.Lines.Add('每个项目写一个sheet ' + IntToStr(aSOPReader_sell.ProjCount));
  

      // 每个项目写一个sheet ///////////////////////////////////////////////////////////////////////////////////////
      for iproj := 0 to aSOPReader_sell.ProjCount - 1 do
      begin


        slMonths := TStringList.Create;

        slver := TStringList.Create;
        slcolor := TStringList.Create;
        slcap := TStringList.Create;
        slall := TStringList.Create;

        try
          aSOPProj := aSOPReader_sell.Projs[iproj];

           
          Memo1.Lines.Add('write project ' + aSOPProj.FName);


          aSOPSimProj := aSOPSimReader.ProjByName[aSOPProj.FName];               
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   000');
          
          aDailyPlanVsAcsSheet := aDailyPlanVsActReader.SheetByName[aSOPProj.FName];
                             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   aaa');
          
          ExcelApp.Sheets[iproj + 2].Activate; 
          ExcelApp.Sheets[iproj + 2].Name := aSOPProj.FName;
                        
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   bbb');

          icol2 := 5;
        
          irow := 1;
          ExcelApp.Cells[irow, 1].Value := '制式';
          ExcelApp.Cells[irow, 2].Value := '物料编码';
          ExcelApp.Cells[irow, 3].Value := '颜色';
          ExcelApp.Cells[irow, 4].Value := '容量';
          ExcelApp.Cells[irow, 5].Value := '计划项';
                             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   ccc');

          MergeCells(ExcelApp, irow, 1, irow + 1, 1); 
          MergeCells(ExcelApp, irow, 2, irow + 1, 2);
          MergeCells(ExcelApp, irow, 3, irow + 1, 3);
          MergeCells(ExcelApp, irow, 4, irow + 1, 4);
          MergeCells(ExcelApp, irow, 5, irow + 1, 5);

          aSOPSimLine := nil;
          aDPVALine := nil;
        
          irow := 3;
          irow1_ver := irow;
          sver0 := '';
          
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111');
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   ' + IntToStr(aSOPProj.LineCount));


          for iline := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iline];
            if aSOPSimProj <> nil then
            begin
              aSOPSimLine := aSOPSimProj.GetLine(aSOPLine.sNumber);
            end;

            if aDailyPlanVsAcsSheet <> nil then
            begin
              aDPVALine := aDailyPlanVsAcsSheet.GetLine(aSOPLine.sNumber);
            end;
          
            if (sver0 = '') or (sver0 <> aSOPLine.sVer) then
            begin                                       
              ExcelApp.Cells[irow, 1].Value := aSOPLine.sVer;
              if sver0 <> '' then
              begin
                MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
              end;
              sver0 := aSOPLine.sVer;
              irow1_ver := irow;
            end;
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sCap;

            MergeCells(ExcelApp, irow, 2, irow + 8, 2);  
            MergeCells(ExcelApp, irow, 3, irow + 8, 3);
            MergeCells(ExcelApp, irow, 4, irow + 8, 4);

            ExcelApp.Cells[irow, 5].Value := '销售计划';
            ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
            ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
            ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
            ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
            ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
            ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
            ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
            ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

            for idate := 0 to aSOPLine.DateCount - 1 do
            begin
              aSOPCol := aSOPLine.Dates[idate];
              icol := idate + 6;

              days := (aSOPLine.DateCount - 1 - idate);
              if days > 4 then
              begin
                days := 4;
              end;

              if iline = 0 then
              begin
                ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
                ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;

                smonth := FormatDateTime('yyyy年MM月', aSOPCol.dt1);
                idx_month := slMonths.IndexOf(smonth);
                if idx_month < 0 then
                begin
                  slmonth_cols := TStringList.Create;
                  slMonths.AddObject(smonth, slmonth_cols);
                end
                else
                begin
                  slmonth_cols := TStringList(slMonths.Objects[idx_month]);
                end;
                slmonth_cols.Add(IntToStr(icol));
              end;
            
              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;    // '销售计划'
              ExcelApp.Cells[irow + 1, icol].Value := aSEOutReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2); // '实际出货';
              ExcelApp.Cells[irow + 2, icol].Value := aDOSPlanReader.GetDosPlan(aSOPProj.FName, aSOPLine.sNumber, aSOPCol.sDate); //'DOS目标';
              if idate = 0 then
              begin
                ExcelApp.Cells[irow + 3, icol].Value := aStockBalReader.GetNumberBal(aSOPLine.sNumber);              // 期初库存
              end
              else
              begin
                ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(icol - 1) + IntToStr(irow + 8);
              end;

//              if idate = 0 then
//              begin
//                s := '=' + GetRef(icol) + IntToStr(irow) + '+SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + 4) + IntToStr(irow) + ')/28*' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 2);
//              end
//              else if idate = aSOPLine.DateCount - 1 then
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(0)/28-' + GetRef(icol) + IntToStr(irow + 2)
//                  + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) +',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,' + GetRef(icol) + IntToStr(irow + 3) +'*SUM(0)/28-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) +',-(' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')))';
//              end
//              else
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
//                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-' + GetRef(icol) + IntToStr(irow + 2) + '+IF(('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) + ',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,'
//                  + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) + ',-(' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')))';
//              end;

              if days = 0 then // 最后一个单元格，取差值，保证总量一致
              begin
                s := 'SUM(' + GetRef(6) + IntToStr(irow) + ':' + GetRef(icol) + IntToStr(irow) + ') - SUM(' + GetRef(6) + IntToStr(irow + 4) + ':' + GetRef(icol - 1) + IntToStr(irow + 4) + ')' + '-' + GetRef(6) + IntToStr(irow + 3);
                s := '=IF(' + s + '<0,0,' + s + ')';
              end
              else
              begin
                s := '=Round(IF(' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntToStr(irow + 3) + '<0,0,'
                  + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntTostr(irow + 3) + '), 0)';
              end;
                                     
              ExcelApp.Cells[irow + 4, icol].Value :=  s; //'S&OP要货计划';
              if aSOPSimLine <> nil then
              begin
                ExcelApp.Cells[irow + 5, icol].Value :=  aSOPSimLine.GetQty(aSOPCol.dt1); //'供应能力';
              end;
              if aDPVALine <> nil then
              begin
                aDPVALine.GetQty(aSOPCol.dt1, aSOPCol.dt2, dSOPQtyPlan, dSOPQtyAct); 
                ExcelApp.Cells[irow + 6, icol].Value := '=MIN(' + GetRef(icol) + IntToStr(irow + 4) + ',' + GetRef(icol) + IntToStr(irow + 5) + ')'; //'S&OP供应计划'; 
                ExcelApp.Cells[irow + 7, icol].Value := dSOPQtyAct;  //'S&OP实际产出';
              end;
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '+' + GetRef(icol) + IntToStr(irow + 4) + '-' + GetRef(icol) + IntToStr(irow) // '期末库存';
            end;                                                                  // 实际库存                         //实际产出

            for idx_month := 0 to slMonths.Count - 1 do
            begin
              slmonth_cols := TStringList(slMonths.Objects[idx_month]);
              icol := aSOPLine.DateCount + idx_month + 6;
              ExcelApp.Cells[irow    , icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow    ) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow    ) + ')';
              ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 1) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 1) + ')';
              ExcelApp.Cells[irow + 2, icol].Value := '=AVERAGE(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 2) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 2) + ')';
              ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 3);  // 取月第一周的期初库存
              ExcelApp.Cells[irow + 4, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 4) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 4) + ')';
              ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 5) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 5) + ')';
              ExcelApp.Cells[irow + 6, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 6) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 6) + ')';
              ExcelApp.Cells[irow + 7, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 7) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 7) + ')';
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 8);  // 取月最后一周的期末库存
            end;

            if iline = 0 then
            begin 
              for idx_month := 0 to slMonths.Count - 1 do
              begin 
                icol := aSOPLine.DateCount + idx_month + 6;
                ExcelApp.Cells[1, icol].Value := slMonths[idx_month];
                MergeCells(ExcelApp, 1, icol, 2, icol);
              end;

              icol2 := aSOPLine.DateCount + slMonths.Count + 5;

            end;

          

            AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
            AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

            ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;


            idx := slver.IndexOfName(aSOPLine.sVer);
            if idx >= 0 then
            begin
              slver.ValueFromIndex[idx] := slver.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slver.Add(aSOPLine.sVer + '=' + IntToStr(irow));
            end;
                     

            idx := slcolor.IndexOfName(aSOPLine.sColor);
            if idx >= 0 then
            begin
              slcolor.ValueFromIndex[idx] := slcolor.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcolor.Add(aSOPLine.sColor + '=' + IntToStr(irow));
            end;

                
            idx := slcap.IndexOfName(aSOPLine.sCap);
            if idx >= 0 then
            begin
              slcap.ValueFromIndex[idx] := slcap.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcap.Add(aSOPLine.sCap + '=' + IntToStr(irow));
            end;

            slall.Add(IntToStr(irow));
            
            irow := irow + 9;
          end;
          if sver0 <> '' then
          begin
            MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
          end;
                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  22222');



          irow1_proj := irow;

          ExcelApp.Cells[irow, 1].Value := aSOPProj.FName;

          for idx := 0 to slver.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slver.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';  
              ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';              


              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slver.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;
             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  33333');


          for idx := 0 to slcolor.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slcolor.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';  
              ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';              


              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slcolor.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;

                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  44444');

              

          for idx := 0 to slcap.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slcap.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';  
              ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slcap.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;

                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  55555');



          for icol := 6 to icol2 do
          begin
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir])) ;
            end;
            ExcelApp.Cells[irow, icol].Value := s;
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 1) ;
            end;
            ExcelApp.Cells[irow + 1, icol].Value := s;   
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 2) ;
            end;
            ExcelApp.Cells[irow + 2, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 3) ;
            end;
            ExcelApp.Cells[irow + 3, icol].Value := s;      
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 4) ;
            end;
            ExcelApp.Cells[irow + 4, icol].Value := s;   
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 5) ;
            end;
            ExcelApp.Cells[irow + 5, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 6) ;
            end;
            ExcelApp.Cells[irow + 6, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 7) ;
            end;
            ExcelApp.Cells[irow + 7, icol].Value := s;    
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 8) ;
            end;
            ExcelApp.Cells[irow + 8, icol].Value := s;
          end;
                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  66666');

          ExcelApp.Cells[irow, 5].Value := '销售计划';
          ExcelApp.Cells[irow + 1, 5].Value := '实际出货';   
          ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
          ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
          ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
          ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
          ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
          ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
          ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

          AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
          AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
          AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
          AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

          ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
          ExcelApp.Range[ExcelApp.Cells[irow + 3, 6], ExcelApp.Cells[irow + 3, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
          ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
          ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

          ExcelApp.Cells[irow, 2].Value := '总计';
          MergeCells(ExcelApp, irow, 2, irow + 8, 4);

          irow := irow + 9;

              
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  77777');




          MergeCells(ExcelApp, irow1_proj, 1, irow - 1, 1);
                   

          AddColor(ExcelApp, 1, 1, 2, icol2 - slMonths.Count, $F1D9C5);  
          AddColor(ExcelApp, 1, icol2 - slMonths.Count + 1, 2, icol2, $D9E9FD);

          AddBorder(ExcelApp, 1, 1, irow - 1, icol2);

          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Name := 'Arial';
          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Size := 9;          
             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  88888');

        finally
          slver.Free;
          slcolor.Free;
          slcap.Free;
                      
//          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  99999');

          for idx_month := 0 to slMonths.Count - 1 do
          begin
            slmonth_cols := TStringList(slMonths.Objects[idx_month]);
            slmonth_cols.Free;
          end; 
          slMonths.Free;
        end;
      end;    
      
          
      Memo1.Lines.Add('write complete ');

      ExcelApp.Sheets[1].Activate;
                                          
      Memo1.Lines.Add('save ');
      WorkBook.SaveAs(sfile_save);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      Memo1.Lines.Add('quit excel ');
      WorkBook.Close;
      ExcelApp.Quit; 
    end;


  

  finally

    
    aDOSPlanReader.Free;
    aStockBalReader.Free;
    aSEOutReader.Free;

    aDailyPlanVsActReader.Free;
    aSOPReader_sell.Free;
//    aSOPReader_demand.Free;

    aSOPSimReader.Free;
    
    slProjYear.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);

end;

procedure TfrmSopSimSum.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    leDOSPlan.Text := ini.ReadString(self.ClassName, leDOSPlan.Name, ''); 
    leStockBal.Text := ini.ReadString(self.ClassName, leStockBal.Name, '');
    leSEOut.Text := ini.ReadString(self.ClassName, leSEOut.Name, '');
    leDailyPlanVsAct.Text := ini.ReadString(self.ClassName, leDailyPlanVsAct.Name, '');
    leSellPlan.Text := ini.ReadString(self.ClassName, leSellPlan.Name, '');
    leDemand.Text := ini.ReadString(self.ClassName, leDemand.Name, '');
    leSopSim.Text := ini.ReadString(self.ClassName, leSopSim.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmSopSimSum.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leDOSPlan.Name, leDOSPlan.Text);
    ini.WriteString(self.ClassName, leStockBal.Name, leStockBal.Text);
    ini.WriteString(self.ClassName, leSEOut.Name, leSEOut.Text);
    ini.WriteString(self.ClassName, leDailyPlanVsAct.Name, leDailyPlanVsAct.Text);
    ini.WriteString(self.ClassName, leSellPlan.Name, leSellPlan.Text);
    ini.WriteString(self.ClassName, leDemand.Name, leDemand.Text);
    ini.WriteString(self.ClassName, leSopSim.Name, leSopSim.Text);
  finally
    ini.Free;
  end;
end;

procedure TfrmSopSimSum.btnSave_demandClick(Sender: TObject);
var
//  aSEOutReader: TSEOutReader;
//  aDailyPlanVsActReader: TDailyPlanVsActReader;
//  aSOPReader_demand: TSOPReader;
   
  aDOSPlanReader: TDOSPlanReader;                              
  aStockBalReader: TStockBalReader;                
  aSOPReader_sell: TSOPReader;
//  aSOPSimReader: TSOPSimReader;

  slProjYear: TStringList;

//  sfile: string;
  sfile_save: string;

  ExcelApp, WorkBook: Variant;
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;

//  aSOPSimProj: TSOPSimProj;
//  aSOPSimLine: TSOPSimLine;

//  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
//  aDPVALine: TDPVALine;

  irow: Integer;
  icol: Integer;
  irow1_ver: Integer;
  irow1_proj: Integer;
  icol2: Integer;
  sver0: string;

  dSOPQtyPlan: Double;
  dSOPQtyAct: Double;
  
  slMonths: TStringList;
  smonth: string;
  idx_month: Integer;
  slmonth_cols: TStringList;

  slver: TStringList;
  slcolor: TStringList;
  slcap: TStringList;    
  slall: TStringList;

  slrow: TStringList;
  idx: Integer;
  ir: Integer;
  s: string;

  dqty: Double;

  days: Integer;
  iMaxDateCount: Integer;
begin
  sfile_save := '要货计划 ' + FormatDateTime('yyyyMMdd hhmmss', Now);
  Memo1.Lines.Add('选择保存文件路径');
  if not ExcelSaveDialog(sfile_save) then Exit;

  Memo1.Lines.Add('open ' + leDOSPlan.Text);

  aDOSPlanReader := TDOSPlanReader.Create(leDOSPlan.Text);

  Memo1.Lines.Add('open ' + leStockBal.Text);
  aStockBalReader := TStockBalReader.Create(leStockBal.Text);

  slProjYear := TfrmProjYear.GetProjYears;

  Memo1.Lines.Add('open ' + leSellPlan.Text);
  aSOPReader_sell := TSOPReader.Create(slProjYear, leSellPlan.Text);
 
  try


    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := True;
      ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '金蝶提示', 0);
        Exit;
      end;
    end;


    WorkBook := ExcelApp.WorkBooks.Add;
 
    try
      while ExcelApp.Sheets.Count < aSOPReader_sell.FProjs.Count + 1 do
      begin
        ExcelApp.Sheets.Add;
      end;



      slMonths := TStringList.Create;

      slver := TStringList.Create;
      slcolor := TStringList.Create;
      slcap := TStringList.Create;
      slall := TStringList.Create;

      try
        ExcelApp.Sheets[1].Activate;
        ExcelApp.Sheets[1].Name := 'S&OP汇总';

        icol2 := 5;
        
        irow := 1;                                 
        ExcelApp.Cells[irow, 1].Value := 'OEM/ODM';
        ExcelApp.Cells[irow, 2].Value := '项目';
        ExcelApp.Cells[irow, 3].Value := '制式';
        ExcelApp.Cells[irow, 4].Value := '物料编码';
        ExcelApp.Cells[irow, 5].Value := '颜色';
        ExcelApp.Cells[irow, 6].Value := '容量';
        ExcelApp.Cells[irow, 7].Value := '计划项';
                                                     
        MergeCells(ExcelApp, irow, 1, irow + 1, 1);
        MergeCells(ExcelApp, irow, 2, irow + 1, 2);
        MergeCells(ExcelApp, irow, 3, irow + 1, 3);
        MergeCells(ExcelApp, irow, 4, irow + 1, 4);
        MergeCells(ExcelApp, irow, 5, irow + 1, 5);
        MergeCells(ExcelApp, irow, 6, irow + 1, 6);
        MergeCells(ExcelApp, irow, 7, irow + 1, 7);

//        aSOPSimLine := nil;
//        aDPVALine := nil;

        irow := 3;
        irow1_ver := irow;
        sver0 := '';


        iMaxDateCount := 0;
        for iproj := 0 to aSOPReader_sell.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader_sell.Projs[iproj];
          for iline := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iline];
            if iMaxDateCount < aSOPLine.DateCount then
            begin
              iMaxDateCount := aSOPLine.DateCount;
            end;
            Break;
          end;
          Break;
        end;
        
        // 所有项目写一个汇总的sheet //////////////////////////////////////////////////////////////////////////////
        for iproj := 0 to aSOPReader_sell.ProjCount - 1 do
        begin
          aSOPProj := aSOPReader_sell.Projs[iproj];

//          aSOPSimProj := aSOPSimReader.ProjByName[aSOPProj.FName];
//          aDailyPlanVsAcsSheet := aDailyPlanVsActReader.SheetByName[aSOPProj.FName];

          for iline := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iline];
//            if aSOPSimProj <> nil then
//            begin
//              aSOPSimLine := aSOPSimProj.GetLine(aSOPLine.sNumber);
//            end;

//            if aDailyPlanVsAcsSheet <> nil then
//            begin
//              aDPVALine := aDailyPlanVsAcsSheet.GetLine(aSOPLine.sNumber);
//            end;

            ExcelApp.Cells[irow,     1].Value := '';
            ExcelApp.Cells[irow + 1, 1].Value := '';
            ExcelApp.Cells[irow + 2, 1].Value := '';
            ExcelApp.Cells[irow + 3, 1].Value := '';
            ExcelApp.Cells[irow + 4, 1].Value := '';
            ExcelApp.Cells[irow + 5, 1].Value := '';
            ExcelApp.Cells[irow + 6, 1].Value := '';
            ExcelApp.Cells[irow + 7, 1].Value := '';
            ExcelApp.Cells[irow + 8, 1].Value := '';
                      
            ExcelApp.Cells[irow,     2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 1, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 2, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 3, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 4, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 5, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 6, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 7, 2].Value := aSOPProj.FName;
            ExcelApp.Cells[irow + 8, 2].Value := aSOPProj.FName;

          
            if (sver0 = '') or (sver0 <> aSOPLine.sVer) then
            begin                                       
              ExcelApp.Cells[irow, 3].Value := aSOPLine.sVer;
              if sver0 <> '' then
              begin
                MergeCells(ExcelApp, irow1_ver, 3, irow - 1, 3);
              end;
              sver0 := aSOPLine.sVer;
              irow1_ver := irow;
            end;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 5].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 6].Value := aSOPLine.sCap;

            MergeCells(ExcelApp, irow, 4, irow + 8, 4);
            MergeCells(ExcelApp, irow, 5, irow + 8, 5);
            MergeCells(ExcelApp, irow, 6, irow + 8, 6);

            ExcelApp.Cells[irow,     7].Value := '销售计划';
            ExcelApp.Cells[irow + 1, 7].Value := '实际出货';
            ExcelApp.Cells[irow + 2, 7].Value := 'DOS目标';
            ExcelApp.Cells[irow + 3, 7].Value := '期初库存';
            ExcelApp.Cells[irow + 4, 7].Value := 'S&OP要货计划';
            ExcelApp.Cells[irow + 5, 7].Value := '供应能力';
            ExcelApp.Cells[irow + 6, 7].Value := 'S&OP供应计划';
            ExcelApp.Cells[irow + 7, 7].Value := 'S&OP实际产出';
            ExcelApp.Cells[irow + 8, 7].Value := '期末库存';


            for idate := 0 to aSOPLine.DateCount - 1 do
            begin
              aSOPCol := aSOPLine.Dates[idate];
              icol := idate + 8;

              days := (aSOPLine.DateCount - 1 - idate);
              if days > 4 then
              begin
                days := 4;
              end;

              if iline = 0 then
              begin
                ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
                ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;

                smonth := FormatDateTime('yyyy年MM月', aSOPCol.dt1);
                idx_month := slMonths.IndexOf(smonth);
                if idx_month < 0 then
                begin
                  slmonth_cols := TStringList.Create;
                  slMonths.AddObject(smonth, slmonth_cols);
                end
                else
                begin
                  slmonth_cols := TStringList(slMonths.Objects[idx_month]);
                end;
                slmonth_cols.Add(IntToStr(icol));
              end;
            
              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;    // '销售计划'      来自营销提供
//              dqty := aSEOutReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2);
//              if dqty > 0 then
//              begin
//                ExcelApp.Cells[irow + 1, icol].Value := dqty; // '实际出货';  来自帐务提供数据
//              end;
                                        
              dqty := aDOSPlanReader.GetDosPlan(aSOPProj.FName, aSOPLine.sNumber, aSOPCol.sDate);
              ExcelApp.Cells[irow + 2, icol].Value := dqty; //'DOS目标';  来自帐务定义
              if idate = 0 then
              begin
                ExcelApp.Cells[irow + 3, icol].Value := aStockBalReader.GetNumberBal(aSOPLine.sNumber);      // 期初库存  等于上一期期末库存
              end
              else
              begin
                ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(icol - 1) + IntToStr(irow + 8);
              end;

              
//              if idate = 0 then
//              begin
//                s := '=' + GetRef(icol) + IntToStr(irow) + '+SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + 4) + IntToStr(irow) + ')/28*' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 2);
//              end
//              else if idate = aSOPLine.DateCount - 1 then
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(0)/28-' + GetRef(icol) + IntToStr(irow + 2)
//                  + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) +',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,' + GetRef(icol) + IntToStr(irow + 3) +'*SUM(0)/28-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) +',-(' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')))';
//              end
//              else
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
//                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-' + GetRef(icol) + IntToStr(irow + 2) + '+IF(('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) + ',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,'
//                  + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) + ',-(' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')))';
//              end;
              if days = 0 then // 最后一个单元格，取差值，保证总量一致
              begin
                s := 'SUM(' + GetRef(8) + IntToStr(irow) + ':' + GetRef(icol) + IntToStr(irow) + ') - SUM(' + GetRef(8) + IntToStr(irow + 4) + ':' + GetRef(icol - 1) + IntToStr(irow + 4) + ')' + '-' + GetRef(8) + IntToStr(irow + 3);
                s := '=IF(' + s + '<0,0,' + s + ')';
              end
              else
              begin
                s := '=Round(IF(' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntToStr(irow + 3) + '<0,0,'
                  + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntTostr(irow + 3) + '), 0)';
              end;

              ExcelApp.Cells[irow + 4, icol].Value :=  s; //'S&OP要货计划';
//              if aSOPSimLine <> nil then
//              begin
//                ExcelApp.Cells[irow + 5, icol].Value :=  aSOPSimLine.GetQtyAvail(aSOPCol.dt1); //'供应能力';
//              end;
//              if aDPVALine <> nil then
//              begin
//                aDPVALine.GetQty(aSOPCol.dt1, aSOPCol.dt2, dSOPQtyPlan, dSOPQtyAct);
//                ExcelApp.Cells[irow + 6, icol].Value := '=MIN(' + GetRef(icol) + IntToStr(irow + 4) + ',' + GetRef(icol) + IntToStr(irow + 5) + ')'; //'S&OP供应计划';
//                ExcelApp.Cells[irow + 7, icol].Value := dSOPQtyAct;  //'S&OP实际产出';
//              end;
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '+' + GetRef(icol) + IntToStr(irow + 4) + '-' + GetRef(icol) + IntToStr(irow) // '期末库存';
            end;                                                             

            for idx_month := 0 to slMonths.Count - 1 do
            begin
              slmonth_cols := TStringList(slMonths.Objects[idx_month]);
              icol := aSOPLine.DateCount + idx_month + 8;
              ExcelApp.Cells[irow    , icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow    ) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow    ) + ')';
              ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 1) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 1) + ')';
              ExcelApp.Cells[irow + 2, icol].Value := '=AVERAGE(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 2) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 2) + ')';
              ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 3);  // 取月第一周的期初库存
              ExcelApp.Cells[irow + 4, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 4) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 4) + ')';
              ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 5) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 5) + ')';
              ExcelApp.Cells[irow + 6, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 6) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 6) + ')';
              ExcelApp.Cells[irow + 7, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 7) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 7) + ')';
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 8);  // 取月最后一周的期末库存 
            end;

            if iline = 0 then
            begin 
              for idx_month := 0 to slMonths.Count - 1 do
              begin 
                icol := iMaxDateCount + idx_month + 8;
                ExcelApp.Cells[1, icol].Value := slMonths[idx_month];
                MergeCells(ExcelApp, 1, icol, 2, icol);
              end;

              icol2 := iMaxDateCount + slMonths.Count + 7;

            end;

          

            AddColor(ExcelApp, irow, 7, irow, icol2, $C4D9DD);
            AddColor(ExcelApp, irow, 7, irow + 1, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 2, 7, irow + 2, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 3, 7, irow + 3, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 4, 7, irow + 4, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 5, 7, irow + 5, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 6, 7, irow + 8, icol2, $DBDCF2);

            ExcelApp.Range[ExcelApp.Cells[irow + 2, 8], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 2, 8], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;

            ExcelApp.Range[ExcelApp.Cells[irow + 8, 8], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 8], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;


            idx := slver.IndexOfName(aSOPLine.sVer);
            if idx >= 0 then
            begin
              slver.ValueFromIndex[idx] := slver.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slver.Add(aSOPLine.sVer + '=' + IntToStr(irow));
            end;
                     

            idx := slcolor.IndexOfName(aSOPLine.sColor);
            if idx >= 0 then
            begin
              slcolor.ValueFromIndex[idx] := slcolor.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcolor.Add(aSOPLine.sColor + '=' + IntToStr(irow));
            end;

                
            idx := slcap.IndexOfName(aSOPLine.sCap);
            if idx >= 0 then
            begin
              slcap.ValueFromIndex[idx] := slcap.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcap.Add(aSOPLine.sCap + '=' + IntToStr(irow));
            end;

            slall.Add(IntToStr(irow));
            
            irow := irow + 9;
          end;
               
        end;
          
        if sver0 <> '' then
        begin
          MergeCells(ExcelApp, irow1_ver, 3, irow - 1, 3);
        end;

 

        AddColor(ExcelApp, 1, 1, 2, icol2 - slMonths.Count, $F1D9C5);  
        AddColor(ExcelApp, 1, icol2 - slMonths.Count + 1, 2, icol2, $D9E9FD);

        AddBorder(ExcelApp, 1, 1, irow - 1, icol2);

        ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Name := 'Arial';
        ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Size := 9;          

      finally
        slver.Free;
        slcolor.Free;
        slcap.Free;

        for idx_month := 0 to slMonths.Count - 1 do
        begin
          slmonth_cols := TStringList(slMonths.Objects[idx_month]);
          slmonth_cols.Free;
        end; 
        slMonths.Free;
      end;   


          
      Memo1.Lines.Add('每个项目写一个sheet ' + IntToStr(aSOPReader_sell.ProjCount));
  

      // 每个项目写一个sheet ///////////////////////////////////////////////////////////////////////////////////////
      for iproj := 0 to aSOPReader_sell.ProjCount - 1 do
      begin


        slMonths := TStringList.Create;

        slver := TStringList.Create;
        slcolor := TStringList.Create;
        slcap := TStringList.Create;
        slall := TStringList.Create;

        try
          aSOPProj := aSOPReader_sell.Projs[iproj];

           
          Memo1.Lines.Add('write project ' + aSOPProj.FName);


//          aSOPSimProj := aSOPSimReader.ProjByName[aSOPProj.FName];               
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   000');
          
//          aDailyPlanVsAcsSheet := aDailyPlanVsActReader.SheetByName[aSOPProj.FName];

          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   aaa');
          
          ExcelApp.Sheets[iproj + 2].Activate; 
          ExcelApp.Sheets[iproj + 2].Name := aSOPProj.FName;
                        
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   bbb');

          icol2 := 5;
        
          irow := 1;
          ExcelApp.Cells[irow, 1].Value := '制式';
          ExcelApp.Cells[irow, 2].Value := '物料编码';
          ExcelApp.Cells[irow, 3].Value := '颜色';
          ExcelApp.Cells[irow, 4].Value := '容量';
          ExcelApp.Cells[irow, 5].Value := '计划项';
                             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111   ccc');

          MergeCells(ExcelApp, irow, 1, irow + 1, 1); 
          MergeCells(ExcelApp, irow, 2, irow + 1, 2);
          MergeCells(ExcelApp, irow, 3, irow + 1, 3);
          MergeCells(ExcelApp, irow, 4, irow + 1, 4);
          MergeCells(ExcelApp, irow, 5, irow + 1, 5);

//          aSOPSimLine := nil;
//          aDPVALine := nil;

          irow := 3;
          irow1_ver := irow;
          sver0 := '';
          
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   11111');
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '   ' + IntToStr(aSOPProj.LineCount));


          for iline := 0 to aSOPProj.LineCount - 1 do
          begin
            aSOPLine := aSOPProj.Lines[iline];
//            if aSOPSimProj <> nil then
//            begin
//              aSOPSimLine := aSOPSimProj.GetLine(aSOPLine.sNumber);
//            end;

//            if aDailyPlanVsAcsSheet <> nil then
//            begin
//              aDPVALine := aDailyPlanVsAcsSheet.GetLine(aSOPLine.sNumber);
//            end;

            if (sver0 = '') or (sver0 <> aSOPLine.sVer) then
            begin                                       
              ExcelApp.Cells[irow, 1].Value := aSOPLine.sVer;
              if sver0 <> '' then
              begin
                MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
              end;
              sver0 := aSOPLine.sVer;
              irow1_ver := irow;
            end;
            ExcelApp.Cells[irow, 2].Value := aSOPLine.sNumber;
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sColor;
            ExcelApp.Cells[irow, 4].Value := aSOPLine.sCap;

            MergeCells(ExcelApp, irow, 2, irow + 8, 2);  
            MergeCells(ExcelApp, irow, 3, irow + 8, 3);
            MergeCells(ExcelApp, irow, 4, irow + 8, 4);

            ExcelApp.Cells[irow, 5].Value := '销售计划';
            ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
            ExcelApp.Cells[irow + 2, 5].Value := 'DOS目标';
            ExcelApp.Cells[irow + 3, 5].Value := '期初库存';
            ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
            ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
            ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
            ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
            ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

            for idate := 0 to aSOPLine.DateCount - 1 do
            begin
              aSOPCol := aSOPLine.Dates[idate];
              icol := idate + 6;

              days := (aSOPLine.DateCount - 1 - idate);
              if days > 4 then
              begin
                days := 4;
              end;

              if iline = 0 then
              begin
                ExcelApp.Cells[1, icol].Value := aSOPCol.sWeek;
                ExcelApp.Cells[2, icol].Value := aSOPCol.sDate;

                smonth := FormatDateTime('yyyy年MM月', aSOPCol.dt1);
                idx_month := slMonths.IndexOf(smonth);
                if idx_month < 0 then
                begin
                  slmonth_cols := TStringList.Create;
                  slMonths.AddObject(smonth, slmonth_cols);
                end
                else
                begin
                  slmonth_cols := TStringList(slMonths.Objects[idx_month]);
                end;
                slmonth_cols.Add(IntToStr(icol));
              end;
            
              ExcelApp.Cells[irow, icol].Value := aSOPCol.iQty;    // '销售计划'
//              ExcelApp.Cells[irow + 1, icol].Value := aSEOutReader.GetQty(aSOPLine.sNumber, aSOPCol.dt1, aSOPCol.dt2); // '实际出货';
              ExcelApp.Cells[irow + 2, icol].Value := aDOSPlanReader.GetDosPlan(aSOPProj.FName, aSOPLine.sNumber, aSOPCol.sDate); //'DOS目标';
              if idate = 0 then
              begin
                ExcelApp.Cells[irow + 3, icol].Value := aStockBalReader.GetNumberBal(aSOPLine.sNumber);              // 期初库存
              end
              else
              begin
                ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(icol - 1) + IntToStr(irow + 8);
              end;

//              if idate = 0 then
//              begin
//                s := '=' + GetRef(icol) + IntToStr(irow) + '+SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + days) + IntToStr(irow) + ')/28*' + GetRef(icol) + IntToStr(irow + 3) + '-' + GetRef(icol) + IntToStr(irow + 2);
//              end
//              else if idate = aSOPLine.DateCount - 1 then
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(0)/28-' + GetRef(icol) + IntToStr(irow + 2)
//                  + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) +',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,' + GetRef(icol) + IntToStr(irow + 3) +'*SUM(0)/28-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) +',-(' + GetRef(icol - 1) + IntToStr(irow + 7) +'-' + GetRef(icol - 1) + IntToStr(irow + 6) +')))';
//              end
//              else
//              begin
//                s := '=IF(' + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
//                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-' + GetRef(icol) + IntToStr(irow + 2) + '+IF(('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,' + GetRef(icol - 1) + IntToStr(irow + 6) + ',-('
//                  + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + '))<0,0,'
//                  + GetRef(icol) + IntToStr(irow + 3) + '*SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':' + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + '-'
//                  + GetRef(icol) + IntToStr(irow + 2) + '+IF((' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')=0,'
//                  + GetRef(icol - 1) + IntToStr(irow + 6) + ',-(' + GetRef(icol - 1) + IntToStr(irow + 7) + '-' + GetRef(icol - 1) + IntToStr(irow + 6) + ')))';
//              end;

              if days = 0 then // 最后一个单元格，取差值，保证总量一致
              begin
                s := 'SUM(' + GetRef(6) + IntToStr(irow) + ':' + GetRef(icol) + IntToStr(irow) + ') - SUM(' + GetRef(6) + IntToStr(irow + 4) + ':' + GetRef(icol - 1) + IntToStr(irow + 4) + ')' + '-' + GetRef(6) + IntToStr(irow + 3);
                s := '=IF(' + s + '<0,0,' + s + ')';
              end
              else
              begin
                s := '=Round(IF(' + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntToStr(irow + 3) + '<0,0,'
                  + GetRef(icol) + IntToStr(irow) + '+' + GetRef(icol) + IntToStr(irow + 2) + '*(SUM(' + GetRef(icol + 1) + IntToStr(irow) + ':'
                  + GetRef(icol + days) + IntToStr(irow) + ')/' + IntToStr(4 * 7) + ')-' + GetRef(icol) + IntTostr(irow + 3) + '), 0)';
              end;
                               
              ExcelApp.Cells[irow + 4, icol].Value :=  s; //'S&OP要货计划';
//              if aSOPSimLine <> nil then
//              begin
//                ExcelApp.Cells[irow + 5, icol].Value :=  aSOPSimLine.GetQty(aSOPCol.dt1); //'供应能力';
//              end;
//              if aDPVALine <> nil then
//              begin
//                aDPVALine.GetQty(aSOPCol.dt1, aSOPCol.dt2, dSOPQtyPlan, dSOPQtyAct); 
//                ExcelApp.Cells[irow + 6, icol].Value := '=MIN(' + GetRef(icol) + IntToStr(irow + 4) + ',' + GetRef(icol) + IntToStr(irow + 5) + ')'; //'S&OP供应计划'; 
//                ExcelApp.Cells[irow + 7, icol].Value := dSOPQtyAct;  //'S&OP实际产出';
//              end;
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(icol) + IntToStr(irow + 3) + '+' + GetRef(icol) + IntToStr(irow + 4) + '-' + GetRef(icol) + IntToStr(irow) // '期末库存';
            end;                                                                  // 实际库存                         //实际产出

            for idx_month := 0 to slMonths.Count - 1 do
            begin
              slmonth_cols := TStringList(slMonths.Objects[idx_month]);
              icol := aSOPLine.DateCount + idx_month + 6;
              ExcelApp.Cells[irow    , icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow    ) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow    ) + ')';
              ExcelApp.Cells[irow + 1, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 1) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 1) + ')';
              ExcelApp.Cells[irow + 2, icol].Value := '=AVERAGE(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 2) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 2) + ')';
              ExcelApp.Cells[irow + 3, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 3);  // 取月第一周的期初库存
              ExcelApp.Cells[irow + 4, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 4) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 4) + ')';
              ExcelApp.Cells[irow + 5, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 5) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 5) + ')';
              ExcelApp.Cells[irow + 6, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 6) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 6) + ')';
              ExcelApp.Cells[irow + 7, icol].Value := '=SUM(' + GetRef(StrToInt(slmonth_cols[0])) + IntToStr(irow + 7) + ':' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 7) + ')';
              ExcelApp.Cells[irow + 8, icol].Value := '=' + GetRef(StrToInt(slmonth_cols[slmonth_cols.Count - 1])) + IntToStr(irow + 8);  // 取月最后一周的期末库存
            end;

            if iline = 0 then
            begin 
              for idx_month := 0 to slMonths.Count - 1 do
              begin 
                icol := aSOPLine.DateCount + idx_month + 6;
                ExcelApp.Cells[1, icol].Value := slMonths[idx_month];
                MergeCells(ExcelApp, 1, icol, 2, icol);
              end;

              icol2 := aSOPLine.DateCount + slMonths.Count + 5;

            end;

          

            AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
            AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
            AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
            AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

            ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
            ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;


            idx := slver.IndexOfName(aSOPLine.sVer);
            if idx >= 0 then
            begin
              slver.ValueFromIndex[idx] := slver.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slver.Add(aSOPLine.sVer + '=' + IntToStr(irow));
            end;
                     

            idx := slcolor.IndexOfName(aSOPLine.sColor);
            if idx >= 0 then
            begin
              slcolor.ValueFromIndex[idx] := slcolor.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcolor.Add(aSOPLine.sColor + '=' + IntToStr(irow));
            end;

                
            idx := slcap.IndexOfName(aSOPLine.sCap);
            if idx >= 0 then
            begin
              slcap.ValueFromIndex[idx] := slcap.ValueFromIndex[idx] + ';' + IntToStr(irow);
            end
            else
            begin
              slcap.Add(aSOPLine.sCap + '=' + IntToStr(irow));
            end;

            slall.Add(IntToStr(irow));
            
            irow := irow + 9;
          end;
          if sver0 <> '' then
          begin
            MergeCells(ExcelApp, irow1_ver, 1, irow - 1, 1);
          end;
                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  22222');



          irow1_proj := irow;

          ExcelApp.Cells[irow, 1].Value := aSOPProj.FName;

          for idx := 0 to slver.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slver.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
              ExcelApp.Cells[irow + 2, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 3, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';              


              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slver.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;
             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  33333');


          for idx := 0 to slcolor.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slcolor.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
              ExcelApp.Cells[irow + 2, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 3, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';              


              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slcolor.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;

                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  44444');

              

          for idx := 0 to slcap.Count - 1 do
          begin
            slrow := TStringList.Create;
            try
              slrow.Text := StringReplace(slcap.ValueFromIndex[idx], ';', #13#10, [rfReplaceAll]);

              for icol := 6 to icol2 do
              begin
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir])) ; 
                end;
                ExcelApp.Cells[irow, icol].Value := s;
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 1) ; 
                end;
                ExcelApp.Cells[irow + 1, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 2) ;
                end;
                ExcelApp.Cells[irow + 2, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 3) ;
                end;
                ExcelApp.Cells[irow + 3, icol].Value := s;      
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 4) ;
                end;
                ExcelApp.Cells[irow + 4, icol].Value := s;   
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 5) ;
                end;
                ExcelApp.Cells[irow + 5, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 6) ;
                end;
                ExcelApp.Cells[irow + 6, icol].Value := s;  
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 7) ;
                end;
                ExcelApp.Cells[irow + 7, icol].Value := s;    
                
                s := '=0';
                for ir := 0 to slrow.Count - 1 do
                begin
                  s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slrow[ir]) + 8) ;
                end;
                ExcelApp.Cells[irow + 8, icol].Value := s;
              end;

              ExcelApp.Cells[irow, 5].Value := '销售计划';
              ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
              ExcelApp.Cells[irow + 2, 5].Value := '期初库存';
              ExcelApp.Cells[irow + 3, 5].Value := 'DOS目标';
              ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
              ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
              ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
              ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
              ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

              AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
              AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
              AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
              AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
              ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

              ExcelApp.Cells[irow, 2].Value := slcap.Names[idx];
              MergeCells(ExcelApp, irow, 2, irow + 8, 4);
            finally
              slrow.Free;
            end;
            irow := irow + 9;
          end;

                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  55555');



          for icol := 6 to icol2 do
          begin
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir])) ;
            end;
            ExcelApp.Cells[irow, icol].Value := s;
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 1) ;
            end;
            ExcelApp.Cells[irow + 1, icol].Value := s;   
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 2) ;
            end;
            ExcelApp.Cells[irow + 2, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 3) ;
            end;
            ExcelApp.Cells[irow + 3, icol].Value := s;      
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 4) ;
            end;
            ExcelApp.Cells[irow + 4, icol].Value := s;   
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 5) ;
            end;
            ExcelApp.Cells[irow + 5, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 6) ;
            end;
            ExcelApp.Cells[irow + 6, icol].Value := s;  
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 7) ;
            end;
            ExcelApp.Cells[irow + 7, icol].Value := s;    
                
            s := '=0';
            for ir := 0 to slall.Count - 1 do
            begin
              s := s + '+' + GetRef(icol) + IntToStr(StrToInt(slall[ir]) + 8) ;
            end;
            ExcelApp.Cells[irow + 8, icol].Value := s;
          end;
                
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  66666');

          ExcelApp.Cells[irow, 5].Value := '销售计划';
          ExcelApp.Cells[irow + 1, 5].Value := '实际出货';
          ExcelApp.Cells[irow + 2, 5].Value := '期初库存';
          ExcelApp.Cells[irow + 3, 5].Value := 'DOS目标';
          ExcelApp.Cells[irow + 4, 5].Value := 'S&OP要货计划';
          ExcelApp.Cells[irow + 5, 5].Value := '供应能力';
          ExcelApp.Cells[irow + 6, 5].Value := 'S&OP供应计划';
          ExcelApp.Cells[irow + 7, 5].Value := 'S&OP实际产出';
          ExcelApp.Cells[irow + 8, 5].Value := '期末库存';

          AddColor(ExcelApp, irow, 5, irow, icol2, $C4D9DD);
          AddColor(ExcelApp, irow, 5, irow + 1, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 2, 5, irow + 2, icol2, $DBDCF2);
          AddColor(ExcelApp, irow + 3, 5, irow + 3, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 4, 5, irow + 4, icol2, $DBDCF2);
          AddColor(ExcelApp, irow + 5, 5, irow + 5, icol2, $C4D9DD);
          AddColor(ExcelApp, irow + 6, 5, irow + 8, icol2, $DBDCF2);

          ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
          ExcelApp.Range[ExcelApp.Cells[irow + 2, 6], ExcelApp.Cells[irow + 2, icol2]].FormatConditions[1].Font.Color := $0000FF;
                                                            
          ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
          ExcelApp.Range[ExcelApp.Cells[irow + 8, 6], ExcelApp.Cells[irow + 8, icol2]].FormatConditions[1].Font.Color := $0000FF;
                            

          ExcelApp.Cells[irow, 2].Value := '总计';
          MergeCells(ExcelApp, irow, 2, irow + 8, 4);

          irow := irow + 9;

              
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  77777');




          MergeCells(ExcelApp, irow1_proj, 1, irow - 1, 1);
                   

          AddColor(ExcelApp, 1, 1, 2, icol2 - slMonths.Count, $F1D9C5);  
          AddColor(ExcelApp, 1, icol2 - slMonths.Count + 1, 2, icol2, $D9E9FD);

          AddBorder(ExcelApp, 1, 1, irow - 1, icol2);

          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Name := 'Arial';
          ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, icol2]].Font.Size := 9;          
             
          Memo1.Lines.Add('write project ' + aSOPProj.FName + '  88888');

        finally
          slver.Free;
          slcolor.Free;
          slcap.Free;
               
          for idx_month := 0 to slMonths.Count - 1 do
          begin
            slmonth_cols := TStringList(slMonths.Objects[idx_month]);
            slmonth_cols.Free;
          end; 
          slMonths.Free;
        end;
      end;    
      
          
      Memo1.Lines.Add('write complete ');

      ExcelApp.Sheets[1].Activate;
                                          
      Memo1.Lines.Add('save ');
      WorkBook.SaveAs(sfile_save);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      Memo1.Lines.Add('quit excel ');
      WorkBook.Close;
      ExcelApp.Quit; 
    end;


  

  finally

    
    aDOSPlanReader.Free;
    aStockBalReader.Free;
//    aSEOutReader.Free;

//    aDailyPlanVsActReader.Free;
    aSOPReader_sell.Free;

//    aSOPSimReader.Free;

    slProjYear.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);

end;

end.
