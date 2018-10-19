unit MergePlansWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, ImgList, ComObj, StdCtrls, ExtCtrls, DateUtils,
  CommUtils, IniFiles, FGStockReader, FGTableReader, SAPHW90Reader;
  
const
  xlCenter = -4108;

type
  TfrmMergePlans = class(TForm)
    ImageList1: TImageList;
    ToolBar1: TToolBar;
    ToolButton5: TToolButton;
    btnSave2: TToolButton;
    ProgressBar1: TProgressBar;
    ToolButton7: TToolButton;
    btnExit: TToolButton;
    leProj: TLabeledEdit;
    leSchedule: TLabeledEdit;
    btnSchedule: TButton;
    leMPS: TLabeledEdit;
    btnMPS: TButton;
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    leMarket: TLabeledEdit;
    btnMarket: TButton;
    Memo1: TMemo;
    leMPSYear: TLabeledEdit;
    leSOPYear: TLabeledEdit;
    leMarketYear: TLabeledEdit;
    leSUM: TLabeledEdit;
    btnSUM: TButton;
    leSUMYear: TLabeledEdit;
    dtpWeekStart: TDateTimePicker;
    Label1: TLabel;
    leFGTable: TLabeledEdit;
    btnFGTable: TButton;
    leStock: TLabeledEdit;
    btnStock: TButton;
    dtpWeekStartStock: TDateTimePicker;
    Label2: TLabel;
    procedure btnScheduleClick(Sender: TObject);
    procedure btnMPSClick(Sender: TObject);
    procedure btnSOPClick(Sender: TObject);
    procedure btnMarketClick(Sender: TObject);
    procedure btnSave2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnSUMClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnFGTableClick(Sender: TObject);
    procedure btnStockClick(Sender: TObject);
  private
    { Private declarations }
    FSchedules: TList;
    FMPSs: TList;
    FSOPs: TList;
    FMarkets: TList;
    FSUMs: TList;
    FLastWeek: string;
    FFirstDate: TDateTime;
    procedure Clear;
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
                                            
var
  frmMergePlans: TfrmMergePlans;
  
class procedure TfrmMergePlans.ShowForm;
begin
  frmMergePlans := TfrmMergePlans.Create(nil);
  frmMergePlans.ShowModal;
  frmMergePlans.Free;
end;

function ExcelOpenDialog(var sfile: string): Boolean;
begin
  with TOpenDialog.Create(nil) do
  try
    Filter := 'Excel Files|*.xls;*.xlsx';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options - [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := FileName;
    end;
  finally
    Free;
  end;
end;    

function ExcelOpenDialogs(var sfile: string): Boolean;
var
  i: Integer;
begin
  with TOpenDialog.Create(nil) do
  try
    Filter := 'Excel Files|*.xls;*.xlsx';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options + [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := Files[0];
      for i := 1 to Files.Count - 1 do
      begin
        sfile := sfile + ';' + Files[i];
      end;
    end;
  finally
    Free;
  end;
end;
     
function ExcelSaveDialog(var sfile: string): Boolean;
begin
  with TSaveDialog.Create(nil) do
  try
    FileName := sfile;
    Filter := 'Excel Files|*.xlsx;*.xls';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options - [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := FileName;
    end;
  finally
    Free;
  end;
end;
 
/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////

type
  TDateQty = packed record
    sweek: string;
    sdate: string;
    dt1, dt2: TDateTime;
    dqty: Double;
    dqtyact: Double;
  end;
  PDateQty = ^TDateQty;

  TLineRecord = class
  private
    FNumber:string;
    FColor: string;
    FCap: string;
    FVer: string;
    FWeek: string;
    FPlan: string;
    FDates: TList;
  public
    constructor Create(const snumber, scolor, scap, sver: string);
    destructor Destroy; override;
    procedure Clear;
    function Add(const sweek, sdate: string; dqty: Double; dt1, dt2: TDateTime): Integer; overload;
    function Add(const sdate: string; dqty, dqtyact: Double; dt1, dt2: TDateTime): Integer; overload;
  end;

{ TLineRecord }


constructor TLineRecord.Create(const snumber, scolor, scap, sver: string);
begin
  FNumber := snumber;
  FColor := scolor;
  FCap := scap;
  FVer := sver;
  FDates := TList.Create;
end;

destructor TLineRecord.Destroy;
begin
  Clear;
  FDates.Free;
  inherited;
end;

procedure TLineRecord.Clear;
var
  i: Integer;
  aDateQtyPtr: PDateQty;
begin
  for i := 0 to FDates.Count - 1 do
  begin
    aDateQtyPtr := PDateQty(FDates[i]);
    Dispose(aDateQtyPtr);
  end;
  FDates.Clear;
end;

function TLineRecord.Add(const sweek, sdate: string; dqty: Double; dt1, dt2: TDateTime): Integer;
var
  aDateQtyPtr: PDateQty;
begin
  aDateQtyPtr := New(PDateQty);
  aDateQtyPtr^.sweek := sweek;
  aDateQtyPtr^.sdate := sdate;
  aDateQtyPtr^.dqty := dqty;
  aDateQtyPtr^.dt1 := dt1;
  aDateQtyPtr^.dt2 := dt2;
  Result := FDates.Add(aDateQtyPtr);
end;
       
function TLineRecord.Add(const sdate: string; dqty, dqtyact: Double; dt1, dt2: TDateTime): Integer;
var
  aDateQtyPtr: PDateQty;
begin
  aDateQtyPtr := New(PDateQty);
  aDateQtyPtr^.sweek := '';
  aDateQtyPtr^.sdate := sdate;
  aDateQtyPtr^.dqty := dqty;
  aDateQtyPtr^.dqtyact := dqtyact;
  aDateQtyPtr^.dt1 := dt1;
  aDateQtyPtr^.dt2 := dt2;
  Result := FDates.Add(aDateQtyPtr);
end;

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
     
procedure TfrmMergePlans.btnScheduleClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialogs(sfile) then Exit;
  if leSchedule.Text <> '' then
    leSchedule.Text := leSchedule.Text + ';';
  leSchedule.Text := leSchedule.Text + sfile;
//  leProj.Text := ExtractFileName(ChangeFileExt(leSchedule.Text, ''));
end;

procedure TfrmMergePlans.btnMPSClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMPS.Text := sfile;
end;

procedure TfrmMergePlans.btnSOPClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSOP.Text := sfile;
end;

procedure TfrmMergePlans.btnMarketClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMarket.Text := sfile;
end;

procedure TfrmMergePlans.btnSUMClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSUM.Text := sfile;
end;



function OpenSchedule(const sfile: string; lst: TList): Boolean;
var
  ExcelApp, WorkBook: Variant;
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;
  stitle: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string; 
  bTitleOk: Boolean;
  irow, icol: Integer;
  sver, snumber, scolor, scap: string;
  sitem1, sitem2: string;
  sitem: string;
  irow1, irow2: Integer;
  icol1, icol2: Integer;
  slDate: TStringList;
  dt: TDateTime;
  dQty, dQtyAct: Double;
  sdate: string;
  aLineRecord: TLineRecord;
  dt1, dt2: TDateTime;
  v: Variant;
  s: string;
begin
  
  
  Result := False;
  if not FileExists(sfile) then
  begin
    MessageBox(frmMergePlans.Handle, '文件不存在', '金蝶提示', 0);
    Exit;
  end;
  
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';

  slDate := TStringList.Create;
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;
        
        sSheet := ExcelApp.Sheets[iSheet].Name;
        frmMergePlans.Memo1.Lines.Add(sSheet);

        if UpperCase(Copy(sSheet, 1, 4)) <> 'CTB-' then
        begin
          frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  名称 格式不符合');
          Continue;
        end;

        bTitleOk := True;
        
        irow := 1;

        stitle1 := Trim(ExcelApp.Cells[irow, 1].Value);
        stitle2 := Trim(ExcelApp.Cells[irow, 2].Value);
        stitle3 := Trim(ExcelApp.Cells[irow, 3].Value);
        stitle4 := Trim(ExcelApp.Cells[irow, 4].Value);
        stitle5 := Trim(ExcelApp.Cells[irow, 5].Value);
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;

        while stitle <> '机器型号物料编码颜色容量项目' do
        begin
          irow := irow + 1;

          if irow > 10 then
          begin
            bTitleOk := False;
            frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  内容 格式不符合 日排产计划 前几列标题必须是: 机器型号 物料编码 颜色 容量 项目');
            Break;
          end;

          stitle1 := Trim(ExcelApp.Cells[irow, 1].Value);
          stitle2 := Trim(ExcelApp.Cells[irow, 2].Value);
          stitle3 := Trim(ExcelApp.Cells[irow, 3].Value);
          stitle4 := Trim(ExcelApp.Cells[irow, 4].Value);
          stitle5 := Trim(ExcelApp.Cells[irow, 5].Value);
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        end;

        // 列标题不符合，停止循环
        if not bTitleOk then
        begin
          Break;
        end;
            
        // 读取日期列范围
        icol1 := 7;
        icol2 := icol1;
        while not IsCellMerged(ExcelApp, irow, icol2, irow + 1, icol2) do
        begin
          s := ExcelApp.Cells[irow + 1, icol2].Value;
          if s = '' then Break;

          v := ExcelApp.Cells[irow + 1, icol2].Value;
          if not VarIsType(v, varDate) then Break;
          
          dt := ExcelApp.Cells[irow + 1, icol2].Value;
          slDate.Add(FormatDatetime('yyyy-MM-dd', dt));
          icol2 := icol2 + 1;
        end;
        icol2 := icol2 - 1;
        frmMergePlans.Memo1.Lines.Add('日期列范围： icol1: ' + IntToStr(icol1) + '  icol2: ' + IntToStr(icol2));

        irow := irow + 2;
                                 
        sitem1 := ExcelApp.Cells[irow, 5].Value;
        sitem2 := ExcelApp.Cells[irow + 1, 5].Value;
        sitem := sitem1 + sitem2;
        while sitem = '计划实际' do
        begin

          if IsCellMerged(ExcelApp, irow, 1, irow + 1, 4) then
          begin
            irow := irow + 2;
            sitem1 := ExcelApp.Cells[irow, 5].Value;
            sitem2 := ExcelApp.Cells[irow + 1, 5].Value;
            sitem := sitem1 + sitem2;

            sver := ExcelApp.Cells[irow, 1].Value;
            frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  ' + sver + ' 跳过');
            Continue;
          end;

          // 一个制式的，例如   全网通公开版
          irow1 := irow;
          irow2 := irow + 2;
          while IsCellMerged(ExcelApp, irow1, 1, irow2 + 1, 1) do
          begin
            irow2 := irow2 + 2;
          end;
          irow2 := irow2 - 2;
                                
          sver := ExcelApp.Cells[irow, 1].Value;
          while irow <= irow2 do
          begin
            snumber := ExcelApp.Cells[irow, 2].Value;
            scolor := ExcelApp.Cells[irow, 3].Value;
            scap := ExcelApp.Cells[irow, 4].Value;

            if ( (Pos('海外', sver) > 0) and (scolor <> '') and (scap <> '') and (sver <> '') )
              or ( (snumber <> '') and (scolor <> '') and (scap <> '') and (sver <> '') ) then
            begin
              aLineRecord := TLineRecord.Create(snumber, scolor, scap, sver);
              lst.Add(aLineRecord);

              try
                for icol := icol1 to icol2 do
                begin
                  dQty := 0;
                  dQtyAct := 0;
                  v := ExcelApp.Cells[irow, icol].Value;
                  if VarIsNumeric(v) then dQty := v;
                  v := ExcelApp.Cells[irow + 1, icol].Value;
                  if VarIsNumeric(v) then dQtyAct := v;
                  sdate := slDate[icol - icol1];
                  dt1 := myStrToDateTime(sdate);
                  dt2 := dt1;
                  aLineRecord.Add(sdate, dQty, dQtyAct, dt1, dt2);
                end;
              except
                on e: Exception do
                begin
                  //frmMergePlans.Memo1.Lines.Add('irow ' + IntToStr(irow) + ', icol ' + IntToStr(icol));
                  raise e;
                end;
              end;
            end;
 
            irow := irow + 2;
 
            sitem1 := ExcelApp.Cells[irow, 5].Value;
            sitem2 := ExcelApp.Cells[irow + 1, 5].Value;
            sitem := sitem1 + sitem2;

          end;      

        end;

        frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  读取完成了  ' + sitem);
        
      end;


    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    slDate.Free;
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;


  Result := True;
end;
   
procedure InsertTo(lstDate: TList; toInsert: PDateQty);
var
  i: Integer;
  p1: PDateQty;
  newptr: PDateQty;
begin
  if lstDate.Count = 0 then
  begin
    newptr := New(PDateQty);
    newptr^ := toInsert^;
    newptr^.dqty := 0;
    lstDate.Add(newptr);
  end
  else if lstDate.Count = 1 then
  begin
    p1 := PDateQty(lstDate[0]);
    if p1^.dt1 < toInsert^.dt1 then
    begin
      newptr := New(PDateQty);
      newptr^ := toInsert^;   
      newptr^.dqty := 0;
      lstDate.Add(newptr);
    end
    else if p1^.dt1 > toInsert^.dt1 then
    begin
      newptr := New(PDateQty);
      newptr^ := toInsert^;       
      newptr^.dqty := 0;
      lstDate.Insert(0, newptr);
    end;
  end
  else
  begin
    for i := 0 to lstDate.Count - 1 do
    begin
      p1 := PDateQty(lstDate[i]);
      if toInsert^.dt1 < p1^.dt1 then // 小于，插入
      begin
        newptr := New(PDateQty);
        newptr^ := toInsert^;       
        newptr^.dqty := 0;
        lstDate.Insert(i, newptr);
        Exit;
      end
      else if toInsert^.dt1 = p1^.dt1 then
      begin
        Exit;  // 如果出现相等，退出，不插入
      end;
    end;
    // 遍历完了， 都小于插入日期， 插入日期添加在最后
    newptr := New(PDateQty);
    newptr^ := toInsert^;     
    newptr^.dqty := 0;
    lstDate.Add(newptr);    
  end;
end;
 
procedure MakeDateSet(lstDate: TList; lst: TList);
var
  i: Integer;
  aLineRecord: TLineRecord;
begin
  if lst.Count = 0 then Exit;

  aLineRecord := TLineRecord(lst[0]);
  for i := 0 to aLineRecord.FDates.Count - 1 do
  begin
    InsertTo(lstDate, aLineRecord.FDates[i]);
  end;
end;
         
procedure DoFillDataToList(lst: TList; toInsert: PDateQty; idx: Integer);     
var                     
  i: Integer;       
  newptr: PDateQty;
  aLineRecord: TLineRecord;
begin
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    newptr := New(PDateQty);
    newptr^ := toInsert^;
    if idx < 0 then
    begin
      aLineRecord.FDates.Add(newPtr);  
    end
    else
    begin
      aLineRecord.FDates.Insert(idx, newPtr);  
    end;
  end;
end;

procedure FillDataToList(lstDate: TList; lst: TList; toInsert: PDateQty);
var                     
  i: Integer;    
  p: PDateQty; 
  aLineRecord: TLineRecord;
begin
  if lst.Count = 0 then Exit;
  
  aLineRecord := TLineRecord(lst[0]);
  for i := aLineRecord.FDates.Count - 1 downto 0 do
  begin
    p := PDateQty(aLineRecord.FDates[i]);

    if i = aLineRecord.FDates.Count - 1 then  // 最后一个week， 日期最大
    begin
      // 大于最大的日期，添加在后面
      if toInsert^.dt1 > p^.dt1 then
      begin
        DoFillDataToList(lst, toInsert, -1);
        Break;  // 插入了，处理下一个week
      end
      else if toInsert^.dt1 = p^.dt1 then
      begin
        // 等于， week已存在，处理下一个week
        Break;
      end
      else
      begin
        // else 小于， 继续循环
      end;
    end
    else if i = 0 then         // 第一个week， 日期最小
    begin
      // 大于最大的日期，添加在后面
      if toInsert^.dt1 > p^.dt1 then
      begin
        DoFillDataToList(lst, toInsert, i + 1); // 如果只有1个元素，不会执行到这里，所以 j + 1 一定有效 
        Break;  // 插入了，处理下一个week
      end
      else if toInsert^.dt1 = p^.dt1 then
      begin
        // 等于， week已存在，处理下一个week
        Break;
      end
      else
      begin
        // else 小于， 继续循环
        DoFillDataToList(lst, toInsert, i);  // 如果只有1个元素，不会执行到这里，所以 j + 1 一定有效 
        Break;  // 插入了，处理下一个week    其实这里不用break了， 已经是循环到最后一个了
      end;
    end
    else
    begin
      // 大于最大的日期，添加在后面
      if toInsert^.dt1 > p^.dt1 then
      begin
        DoFillDataToList(lst, toInsert, i + 1);  // 如果只有1个元素，不会执行到这里，所以 j + 1 一定有效
        Break;  // 插入了，处理下一个week
      end
      else if toInsert^.dt1 = p^.dt1 then
      begin
        // 等于， week已存在，处理下一个week
        Break;
      end
      else
      begin
          
      end;
    end;
 
  end;
end;

function OpenMPS(const sfile, sProj, syearIn: string; plst: TList): Boolean;
var
  ExcelApp, WorkBook: Variant;
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;
  stitlex4, stitlex5, stitlex8, stitlex9: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8, stitle9: string;
  stitle: string;
  irow, icol: Integer; 
  sver, snumber, scolor, scap: string;
  sitem: string;
  irow1, irow2: Integer; 
  slDate: TStringList; 
  dQty: Double;
  sdate: string;
  sweek: string;
  sdate1, sdate2: string;
  dt0: TDateTime;
  dt1, dt2: TDateTime;
  aLineRecord: TLineRecord;
  i: Integer;
  iList: Integer;
  s: string;
  idx: Integer;
  syear: string;
  icolVer, icolNumber, icolColor, icolCap: Integer;
  lsts: TList;
  lst: TList;
  lstDate: TList;
  newptr: PDateQty;
  toInsert: PDateQty;
  v: Variant;
begin

  
  Result := False;
  if not FileExists(sfile) then
  begin
    MessageBox(frmMergePlans.Handle, '文件不存在', '金蝶提示', 0);
    Exit;
  end;
  
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';

  lsts := TList.Create;
  
  slDate := TStringList.Create;
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;
        
        sSheet := ExcelApp.Sheets[iSheet].Name;
        sSheet := Trim(sSheet);
        frmMergePlans.Memo1.Lines.Add(sSheet);

        // 判断是当前要读取的项目   
        s := StringReplace(sSheet, '-', ' ', []);
        idx := Pos(' ', s);
        if idx > 0 then      // 没有空格
        begin
          s := Copy(s, 1, idx - 1);
        end;

        if sProj <> s then
        begin
          frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + ' is not  ' + sProj + ' , skip');
          Continue;
        end;


        try
          stitle1 := Trim(ExcelApp.Cells[1, 1].Value);
          stitle2 := Trim(ExcelApp.Cells[1, 2].Value);
          stitle3 := Trim(ExcelApp.Cells[1, 3].Value);
          stitle4 := Trim(ExcelApp.Cells[1, 4].Value);    
          stitle5 := Trim(ExcelApp.Cells[1, 5].Value);
          stitle6 := Trim(ExcelApp.Cells[1, 6].Value);
          stitle7 := Trim(ExcelApp.Cells[1, 7].Value);
          stitle8 := Trim(ExcelApp.Cells[1, 8].Value);
          stitle9 := Trim(ExcelApp.Cells[1, 9].Value);
          stitlex4 := stitle1 + stitle2 + stitle3 + stitle4;  
          stitlex5 := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
          stitlex8 := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8;
          stitlex9 := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8 + stitle9;
           


          if stitlex4 = '制式物料编码颜色容量' then
          begin
            icol := 5;
            icolVer := 1;
            icolNumber := 2;
            icolColor := 3;
            icolCap := 4;
          end
          else if stitlex8 = '项目整机/裸机包装标准制式制式物料编码颜色容量' then
          begin
            icol := 9;
            icolVer := 5;
            icolNumber := 6;
            icolColor := 7;
            icolCap := 8;
          end
          else if (stitlex5 = 'MRP范围制式物料编码颜色容量')
            or (stitlex5 = 'MRP区域制式物料编码颜色容量') then
          begin
            icol := 6;
            icolVer := 2;
            icolNumber := 3;
            icolColor := 4;
            icolCap := 5;
          end
          else if ( stitlex9 = '项目整机/裸机包装标准制式MRP区域制式物料编码颜色容量')
            or (stitlex9 = '项目整机/裸机包装标准制式MRP范围制式物料编码颜色容量') then
          begin
            icol := 10;
            icolVer := 6;
            icolNumber := 7;
            icolColor := 8;
            icolCap := 9;
          end
          else
          begin
            frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  内容 格式不符合 MPS计划 前几列标题必须是: 制式 物料编码 颜色 容量');
            Continue;
          end;

          slDate.Clear;
          lst := TList.Create;
          lsts.Add(lst);

          stitle1 := Trim(ExcelApp.Cells[1, icol].Value);
          stitle2 := Trim(ExcelApp.Cells[2, icol].Value);
          stitle := stitle1 + stitle2;
          while stitle <> '' do
          begin
            if not IsCellMerged(ExcelApp, 1, icol, 2, icol) then
            begin                              
              sweek := ExcelApp.Cells[1, icol].Value;
              sdate := ExcelApp.Cells[2, icol].Value;
              slDate.AddObject(sdate + '=' + sweek, TObject(icol));
            end;
            icol := icol + 1;    
            stitle1 := Trim(ExcelApp.Cells[1, icol].Value);
            stitle2 := Trim(ExcelApp.Cells[2, icol].Value);
            stitle := stitle1 + stitle2;
          end;

          irow := 3;
          while not IsCellMerged(ExcelApp, irow, icolNumber, irow, icolColor) do
          begin
            irow1 := irow;
            irow2 := irow1 + 1;
            while IsCellMerged(ExcelApp, irow1, icolVer, irow2, icolVer) do
            begin
              irow2 := irow2 + 1;
            end;
            irow2 := irow2 - 1;

            sver := ExcelApp.Cells[irow, icolVer].Value;
            for irow := irow1 to irow2 do
            begin
              snumber := ExcelApp.Cells[irow, icolNumber].Value;
              scolor := ExcelApp.Cells[irow, icolColor].Value;
              scap := ExcelApp.Cells[irow, icolCap].Value;

              if Trim(snumber) = '' then
              begin
                snumber:= sver + scolor + scap;
              end;

              aLineRecord := TLineRecord.Create(snumber, scolor, scap, sver);
              lst.Add(aLineRecord);

              syear := syearIn;  
              dt0 := 0;

              for i := 0 to slDate.Count - 1 do
              begin
                icol := Integer(slDate.Objects[i]);
                v := ExcelApp.Cells[irow, icol].Value;
                if not VarIsNumeric(v) then
                begin  
                  dQty := 0;       
                  s := v;
                  if Trim(s) <> '' then
                  begin
                    MessageBox(0, PChar(sfile + #13#10 + sSheet + ' 单元格' + IntToStr(irow) + GetRef(icol) + '不是有效数字'), '提示', 0);
                    raise Exception.Create(sfile + #13#10 + sSheet + ' 单元格' + IntToStr(irow) + GetRef(icol) + '不是有效数字');
                  end;
                end
                else
                begin
                  dQty := v;
                end;
                sdate := slDate.Names[i];
                sweek := slDate.ValueFromIndex[i];
                if Pos('-', sdate) > 0 then
                begin
                  idx := Pos('-', sdate);

                  sdate1 := Copy(sdate, 1, idx - 1);
                  sdate2 := Copy(sdate, idx + 1, Length(sdate) - Pos('-', sdate));

                  sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                  sdate2 := syear + '-' + StringReplace(sdate2, '/', '-', [rfReplaceAll]);

                  dt1 := myStrToDateTime(sdate1);
                  dt2 := myStrToDateTime(sdate2);

                  if dt1 < dt0 then
                  begin
                    syear := IntToStr( StrToInt(syear) + 1 );
                    sdate1 := Copy(sdate, 1, idx - 1);
                    sdate2 := Copy(sdate, idx + 1, Length(sdate) - Pos('-', sdate));

                    sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                    sdate2 := syear + '-' + StringReplace(sdate2, '/', '-', [rfReplaceAll]);

                    dt1 := myStrToDateTime(sdate1);
                    dt2 := myStrToDateTime(sdate2);
                  end;

                  dt0 := dt1;
                end
                else
                begin
                  sdate1 := sdate; 
                  sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                  dt1 := myStrToDateTime(sdate1);
                  if dt1 < dt0 then
                  begin
                    syear := IntToStr( StrToInt(syear) + 1 );
                    sdate1 := sdate;
                    sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                    dt1 := myStrToDateTime(sdate1); 
                  end;
                  dt2 := dt1;

                  dt0 := dt1;
                end;

                aLineRecord.Add(sweek, sdate, dQty, dt1, dt2)
              end;

            end;

            irow := irow2 + 1;
          end;
        except
          on e: Exception do
          begin
            frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  读取错误 irow: ' + IntToStr(irow) + '   icol: ' + IntToStr(icol) );
            raise e;
          end;
        end;

        frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  读取完成了  ' + sitem);
        
      end;


    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

    lstDate := TList.Create;
    try
      for i := 0 to lsts.Count - 1 do
      begin
        lst := TList(lsts[i]);
        MakeDateSet(lstDate, lst);
      end;
      

      for iList := 0 to lsts.Count - 1 do
      begin
        lst := TList(lsts[iList]);
        for i := 0 to lstDate.Count - 1 do
        begin
          toInsert := PDateQty(lstDate[i]);
          FillDataToList(lstDate, lst, toInsert);
        end;
      end;

      for iList := 0 to lsts.Count - 1 do
      begin
        lst := TList(lsts[iList]);   
        for i := 0 to lst.Count - 1 do
        begin
          plst.Add(lst[i]);
        end; 
      end;


      for i := 0 to lsts.Count - 1 do
      begin
        lst := TList(lsts[i]);
        lst.Free;
      end;
    finally
      for i := 0 to lstDate.Count - 1 do
      begin
        newptr := PDateQty(lstDate[i]);
        Dispose(newptr);
      end;
      lstDate.Free;
    end;

  finally

    lsts.Free;
    slDate.Free;
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;


  Result := True;
end;

function OpenSUM(const sfile, sProj, syearIn: string; lst: TList): Boolean;
var
  ExcelApp, WorkBook: Variant;
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;
  stitle: string;
  stitle1, stitle2, stitle3, stitle4: string;
  irow, icol: Integer;
  sweek: string;
  splan: string; 
  sver, snumber, scolor, scap: string;
  sitem: string;
  irow1, irow2: Integer; 
  slDate: TStringList; 
  dQty: Double;
  sdate: string;
  sdate1, sdate2: string;
  dt0: TDateTime;
  dt1, dt2: TDateTime;
  aLineRecord: TLineRecord;
  i: Integer;
  s: string;
  idx: Integer;
  syear: string;
  sweekTitle: string;
begin
  
  
  Result := False;
  if not FileExists(sfile) then
  begin
    //MessageBox(frmMergePlans.Handle, '文件不存在', '金蝶提示', 0);
    Exit;
  end;
  
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';

  
  slDate := TStringList.Create;
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        frmMergePlans.Memo1.Lines.Add(sSheet);
                                                       
        stitle1 := Trim(ExcelApp.Cells[1, 2].Value);
        stitle2 := Trim(ExcelApp.Cells[1, 3].Value);
        stitle3 := Trim(ExcelApp.Cells[1, 4].Value);
        stitle4 := Trim(ExcelApp.Cells[1, 5].Value);
        stitle := stitle1 + stitle2 + stitle3 + stitle4;

        if stitle <> '物料编码颜色容量制式' then
        begin
          frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  内容 格式不符合 MPS计划 前几列标题必须是: 物料编码 颜色 容量 制式');
          Continue;
        end;

        irow := 1;
        icol := 7;                                     
        sweekTitle := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        while sdate <> '' do
        begin
          slDate.AddObject(sdate + '=' + sweekTitle, TObject(icol));   
          icol := icol + 1;                               
          sweekTitle := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
        end;

        irow := irow + 2;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin
          sweek := ExcelApp.Cells[irow, 1].Value;
          scolor := ExcelApp.Cells[irow, 3].Value;
          scap := ExcelApp.Cells[irow, 4].Value;
          sver := ExcelApp.Cells[irow, 5].Value;
          splan := ExcelApp.Cells[irow, 6].Value;

          aLineRecord := TLineRecord.Create(snumber, scolor, scap, sver);
          aLineRecord.FWeek := sweek;
          aLineRecord.FPlan := splan;
          lst.Add(aLineRecord);

          dt0 := 0;
          syear := syearIn;

          for i := 0 to slDate.Count - 1 do
          begin 
            icol := Integer(slDate.Objects[i]);
            dQty := ExcelApp.Cells[irow, icol].Value;
            sdate := slDate.Names[i];
            sweekTitle := slDate.ValueFromIndex[i];
 
            if Pos('-', sdate) > 0 then
            begin
              idx := Pos('-', sdate);

              sdate1 := Copy(sdate, 1, idx - 1);
              sdate2 := Copy(sdate, idx + 1, Length(sdate) - Pos('-', sdate));

              sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
              sdate2 := syear + '-' + StringReplace(sdate2, '/', '-', [rfReplaceAll]);

              dt1 := myStrToDateTime(sdate1);
              dt2 := myStrToDateTime(sdate2);

              if dt1 < dt0 then
              begin
                syear := IntToStr( StrToInt(syear) + 1 );
                sdate1 := Copy(sdate, 1, idx - 1);
                sdate2 := Copy(sdate, idx + 1, Length(sdate) - Pos('-', sdate));

                sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                sdate2 := syear + '-' + StringReplace(sdate2, '/', '-', [rfReplaceAll]);

                dt1 := myStrToDateTime(sdate1);
                dt2 := myStrToDateTime(sdate2);
              end;

              dt0 := dt1;
            end
            else
            begin
              sdate1 := sdate; 
              sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
              dt1 := myStrToDateTime(sdate1);
              if dt1 < dt0 then
              begin
                syear := IntToStr( StrToInt(syear) + 1 );
                sdate1 := sdate;
                sdate1 := syear + '-' + StringReplace(sdate1, '/', '-', [rfReplaceAll]);
                dt1 := myStrToDateTime(sdate1); 
              end;
              dt2 := dt1;

              dt0 := dt1;
            end;

            aLineRecord.Add(sweekTitle, sdate, dQty, dt1, dt2)

          end;
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 2].Value;
        end;
 
        frmMergePlans.Memo1.Lines.Add('Sheet ' + sSheet + '  读取完成了  ' + sitem);
        
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    slDate.Free;
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;


  Result := True;
end;
 
procedure GroupByNumber(lst: TList);
var
  i: Integer;
  idx: Integer;
  idate: Integer;
  aLineRecord: TLineRecord;
  adest: TLineRecord;
  sl: TStringList;  
  p1: PDateQty;
  p2: PDateQty;
  s: string;
begin
  sl := TStringList.Create;
  try
    for i := 0 to lst.Count - 1 do
    begin
      aLineRecord := TLineRecord(lst[i]);
      if Pos('海外', aLineRecord.FVer) > 0 then  // 海外的，按颜色、容量、制式汇总
      begin
        s := aLineRecord.FColor + aLineRecord.FCap + aLineRecord.FVer;
      end
      else
      begin
        s := aLineRecord.FNumber;
      end;
      idx := sl.IndexOf(s);
      if idx < 0 then
      begin
        sl.AddObject(s, aLineRecord);
      end
      else
      begin
        adest := TLineRecord(sl.Objects[idx]);
        for idate := 0 to adest.FDates.Count - 1 do
        begin
          p1 := PDateQty(adest.FDates[idate]);
          p2 := PDateQty(aLineRecord.FDates[idate]);
          p1^.dqty := p1^.dqty + p2^.dqty;
          p1^.dqtyact := p1^.dqtyact + p2^.dqtyact;
        end;
        aLineRecord.Free; // 释放掉重复的料号
      end;
    end;
    
    lst.Clear;
    for i := 0 to sl.Count - 1 do
    begin
      lst.Add(Pointer(sl.Objects[i]));
    end;
  finally
    sl.Free;
  end;
end;             
 
procedure GroupByOverSea(lst: TList);
  function IndexOf(sl: TStringList; const scolor, scap, sver: string): Integer;
  var
    ix: Integer;  
    lr: TLineRecord;
  begin
    Result := -1;
    for ix := 0 to sl.Count - 1 do
    begin
      lr := TLineRecord(sl.Objects[ix]);
      if (lr.FColor = scolor) and (lr.FCap = scap) and (lr.FVer = sver) then
      begin
        Result := ix;
        Break;
      end;
    end;
  end;
var
  i: Integer;
  idx: Integer;
  idate: Integer;
  aLineRecord: TLineRecord;
  adest: TLineRecord;
  sl: TStringList;  
  p1: PDateQty;
  p2: PDateQty;
begin
  sl := TStringList.Create;
  try
    for i := 0 to lst.Count - 1 do
    begin                               
      aLineRecord := TLineRecord(lst[i]);
      if Pos('海外', aLineRecord.FVer) > 0 then
      begin
        idx := IndexOf( sl, aLineRecord.FColor, aLineRecord.FCap, aLineRecord.FVer );
        if idx < 0 then
        begin
          sl.AddObject(aLineRecord.FNumber, aLineRecord);
        end
        else
        begin
          adest := TLineRecord(sl.Objects[idx]);
          for idate := 0 to adest.FDates.Count - 1 do
          begin
            p1 := PDateQty(adest.FDates[idate]);
            p2 := PDateQty(aLineRecord.FDates[idate]);
            p1^.dqty := p1^.dqty + p2^.dqty;
            p1^.dqtyact := p1^.dqtyact + p2^.dqtyact;
          end;
          aLineRecord.Free; // 释放掉重复的料号
        end;
      end
      else
      begin
        sl.AddObject(aLineRecord.FNumber, aLineRecord);
      end;
    end;
    
    lst.Clear;
    for i := 0 to sl.Count - 1 do
    begin
      lst.Add(Pointer(sl.Objects[i]));
    end;
  finally
    sl.Free;
  end;
end;

procedure MakeNumberSet(lstNumber: TStringList; lst: TList);
var
  i: Integer;
  aLineRecord: TLineRecord;
  newNode: TLineRecord;
  s: string;
begin
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if Pos('海外', aLineRecord.FVer) > 0 then
    begin
      s := aLineRecord.FColor + aLineRecord.FCap + aLineRecord.FVer;
    end
    else
    begin
      s := aLineRecord.FNumber;
    end;
    if lstNumber.IndexOf( s ) < 0 then
    begin
      newNode := TLineRecord.Create(aLineRecord.FNumber, aLineRecord.FColor,
        aLineRecord.FCap, aLineRecord.FVer);
      lstNumber.AddObject( s, newNode );
    end;
  end;
end;

function GetByNumber(lst: TList; const snumber: string): TLineRecord;
var
  i: Integer; 
  aLineRecord: TLineRecord;
begin
  Result := nil;
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if aLineRecord.FNumber = snumber then
    begin
      Result := aLineRecord;
      Break;
    end;
  end;
end;

function GetByClrCapVerWeekPlan(lst: TList; const
  scolor, scap, sver, slastweek, splan: string): TLineRecord;
var
  i: Integer; 
  aLineRecord: TLineRecord;
begin
  Result := nil;
  for i := lst.Count - 1 downto 0 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if (aLineRecord.FColor = scolor) and (aLineRecord.FCap = scap) and (aLineRecord.FVer = sver)
      and (aLineRecord.FWeek = slastweek)
      and (aLineRecord.FPlan = splan) then
    begin
      Result := aLineRecord;
      Break;
    end;
  end;
end;

function GetByNumberWeekPlan(lst: TList; const
  sNumber, slastweek, splan: string): TLineRecord;
var
  i: Integer; 
  aLineRecord: TLineRecord;
begin
  Result := nil;
  for i := lst.Count - 1 downto 0 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if (aLineRecord.FNumber = snumber) and (aLineRecord.FWeek = slastweek)
      and (aLineRecord.FPlan = splan) then
    begin
      Result := aLineRecord;
      Break;
    end;
  end;
end;

procedure GetAllByNumber(lst: TList; const snumber: string; lstResult: TList);
var
  i: Integer; 
  aLineRecord: TLineRecord;
begin
  lstResult.Clear;
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if aLineRecord.FNumber = snumber then
    begin
      lstResult.Add(aLineRecord);
    end;
  end;
end;

procedure GetAllByClrCapVer(lst: TList; const scolor, scap, sver: string; lstResult: TList);
var
  i: Integer; 
  aLineRecord: TLineRecord;
begin
  lstResult.Clear;
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    if (aLineRecord.FColor = scolor) and (aLineRecord.FCap = scap)
      and (aLineRecord.FVer = sver) then
    begin
      lstResult.Add(aLineRecord);
    end;
  end;
end;

function GetFGStock( const snumber: string; aSAPHW90Reader: TSAPHW90Reader;
  aFGStockReader: TFGStockReader; Memo1: TMemo): Double;
var
  iStock: Integer;
  ifg: Integer;
  sl: TStringList;
begin
  Result := 0;

  // 用成品料号找库存
  for iStock := 0 to aFGStockReader.Count - 1 do
  begin
    if UpperCase(Trim(snumber)) = UpperCase(Trim(aFGStockReader.Items[iStock]^.snumber)) then
    begin
      with aFGStockReader.Items[iStock]^ do
      begin
        Result := fqty + fqty_sf + fqty_kj + fqty_ort;
      end;
      Exit;
    end;
  end;

  // 成品料号找不到库存，认为是裸机料号，汇总裸机料号上层的成品库存（海外的成品料号）

  sl := TStringList.Create;
  try
    // 1、先找出成品料号
    for ifg := 0 to aSAPHW90Reader.Count - 1 do
    begin
      if Trim(aSAPHW90Reader.Items[ifg]^.snumber90) = Trim(snumber) then
      begin
        // 名称 带 H 的认为是海外的
        if  (Pos('H_V', aSAPHW90Reader.Items[ifg]^.sname) > 0)
          or (Pos('H/V', aSAPHW90Reader.Items[ifg]^.sname) > 0)
          or (Pos('H-V', aSAPHW90Reader.Items[ifg]^.sname) > 0) then
        begin
          // 防止重复料号
          if sl.IndexOf(Trim(aSAPHW90Reader.Items[ifg]^.snumber)) < 0 then
          begin
            sl.Add(Trim(aSAPHW90Reader.Items[ifg]^.snumber));
          end;
        end;
      end;
    end;

    Memo1.Lines.Add(sl.Text);
 
    for ifg := 0 to sl.Count - 1 do
    begin
      // 汇总海外成品料号库存
      for iStock := 0 to aFGStockReader.Count - 1 do
      begin
        if sl[ifg] = Trim(aFGStockReader.Items[iStock]^.snumber) then
        begin
          with aFGStockReader.Items[iStock]^ do
          begin
            Result := Result + fqty + fqty_sf + fqty_kj + fqty_ort;
          end;
        end;
      end;
    end;
  finally
    sl.Free;
  end;
end;

procedure WriteAStock(ExcelApp: Variant; const snumber: string; lst: TList;
  irow, icol: Integer; dtWeekStart: TDateTime; FSUMs: TList;
  const slastweek: string; const splan: string; dqty: Double);
var
  i: Integer;
  p: PDateQty;
  aLineRecord: TLineRecord;
  aLineLastWeek: TLineRecord;
begin
  aLineRecord := GetByNumber(lst, snumber);
  if aLineRecord = nil then Exit;

  if Pos('海外', aLineRecord.FVer) > 0 then
  begin
    aLineLastWeek := GetByClrCapVerWeekPlan(FSUMs, aLineRecord.FColor, aLineRecord.FCap, aLineRecord.FVer, slastweek, splan);
  end
  else
  begin
    aLineLastWeek := GetByNumberWeekPlan(FSUMs, aLineRecord.FNumber, slastweek, splan);
  end;

  for i := 0 to aLineRecord.FDates.Count - 1 do
  begin
    p := PDateQty(aLineRecord.FDates[i]);     
    if p^.dt1 < dtWeekStart then    // 过去的周继承过去的期初库存
    begin
      if aLineLastWeek <> nil then
      begin
        p := PDateQty(aLineLastWeek.FDates[i]);
        ExcelApp.Cells[irow, icol + i].Value := p^.dqty;
      end;
    end
    else if p^.dt1 = dtWeekStart then  //  当周库存从库存Excel表格取。
    begin
      ExcelApp.Cells[irow, icol + i].Value := dqty;
      Break; //  往后的不填，跳出
    end;
  end;
end;

procedure WriteASchedule(ExcelApp: Variant; const alr: TLineRecord; lstDate: TList;
  lst: TList; irow, icol: Integer; aFirstDate, dtWeekStart: TDateTime; const slastweek: string;
  lstSUM: TList);
var
  ir: Integer;
  i: Integer;
  j: Integer;
  p: PDateQty;
  psch: PDateQty;
  aLineRecord: TLineRecord;
  dQty, dQtyAct: Double;
  lstResult: TList;
  alrLastWeekPlan: TLineRecord;
  alrLastWeekAct: TLineRecord;
  idate: Integer;
  pDQLastWeek: PDateQty;
begin
  lstResult := TList.Create;
  try
    if Pos('海外', alr.FVer) > 0 then
    begin
      GetAllByClrCapVer(lst, alr.FColor, alr.FCap, alr.FVer, lstResult);
    end
    else
    begin
      GetAllByNumber(lst, alr.FNumber, lstResult);
    end;
    if lstResult.Count = 0 then Exit;

    for ir := 0 to lstResult.Count - 1 do
    begin
      aLineRecord := TLineRecord(lstResult[ir]);
      for i := 0 to lstDate.Count - 1 do
      begin
        p := PDateQty(lstDate[i]);                 


        dQty := 0;
        if p^.dt1 < dtWeekStart then // 找上一周数据
        begin
          if Pos('海外', alr.FVer) > 0 then
          begin
            alrLastWeekPlan := GetByClrCapVerWeekPlan(lstSUM, alr.FColor, alr.FCap, alr.FVer, slastweek, '排产计划');
          end
          else
          begin
            alrLastWeekPlan := GetByNumberWeekPlan(lstSUM, alr.FNumber, slastweek, '排产计划');
          end;
          if alrLastWeekPlan <> nil then
          begin
            for idate := 0 to alrLastWeekPlan.FDates.Count - 1 do
            begin
              pDQLastWeek := PDateQty(alrLastWeekPlan.FDates[idate]);
              if pDQLastWeek^.dt1 = p^.dt1 then
              begin
                dQty := pDQLastWeek^.dqty;
                Break;
              end;
            end;
          end;
        end
        else
        begin            
          for j := 0 to  aLineRecord.FDates.Count - 1 do
          begin
            psch := PDateQty(aLineRecord.FDates[j]);
            if (psch^.dt1 >= p^.dt1) and (psch^.dt1 <= p^.dt2) then
            begin
              dQty := dQty + psch^.dqty; 
            end;
            if psch^.dt1 > p^.dt2 then Break;
          end;
        end;
                
        dQtyAct := 0;
        if p^.dt1 < aFirstDate then // 找上一周数据
        begin
          if Pos('海外', alr.FVer) > 0 then
          begin
            alrLastWeekAct := GetByClrCapVerWeekPlan(lstSUM, alr.FColor, alr.FCap, alr.FVer, slastweek, '实际产出');
          end
          else
          begin
            alrLastWeekAct := GetByNumberWeekPlan(lstSUM, alr.FNumber, slastweek, '实际产出');
          end;
          if alrLastWeekAct <> nil then
          begin
            for idate := 0 to alrLastWeekAct.FDates.Count - 1 do
            begin
              pDQLastWeek := PDateQty(alrLastWeekAct.FDates[idate]);
              if pDQLastWeek^.dt1 = p^.dt1 then
              begin
                dQtyAct := pDQLastWeek^.dqty; // 这里仍然用 dqty
                Break;
              end;
            end;
          end; 
        end
        else
        begin            
          for j := 0 to  aLineRecord.FDates.Count - 1 do
          begin
            psch := PDateQty(aLineRecord.FDates[j]);
            if (psch^.dt1 >= p^.dt1) and (psch^.dt1 <= p^.dt2) then
            begin 
              dQtyAct := dQtyAct + psch^.dQtyAct;
            end;
            if psch^.dt1 > p^.dt2 then Break;
          end;
        end;
 
        ExcelApp.Cells[irow, icol + i].Value := dQty;
        ExcelApp.Cells[irow + 1, icol + i].Value := dQtyAct;
      end;
    end;
  finally
    lstResult.Free;
  end;
end;

procedure WriteAPlan(ExcelApp: Variant; const snumber: string; lst: TList;
  irow, icol: Integer; dtWeekStart: TDateTime; FSUMs: TList;
  const slastweek: string; const splan: string);
var
  i: Integer;
  p: PDateQty;
  aLineRecord: TLineRecord;
  aLineLastWeek: TLineRecord;
begin
  aLineRecord := GetByNumber(lst, snumber);
  if aLineRecord = nil then Exit;

  if Pos('海外', aLineRecord.FVer) > 0 then
  begin
    aLineLastWeek := GetByClrCapVerWeekPlan(FSUMs, aLineRecord.FColor, aLineRecord.FCap, aLineRecord.FVer, slastweek, splan);
  end
  else
  begin
    aLineLastWeek := GetByNumberWeekPlan(FSUMs, aLineRecord.FNumber, slastweek, splan);
  end;

  for i := 0 to aLineRecord.FDates.Count - 1 do
  begin
    p := PDateQty(aLineRecord.FDates[i]);
    if aLineLastWeek <> nil then
    begin
      if p^.dt1 < dtWeekStart then
      begin
        p := PDateQty(aLineLastWeek.FDates[i]);
      end;
    end;
    ExcelApp.Cells[irow, icol + i].Value := p^.dqty;
  end;
end;

procedure TfrmMergePlans.btnSave2Click(Sender: TObject);
const
  CIColDate = 7;
  CSPlans: array[0..5] of string = ('销售计划', 'S&OP供应计划', 'MPS', '排产计划', '实际产出', '期初库存');
var 
  i: Integer;
  j: Integer;
  lstDate: TList;
  lstNumber: TStringList;  
  aLineRecord: TLineRecord;
  p: PDateQty;
  toInsert: PDateQty;
  ExcelApp, WorkBook: Variant;
  sfile: string;
  irow: Integer;
  sweek: string;
  slfile: TStringList;
  dqty: Double;
//  aFGTableReader: TFGTableReader ;
  aFGStockReader: TFGStockReader ;
  aSAPHW90Reader: TSAPHW90Reader;
begin
  Clear;

  if not ExcelSaveDialog(sfile) then Exit;

  Memo1.Lines.Add('开始读排产计划================================================');
  slfile := TStringList.Create;
  slfile.Text := StringReplace(leSchedule.Text, ';', #13#10, [rfReplaceAll]);
  for i := 0 to slfile.Count - 1 do
  begin
    OpenSchedule(slfile[i], FSchedules);
  end;

  aSAPHW90Reader := TSAPHW90Reader.Create(leFGTable.Text);

  Memo1.Lines.Add('开始读成品裸机料号列表================================================');
//  aFGTableReader := TFGTableReader.Create(leFGTable.Text);

  Memo1.Lines.Add('开始读期初库存================================================');
  aFGStockReader := TFGStockReader.Create(leStock.Text);

  Memo1.Lines.Add('开始读MPS================================================');
  OpenMPS(leMPS.Text, leProj.Text, leMPSYear.Text, FMPSs);
  Memo1.Lines.Add('开始读S&OP供应计划================================================');
  OpenMPS(leSOP.Text, leProj.Text, leSOPYear.Text, FSOPs);
  Memo1.Lines.Add('开始读销售计划================================================');
  OpenMPS(leMarket.Text, leProj.Text, leMarketYear.Text, FMarkets);

  Memo1.Lines.Add('开始读汇总文件================================================');
  OpenSUM(leSUM.Text, leProj.Text, leSUMYear.Text, FSUMs);

  FLastWeek := '';
  if FSUMs.Count > 0 then
  begin
    aLineRecord := TLineRecord(FSUMs[FSUMs.Count - 1]);
    FLastWeek := aLineRecord.FWeek;
  end;

  FFirstDate := myStrToDateTime('1900-01-01');
  if FSchedules.Count > 0 then
  begin
    aLineRecord := TLineRecord(FSchedules[0]);
    if aLineRecord.FDates.Count > 0 then
    begin
      p := PDateQty(aLineRecord.FDates[0]);
      FFirstDate := p^.dt1;
    end;
  end;

  sweek := ExtractFileName(ChangeFileExt(leSOP.Text, '') ) ;

  lstDate := TList.Create;
  lstNumber := TStringList.Create;
  try
    MakeDateSet(lstDate, FMPSs);
    MakeDateSet(lstDate, FSOPs);
    MakeDateSet(lstDate, FMarkets);   
    MakeDateSet(lstDate, FSUMs);

    MakeNumberSet(lstNumber, FMPSs);  
    MakeNumberSet(lstNumber, FSOPs);
    MakeNumberSet(lstNumber, FMarkets);

    MakeNumberSet(lstNumber, FSUMs);
        
    for i := 0 to lstDate.Count - 1 do
    begin
      toInsert := PDateQty(lstDate[i]);
      Memo1.Lines.Add( toInsert^.sdate );
    end;


    for i := 0 to lstDate.Count - 1 do
    begin
      toInsert := PDateQty(lstDate[i]);

      FillDataToList(lstDate, FMPSs, toInsert);
      FillDataToList(lstDate, FSOPs, toInsert);
      FillDataToList(lstDate, FMarkets, toInsert);
       
      FillDataToList(lstDate, FSUMs, toInsert);
    end;

    
    // 一个料号可能有多行， 合并数量
    GroupByNumber(FMPSs);
    GroupByNumber(FSOPs);
    GroupByNumber(FMarkets);    
    GroupByNumber(FSchedules);
    GroupByOverSea(FSchedules);
     


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

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[1].Delete;
    end;
    ExcelApp.Sheets[1].Name := '数据集成';

    try
      ProgressBar1.Max := FSUMs.Count + lstNumber.Count;
      ProgressBar1.Position := 1;

      irow := 1;
      ExcelApp.Cells[irow, 1].Value := 'week';
      ExcelApp.Cells[irow, 2].Value := '物料编码';
      ExcelApp.Cells[irow, 3].Value := '颜色';
      ExcelApp.Cells[irow, 4].Value := '容量';
      ExcelApp.Cells[irow, 5].Value := '制式';
      ExcelApp.Cells[irow, 6].Value := '计划';

      ExcelApp.Range[ ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow + 1, 1] ].MergeCells := True;
      ExcelApp.Range[ ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow + 1, 2] ].MergeCells := True;
      ExcelApp.Range[ ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow + 1, 3] ].MergeCells := True;
      ExcelApp.Range[ ExcelApp.Cells[irow, 4], ExcelApp.Cells[irow + 1, 4] ].MergeCells := True;
      ExcelApp.Range[ ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow + 1, 5] ].MergeCells := True;
      ExcelApp.Range[ ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + 1, 6] ].MergeCells := True;

      ExcelApp.Columns[1].ColumnWidth := 14;
      ExcelApp.Columns[2].ColumnWidth := 16;
      ExcelApp.Columns[3].ColumnWidth := 7;
      ExcelApp.Columns[4].ColumnWidth := 6;
      ExcelApp.Columns[5].ColumnWidth := 13;
      ExcelApp.Columns[6].ColumnWidth := 8;

      for i := 0 to lstDate.Count - 1 do
      begin
        p := PDateQty(lstDate[i]);
        ExcelApp.Cells[irow, CIColDate + i].Value := p^.sweek;  
        ExcelApp.Cells[irow + 1, CIColDate + i].Value := p^.sdate;
        ExcelApp.Columns[CIColDate + i].ColumnWidth := 11;
      end;
       
      irow := irow + 2;     // 跳过两行标题行
      for i := 0 to FSUMs.Count - 1 do
      begin
        aLineRecord := TLineRecord(FSUMs[i]);

        ExcelApp.Cells[irow, 1].Value := aLineRecord.FWeek;
        ExcelApp.Cells[irow, 2].Value := aLineRecord.FNumber;
        ExcelApp.Cells[irow, 3].Value := aLineRecord.FColor;
        ExcelApp.Cells[irow, 4].Value := aLineRecord.FCap;
        ExcelApp.Cells[irow, 5].Value := aLineRecord.FVer;
        ExcelApp.Cells[irow, 6].Value := aLineRecord.FPlan;

        for j := 0 to aLineRecord.FDates.Count - 1 do
        begin
          p := PDateQty(aLineRecord.FDates[j]);
          ExcelApp.Cells[irow, CIColDate + j].Value := p^.dqty;
        end;
 
        irow := irow + 1;

        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;

            
      for i := 0 to lstNumber.Count - 1 do
      begin
        aLineRecord := TLineRecord(lstNumber.Objects[i]);
        for j := 0 to Length(CSPlans) - 1 do
        begin
          ExcelApp.Cells[irow + j, 1].Value := sweek;
          ExcelApp.Cells[irow + j, 2].Value := aLineRecord.FNumber;
          ExcelApp.Cells[irow + j, 3].Value := aLineRecord.FColor;
          ExcelApp.Cells[irow + j, 4].Value := aLineRecord.FCap;
          ExcelApp.Cells[irow + j, 5].Value := aLineRecord.FVer;
          ExcelApp.Cells[irow + j, 6].Value := CSPlans[j];
        end;
                                                                             
        WriteAPlan(ExcelApp, aLineRecord.FNumber, FMarkets, irow, CIColDate, dtpWeekStart.DateTime, FSUMs, FLastWeek, CSPlans[0]);
        WriteAPlan(ExcelApp, aLineRecord.FNumber, FSOPs, irow + 1, CIColDate, dtpWeekStart.DateTime, FSUMs, FLastWeek, CSPlans[1]);
        WriteAPlan(ExcelApp, aLineRecord.FNumber, FMPSs, irow + 2, CIColDate, dtpWeekStart.DateTime, FSUMs, FLastWeek, CSPlans[2]);

        WriteASchedule(ExcelApp, aLineRecord, lstDate, FSchedules, irow + 3,
          CIColDate, FFirstDate, dtpWeekStart.DateTime, FLastWeek, FSUMs);

        dqty := GetFGStock(aLineRecord.FNumber, aSAPHW90Reader, aFGStockReader, Memo1);
        WriteAStock(ExcelApp, aLineRecord.FNumber, FMPSs, irow + 5, CIColDate, dtpWeekStartStock.DateTime, FSUMs, FLastWeek, CSPlans[5], dqty);
 
        
        irow := irow + 6;

        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;                                      

      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[2, CIColDate + lstDate.Count - 1] ].Interior.Color := $DBDCF2;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[2, CIColDate + lstDate.Count - 1] ].HorizontalAlignment := xlCenter;
      ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow - 1, CIColDate + lstDate.Count - 1] ].Borders.LineStyle := 1; //加边框

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end; 
    
    MessageBox(Handle, '完成', '金蝶提示', 0);

  finally
    for i := 0 to lstDate.Count - 1 do
    begin
      p := PDateQty(lstDate[i]);
      Dispose(p);
    end;
    lstDate.Clear;
    lstDate.Free;

    for i := 0 to lstNumber.Count - 1 do
    begin                
      aLineRecord := TLineRecord(lstNumber.Objects[i]);
      aLineRecord.Free;
    end;
    lstNumber.Clear;
    lstNumber.Free;

    aSAPHW90Reader.Free;
    aFGStockReader.Free;
  end; 
end;

procedure ClearList(lst: TList);
var
  i: Integer;
  aLineRecord: TLineRecord;
begin
  for i := 0 to lst.Count - 1 do
  begin
    aLineRecord := TLineRecord(lst[i]);
    aLineRecord.Free;
  end;
  lst.Clear;
end;

procedure TfrmMergePlans.Clear;
begin
  ClearList(FSchedules);
  ClearList(FMPSs);
  ClearList(FSOPs);
  ClearList(FMarkets);
  ClearList(FSUMs);
end;


function GetFileVersion(FileName: string): string; 
   type 
     PVerInfo = ^TVS_FIXEDFILEINFO; 
     TVS_FIXEDFILEINFO = record 
       dwSignature: longint; 
       dwStrucVersion: longint; 
       dwFileVersionMS: longint; 
       dwFileVersionLS: longint; 
       dwFileFlagsMask: longint; 
       dwFileFlags: longint; 
       dwFileOS: longint; 
       dwFileType: longint; 
       dwFileSubtype: longint; 
       dwFileDateMS: longint; 
       dwFileDateLS: longint; 
     end; 
   var 
     ExeNames: array[0..255] of char;  
     VerInfo: PVerInfo; 
     Buf: pointer; 
     Sz: word; 
     L, Len: Cardinal; 
   begin 
     StrPCopy(ExeNames, FileName); 
     Sz := GetFileVersionInfoSize(ExeNames, L); 
     if Sz=0 then 
     begin 
       Result:=''; 
       Exit; 
     end; 

     try
       GetMem(Buf, Sz); 
       try 
         GetFileVersionInfo(ExeNames, 0, Sz, Buf); 
         if VerQueryValue(Buf, '\', Pointer(VerInfo), Len) then 
         begin 
           Result := IntToStr(HIWORD(VerInfo.dwFileVersionMS)) + '.' + 
                     IntToStr(LOWORD(VerInfo.dwFileVersionMS)) + '.' + 
                     IntToStr(HIWORD(VerInfo.dwFileVersionLS)) + '.' + 
                     IntToStr(LOWORD(VerInfo.dwFileVersionLS)); 

         end; 
       finally 
         FreeMem(Buf); 
       end; 
     except 
       Result := '-1'; 
     end; 
   end;

procedure TfrmMergePlans.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin
  dtpWeekStart.DateTime := myStrToDateTime('1900-01-01');
  dtpWeekStartStock.DateTime := myStrToDateTime('1900-01-01');
  FSchedules := TList.Create;
  FMPSs := TList.Create;
  FSOPs := TList.Create;
  FMarkets := TList.Create;
  FSUMs := TList.Create;

  ini := TIniFile.Create(AppIni);
  try
    leProj.Text := ini.ReadString(self.ClassName, leProj.Name, '');
    leSchedule.Text := ini.ReadString(self.ClassName, leSchedule.Name, '');
    leMPS.Text := ini.ReadString(self.ClassName, leMPS.Name, '');
    leSOP.Text := ini.ReadString(self.ClassName, leSOP.Name, '');
    leMarket.Text := ini.ReadString(self.ClassName, leMarket.Name, '');
    leSUM.Text := ini.ReadString(self.ClassName, leSUM.Name, '');
    leMPSYear.Text := ini.ReadString(self.ClassName, leMPSYear.Name, '');
    leSOPYear.Text := ini.ReadString(self.ClassName, leSOPYear.Name, '');
    leMarketYear.Text := ini.ReadString(self.ClassName, leMarketYear.Name, '');
    leSUMYear.Text := ini.ReadString(self.ClassName, leSUMYear.Name, '');       
    leFGTable.Text := ini.ReadString(self.ClassName, leFGTable.Name, '');
    leStock.Text := ini.ReadString(self.ClassName, leStock.Name, '');
    dtpWeekStart.DateTime := myStrToDateTime( StringReplace( ini.ReadString(self.ClassName, dtpWeekStart.Name, '1900-01-01'), '/', '-', [rfReplaceAll]) );
    dtpWeekStartStock.DateTime := myStrToDateTime( StringReplace( ini.ReadString(self.ClassName, dtpWeekStartStock.Name, '1900-01-01'), '/', '-', [rfReplaceAll]) );
  finally
    ini.Free;
  end;
end;
  
procedure TfrmMergePlans.FormDestroy(Sender: TObject);  
var
  ini: TIniFile;
begin
  Clear;
  FSchedules.Free;
  FMPSs.Free;  
  FSOPs.Free;
  FMarkets.Free;
  FSUMs.Free;

  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leProj.Name, leProj.Text);
    ini.WriteString(self.ClassName, leSchedule.Name, leSchedule.Text);
    ini.WriteString(self.ClassName, leMPS.Name, leMPS.Text);
    ini.WriteString(self.ClassName, leSOP.Name, leSOP.Text);
    ini.WriteString(self.ClassName, leMarket.Name, leMarket.Text);
    ini.WriteString(self.ClassName, leSUM.Name, leSUM.Text);
    ini.WriteString(self.ClassName, leMPSYear.Name, leMPSYear.Text);
    ini.WriteString(self.ClassName, leSOPYear.Name, leSOPYear.Text);
    ini.WriteString(self.ClassName, leMarketYear.Name, leMarketYear.Text);
    ini.WriteString(self.ClassName, leSUMYear.Name, leSUMYear.Text);
    ini.WriteString(self.ClassName, dtpWeekStart.Name, FormatDateTime('YYYY-MM-DD', dtpWeekStart.DateTime));  
    ini.WriteString(self.ClassName, dtpWeekStartStock.Name, FormatDateTime('YYYY-MM-DD', dtpWeekStartStock.DateTime));
    ini.WriteString(self.ClassName, leFGTable.Name, leFGTable.Text);
    ini.WriteString(self.ClassName, leStock.Name, leStock.Text);
  finally
    ini.Free;
  end;  
end;

procedure TfrmMergePlans.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmMergePlans.btnFGTableClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leFGTable.Text := sfile;
end;

procedure TfrmMergePlans.btnStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leStock.Text := sfile;
end;

end.
