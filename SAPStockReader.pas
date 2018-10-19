unit SAPStockReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TSAPStockSum = class;
  
  TSAPStockRecord = packed record
    snumber: string;
    sstock: string;
    dqty: Double;
    dQty_Alloc: Double;
  end;
  PSAPStockRecord = ^TSAPStockRecord;
 
  TSAPStockReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string); 
  public
    FList: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function AllocStockSum(const snumber: string): Double;
    function GetAvailStock(const sNumber: string): Double;
    procedure SumTo(aSAPStockSum: TSAPStockSum);
    property ReadOk: Boolean read FReadOk;
  end;

  TSAPStockSum = class
  public
    FList: TStringList;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    function GetAvailStock(const snumber: string): Double;
    procedure Alloc(const snumber: string; dQty: Double);
    function GetStock(const snumber: string): Double;


    function Alloc2(const snumber: string; dQty: Double): Double;
  end;

implementation
      
{ TSAPStockSum }

constructor TSAPStockSum.Create;
begin
  FList := TStringList.Create;
end;

destructor TSAPStockSum.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPStockSum.Clear;
var
  i: Integer;
  p: PSAPStockRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPStockSum.GetAvailStock(const snumber: string): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    Result := p^.dqty - p^.dQty_Alloc;
  end
  else Result := 0;
end;

function TSAPStockSum.GetStock(const snumber: string): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    Result := p^.dqty;
  end
  else Result := 0;
end;

function TSAPStockSum.Alloc2(const snumber: string; dQty: Double): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    // 可用量够分配
    if DoubleGE( p^.dqty - p^.dQty_Alloc , dQty ) then
    begin
      Result := dQty;
      p^.dQty_Alloc := p^.dQty_Alloc + dQty;
    end
    else   // 部分满足
    begin
      Result := p^.dqty - p^.dQty_Alloc;
      p^.dQty_Alloc := p^.dqty;
    end; 
  end
  else Result := 0;
end;  

procedure TSAPStockSum.Alloc(const snumber: string; dQty: Double);
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  if DoubleE(dQty, 0) then Exit;
  
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    if DoubleL(p^.dqty, p^.dQty_Alloc + dQty) then // 不够分配
    begin
      raise Exception.Create(snumber + ' 库存不足，无法分配库存  TSAPStockSum');
    end
    else
    begin
      p^.dQty_Alloc := p^.dQty_Alloc + dQty;
    end;
  end
  else raise Exception.Create(snumber + ' 库存不存在，无法分配库存  TSAPStockSum');
end;  

{ TSAPStockReader }

constructor TSAPStockReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPStockReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPStockReader.Clear;
var
  i: Integer;
  p: PSAPStockRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPStockReader.AllocStockSum(const snumber: string): Double;
var
  i: Integer;
  p: PSAPStockRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    if p^.snumber <> snumber then Continue;
    Result := Result + p^.dqty;
    p^.dqty := 0; // 分配了，置0
  end;
end;

function TSAPStockReader.GetAvailStock(const sNumber: string): Double;
var
  i: Integer;
  p: PSAPStockRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    if p^.snumber = snumber then
    begin
      Result := Result + p^.dqty;
    end;
  end;
end;

procedure TSAPStockReader.SumTo(aSAPStockSum: TSAPStockSum);
var
  i: Integer;
  p: PSAPStockRecord;
  idx: Integer;
  pSum: PSAPStockRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    idx := aSAPStockSum.FList.IndexOf(p^.snumber);
    if idx >= 0 then
    begin
      pSum := PSAPStockRecord(aSAPStockSum.FList.Objects[idx]);
      pSum^.dqty := pSum^.dqty + p^.dqty;
    end
    else
    begin
      pSum := New(PSAPStockRecord);
      aSAPStockSum.FList.AddObject(p^.snumber, TObject(pSum));
      pSum^ := p^;
    end;
  end;
  aSAPStockSum.FList.Sort;
end;  

procedure TSAPStockReader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function IndexOfCol(ExcelApp: Variant; irow: Integer; const scol: string): Integer;
var
  i: Integer;
  s: string;
begin
  Result := -1;
  for i := 1 to 50 do
  begin
    s := ExcelApp.Cells[irow, i].Value;
    if s = scol then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TSAPStockReader.Open;
const
  CSNumber = '物料';
  CSStock = '库存地点';
  CSQty = '非限制使用的库存';
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PSAPStockRecord;
  iColNumber: Integer;
  iColStock: Integer;
  iColQty: Integer;
begin
  Clear;
          
  FReadOk := False;

  if not FileExists(FFile) then Exit;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(FFile);   

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
        if stitle <> '工厂库存地点仓储地点的描述物料物料描述非限制使用的库存' then
        begin        

          Log(sSheet +'  不是  SAP导出库存  格式  ( 工厂库存地点仓储地点的描述物料物料描述非限制使用的库存 )');
          Continue;
        end;

        iColNumber := IndexOfCol(ExcelApp, irow, CSNumber);
        iColStock := IndexOfCol(ExcelApp, irow, CSStock);
        iColQty := IndexOfCol(ExcelApp, irow, CSQty);

        if (iColNumber = -1) or (iColStock = -1) or (iColQty = -1) then
        begin
          Log(sSheet +'  不是  SAP导出库存  格式');
          Continue;
        end;
                      
        FReadOk := True;

        irow := 2;
        snumber := ExcelApp.Cells[irow, iColNumber].Value;
        while snumber <> '' do
        begin                                
          p := New(PSAPStockRecord);
          FList.AddObject(snumber, TObject(p));

          p^.snumber := snumber;
          p^.sstock := ExcelApp.Cells[irow, iColStock].Value;
          p^.dqty := ExcelApp.Cells[irow, iColQty].Value; 
          p^.dQty_Alloc := 0;
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, iColNumber].Value;
        end;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;  
end;
 
end.
