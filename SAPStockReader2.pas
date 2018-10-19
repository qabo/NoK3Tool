unit SAPStockReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
      
  TTransferRecord = packed record
    snumber: string;
    sname: string;
    dqty: Double;
    dt: TDateTime;
    sfrom: string;
    sto: string;
  end;
  PTransferRecord = ^TTransferRecord;
  
  TAllocUnit = packed record
    sMrpArea: string;
    sStock: string;
    dQty: Double;
  end;
  PAllocUnit = ^TAllocUnit;
         
  TSAPStockSum = class;
  
  TSAPStockRecord = packed record
    snumber: string;
    sname: string;
    sstock: string;
    dqty: Double;
    dQty_Alloc: Double;
    dDemand: Double;
  end;
  PSAPStockRecord = ^TSAPStockRecord;
 
  TSAPStockReader2 = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPStockRecord;
  public
    FList: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function AllocStockSum(const snumber: string): Double;
    function GetAvailStock(const sNumber: string): Double;
    procedure SumTo(aSAPStockSum: TSAPStockSum);
    function GetStocks(const snumber: string; sl: TStringList): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPStockRecord read GetItems;
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


    function Alloc2(const smrparea, snumber: string; dQty: Double): Double;
    function AllocStock_area(const snumber: string; dQty: Double): Double;
    function AccDemand(const snumber: string; dQty: Double): Double;
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

function TSAPStockSum.Alloc2(const smrparea, snumber: string; dQty: Double): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
  aAllocUnitPtr: PAllocUnit;
begin
  Result := 0;
  if DoubleE(dQty, 0) then Exit;
  
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    if DoubleE( p^.dqty, p^.dQty_Alloc ) then Exit; // 分配完了

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
  end;
end;

function TSAPStockSum.AllocStock_area(const snumber: string; dQty: Double): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    if DoubleE(p^.dDemand, 0) then  // 本身没需求，允许分配给其他区域
    begin 
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
    end; 
  end
  else Result := 0;
end;

function TSAPStockSum.AccDemand(const snumber: string; dQty: Double): Double;
var
  idx: Integer;
  p: PSAPStockRecord;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    p := PSAPStockRecord(FList.Objects[idx]);
    p^.dDemand := p^.dDemand + dQty;
  end;
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

{ TSAPStockReader2 }

constructor TSAPStockReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPStockReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPStockReader2.Clear;
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

function TSAPStockReader2.AllocStockSum(const snumber: string): Double;
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

function TSAPStockReader2.GetAvailStock(const sNumber: string): Double;
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

procedure TSAPStockReader2.SumTo(aSAPStockSum: TSAPStockSum);
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
end;  

function TSAPStockReader2.GetStocks(const snumber: string; sl: TStringList): Double;
var
  i: Integer;
  p: PSAPStockRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPStockRecord(FList.Objects[i]);
    if (p^.snumber = snumber) and ( sl.IndexOfName(p^.sstock) >= 0 ) then
    begin
      Result := Result + p^.dqty;
    end;
  end;
end;  

procedure TSAPStockReader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TSAPStockReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPStockReader2.GetItems(i: Integer): PSAPStockRecord;
begin
  Result := PSAPStockRecord(FList.Objects[i]);
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

procedure TSAPStockReader2.Open;
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


  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
begin
  Clear;

  if not FileExists(FFile) then Exit;

       
  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'Sheet1'+'$]';

    ADOTabXLS.Active:=true;


    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      snumber := ADOTabXLS.FieldByName('物料').AsString; // ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;

      p := New(PSAPStockRecord);
      FList.AddObject(snumber, TObject(p));

      p^.snumber := snumber;
      p^.sname :=  ADOTabXLS.FieldByName('物料描述').AsString; 
      p^.sstock := ADOTabXLS.FieldByName('库存地点').AsString; //  ExcelApp.Cells[irow, iColStock].Value;
      p^.dqty := ADOTabXLS.FieldByName('非限制使用的库存').AsFloat; //  ExcelApp.Cells[irow, iColQty].Value;
      p^.dQty_Alloc := 0;
      p^.dDemand := 0;

      snumber := ADOTabXLS.FieldByName('物料').AsString; //  ExcelApp.Cells[irow, iColNumber].Value;
      
      ADOTabXLS.Next;
    end;

    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;    
end;
 
end.
