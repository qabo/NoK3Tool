unit SAPOPOReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TSAPOPOAlloc = packed record
    dQty: Double;
    dt: TDateTime;
    sMrpAreaNo: string; 
  end;
  PSAPOPOAlloc= ^TSAPOPOAlloc;
  
  TSAPOPOLine = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPOPOAlloc;    
  public
    FNumber: string;
    FName: string;
    FDate: TDateTime;
    FStock: string; // 库存地点
    FQty: Double;
    FQtyAlloc: Double;
    FBillNo: string;
    FLine: string;
    FPlanLine: string;
    FBillDate: TDateTime;
    FSupplier: string;
    Tag: TObject;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;              
    function Alloc(dt: TDateTime; var dQty: Double;
      const sAreaNo: string): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPOPOAlloc read GetItems;
  end;
  
  TSAPOPOReader2 = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPOPOLine;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                                                           
    procedure GetOPOs(slNumber: TStringList; lst: TList);
    function Alloc(slNumber: TStringList; dt: TDateTime; dQty: Double): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPOPOLine read GetItems;
  end;

    function ListSortCompare_DateTime_PO(Item1, Item2: Pointer): Integer;

implementation
      
{ TSAPOPOLine }

constructor TSAPOPOLine.Create;
begin
  FList := TList.Create;
  FQtyAlloc := 0;
end;

destructor TSAPOPOLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPOPOLine.Clear;
var
  p: PSAPOPOAlloc;
  i: Integer;
begin
  for i := 0 to FList.Count -1 do
  begin
    p := PSAPOPOAlloc(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPOPOLine.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPOPOLine.GetItems(i: Integer): PSAPOPOAlloc;
begin
  Result := PSAPOPOAlloc(FList[i]);
end;

function TSAPOPOLine.Alloc(dt: TDateTime; var dQty: Double;
  const sAreaNo: string): Double;
var                      
  p: PSAPOPOAlloc;
begin
  Result := 0;
  if DoubleG( FQty, FQtyAlloc ) then //有可分配
  begin
    p := New(PSAPOPOAlloc);
    FList.Add(p);
    if DoubleGE( FQty - FQtyAlloc, dQty ) then //数量够分配
    begin                             
      Result := dQty;
      FQtyAlloc := FQtyAlloc + dQty;

      dQty := 0;
    end
    else  // 数量不够分配
    begin                         
      Result := FQty - FQtyAlloc;
      FQtyAlloc := FQty;
      
      dQty := dQty - Result;
    end;                            

    p^.dQty := Result;
    p^.dt := dt;
    p^.sMrpAreaNo := sAreaNo; 
  end;
end;  

{ TSAPOPOReader2 }

constructor TSAPOPOReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPOPOReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPOPOReader2.Clear;
var
  i: Integer;
  p: TSAPOPOLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := TSAPOPOLine(FList.Objects[i]);
    p.Free;
  end;
  FList.Clear;
end;

procedure TSAPOPOReader2.GetOPOs(slNumber: TStringList; lst: TList);
var
  i: Integer;
begin
  for i := 0 to FList.Count - 1 do
  begin
    if slNumber.IndexOf(FList[i]) >= 0 then
    begin
      lst.Add( FList.Objects[i] );
    end;
  end;
end;

function TSAPOPOReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPOPOReader2.GetItems(i: Integer): TSAPOPOLine;
begin
  Result := TSAPOPOLine(FList.Objects[i]);
end;

              
    function ListSortCompare_DateTime_PO(Item1, Item2: Pointer): Integer;
    var
      p1, p2: TSAPOPOLine;
    begin
      p1 := TSAPOPOLine(Item1);
      p2 := TSAPOPOLine(Item2);
      
      if DoubleG(p1.FDate, p2.FDate) then
        Result := 1
      else if DoubleL(p1.FDate, p2.FDate) then
        Result := -1
      else Result := 0;
    end;

function TSAPOPOReader2.Alloc(slNumber: TStringList; dt: TDateTime;
  dQty: Double): Double;
var
  lst: TList;
  i: Integer;
  p: TSAPOPOLine;
begin
  Result := 0;
  
  lst := TList.Create;
  GetOPOs(slNumber, lst); // 找到所有替代料的可用采购订单

  lst.Sort(ListSortCompare_DateTime_PO);  // 按日期排序
  for i := 0 to lst.Count - 1 do
  begin
    p := TSAPOPOLine(lst[i]);
    Result := Result + p.Alloc(dt, dQty, '');
    if DoubleE( dQty, 0) then // 需求满足分配了
    begin
      Break;
    end;
  end;

  lst.Free;
end;

procedure TSAPOPOReader2.Log(const str: string);
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

procedure TSAPOPOReader2.Open;
const
  CSNumber = '物料';
  CSName = '短文本';   
  CSColDate = '凭证日期';
  CSStock = '库存地点';
  CSQty = '计划数量';
  CSBillNo = '采购凭证';
  CSLine = '项目';
  CSPlanLine = '计划行';

var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  aSAPOPOLine: TSAPOPOLine;
  iColNumber: Integer;
  iColName: Integer;
  iColDate: Integer;
  iColStock: Integer;
  iColQty: Integer;
  iColBillNo: Integer;
  iColLine: Integer;
  iColPlanLine: Integer;
  
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
  s: string;
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

      snumber := ADOTabXLS.FieldByName('物料').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;

      s := ADOTabXLS.FieldByName('删除标识').AsString;  //  ExcelApp.Cells[irow, iColName].Value;
      if s = 'L' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
            
      s := ADOTabXLS.FieldByName('工厂').AsString;  //  ExcelApp.Cells[irow, iColName].Value;
      if s <> '1001' then
      begin
        ADOTabXLS.Next;
        Continue;
      end;
           
      aSAPOPOLine := TSAPOPOLine.Create;
      FList.AddObject(snumber, aSAPOPOLine);

      aSAPOPOLine.FNumber := snumber;
      aSAPOPOLine.FName := ADOTabXLS.FieldByName('短文本').AsString;  //  ExcelApp.Cells[irow, iColName].Value;
      aSAPOPOLine.FDate := ADOTabXLS.FieldByName('交货日期').AsDateTime;  //  ExcelApp.Cells[irow, iColDate].Value;
      aSAPOPOLine.FStock := ADOTabXLS.FieldByName('库存地点').AsString;  //  ExcelApp.Cells[irow, iColStock].Value;
      aSAPOPOLine.FQty := ADOTabXLS.FieldByName('仍要交货(数量)').AsFloat;      
      aSAPOPOLine.FBillNo := ADOTabXLS.FieldByName('采购凭证').AsString;  //  ExcelApp.Cells[irow, iColBillNo].Value;
      aSAPOPOLine.FLine := ADOTabXLS.FieldByName('项目').AsString;  //  ExcelApp.Cells[irow, iColLine].Value;
      aSAPOPOLine.FPlanLine := ADOTabXLS.FieldByName('计划行').AsString;  //   ExcelApp.Cells[irow, iColPlanLine].Value;
      aSAPOPOLine.FBillDate := ADOTabXLS.FieldByName('凭证日期').AsDateTime;  //  ExcelApp.Cells[irow, iColDate].Value;
      aSAPOPOLine.FSupplier := ADOTabXLS.FieldByName('供应商/供货工厂').AsString;  //  ExcelApp.Cells[irow, iColDate].Value;


      irow := irow + 1;
      snumber := ADOTabXLS.FieldByName('物料').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;

          
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
