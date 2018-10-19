unit SAPMrpAreaStockReader;

interface
          
uses
  Classes, SysUtils, ComObj, SAPOPOReader2, SAPStockReader2, CommUtils;

type
  TMyStringList = class(TStringList)
  public
    tag: TObject;
  end;
  
  TSAPMrpAREA = class
  private
    procedure GetOPOs(slNumber: TStringList; lst: TList);
  public 
    sAreaNo: string;
    sAreaName: string; 
    FList: TStringList;
    FOPOList: TMyStringList;
    FSAPStockSum: TSAPStockSum;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;        
    function Alloc(slNumber: TStringList; dt: TDateTime; dQty: Double): Double;       
  end;

  TSAPMrpAreaStockReader = class
  private
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPMrpAREA;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                                                           

    procedure SetOPOList(aSAPOPOReader2: TSAPOPOReader2);
    function Alloc(slNumber: TStringList; dt: TDateTime; dQty: Double;
      const sMrpArea: string): Double;
      
    procedure SetStock(aSAPStockReader: TSAPStockReader2);
    function AllocStock(const snumber: string; dQty: Double;
      const smrparea: string): Double;   
    function AllocStock2(const snumber: string; dQty: Double;
      const smrparea: string): Double;
    function AllocStock_area(const snumber, sname: string; dt: TDateTime;
      dQty: Double; const smrparea: string; slTransfer: TList): Double;
    function AccDemand(const snumber: string; dQty: Double;
      const smrparea: string): Double;
    
    function GetSAPMrpAREA(const sMrpArea: string): TSAPMrpAREA;
    procedure GetOPOs(slNumber: TStringList; lst: TList; const sMrpArea: string);
    function MrpAreaNo2Name(const sMrpAreaNo: string): string;
    function MrpAreaOfStockNo(const sstock: string): string;

    function GetStockSum(const snumber: string): Double;
    function GetOPOSum(const snumber: string): Double;

    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPMrpAREA read GetItems;
    property ReadOk: Boolean read FReadOk;
  end; 

implementation

procedure QuickSort(sl: TStringList; L, R: Integer;
  SCompare: TStringListSortCompare);
var
  I, J: Integer;
  P, T: Pointer;
  iP: Integer;
  o: TObject;
  s: string;
begin
  repeat
    I := L;
    J := R;
    iP := (L + R) shr 1;
    repeat
      while SCompare(sl, I, iP) < 0 do
        Inc(I);
      while SCompare(sl, J, iP) > 0 do
        Dec(J);
      if I <= J then
      begin
        o := sl.Objects[I];
        sl.Objects[I] := sl.Objects[J];
        sl.Objects[J] := o;

        s := sl[I];
        sl[I] := sl[J];
        sl[J] := s;
        
        Inc(I);
        Dec(J);
      end;
    until I > J;
    if L < J then
      QuickSort(sl, L, J, SCompare);
    L := I;
  until I >= R;
end;
       
  function StringListSortCompare_OPO(List: TStringList; Index1, Index2: Integer): Integer;
  var                           
    aSAPMrpAREA: TSAPMrpAREA;
    aSAPOPOLine1: TSAPOPOLine;
    aSAPOPOLine2: TSAPOPOLine;
  begin        
    aSAPMrpAREA := TSAPMrpAREA( TMyStringList(List).tag );

    Result := 0;
//    if List[Index1] <> List[Index2] then Exit; // 料号不同，不变顺序

    aSAPOPOLine1 := TSAPOPOLine(List.Objects[Index1]);
    aSAPOPOLine2 := TSAPOPOLine(List.Objects[Index2]);

    if aSAPOPOLine1.FNumber > aSAPOPOLine2.FNumber then
    begin
      Result := 1
    end
    else if aSAPOPOLine1.FNumber < aSAPOPOLine2.FNumber then
    begin
      Result := -1
    end
    else
    begin
      if DoubleG( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
      begin
        Result := 1;
      end
      else if DoubleL( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
      begin
        Result := -1
      end
      else
      begin
        if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) >= 0) and   // 仓库在本区域， 排前面
          (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) < 0) then
        begin
          Result := -1;
        end
        else if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) < 0) and  // 仓库不在本区域， 排后面
          (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) >= 0) then
        begin
          Result := 1;
        end
        else
        begin
          Result := 0;
        end;
      end;
    end;
  end;

{ TSAPMrpAREA }

constructor TSAPMrpAREA.Create;
begin
  FOPOList := TMyStringList.Create;
  FOPOList.CaseSensitive := False;
  FOPOList.tag := Self;
  FList := TStringList.Create;
  FSAPStockSum := TSAPStockSum.Create;
end;

destructor TSAPMrpAREA.Destroy;
begin
  Clear;
  FList.Free;
  FOPOList.Free;
  FSAPStockSum.Free;
  inherited;
end;

procedure TSAPMrpAREA.Clear;
begin
  FOPOList.Clear;
  FList.Clear;
  FSAPStockSum.Clear;
end;
                                                             
function ListSortCompare(Item1, Item2: Pointer): Integer;
var
  aSAPOPOLine1: TSAPOPOLine;
  aSAPOPOLine2: TSAPOPOLine;
  aSAPMrpAREA: TSAPMrpAREA;
begin
  aSAPOPOLine1 := TSAPOPOLine(Item1);
  aSAPOPOLine2 := TSAPOPOLine(Item2);

  if DoubleG( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
  begin
    Result := 1;
  end
  else if DoubleL( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
  begin
    Result := -1
  end
  else
  begin
    aSAPMrpAREA := TSAPMrpAREA(aSAPOPOLine1.Tag);
    
    if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) >= 0) and   // 仓库在本区域， 排前面
      (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) < 0) then
    begin
      Result := -1;
    end
    else if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) < 0) and  // 仓库不在本区域， 排后面
      (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) >= 0) then
    begin
      Result := 1;
    end
    else
    begin
      Result := 0;
    end;
  end;
end;

procedure TSAPMrpAREA.GetOPOs(slNumber: TStringList; lst: TList);
var
  i: Integer;
  idx: Integer;
begin
  for i := 0 to slNumber.Count - 1 do
  begin
    idx := FOPOList.IndexOf(slNumber[i]);
    if idx >= 0 then
    begin
      while (idx >= 0) and  (FOPOList[idx] = slNumber[i]) do
      begin
        idx := idx - 1;
      end;
      idx := idx + 1;
      while (idx < FOPOList.Count) and (FOPOList[idx] = slNumber[i]) do
      begin
        lst.Add(FOPOList.Objects[idx]);
        idx := idx + 1;
      end;
    end;
  end;

  lst.Sort(ListSortCompare);
//

//  for i := 0 to FOPOList.Count - 1 do
//  begin
//    if slNumber.IndexOf(FOPOList[i]) >= 0 then
//    begin
//      lst.Add( FOPOList.Objects[i] );
//    end;
//  end;
end;

function TSAPMrpAREA.Alloc(slNumber: TStringList; dt: TDateTime; dQty: Double): Double;
var
  lst: TList;
  i: Integer;
  p: TSAPOPOLine;
begin
  Result := 0;
  
  lst := TList.Create;
  GetOPOs(slNumber, lst); // 找到所有替代料的可用采购订单
 
  for i := 0 to lst.Count - 1 do
  begin
    p := TSAPOPOLine(lst[i]);            
    Result := Result + p.Alloc(dt, dQty, sAreaNo);
    if DoubleE( dQty, 0) then // 需求满足分配了
    begin
      Break;
    end;
  end;

  lst.Free;
end;  
 
{ TSAPMrpAreaStockReader }

constructor TSAPMrpAreaStockReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPMrpAreaStockReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPMrpAreaStockReader.Clear;
var
  i: Integer;
  p: TSAPMrpAREA;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := TSAPMrpAREA(FList.Objects[i]);
    p.Free;
  end;
  FList.Clear;
end;

  function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
  var                           
    aSAPMrpAREA: TSAPMrpAREA;
    aSAPOPOLine1: TSAPOPOLine;
    aSAPOPOLine2: TSAPOPOLine;
  begin        
    aSAPMrpAREA := TSAPMrpAREA( TMyStringList(List).tag );
    
    aSAPOPOLine1 := TSAPOPOLine(List.Objects[Index1]);
    aSAPOPOLine2 := TSAPOPOLine(List.Objects[Index2]);
    if DoubleG( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
    begin
      Result := 1;
    end
    else if DoubleL( aSAPOPOLine1.FDate, aSAPOPOLine2.FDate ) then
    begin
      Result := -1
    end
    else
    begin
      if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) >= 0) and   // 仓库在本区域， 排前面
        (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) < 0) then
      begin
        Result := -1;
      end
      else if (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine1.FStock) < 0) and  // 仓库不在本区域， 排后面
        (aSAPMrpAREA.FList.IndexOfName(aSAPOPOLine2.FStock) >= 0) then
      begin
        Result := 1;
      end
      else
      begin
        Result := 0;
      end;
    end;
  end;
     
procedure TSAPMrpAreaStockReader.SetOPOList(aSAPOPOReader2: TSAPOPOReader2);
var
  iLine: Integer;
  aSAPOPOLine: TSAPOPOLine;
  iArea: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
begin
  for iLine := 0 to aSAPOPOReader2.Count - 1 do
  begin
    aSAPOPOLine := aSAPOPOReader2.Items[iLine];
    for iArea := 0 to Self.Count - 1 do
    begin
      aSAPMrpAREA := Items[iArea];
      aSAPOPOLine.Tag := aSAPMrpAREA;
      aSAPMrpAREA.FOPOList.AddObject(aSAPOPOLine.FNumber, aSAPOPOLine);
    end;
  end;
  for iArea := 0 to Self.Count - 1 do
  begin
    aSAPMrpAREA := Items[iArea];
    aSAPMrpAREA.FOPOList.Sort;
    QuickSort(aSAPMrpAREA.FOPOList, 0, aSAPMrpAREA.FOPOList.Count - 1, StringListSortCompare_OPO);
//    aSAPMrpAREA.FOPOList.CustomSort( StringListSortCompare );
  end;
end;

procedure TSAPMrpAreaStockReader.SetStock(aSAPStockReader: TSAPStockReader2);
var
  iLine: Integer;
  aSAPStockRecordPtr: PSAPStockRecord;
  aSAPMrpAREA: TSAPMrpAREA;
  idx: Integer;
  p: PSAPStockRecord;
  sArea: string;
begin
  for iLine := 0 to aSAPStockReader.Count - 1 do
  begin
    aSAPStockRecordPtr := aSAPStockReader.Items[iLine];

    if aSAPStockRecordPtr^.snumber = '01.01.1013014' then
    begin
      Sleep(10);
    end;

    if aSAPStockRecordPtr^.sstock = 'AM0M' then
    begin
      Sleep(10);
    end;

    sArea := MrpAreaOfStockNo( aSAPStockRecordPtr^.sstock );

    if sArea = '' then
    begin
      Log('仓库 ' +  aSAPStockRecordPtr^.sstock + ' 没有对应的MRP区域');
      Continue;
    end;

    aSAPMrpAREA := GetSAPMrpAREA(sArea);
    if aSAPMrpAREA = nil then
    begin
      Log( 'mrp area of stock ' + aSAPStockRecordPtr^.sstock + ' not found (TSAPMrpAreaStockReader.SetStock)' );
      //raise Exception.Create('mrp area of stock ' + aSAPStockRecordPtr^.sstock + ' not found');
    end;

    try
      idx := aSAPMrpAREA.FSAPStockSum.FList.IndexOf(aSAPStockRecordPtr^.snumber);
      if idx < 0 then
      begin
        p := New(PSAPStockRecord);
        p^ := aSAPStockRecordPtr^;
        aSAPMrpAREA.FSAPStockSum.FList.AddObject(aSAPStockRecordPtr^.snumber, TObject(p));
      end
      else
      begin
        p := PSAPStockRecord(aSAPMrpAREA.FSAPStockSum.FList.Objects[idx]);
        p^.dqty := p^.dqty + aSAPStockRecordPtr^.dqty;
      end;
    except
      raise Exception.Create('error');
    end;
 
  end;
end;

function TSAPMrpAreaStockReader.AllocStock(const snumber: string; dQty: Double;
  const smrparea: string): Double;
var                 
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  dQtyAlloc: Double;
begin
  Result := 0;

//  if snumber = '01.01.1013014A' then
//  begin
//    Sleep(1);
//  end;

//  aSAPMrpAREA := GetSAPMrpAREA(sMrpArea);
//  if aSAPMrpAREA = nil then
//  begin
//    raise Exception.Create('MRP Area not exists ' + sMrpArea);
//  end;

  // 不分区域， 有库存就分配。 分配了哪个区域、仓库的库存，记录下来 
  for i := 0 to Self.Count - 1 do
  begin
    aSAPMrpAREA := Items[i];
    dQtyAlloc := aSAPMrpAREA.FSAPStockSum.Alloc2(sMrpArea, snumber, dQty);   
    dQty := dQty - dQtyAlloc;
    Result := Result + dQtyAlloc;
  end;

end;     

//只考虑本代工厂库存
function TSAPMrpAreaStockReader.AllocStock2(const snumber: string; dQty: Double;
  const smrparea: string): Double;
var                 
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  dQtyAlloc: Double;
begin
  Result := 0;
 
  aSAPMrpAREA := GetSAPMrpAREA(sMrpArea);
  if aSAPMrpAREA = nil then
  begin
    raise Exception.Create('MRP Area not exists ' + sMrpArea);
  end;

  Result := aSAPMrpAREA.FSAPStockSum.Alloc2(sMrpArea, snumber, dQty);

end;

function TSAPMrpAreaStockReader.AllocStock_area(const snumber, sname: string;
  dt: TDateTime; dQty: Double; const smrparea: string; slTransfer: TList): Double;
var
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  aTransferRecordPtr: PTransferRecord;
  dAlloc: Double;
begin
  Result := 0;

  for i := 0 to Self.Count - 1 do
  begin
    if self.Items[i].sAreaNo = sMrpArea then  // 本区域已经分配过了，无需再计算
    begin
      Continue;
    end;

    aSAPMrpAREA := self.Items[i];

    dAlloc := aSAPMrpAREA.FSAPStockSum.AllocStock_area(snumber, dQty);
    if dAlloc > 0 then
    begin
      aTransferRecordPtr := New(PTransferRecord);
      aTransferRecordPtr^.snumber := snumber;
      aTransferRecordPtr^.sname := sname;
      aTransferRecordPtr^.dt := dt;
      aTransferRecordPtr^.sfrom := aSAPMrpAREA.sAreaNo;
      aTransferRecordPtr^.sto := smrparea;
      aTransferRecordPtr^.dqty := dAlloc;
      slTransfer.Add(aTransferRecordPtr);
      Result := Result + dAlloc;
    end;
  end; 
end;

function TSAPMrpAreaStockReader.AccDemand(const snumber: string; dQty: Double;
  const smrparea: string): Double;
var
  aSAPMrpAREA: TSAPMrpAREA;
begin
  Result := 0;
 
  aSAPMrpAREA := GetSAPMrpAREA(sMrpArea);
  if aSAPMrpAREA = nil then
  begin
    raise Exception.Create('MRP Area not exists ' + sMrpArea);
  end;
 
  Result := aSAPMrpAREA.FSAPStockSum.AccDemand(snumber, dQty); 
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

function TSAPMrpAreaStockReader.Alloc(slNumber: TStringList; dt: TDateTime;
  dQty: Double; const sMrpArea: string): Double;
var
  aSAPMrpAREA: TSAPMrpAREA;
begin
  Result := 0;


  aSAPMrpAREA := GetSAPMrpAREA(sMrpArea);
  if aSAPMrpAREA = nil then
  begin
    raise Exception.Create('MRP Area not exists ' + sMrpArea);
  end;

  Result := aSAPMrpAREA.Alloc(slNumber, dt, dQty); 
end;

function TSAPMrpAreaStockReader.GetSAPMrpAREA(const sMrpArea: string): TSAPMrpAREA;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Self.Count - 1 do
  begin
    if self.Items[i].sAreaNo = sMrpArea then
    begin
      Result := self.Items[i];
      Break;
    end;
  end;
end;
                          
procedure TSAPMrpAreaStockReader.GetOPOs(slNumber: TStringList; lst: TList;
  const sMrpArea: string);
var
  aSAPMrpAREA: TSAPMrpAREA;
begin
  aSAPMrpAREA := GetSAPMrpAREA(sMrpArea);
  if aSAPMrpAREA = nil then
  begin
    raise Exception.Create('TSAPMrpAreaStockReader.GetOPOs  mrp area not exists ' + sMrpArea);
  end;
  aSAPMrpAREA.GetOPOs(slNumber, lst);
end;

function TSAPMrpAreaStockReader.MrpAreaNo2Name(const sMrpAreaNo: string): string;
var
  aSAPMrpAREA: TSAPMrpAREA;
begin
  Result := '';
  
  if sMrpAreaNo = '' then
  begin
    Exit;
  end;

  aSAPMrpAREA := GetSAPMrpAREA(sMrpAreaNo);
  //  if aSAPMrpAREA = nil then
  //  begin
  //    raise Exception.Create('TSAPMrpAreaStockReader.MrpAreaNo2Name  mrp area not exists ' + sMrpAreaNo);
  //  end;
  Result := aSAPMrpAREA.sAreaName;

end;

function TSAPMrpAreaStockReader.MrpAreaOfStockNo(const sstock: string): string;
var
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
begin
  Result := '';
  for i := 0 to self.Count - 1 do
  begin
    aSAPMrpAREA := Items[i];
    if aSAPMrpAREA.FList.IndexOfName(sstock) >= 0 then
    begin
      Result := aSAPMrpAREA.sAreaNo;
      Break;
    end;
  end;
end;

function TSAPMrpAreaStockReader.GetStockSum(const snumber: string): Double;
var
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
begin
  Result := 0;
  for i := 0 to self.Count - 1 do
  begin
    aSAPMrpAREA := Items[i];
    Result := Result + aSAPMrpAREA.FSAPStockSum.GetStock(snumber);
  end;
end;      

function TSAPMrpAreaStockReader.GetOPOSum(const snumber: string): Double;
var
  i: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  iNumber: Integer;
  aSAPOPOLine: TSAPOPOLine;
begin
  Result := 0;
  for i := 0 to self.Count - 1 do
  begin
    aSAPMrpAREA := Items[i];
    for iNumber := 0 to aSAPMrpAREA.FOPOList.Count - 1 do
    begin
      if aSAPMrpAREA.FOPOList[iNumber] = snumber then
      begin
        aSAPOPOLine := TSAPOPOLine(aSAPMrpAREA.FOPOList.Objects[iNumber]);
        Result := Result + aSAPOPOLine.FQty;
      end;
    end;
    Break;
  end;
end;  

function TSAPMrpAreaStockReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMrpAreaStockReader.GetItems(i: Integer): TSAPMrpAREA;
begin
  Result := TSAPMrpAREA(FList.Objects[i]);
end;
  
procedure TSAPMrpAreaStockReader.Log(const str: string);
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

procedure TSAPMrpAreaStockReader.Open; 
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  sAreaNo: string;
  sMrp: string;   
  aSAPMrpAREA: TSAPMrpAREA;
  sStockNo: string;
  sStockName: string;

  icolFac: Integer;
  icolAreaNo: Integer;
  icolAreaName: Integer;
  icolStockNo: Integer;
  icolStockName: Integer;
  icolMrp: Integer;

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
        if (stitle = 'MRP区域MRP区域描述仓库仓库描述')
          or (stitle = 'MRP区域MRP区域描述仓库仓库描述是否参于MRP计算') then
        begin
          icolFac := -1;
          icolAreaNo := 1;
          icolAreaName := 2;
          icolStockNo := 3;
          icolStockName := 4;    
          icolMrp := -1;
        end
        else if stitle = '工厂库位仓储地点描述MMRP 范围' then
        begin
          icolFac := 1;
          icolAreaNo := 5;
          icolAreaName := -1;
          icolStockNo := 2;
          icolStockName := 3;
          icolMrp := -1;
        end
        else if stitle = '工厂MRP区域MRP区域描述仓库仓储描述是否参与MRP计算' then
        begin              
          icolFac := 1;
          icolAreaNo := 2;
          icolAreaName := 3;
          icolStockNo := 4;
          icolStockName := 5; 
          icolMrp := 6;
        end
        else
        begin

          Log(sSheet +'  不是  MRP区域  格式');
          Continue;
        end;

        FReadOk := True;

        irow := 2;
        sStockNo := ExcelApp.Cells[irow, icolStockNo].Value;
        while sStockNo <> '' do
        begin                                            
          sAreaNo := ExcelApp.Cells[irow, icolAreaNo].Value;
          if sAreaNo = '' then
          begin             
            irow := irow + 1;
            sStockNo := ExcelApp.Cells[irow, icolStockNo].Value;
            Continue;
          end;

          if icolMrp <> -1 then
          begin
            sMrp := ExcelApp.Cells[irow, icolMrp].Value;
            if sMrp <> 'Y' then
            begin
              irow := irow + 1;
              sStockNo := ExcelApp.Cells[irow, icolStockNo].Value;
              Continue;
            end;
          end;

          aSAPMrpAREA := GetSAPMrpAREA( sAreaNo );
          if aSAPMrpAREA = nil then
          begin
            aSAPMrpAREA := TSAPMrpAREA.Create;
            aSAPMrpAREA.sAreaNo := sAreaNo;
            if icolAreaName <> - 1 then
            begin
              aSAPMrpAREA.sAreaName := ExcelApp.Cells[irow, icolAreaName].Value;
            end
            else aSAPMrpAREA.sAreaName := '';      
            FList.AddObject(sAreaNo, aSAPMrpAREA);
          end;
  
          sStockName := ExcelApp.Cells[irow, icolStockName].Value;

          aSAPMrpAREA.FList.Add(sStockNo + '=' + sStockName);
 

          irow := irow + 1;
          sStockNo := ExcelApp.Cells[irow, icolStockNo].Value;
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
