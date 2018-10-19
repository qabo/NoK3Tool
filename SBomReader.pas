unit SBomReader;

interface

uses
  Classes, SysUtils, ComObj, CommUtils, KeyICItemSupplyReader;

type
  TSBom = class;
  
  TBomGroupChild = packed record
    dStockSemi: Double;
    spriority: string;
    supp: TKeyICItemSupplyLine;
  end;
  PBomGroupChild = ^TBomGroupChild;

  TSBomChild = class
  private
    procedure SortPriority;
  public
    FGroup: string;
    FLT: Integer;
    dUsage: Double;
    
    FStockSemi: Double;
    FList: TStringList;
    FParent: TSBom;
    constructor Create(aSBom: TSBom);
    destructor Destroy; override;    
    procedure Clear;

    function GetQtyAvail(dt: TDateTime): Double;    
    function GetQtyAvail2(dt: TDateTime): Double;
    procedure AllocQty(const snumber_cp: string; dt: TDateTime; dqty: Double);
    procedure AllocQty2( dt: TDateTime; dqty: Double );
    function GetAvailStockSemi(const sNumber: string): Double;
  end;
  
  TSBom = class
  private
    procedure Clear;   
    procedure SortPriority;
  public
    FNumber: string;
    FName: string;
    FList: TStringList;
    constructor Create;
    destructor Destroy; override;
    function GetAvailStockSemi(const sNumber: string): Double;
  end;

  TSBomReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    procedure SortPriority;
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetBom(const sNumber: string): TSBom;
    function GetAvailStockSemi(const sNumber: string): Double;
  end;

implementation
         
{ TSBomChild }

constructor TSBomChild.Create(aSBom: TSBom);
begin
  FList := TStringList.Create;
  FParent := aSBom;
end;

destructor TSBomChild.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSBomChild.Clear;
var
  i: Integer;
  p: PBomGroupChild;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomGroupChild(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function StringListSortCompare_BomGroupChild(List: TStringList; Index1, Index2: Integer): Integer;
var
  p1, p2: PBomGroupChild;
begin
  p1 := PBomGroupChild(List.Objects[Index1]);
  p2 := PBomGroupChild(List.Objects[Index2]);
  if p1^.spriority > p2^.spriority then
    Result := 1
  else if p1^.spriority = p2^.spriority then
    Result := 0
  else // <
    Result := -1;
end;

procedure TSBomChild.SortPriority;
begin
  if FList.Count <= 1 then Exit; // 没有替代率，无需排序
  FList.CustomSort(StringListSortCompare_BomGroupChild);
end;

function TSBomChild.GetQtyAvail(dt: TDateTime): Double;
var
  igroup: Integer;
  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
  dqty_a: Double;
begin
  Result := 0;
  for igroup := 0 to FList.Count - 1 do
  begin
    aKeyICItemSupplyLine := PBomGroupChild(FList.Objects[igroup])^.supp;
    if aKeyICItemSupplyLine = nil then
    begin
      Continue;
    end;

    // 取可供应数量
    aKeyICItemSupplyLine.GetQtyAvailx(dt, dqty_a);
    Result := Result + dqty_a;
  end;
end;

function TSBomChild.GetQtyAvail2(dt: TDateTime): Double;     
var
  igroup: Integer;
  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
  dqty_a: Double;
begin       
  Result := 0;
  for igroup := 0 to FList.Count - 1 do
  begin
    aKeyICItemSupplyLine := PBomGroupChild(FList.Objects[igroup])^.supp;
    if aKeyICItemSupplyLine = nil then
    begin
      Continue;
    end;

    // 取可供应数量
    aKeyICItemSupplyLine.GetQtyAvail2x(dt, dqty_a);
    
    Result := Result + dqty_a;
  end;
end;

procedure TSBomChild.AllocQty(const snumber_cp: string; dt: TDateTime; dqty: Double);
var
  igroup: Integer;
  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
  dqty_a: Double;
  d: Double;
begin
  d := dqty;

//  if PBomGroupChild(FList.Objects[0])^.supp.sNumber99 = '01.01.117397602A' then
//    if (DoubleE(dt , 43093) )and (DoubleE(dqty , 301.001)) then
//    begin
//        asm int 3 end;
//    end;

  for igroup := 0 to FList.Count - 1 do
  begin
    aKeyICItemSupplyLine := PBomGroupChild(FList.Objects[igroup])^.supp;
    if aKeyICItemSupplyLine = nil then
    begin
      Continue;
    end;

    if Copy( Self.FParent.FNumber, 3, 4) <> '.55.' then
    begin
      sleep(1);
    end;
    
    if aKeyICItemSupplyLine.sNumber99 = '01.01.1013089' then
    begin
      savelogtoexe( Format('snumber_cp: %s, dqty: %0.0f, d: %0.0f', [snumber_cp, dqty, d]) );
    end;

    // 取可供应数量
    aKeyICItemSupplyLine.GetQtyAvailx(dt, dqty_a);

    if DoubleLE( d , dqty_a) then
    begin
      if not aKeyICItemSupplyLine.AllocQtyx(dt,  d ) then
      begin
//        asm int 3 end;
//        aKeyICItemSupplyLine.GetQtyAvail(dt, dqty_a);
//
//        aKeyICItemSupplyLine.AllocQty( d , dt)
      end;
      
      d := 0;
    end
    else
    begin
      aKeyICItemSupplyLine.AllocQtyx(dt, dqty_a );
      d := d - dqty_a;
    end;

    if DoubleLE(d , 0) then Break;
  end;
end;

procedure TSBomChild.AllocQty2( dt: TDateTime; dqty: Double );
var
  igroup: Integer;
  aKeyICItemSupplyLine: TKeyICItemSupplyLine;
  dqty_a: Double;
  d: Double;
begin
  d := dqty;

  for igroup := 0 to FList.Count - 1 do
  begin
    aKeyICItemSupplyLine := PBomGroupChild(FList.Objects[igroup])^.supp;
    if aKeyICItemSupplyLine = nil then
    begin
      Continue;
    end;

    // 取可供应数量
    aKeyICItemSupplyLine.GetQtyAvail2x(dt, dqty_a);

    if DoubleLE( d , dqty_a) then
    begin
      aKeyICItemSupplyLine.AllocQty2x(dt,  d );
      d := 0;
    end
    else
    begin
      aKeyICItemSupplyLine.AllocQty2x(dt,  dqty_a);
      d := d - dqty_a;
    end;

    if DoubleLE( d , 0) then Break;
  end;
end;

function TSBomChild.GetAvailStockSemi(const sNumber: string): Double;
var
  i: Integer;
  p: PBomGroupChild;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    if FList[i] = sNumber then
    begin
      p := PBomGroupChild(FList.Objects[i]);
      Result := Result + p^.dStockSemi;
    end;
  end;
end;

{ TSBom }

constructor TSBom.Create;
begin
  FList := TStringList.Create;
end;

destructor TSBom.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TSBom.GetAvailStockSemi(const sNumber: string): Double;
var
  i: Integer;
  abc: TSBomChild;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    abc := TSBomChild(FList.Objects[i]);
    Result := Result + abc.GetAvailStockSemi(sNumber);
  end;
end;

procedure TSBom.Clear;
var
  i: Integer;
  abc: TSBomChild;
begin
  for i := 0 to FList.Count - 1 do
  begin
    abc := TSBomChild(FList.Objects[i]);
    abc.Free;
  end;
  FList.Clear;
end;

procedure TSBom.SortPriority;
var
  i: Integer;
  abc: TSBomChild;
begin
  for i := 0 to FList.Count - 1 do
  begin
    abc := TSBomChild(FList.Objects[i]);
    abc.SortPriority;
  end;
end;

{ TSBomReader }

constructor TSBomReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
  SortPriority;
end;

destructor TSBomReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSBomReader.Clear;
var
  i: Integer;
  aSBom: TSBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSBom := TSBom(FList.Objects[i]);
    aSBom.Free;
  end;
  FList.Clear;
end;

procedure TSBomReader.Log(const str: string);
begin

end;

procedure TSBomReader.SortPriority;
var
  i: Integer;
  aSBom: TSBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSBom := TSBom(FList.Objects[i]);
    aSBom.SortPriority;
  end;
end;

function TSBomReader.GetBom(const sNumber: string): TSBom;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    if FList[i] = sNumber then
    begin
      Result := TSBom(FList.Objects[i]);
      Break;
    end;
  end;
end;

function TSBomReader.GetAvailStockSemi(const sNumber: string): Double;
var
  i: Integer;
  aSBom: TSBom;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aSBom := TSBom(FList.Objects[i]);
    Result := Result + aSBom.GetAvailStockSemi(sNumber);
  end;
end;  

procedure TSBomReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9, stitle10: string;
  stitle: string;
  irow: Integer;
  irow1: Integer;
  snumber: string;
  aSBom: TSBom;
  aSBomChild: TSBomChild;
  sgroup, sgroup0: string;
  snumber_child: string;
  dStockSemi: Double;
  aBomGroupChildPtr: PBomGroupChild;
begin
  Clear;


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
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle8 := ExcelApp.Cells[irow, 8].Value;                                                                   
        stitle9 := ExcelApp.Cells[irow, 9].Value;
        stitle10 := ExcelApp.Cells[irow, 10].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8 + stitle9 + stitle10;
        if stitle <> '产品编码产品名称物料编码物料名称提前期替代组半成品库存库存用量优先级' then
        begin
          Log(sSheet +'  不是简易BOM格式');
          Continue;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          aSBom := TSBom.Create;
          FList.AddObject(snumber, aSBom);

          aSBom.FNumber := snumber;
          aSBom.FName := ExcelApp.Cells[irow, 2].Value;

          aSBomChild := nil;
          sgroup0 := '';
          irow1 := irow;
          while IsCellMerged(ExcelApp, irow1, 1, irow, 1) do
          begin                                       
            snumber_child := ExcelApp.Cells[irow, 3].Value;
            sgroup := ExcelApp.Cells[irow, 6].Value;
            dStockSemi := ExcelApp.Cells[irow, 7].Value;

            if (sgroup = '') or ( sgroup <> sgroup0 ) then
            begin
              aSBomChild := TSBomChild.Create(aSBom);
              aSBomChild.FGroup := sgroup;
              aSBomChild.FLT := ExcelApp.Cells[irow, 5].Value;
              aSBomChild.dUsage := ExcelApp.Cells[irow, 9].Value;
              aSBom.FList.AddObject(aSBomChild.FGroup, aSBomChild);
            end;
            aBomGroupChildPtr := New(PBomGroupChild);
            aBomGroupChildPtr^.dStockSemi := dStockSemi;
            aBomGroupChildPtr^.sPriority := ExcelApp.Cells[irow, 10].Value;
            aBomGroupChildPtr^.supp := nil;
            aSBomChild.FList.AddObject(snumber_child, TObject(aBomGroupChildPtr));
            sgroup0 := sgroup;
            
            irow := irow + 1;
          end;   
          snumber := ExcelApp.Cells[irow, 1].Value;
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
