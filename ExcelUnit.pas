unit ExcelUnit;

interface

uses
  Windows, Graphics, Classes, Variants, ComObj, SysUtils, CommVars, CommUtils,
  ADODB, ExcelConsts, DBGridEh, DBClient;

type
  TPlanRecord = packed record
    fdate: TDateTime;
    fqty: Double;
  end;
  PPlanRecord = ^TPlanRecord;

  TExcelBase = class
  protected
    FFile: string;
    ExcelApp, WorkBook: Variant;
    function LoadData: Boolean; virtual; abstract;
  public
    ErrMsg: string;
    constructor Create;
    destructor Destroy; override;
    function Open(const sfile: string): Boolean;
    procedure Close(); 
  end;

  TRoutPlan = class
  private
    frout: string;
    fcolor: string;
    FPlanList: TList;
    procedure ClearPlanList;
    function GetCount: Integer;
    function GetItems(index: Integer): PPlanRecord;
  public
    constructor Create(const srout, scolor: string);
    destructor Destroy; override;
    procedure LoadData(ExcelApp, WorkBook: Variant; irow: Integer);
    function IndexOf(fdate: TDateTime): Integer;
    property Count: Integer read GetCount;
    property Items[index: Integer]: PPlanRecord read GetItems;
    property Rout: string read frout;
    property Color: string read fcolor;
  end;

  TExcelPPlan = class(TExcelBase)    
  private
    FPPlanList: TList;
    procedure ClearPPlanList;
    function GetRoutCount: Integer;
    function GetDateCount: Integer;
    function GetRouts(index: Integer): TRoutPlan;
  protected
    function LoadData: Boolean; override;
  public
    constructor Create;
    destructor Destroy; override;
    function IndexOfDate(fdate: TDateTime): Integer;
    property RoutCount: Integer read GetRoutCount;
    property DateCount: Integer read GetDateCount;
    property Routs[index: Integer]: TRoutPlan read GetRouts;
  end;
       
  TBomItemModels = class
  private
    fnumber: string;
    fName: string;
    fPO: string;
    fPE: string;
    fSupplier: string;
    fLT: string;
    fMC: string;
    fDepts: TStringList;
    fUsages: TStringList;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddUsage(const sModel: string; dUsage: Double; const sDept: string);
    function Clone: TBomItemModels;
  end;

  TBomItem = class
  private
    FItems: TList;
    procedure ClearItems;
    function GetCount: Integer;
    function GetItems(i: Integer): TBomItemModels;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddItem(const sNumber, sName, sPO, sPE, sSupplier, sLT, sMC,
      sDept: string; const sModel: string; dUsage: Double); overload;
    procedure AddItem(bim: TBomItemModels); overload;
    function IndexOf(const snumber: string): Integer;
    function Clone: TBomItem;
    procedure WriteData(const sType: string; aModels: TStringList;
      ExcelApp: Variant; FColUsage, FColStockML, FColStockFOX,
      FColStockWW, FColWin, FColOrder: Integer; var irow: Integer);
    property Count: Integer read GetCount;
    property Items[i: Integer]: TBomItemModels read GetItems;
  end;
       
  TExcelBom = class(TExcelBase)
  private
    FModel: string;
    FCPNumber: string;       
    FType: string;
    FICMOQtyFOX: Integer;
    FICMOQtyML: Integer;
    FBomItems: TList;
    FMergedModels: TStringList;
    FMergedNumbers: TStringList;
    FColUsage: Integer;
    FColStockML, FColStockFOX, FColStockWW, FColWin, FColOrder: Integer;
    function IndexOfCol(irow: Integer; const scolname: string): Integer;
    function GetCellValueAsStr(irow, icol: Integer): string;
    function GetCellValueAsDouble(irow, icol: Integer): Double;
    procedure ClearBomItems;
    procedure MergeBomItem(bi: TBomItem);
    procedure MergeBomItemTo(bi: TBomItem; idx: Integer);
    function GetCount: Integer;
    function GetItems(i: Integer): TBomItem;
  protected
    function LoadData: Boolean; override;
    procedure WriteData;
  public
    constructor Create(const sModel, sType: string; iICMOQtyFOX, iICMOQtyML: Integer);
    destructor Destroy; override;
    procedure MergeBom(aExcelBom: TExcelBom);
    procedure SaveAs(const sfile: string);
    function IndexByNumber(const snumber: string): Integer;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TBomItem read GetItems;
    property Model: string read FModel;
  end;

  TExcelDBGridEH = class(TExcelBase)
  private
    procedure WriteData(ClientDataSet1: TClientDataSet);
  protected
  public
    constructor Create;
    destructor Destroy; override;
    procedure SaveAs(ClientDataSet1: TClientDataSet; const sfile: string);
  end;

implementation
     
const
  CSColNumber = 1;
  CSColName = 2;  

{ TExcelBase }

constructor TExcelBase.Create;
begin
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
end;

destructor TExcelBase.Destroy;
begin               
  ExcelApp.Visible := True;
  ExcelApp.Quit; 
  inherited;
end;

function TExcelBase.Open(const sfile: string): Boolean;
begin
  Result := False;              
  WorkBook := null;
  FFile := sfile;
  if not FileExists(sFile) then Exit;
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    Result := LoadData;
  finally

  end;
end;
      
procedure TExcelBase.Close();
begin
  if VarIsNull(WorkBook) then Exit;
  ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
  WorkBook.Close;
end;

{ TExcelPPlan }
         
constructor TExcelPPlan.Create;
begin
  FPPlanList := TList.Create;
end;

destructor TExcelPPlan.Destroy;
begin
  ClearPPlanList;
  FPPlanList.Free;
  inherited;
end;

function TExcelPPlan.IndexOfDate(fdate: TDateTime): Integer;
var
  aPPlan: TRoutPlan;
begin
  if FPPlanList.Count > 0 then
  begin
    aPPlan := TRoutPlan(FPPlanList[0]);
    Result := aPPlan.IndexOf(fdate);
  end
  else
  begin
    Result := -1;
  end;
end;

procedure TExcelPPlan.ClearPPlanList;
var
  i: Integer;
  aPPlan: TRoutPlan;
begin
  for i := 0 to FPPlanList.Count - 1 do
  begin
    aPPlan := TRoutPlan(FPPlanList[i]);
    aPPlan.Free;
  end;
  FPPlanList.Clear;
end;
  
function TExcelPPlan.LoadData: Boolean;
const
  CIColRout = 1;
  CIColColor = 2;
var
  irow: Integer;
  frout, fcolor: string;
  aPPlan: TRoutPlan;
begin
  ExcelApp.WorkSheets['生产计划'].Activate;

  irow := 2;
  frout := ExcelApp.Cells[irow, CIColRout].Value;
  fcolor := ExcelApp.Cells[irow, CIColColor].Value;
  while frout <> EmptyStr do
  begin
    aPPlan := TRoutPlan.Create(frout, fcolor);
    aPPlan.LoadData(ExcelApp, WorkBook, irow);
    FPPlanList.Add(aPPlan);

    irow := irow + 1;
    frout := ExcelApp.Cells[irow, 1].Value;
  end;
  Result := True;
end;
    
function TExcelPPlan.GetRoutCount: Integer;
begin
  Result := FPPlanList.Count;
end;

function TExcelPPlan.GetDateCount: Integer;
var
  aPPlan: TRoutPlan;
begin
  if FPPlanList.Count > 0 then
  begin
    aPPlan := TRoutPlan(FPPlanList[0]);
    Result := aPPlan.Count;
  end
  else
  begin
    Result := 0;
  end;
end;

function TExcelPPlan.GetRouts(index: Integer): TRoutPlan;
begin
  Result := TRoutPlan(FPPlanList[index]);
end;

{ TRoutPlan }

constructor TRoutPlan.Create(const srout, scolor: string);
begin
  frout := srout;
  fcolor := scolor;
  FPlanList := TList.Create;
end;

destructor TRoutPlan.Destroy;
begin
  ClearPlanList;
  FPlanList.Free;
  inherited;
end;

procedure TRoutPlan.LoadData(ExcelApp, WorkBook: Variant; irow: Integer);
const
  CIRowDate = 1;
var
  icol: Integer;  
  pr: PPlanRecord;
  fdate: TDateTime;
begin
  ClearPlanList;
  icol := 3;
  fdate := ExcelApp.Cells[CIRowDate, icol].Value;
  while fdate <> 0 do
  begin
    pr := New(PPlanRecord);
    pr^.fdate := fdate;
    pr^.fqty := ExcelApp.Cells[irow, icol].Value;
    FPlanList.Add(pr);
  end;
end;

function TRoutPlan.IndexOf(fdate: TDateTime): Integer;
var
  i: Integer;
  pr: PPlanRecord;
begin
  Result := -1;
  for i := 0 to FPlanList.Count - 1 do
  begin    
    pr := PPlanRecord(FPlanList[i]);
    if pr^.fdate = fdate then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function TRoutPlan.GetCount: Integer;
begin
  Result := FPlanList.Count;
end;

procedure TRoutPlan.ClearPlanList;
var
  i: Integer;
  pr: PPlanRecord;
begin
  for i := 0 to FPlanList.Count - 1 do
  begin
    pr := PPlanRecord(FPlanList[i]);
    Dispose(pr);
  end;
  FPlanList.Clear;
end;

function TRoutPlan.GetItems(index: Integer): PPlanRecord;
begin
  Result := PPlanRecord(FPlanList[index]);
end;
             
{ TBomItemModels }

constructor TBomItemModels.Create;
begin
  inherited;
  fUsages := TStringList.Create;
  fDepts := TStringList.Create;
end;

destructor TBomItemModels.Destroy;
begin
  fUsages.Free;
  fDepts.Free;
  inherited;
end;

procedure TBomItemModels.AddUsage(const sModel: string; dUsage: Double; const sDept: string);
begin
  //if dUsage <> -9999 then
  fUsages.Values[sModel] := Format('%0.6f', [dUsage]);
  if sDept = '' then
  begin
    fDepts.Values[sModel] := '-';
  end
  else
  begin
    fDepts.Values[sModel] := sDept;
  end;
end;

function TBomItemModels.Clone: TBomItemModels;
var
  bim: TBomItemModels;
begin
  bim := TBomItemModels.Create;
  bim.fnumber := Self.fnumber;
  bim.fName := Self.fName;
  bim.fPO := Self.fPO;
  bim.fPE := Self.fPE;
  bim.fSupplier := Self.fSupplier;
  bim.fLT := Self.fLT;
  bim.fMC := Self.fMC;
  bim.fDepts := TStringList.Create;
  bim.fDepts.Text := Self.fDepts.Text;
  bim.fUsages := TStringList.Create;
  bim.fUsages.Text := Self.fUsages.Text;
  Result := bim;
end;

{ TBomItem }

constructor TBomItem.Create;
begin
  FItems := TList.Create;
end;

destructor TBomItem.Destroy;
begin
  ClearItems;
  FItems.Free;
  inherited;
end;

procedure TBomItem.ClearItems;
var
  i: Integer;
  pbir: TBomItemModels;
begin
  for i := 0 to FItems.Count - 1 do
  begin
    pbir := TBomItemModels(FItems[i]);
    pbir.Free;
  end;
  FItems.Clear;
end;

function TBomItem.GetCount: Integer;
begin
  Result := FItems.Count;
end;

function TBomItem.GetItems(i: Integer): TBomItemModels;
begin
  Result := TBomItemModels(FItems[i]);
end;

procedure TBomItem.AddItem(const sNumber, sName, sPO, sPE, sSupplier, sLT, sMC,
  sDept: string; const sModel: string; dUsage: Double);
var
  pbir: TBomItemModels;
begin
  pbir := TBomItemModels.Create;
  pbir.fnumber := sNumber;
  pbir.fName := sName;
  pbir.fPO := sPO;
  pbir.fPE := sPE;
  pbir.fSupplier := sSupplier;
  pbir.fLT := sLT;
  pbir.fMC := sMC;
  pbir.AddUsage(sModel, dUsage, sDept);
  FItems.Add(pbir);
end;

procedure TBomItem.AddItem(bim: TBomItemModels);
begin
  FItems.Add(bim);
end;

function TBomItem.IndexOf(const snumber: string): Integer;
var
  i: Integer;
  pbir: TBomItemModels;
begin      
  Result := -1;
  for i := 0 to Self.Count - 1 do
  begin
    pbir := Self.Items[i];
    if pbir.fnumber = snumber then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function TBomItem.Clone: TBomItem;
var
  bi: TBomItem;
  i: Integer;
  bim: TBomItemModels;
begin
  bi := TBomItem.Create;
  for i := 0 to Self.Count - 1 do
  begin
    bim := Self.Items[i];
    bi.AddItem(bim.Clone);
  end;
  Result := bi;
end;

procedure TBomItem.WriteData(const sType: string; aModels: TStringList;
  ExcelApp: Variant; FColUsage, FColStockML, FColStockFOX,
  FColStockWW, FColWin, FColOrder: Integer; var irow: Integer);
var
  i, j: Integer;
  icol: Integer;
  irow1, irow2: Integer;
  bim, bim2: TBomItemModels;
  sModel: string;
  im: Integer;
  iModel: Integer;
  svalue1, svalue2: string;
  icolNeedSumML, icolNeedSumFOX: Integer;
  idxJ: Integer;    
begin
  for i := 0 to Self.Count - 1 do
  begin
    bim := Self.Items[i];
    icol := CSColNumber;
    ExcelApp.Cells[irow + i, icol].Value := bim.fNumber;
    icol := CSColName;
    ExcelApp.Cells[irow + i, icol].Value := bim.fName;

    icol := icol + 1;
    if i = 0 then
    begin
      for iModel := 0 to aModels.Count - 1 do
      begin
        for j := 0 to Self.Count - 1 do
        begin
          bim2 := Self.Items[j];
          idxJ := bim2.fUsages.IndexOfName(aModels[iModel]);
          if (idxJ >= 0) and (bim2.fUsages.Values[aModels[iModel]] <> '0.000000')
            and (bim2.fUsages.Values[aModels[iModel]] <> '-9999.000000') then
          begin
            ExcelApp.Cells[irow, icol + iModel].Value := bim2.fUsages.Values[aModels[iModel]];
            Break;
          end;
        end;
      end;
    end;
    icol := icol + aModels.Count;
    
    sModel := bim.fUsages.Names[0];
    for im := 1 to bim.fUsages.Count - 1 do
    begin
      sModel := sModel + '/' + bim.fUsages.Names[im];
    end;
    ExcelApp.Cells[irow + i, icol].Value := sModel;

    icol := icol + 1 + (aModels.Count + 1) * 2 + 9 + 6;
    if sType='S' then
    begin
      icol := icol + aModels.Count * 2;
    end;
    
    ExcelApp.Cells[irow + i, icol].Value := bim.fPO;
    
    icol := icol + 1;
    ExcelApp.Cells[irow + i, icol].Value := bim.fPE;

    icol := icol + 1;
    ExcelApp.Cells[irow + i, icol].Value := bim.fSupplier;
    
    icol := icol + 1;
    ExcelApp.Cells[irow + i, icol].Value := bim.fLT;
    
    icol := icol + 1;
    ExcelApp.Cells[irow + i, icol].Value := bim.fMC;
    
    icol := icol + 1;
    for iModel := 0 to aModels.Count - 1 do
    begin
      ExcelApp.Cells[irow + i, icol + iModel].Value := bim.fDepts.Values[aModels[iModel]];
    end;

    ExcelApp.Cells[irow + i, FColStockML].Value := 0;
    ExcelApp.Cells[irow + i, FColStockFOX].Value := 0;
    ExcelApp.Cells[irow + i, FColStockWW].Value := 0;
    ExcelApp.Cells[irow + i, FColWin].Value := 0;
    ExcelApp.Cells[irow + i, FColWin + 1].Value := 0;
    ExcelApp.Cells[irow + i, FColWin + 2].Value := 0;
    ExcelApp.Cells[irow + i, FColOrder].Value := 0;

    ExcelApp.Cells[irow + i, FColStockML].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Cells[irow + i, FColStockFOX].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Cells[irow + i, FColStockWW].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Cells[irow + i, FColWin].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';      
    ExcelApp.Cells[irow + i, FColWin + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Cells[irow + i, FColWin + 2].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';   
    ExcelApp.Cells[irow + i, FColOrder].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

  end;       
  irow1 := irow;
  irow2 := irow + Self.Count - 1;

  icol := CSColName + 1;
  for iModel := 0 to aModels.Count - 1 do
  begin
    ExcelApp.Range[ExcelApp.Cells[irow1, icol + iModel], ExcelApp.Cells[irow2, icol + iModel]].MergeCells := true;
  end;
  icol := icol + aModels.Count;
     
  icol := icol + 1; //跳过 所属产品型号
  for iModel := 0 to aModels.Count - 1 do
  begin
    //生产需求数量 魅力
    ExcelApp.Cells[irow1, icol + iModel * 2].Value := '=' + GetRef(3 + iModel) + IntToStr(irow1) + '*$' + GetRef(icol + iModel * 2) + '$3';   
    ExcelApp.Cells[irow1, icol + iModel * 2].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Range[ExcelApp.Cells[irow1, icol + iModel * 2], ExcelApp.Cells[irow2, icol + iModel * 2]].MergeCells := true;
    //生产需求数量 富士康
    ExcelApp.Cells[irow1, icol + iModel * 2 + 1].Value := '=' + GetRef(3 + iModel) + IntToStr(irow1) + '*$' + GetRef(icol + iModel * 2 + 1) + '$3';
    ExcelApp.Cells[irow1, icol + iModel * 2 + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
    ExcelApp.Range[ExcelApp.Cells[irow1, icol + iModel * 2 + 1], ExcelApp.Cells[irow2, icol + iModel * 2 + 1]].MergeCells := true;
  end;

  svalue1 := '=' + GetRef(icol) + IntToStr(irow1);
  svalue2 := '=' + GetRef(icol + 1) + IntToStr(irow1);
  for iModel := 1 to aModels.Count - 1 do
  begin
    svalue1 := svalue1 + '+' + GetRef(icol + iModel * 2) + IntToStr(irow1);
    svalue2 := svalue2 + '+' + GetRef(icol + iModel * 2 + 1) + IntToStr(irow1);
  end;
  
  icol := icol + aModels.Count * 2; 
  //总生产需求数量 魅力
  icolNeedSumML := icol;
  ExcelApp.Cells[irow1, icolNeedSumML].Value := svalue1;
  ExcelApp.Cells[irow1, icolNeedSumML].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icolNeedSumML], ExcelApp.Cells[irow2, icolNeedSumML]].MergeCells := true;
  //总生产需求数量 富士康   
  icolNeedSumFOX := icol + 1;
  ExcelApp.Cells[irow1, icolNeedSumFOX].Value := svalue2;
  ExcelApp.Cells[irow1, icolNeedSumFOX].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icolNeedSumFOX], ExcelApp.Cells[irow2, icolNeedSumFOX]].MergeCells := true;

  icol := icol + 2;

  //魅力料况  魅力库存
  //魅力料况  总库存
  svalue1 := '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
  ExcelApp.Cells[irow1, icol + 1].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 1], ExcelApp.Cells[irow2, icol + 1]].MergeCells := true;
  //魅力料况  Yes/No
  svalue1 := '=IF(' + GetRef(icol + 3) + IntToStr(irow1) + '<0,"No","Yes")';
  ExcelApp.Cells[irow1, icol + 2].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 2].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 2], ExcelApp.Cells[irow2, icol + 2]].MergeCells := true;
  //魅力料况  差异数量
  svalue1 := '=' + GetRef(icol + 1) + IntToStr(irow1) + '-' + GetRef(icolNeedSumML) + IntToStr(irow1);
  ExcelApp.Cells[irow1, icol + 3].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 3].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].FormatConditions[1].Interior.ColorIndex := 3;

  icol := icol + 4;

  //富士康料况  魅力库存
  //富士康料况  总库存
  svalue1 := '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
  ExcelApp.Cells[irow1, icol + 1].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 1], ExcelApp.Cells[irow2, icol + 1]].MergeCells := true;
  //富士康料况  Yes/No
  svalue1 := '=IF(' + GetRef(icol + 3) + IntToStr(irow1) + '<0,"No","Yes")';
  ExcelApp.Cells[irow1, icol + 2].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 2].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 2], ExcelApp.Cells[irow2, icol + 2]].MergeCells := true;
  //富士康料况  差异数量                            
  svalue1 := '=' + GetRef(icol + 1) + IntToStr(irow1) + '-' + GetRef(icolNeedSumFOX) + IntToStr(irow1);
  ExcelApp.Cells[irow1, icol + 3].Value := svalue1;
  ExcelApp.Cells[irow1, icol + 3].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].FormatConditions.Delete;
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  ExcelApp.Range[ExcelApp.Cells[irow1, icol + 3], ExcelApp.Cells[irow2, icol + 3]].FormatConditions[1].Interior.ColorIndex := 3;

  icol := FColWin + 3;
  
  if sType = 'S' then
  begin
    //累计生产数量
    for iModel := 0 to aModels.Count - 1 do
    begin
      svalue1 := '=$' + GetRef(icol) + '$3*' + GetRef(FColUsage + iModel) + IntToStr(irow1);
      ExcelApp.Cells[irow1, icol].Value := svalue1;
      ExcelApp.Cells[irow1, icol].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      ExcelApp.Range[ExcelApp.Cells[irow1, icol], ExcelApp.Cells[irow2, icol]].MergeCells := true;

      svalue1 := '=$' + GetRef(icol + 1) + '$3*' + GetRef(FColUsage + iModel) + IntToStr(irow1);
      ExcelApp.Cells[irow1, icol + 1].Value := svalue1;
      ExcelApp.Cells[irow1, icol + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
      ExcelApp.Range[ExcelApp.Cells[irow1, icol + 1], ExcelApp.Cells[irow2, icol + 1]].MergeCells := true;

      icol := icol + 2;
    end;
  end;

  irow := irow + Self.Count;
end;

{ TExcelBom }

constructor TExcelBom.Create(const sModel, sType: string; iICMOQtyFOX, iICMOQtyML: Integer);
begin
  inherited Create;
  FModel := sModel;
  FType := sType;
  FICMOQtyFOX := iICMOQtyFox;
  FICMOQtyML := iICMOQtyML;
  FBomItems := TList.Create;
  FMergedModels := TStringList.Create;
  FMergedNumbers := TStringList.Create;
end;

destructor TExcelBom.Destroy;
begin
  FMergedModels.Free;
  FMergedNumbers.Free;
  ClearBomItems;
  FBomItems.Free;
  inherited;
end;

procedure TExcelBom.MergeBom(aExcelBom: TExcelBom);
var
  i: Integer;
  bi: TBomItem;
begin
  FMergedModels.Add(aExcelBom.Model + '=' + IntToStr(aExcelBom.FICMOQtyFOX) + '/' + IntToStr(aExcelBom.FICMOQtyML));
  FMergedNumbers.Add(aExcelBom.Model + '=' + aExcelBom.FCPNumber);
  for i := 0 to aExcelBom.Count - 1 do
  begin
    bi := aExcelBom.Items[i];
    MergeBomItem(bi);
  end;
end;

procedure TExcelBom.SaveAs(const sfile: string);
begin
  WorkBook := ExcelApp.WorkBooks.Add;

  WriteData;

  WorkBook.SaveAs(sfile);
  WorkBook.Close;
end;

function TExcelBom.IndexByNumber(const snumber: string): Integer;
var
  i: Integer;
  idx: Integer;
  bi: TBomItem;
begin
  Result := -1;
  for i := 0 to Self.Count - 1 do
  begin
    bi := Self.Items[i];
    idx := bi.IndexOf(snumber);
    if idx >= 0 then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function TExcelBom.IndexOfCol(irow: Integer; const scolname: string): Integer;
var
  icol: Integer;
  svalue: string;
begin
  Result := -1;
  for icol := 1 to 20 do
  begin
    svalue := ExcelApp.Cells[irow, icol].Value;
    if svalue = scolname then
    begin
      Result := icol;
      Break;
    end;
  end;
end;

function TExcelBom.GetCellValueAsStr(irow, icol: Integer): string;
var
  v: Variant;
begin
  if icol = -1 then
  begin
    Result := EmptyStr;
  end
  else
  begin
    v := ExcelApp.Cells[irow, icol].Value;
    if VarIsStr(v) then
      Result := v
    else
      Result := '';
  end;
end;

function TExcelBom.GetCellValueAsDouble(irow, icol: Integer): Double;
var
  vv: Variant;
begin
  if icol = -1 then
  begin
    Result := -9999;
  end
  else
  begin
    try
      vv := ExcelApp.Cells[irow, icol].Value;
      if vv = Unassigned then
        Result := -9999
      else
        Result := vv;
    except
      Result := -9999;
    end;
  end;
end;

procedure TExcelBom.ClearBomItems;
var
  aBomItem: TBomItem;
  i: integer;
begin
  for i := 0 to FBomItems.Count - 1 do
  begin
    aBomItem := TBomItem(FBomItems[i]);
    aBomItem.Free;
  end;
  FBomItems.Clear;
end;

procedure TExcelBom.MergeBomItem(bi: TBomItem);
var
  i: Integer;
  idx: Integer;
  pbir: TBomItemModels;
begin
  idx := -1;
  for i := 0 to bi.Count - 1 do
  begin
    pbir := TBomItemModels(bi.Items[i]);
    idx := IndexByNumber(pbir.fnumber);
    if idx >= 0 then Break;
  end;      
  if idx >= 0 then
  begin
    MergeBomItemTo(bi, idx);
  end
  else
  begin
    FBomItems.Add(bi.Clone);
  end;
end;

procedure TExcelBom.MergeBomItemTo(bi: TBomItem; idx: Integer);
var
  i: Integer;
  idx_bim: Integer;
  iu: Integer;
  pbir: TBomItemModels;
  biDest: TBomItem;
  bim_dest: TBomItemModels;
  dUsage: Double;
begin
  biDest := Self.Items[idx];
  for i := 0 to bi.Count - 1 do
  begin
    pbir := TBomItemModels(bi.Items[i]);
    idx_bim := biDest.IndexOf(pbir.fnumber);
    if idx_bim >= 0 then
    begin
      bim_dest := biDest.Items[idx_bim];
      for iu := 0 to pbir.fUsages.Count - 1 do
      begin
        dUsage := StrToFloat(pbir.fUsages.ValueFromIndex[iu]);
        bim_dest.AddUsage(pbir.fUsages.Names[iu], dUsage, pbir.fDepts.ValueFromIndex[iu]);
      end;
    end
    else
    begin
      biDest.AddItem(pbir.Clone);
    end;
  end;
end;

function TExcelBom.GetCount: Integer;
begin
  Result := FBomItems.Count;
end;

function TExcelBom.GetItems(i: Integer): TBomItem;
begin
  if (i >= 0) and (i < Count) then
  begin
    Result := TBomItem(FBomItems[i]);
  end
  else Result := nil;
end;

function TExcelBom.LoadData: Boolean;
const
  CSNumber = '物料编码';
  CSName = '物料名称';
  CSUsage = '用量';
  CSPO = '采购员';
  CSPE = '采购工程师';
  CSSupplier ='厂家';
  CSLT ='L/T';
  CSMC = 'MC';
  CSDept = '所属组件';  
  CSCPNumber = '所属料号';

  CIColRout = 1;
  CIColColor = 2;
var
  irow: Integer;

  iNumber, iName, iUsage, iPO, iPE,
  iSupplier, iLT, iMC, iDept, iCPNumber: Integer;
  sNumber, sName, sPO, sPE,
  sSupplier, sLT, sMC, sDept, sCPNumber: string;
  dUsage: Double;
  
  aBomItem: TBomItem;
begin
  ClearBomItems;

  ExcelApp.WorkSheets[1].Activate;

  irow := 1;
  iNumber := IndexOfCol(irow, CSNumber);
  iName := IndexOfCol(irow, CSName);
  iUsage := IndexOfCol(irow, CSUsage);
  iPO := IndexOfCol(irow, CSPO);
  iPE := IndexOfCol(irow, CSPE);
  iSupplier := IndexOfCol(irow, CSSupplier);
  iLT := IndexOfCol(irow, CSLT);
  iMC := IndexOfCol(irow, CSMC);
  iDept := IndexOfCol(irow, CSDept);
  iCPNumber := IndexOfCol(irow, CSCPNumber);

  if iUsage = -1 then
  begin
    MessageBox(0, PChar('文件 ' + self.FFile + ' 没有 用量 列'), '错误', 0);
    Exit;
  end;

  irow := 2;
  aBomItem := nil;
  sNumber := ExcelApp.Cells[irow, iNumber].Value;
  while sNumber <> EmptyStr do
  begin
    if sNumber = '01.01.1020327' then
    begin
      Sleep(10);
    end;
    sName := GetCellValueAsStr(irow, iName);
    dUsage := GetCellValueAsDouble(irow, iUsage);
    sPO := GetCellValueAsStr(irow, iPO);
    sPE := GetCellValueAsStr(irow, iPE);
    sSupplier := GetCellValueAsStr(irow, iSupplier);
    sLT := GetCellValueAsStr(irow, iLT);
    sMC := GetCellValueAsStr(irow, iMC);
    sDept := GetCellValueAsStr(irow, iDept);
    sCPNumber := GetCellValueAsStr(irow, iCPNumber);

    if dUsage <> -9999 then
    begin
      aBomItem := TBomItem.Create;
      FBomItems.Add(aBomItem);
    end;

    aBomItem.AddItem(sNumber, sName, sPO, sPE, sSupplier, sLT, sMC, sDept, FModel, dUsage);
    irow := irow + 1;
    sNumber := GetCellValueAsStr(irow, iNumber);
  end;
  FCPNumber := sCPNumber;
  Result := True;
end;

procedure ExtractQty(const svalue: string; var iqtyFOX: Integer; var iqtyML: Integer);
var
  sv1, sv2: string;
begin
  sv1 := Copy(svalue, 1, Pos('/', svalue) - 1);
  sv2 := Copy(svalue, Pos('/', svalue) + 1, Length(svalue) - Pos('/', svalue));
  iqtyFOX := StrToInt(sv1);
  iqtyML := StrToInt(sv2);
end;
           
procedure TExcelBom.WriteData;
const
  xlCenter = -4108;
var
  bi: TBomItem;
  i: Integer;
  irow: Integer;
  icol: Integer;
  aModels: TStringList;
  iModel: Integer;
  iqtyFOX, iqtyML: Integer;
  sicmosumFOX, sicmosumML: string;
begin
  aModels := TStringList.Create;
  aModels.Add(FModel);
  for i := 0 to FMergedModels.Count - 1 do
  begin
    aModels.Add(FMergedModels.Names[i]);
  end;
  
  ExcelApp.ActiveSheet.Rows[1].Font.Bold := True;
  ExcelApp.ActiveSheet.Rows[2].Font.Bold := True;
  ExcelApp.ActiveSheet.Rows[3].Font.Bold := True;

  icol := CSColNumber;
  ExcelApp.Cells[1, icol].Value := '物料编码';
  ExcelApp.Columns[icol].ColumnWidth := 16;
  ExcelApp.ActiveSheet.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := CSColName;
  ExcelApp.Cells[1, icol].Value := '物料名称';
  ExcelApp.Columns[icol].ColumnWidth := 30;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  //用量///////////////////////
  icol := icol + 1;
  FColUsage := icol;
  ExcelApp.Cells[1, icol].Value := '用量';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;                                            
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + FMergedModels.Count]].MergeCells := true;

  ExcelApp.Cells[2, icol].Value := FModel;
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[3, icol].Value := FCPNumber;
  ExcelApp.Cells[3, icol].HorizontalAlignment := xlCenter;
  //ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  
  for i := 0 to FMergedModels.Count - 1 do
  begin
    icol := icol + 1;
    ExcelApp.Cells[2, icol].Value := FMergedModels.Names[i];
    ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Cells[3, icol].Value := FMergedNumbers.ValueFromIndex[i];
    ExcelApp.Cells[3, icol].HorizontalAlignment := xlCenter;
    //ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  end;
       
  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '所属产品型号';
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;                                      
  ExcelApp.Cells[1, icol].Value := FModel + '生产需求数量';
  ExcelApp.Columns[icol].ColumnWidth := 11;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + 1]].MergeCells := true;     
  ExcelApp.Cells[2, icol].Value := '魅力';
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[3, icol].Value := FICMOQtyML;
  ExcelApp.Cells[3, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[2, icol + 1].Value := '富士康';
  ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[3, icol + 1].Value := FICMOQtyFOX;
  ExcelApp.Cells[3, icol + 1].HorizontalAlignment := xlCenter;
  ExcelApp.Columns[icol + 1].ColumnWidth := 11;

  sicmosumML := '=' + GetRef(icol) + '3';
  sicmosumFOX := '=' + GetRef(icol + 1) + '3';

  for i := 0 to FMergedModels.Count - 1 do
  begin
    icol := icol + 2;
    ExtractQty(FMergedModels.ValueFromIndex[i], iqtyFOX, iqtyML);
    ExcelApp.Cells[1, icol].Value := FMergedModels.Names[i] + '生产需求数量';
    ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + 1]].MergeCells := true;
    ExcelApp.Cells[2, icol].Value := '魅力';
    ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;                                
    ExcelApp.Cells[3, icol].Value := iqtyML;
    ExcelApp.Cells[3, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Cells[2, icol + 1].Value := '富士康';
    ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;
    ExcelApp.Cells[3, icol + 1].Value := iqtyFOX;
    ExcelApp.Cells[3, icol + 1].HorizontalAlignment := xlCenter;
    ExcelApp.Columns[icol].ColumnWidth := 11;     
    ExcelApp.Columns[icol + 1].ColumnWidth := 11;

    sicmosumML := sicmosumML + '+' + GetRef(icol) + '3';
    sicmosumFOX := sicmosumFOX + '+' + GetRef(icol + 1) + '3';;
  end;

  icol := icol + 2;
  ExcelApp.Cells[1, icol].Value := '总生产需求数量';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + 1]].MergeCells := true;
  ExcelApp.Cells[2, icol].Value := '魅力';
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;      
  //ExcelApp.Cells[3, icol].Value := sicmosumML;
  ExcelApp.Cells[3, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[2, icol + 1].Value := '富士康';
  ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;
  //ExcelApp.Cells[3, icol + 1].Value := sicmosumFOX;
  ExcelApp.Cells[3, icol + 1].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol + 1]].Interior.Color := $0099FFFF;

  icol := icol + 2;
  ExcelApp.Cells[1, icol].Value := '魅力料况';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + 3]].MergeCells := true;
  FColStockML := icol;
  ExcelApp.Cells[2, icol].Value := '魅力库存';
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Cells[2, icol + 1].Value := '总库存（PCS）';
  ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 1], ExcelApp.Cells[3, icol + 1]].MergeCells := true;
  ExcelApp.Cells[2, icol + 2].Value := 'Yes/No';
  ExcelApp.Cells[2, icol + 2].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 2], ExcelApp.Cells[3, icol + 2]].MergeCells := true;      
  ExcelApp.Cells[2, icol + 3].Value := '差异数量';
  ExcelApp.Cells[2, icol + 3].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[2, icol + 3].Font.Color := clRed;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 3], ExcelApp.Cells[3, icol + 3]].MergeCells := true;

  icol := icol + 4;
  ExcelApp.Cells[1, icol].Value := '富士康料况';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + 3]].MergeCells := true;
  FColStockFOX := icol;
  ExcelApp.Cells[2, icol].Value := '富士康库存';
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Cells[2, icol + 1].Value := '总库存（PCS）';
  ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 1], ExcelApp.Cells[3, icol + 1]].MergeCells := true;
  ExcelApp.Cells[2, icol + 2].Value := 'Yes/No';
  ExcelApp.Cells[2, icol + 2].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 2], ExcelApp.Cells[3, icol + 2]].MergeCells := true;
  ExcelApp.Cells[2, icol + 3].Value := '差异数量';
  ExcelApp.Cells[2, icol + 3].HorizontalAlignment := xlCenter;
  ExcelApp.Cells[2, icol + 3].Font.Color := clRed;
  ExcelApp.Range[ExcelApp.Cells[2, icol + 3], ExcelApp.Cells[3, icol + 3]].MergeCells := true;

  icol := icol + 4;
  FColStockWW := icol;
  ExcelApp.Cells[1, icol].Value := '委外库存';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;
  FColWin := icol;
  ExcelApp.Cells[1, icol].Value := '累计进料数量';    
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].Interior.Color := $00CCFF99;
                    
  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '累计进料魅力';    
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].Interior.Color := $00CCFF99;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '累计进料富士康';    
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].Interior.Color := $00CCFF99;

  if FType = 'S' then
  begin
    icol := icol + 1;
          
    ExcelApp.Cells[1, icol].Value := '累计生产数量';
    ExcelApp.Cells[1, icol].WrapText := True;
    ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + aModels.Count * 2 - 1]].MergeCells := true;
                        
    for iModel := 0 to aModels.Count - 1 do
    begin
      ExcelApp.Cells[2, icol].Value := aModels[iModel] + '魅力';
      ExcelApp.Cells[2, icol].WrapText := True;
      ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;

      ExcelApp.Cells[2, icol + 1].Value := aModels[iModel] + '富士康';
      ExcelApp.Cells[2, icol + 1].WrapText := True;
      ExcelApp.Cells[2, icol + 1].HorizontalAlignment := xlCenter;   
      icol := icol + 2;
    end;
  end;

  FColOrder := icol;
  ExcelApp.Cells[1, icol].Value := '未结订单';    
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].Interior.Color := $00CCFF99;

  icol := icol + 3;                                              
  ExcelApp.Cells[1, icol].Value := '采购回复交期';
  ExcelApp.Columns[icol].ColumnWidth := 30;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].Interior.Color := $0000FFFF;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '采购员';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '采购工程师';   
  ExcelApp.Cells[1, icol].WrapText := True;
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := '厂家';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := 'L/T';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;
  ExcelApp.Cells[1, icol].Value := 'MC';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[3, icol]].MergeCells := true;

  icol := icol + 1;   
  ExcelApp.Cells[1, icol].Value := '归属组件';
  ExcelApp.Cells[1, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[1, icol], ExcelApp.Cells[1, icol + FMergedModels.Count]].MergeCells := true;
  ExcelApp.Cells[2, icol].Value := FModel;
  ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  for i := 0 to FMergedModels.Count - 1 do
  begin
    icol := icol + 1;
    ExcelApp.Cells[2, icol].Value := FMergedModels.Names[i];
    ExcelApp.Cells[2, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Range[ExcelApp.Cells[2, icol], ExcelApp.Cells[3, icol]].MergeCells := true;
  end;

  icol := icol + 1;
    
  irow := 4;
  for i := 0 to Self.Count - 1 do
  begin
    bi := Self.Items[i];
    bi.WriteData(FType, aModels, ExcelApp, FColUsage,
      FColStockML, FColStockFOX, FColStockWW, FColWin, FColOrder, irow);
  end;                                                               

  ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, icol]].Borders.LineStyle := 1; //加边框
end;
    
{ TExcelDBGridEH }

constructor TExcelDBGridEH.Create;
begin
  inherited Create;
end;

destructor TExcelDBGridEH.Destroy;
begin
  inherited;
end;

procedure TExcelDBGridEH.SaveAs(ClientDataSet1: TClientDataSet; const sfile: string);
begin
  WorkBook := ExcelApp.WorkBooks.Add;

  WriteData(ClientDataSet1);

  WorkBook.SaveAs(sfile);
  WorkBook.Close;
end;

procedure TExcelDBGridEH.WriteData(ClientDataSet1: TClientDataSet);
var
  icol, irow: Integer;
begin
  ClientDataSet1.DisableControls;
  try
    for icol := 0 to ClientDataSet1.FieldCount - 1 do
    begin
      ExcelApp.Cells[1, icol + 1].Value := ClientDataSet1.Fields[icol].FieldName;
    end;

    irow := 2;
    ClientDataSet1.First;
    while not ClientDataSet1.Eof do
    begin
      for icol := 0 to ClientDataSet1.FieldCount - 1 do
      begin
        if (ClientDataSet1.Fields[icol].FieldName = '物料编码') or
          (ClientDataSet1.Fields[icol].FieldName = '仓库编码') or
          (ClientDataSet1.Fields[icol].FieldName = '仓库名称') or
          (ClientDataSet1.Fields[icol].FieldName = '物料名称') then
        begin
          ExcelApp.Cells[irow, icol + 1].Value := '''' + ClientDataSet1.Fields[icol].AsString;
        end
        else
        begin
          ExcelApp.Cells[irow, icol + 1].Value := ClientDataSet1.Fields[icol].AsFloat;
          ExcelApp.Cells[irow, icol + 1].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
        end;
      end;
      ClientDataSet1.Next;
      irow := irow + 1;
    end;
  finally
    ClientDataSet1.EnableControls;
  end;
end;

end.
