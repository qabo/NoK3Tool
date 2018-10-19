unit FGStockRptReader;

interface

uses
  Classes, CommUtils, SysUtils, Variants, ComObj, MakeFGReportCommon;

type
  TFGStockRptRecord = packed record
    snumber: string; //物料长代码
    sname: string; //物料名称     
    sBatchNo: string; // 批次
    dqty: Double; //库存总数
    drework: Double; //返工
    duncheck: Double; //待检验
    saddr: string; //存货地点
    snote: string; //备注
    bSum: Boolean;
    ssheet: string;
    ptr: Pointer;
  end;
  PFGStockRptRecord = ^TFGStockRptRecord;
  
  TFGStockRptReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    FReadOk: Boolean;
    FProjs: TStringList;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PFGStockRptRecord;
    function GetProjCount: Integer;
    function GetProjs(i: Integer): string;
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    procedure GetNumberSet(slNumber: TStringList);
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGStockRptRecord read GetItems;
    property ReadOk: Boolean read FReadOk;
    property ProjCount: Integer read GetProjCount;
    property Projs[i: Integer]: string read GetProjs;
  end;

implementation

{ TFGStockRptReader }

constructor TFGStockRptReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  FProjs := TStringList.Create;
  Open;
end;

destructor TFGStockRptReader.Destroy;
begin
  Clear;
  FList.Free;
  FProjs.Free;
  inherited;
end;

procedure TFGStockRptReader.Clear;
var
  i: Integer;
  p: PFGStockRptRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGStockRptRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;

  FProjs.Clear;
end;

procedure TFGStockRptReader.GetNumberSet(slNumber: TStringList);
var
  i: Integer;
  p: PFGStockRptRecord;    
  pn: PFGNumberRecord;
begin
  for i := 0 to self.Count - 1 do
  begin
    p := Items[i];
    if slNumber.IndexOf(p^.snumber + '=' + p^.sBatchNo ) < 0 then
    begin                      
      pn := New(PFGNumberRecord);
      pn^.snumber := p^.snumber;
      pn^.sname := p^.sname;
      pn^.sBatchNo := p^.sBatchNo;
      slNumber.AddObject(p^.snumber + '=' + p^.sBatchNo, TObject(pn) );
    end;
  end;
end;

procedure TFGStockRptReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function TFGStockRptReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGStockRptReader.GetItems(i: Integer): PFGStockRptRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGStockRptRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

function TFGStockRptReader.GetProjCount: Integer;
begin
  Result := FProjs.Count;
end;

function TFGStockRptReader.GetProjs(i: Integer): string;
begin
  Result := FProjs[i];
end;  

procedure TFGStockRptReader.Open;
const
  CINumber = 1; //料号
  CIName = 2; //产品名称
  CIQty = 3; //库存总数 
  CIRework = 4; //返工
  CIUncheck = 5; //待检验
  CIAddr = 6; //存货地点
  CIBatchNo = 7; //批次
  CINote = 8; //备注

var
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;     
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8: string;
  stitle: string;
  irow: Integer;
  snumber, sname: string;
  p: PFGStockRptRecord;
  v: Variant;
begin
  log(FFile);

  FReadOk := False;

  Clear;

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
        stitle4 := ExcelApp.Cells[irow, 4]. Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle8 := ExcelApp.Cells[irow, 8].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 +
          stitle7 + stitle8;
        if stitle <> '料号产品名称库存总数其中存货地点批次备注' then
        begin
          Log(sSheet +'  不是 成品报表 库存 格式');
          Continue;
        end;        

        Log(sSheet +'  开始读取库存sheet');

        FReadOk := True;

        irow := 3;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        v := ExcelApp.Cells[irow, CIName].Value;
        sname := v;
        while (snumber <> '') and (sname <> '') do
        begin
          p := New(PFGStockRptRecord);
          p^.bSum := False;
          p^.ssheet := sSheet;
          p^.ptr := nil;
          
          p^.snumber := snumber;
          p^.sname := sname;
          p^.dqty := ExcelApp.Cells[irow, CIQty].Value;
          p^.drework := ExcelApp.Cells[irow, CIRework].Value;
          p^.duncheck := ExcelApp.Cells[irow, CIUncheck].Value;
          p^.saddr := ExcelApp.Cells[irow, CIAddr].Value;        
          p^.sBatchNo := ExcelApp.Cells[irow, CIBatchNo].Value;
          p^.snote := ExcelApp.Cells[irow, CINote].Value;

          if FProjs.IndexOf( Copy(snumber, 1, 5) ) < 0 then
          begin
            FProjs.Add( Copy(snumber, 1, 5) );
          end;
          
          FList.AddObject(snumber, TObject(p));

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
          v := ExcelApp.Cells[irow, CIName].Value;
          if VarIsError(v) then
            sname := snumber
          else
            sname := v;
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
