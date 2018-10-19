unit FGToProduceReader;

interface

uses
  Classes, CommUtils, SysUtils, ComObj, MakeFGReportCommon, FGRptWinReader,
  ZMDR001Reader;

type
  TCountrySet = (csC, csR, csU, csI, csO); // 中国，俄罗斯，乌克兰，印尼，其他
  
  TFGToProduceRecord = packed record 
    sproj: string; //机型
    sprojname: string; //产品名称
    sitem: string; //颜色\制式\容量
//    dtotal_in: Double; //项目总量         国内
//    dtotal_out: Double; //项目总量        海外
//    dproduct_in: Double; //已生产数量     国内
//    dproduct_out: Double; //已生产数量    海外
    dtoproduce_in: Double; //待产数量		    海外        10
    dtoproduce_out: Double; //待产数量      小计        11
  end;
  PFGToProduceRecord = ^TFGToProduceRecord;
  
  TFGToProduceReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PFGToProduceRecord;
    procedure SubWinLine(aZMDR001RecordPtr: PZMDR001Record;
      aFGRptWinRecordPtr: PFGRptWinRecord);
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    procedure SubWin( aFGRptWinReader: TFGRptWinReader;
      aSAPMaterialReader: TSAPMaterialReader; dt1, dt2: TDateTime);
    function GetQty(const sProjName,
      sitem: string; slMap: TStrings; bHW: Boolean; acs: TCountrySet): Double;
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGToProduceRecord read GetItems;
    property ReadOk: Boolean read FReadOk;
  end;

implementation

{ TFGToProduceReader}

constructor TFGToProduceReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  Open;
end;

destructor TFGToProduceReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TFGToProduceReader.Clear;
var
  i: Integer;
  p: PFGToProduceRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGToProduceRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
procedure TFGToProduceReader.SubWinLine(aZMDR001RecordPtr: PZMDR001Record;
  aFGRptWinRecordPtr: PFGRptWinRecord);
var
  i: Integer;
  p: PFGToProduceRecord;
begin
  for i := 0 to Count - 1 do
  begin
    p := Items[i];
    
    if (p^.sproj = aZMDR001RecordPtr^.sProj) and
      ((p^.sitem = aZMDR001RecordPtr^.sVer) or
       (p^.sitem = aZMDR001RecordPtr^.scap) or
       (p^.sitem = aZMDR001RecordPtr^.sColor)
       ) then
    begin
      if TSAPMaterialReader.IsHW(aZMDR001RecordPtr) then
      begin
        p^.dtoproduce_out := p^.dtoproduce_out - aFGRptWinRecordPtr^.dQty;
      end
      else
      begin
        p^.dtoproduce_in := p^.dtoproduce_in - aFGRptWinRecordPtr^.dQty;
      end;
    end; 
  end;
end;  

procedure TFGToProduceReader.SubWin( aFGRptWinReader: TFGRptWinReader;
  aSAPMaterialReader: TSAPMaterialReader; dt1, dt2: TDateTime);
var
  i: Integer;
  aFGRptWinRecordPtr: PFGRptWinRecord;
  aZMDR001RecordPtr: PZMDR001Record;
begin
  dt1 := myStrToDateTime( FormatDateTime('yyyy-MM-dd', dt1) );
  dt2 := myStrToDateTime( FormatDateTime('yyyy-MM-dd', dt2) );

  for i := 0 to aFGRptWinReader.Count - 1 do
  begin
    aFGRptWinRecordPtr := aFGRptWinReader.Items[i];
    if (aFGRptWinRecordPtr^.dt < dt1) or (aFGRptWinRecordPtr^.dt > dt2) then
    begin
      Continue;
    end;
    aZMDR001RecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aFGRptWinRecordPtr^.sNumber);
    if aZMDR001RecordPtr = nil then
    begin
      Log(aFGRptWinRecordPtr^.sNumber + ' 找不到物料基础资料');
      Continue;
    end;

    SubWinLine(aZMDR001RecordPtr, aFGRptWinRecordPtr);
  end;  
end;  

function TFGToProduceReader.GetQty(
  const sProjName, sitem: string; slMap: TStrings; bHW: Boolean; acs: TCountrySet): Double;
var
  i: Integer;
  p: PFGToProduceRecord;
  sitem_p: string;
  idx: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    p := Items[i];

    sitem_p := p^.sitem;
    if slMap <> nil then
    begin
      idx := slMap.IndexOfName(sitem_p);
      if idx >= 0 then
      begin
        sitem_p := slMap.ValueFromIndex[idx];
      end;
    end;
    
    if (p^.sproj = sProjName) and (sitem_p = sitem) then
    begin
      if bHW then
      begin

        case acs of
          csC: Continue;
          csR:
            if (Pos('俄文', p^.sitem) = 0)
              and (Pos('俄罗斯', p^.sitem) = 0) then Continue;
          csU:
            if Pos('乌克兰', p^.sitem) = 0 then Continue;
          csI:
            if Pos('印尼', p^.sitem) = 0 then Continue;
          csO:
            if (Pos('俄文', p^.sitem) > 0)
              or (Pos('俄罗斯', p^.sitem) > 0)
              or (Pos('乌克兰', p^.sitem) > 0)
              or (Pos('印尼', p^.sitem) > 0) then Continue;
        end;
        Result := Result + p^.dtoproduce_out;
      end
      else
      begin
        Result := Result + p^.dtoproduce_in;
      end;
    end; 
  end;
end;  

procedure TFGToProduceReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function TFGToProduceReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGToProduceReader.GetItems(i: Integer): PFGToProduceRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGToProduceRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TFGToProduceReader.Open;
const
  CIDate = 1; //入库日期
  CINumber = 2; //物料长代码
  CIName = 3; //物料名称
  CIQty = 4; //数量
  CIFac = 5; //代工厂
  CIBatchNo = 6; //批次
  CINote = 7; //备注
 
var
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;     
  stitle1, stitle2, stitle3, stitle4, stitle7, stitle10: string;
  stitle: string;
  irow: Integer;
  sitem: string;
  p: PFGToProduceRecord;
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
        stitle7 := ExcelApp.Cells[irow, 7].Value;   
        stitle10 := ExcelApp.Cells[irow, 10].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle7 + stitle10;

        if stitle <> '机型产品名称颜色\制式\容量项目总量已生产数量待产数量' then
        begin

          Log(sSheet +'  不是 成品报表 入库 格式');
          Continue;
        end;
             
        FReadOk := True;
  
        irow := 3;
        sitem := ExcelApp.Cells[irow, 3].Value; 
        while sitem <> '' do
        begin
          if sitem <> 'TOTAL' then
          begin
            p := New(PFGToProduceRecord);
            if not IsCellMerged(ExcelApp, irow, 1, irow - 1, 1) then
            begin
              p^.sproj := ExcelApp.Cells[irow, 1].Value; //机型
              p^.sprojname := ExcelApp.Cells[irow, 2].Value; //产品名称
            end;
            p^.sitem := sitem; //颜色\制式\容量
            p^.dtoproduce_in := ExcelApp.Cells[irow, 10].Value; //待产数量		    海外        10
            p^.dtoproduce_out := ExcelApp.Cells[irow, 11].Value; //待产数量      小计        11

            FList.AddObject(sitem, TObject(p));
          end;

          irow := irow + 1;
          sitem := ExcelApp.Cells[irow, 3].Value; 
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
