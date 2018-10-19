unit FGDBRptReader;

interface

uses
  Classes, CommUtils, SysUtils, ComObj, MakeFGReportCommon;

type
  TFGDBRptRecord = packed record
    snumber: string; //料号
    sname: string; //产品名称
    sBatchNo: string; // 批次
    sstock: string; 
    dqty: Double;
  end;
  PFGDBRptRecord = ^TFGDBRptRecord;
  
  TFGDBRptReader = class
  private
    FFile: string;
    FList: TStringList;               
    FLogEvent: TLogEvent;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);       
    function GetCount: Integer;
    function GetItems(i: Integer): PFGDBRptRecord;
  public
    constructor Create(const sfile: string; const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                  
    procedure GetNumberSet(slNumber: TStringList);    
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;    
    property Count: Integer read GetCount;
    property Items[i: Integer]: PFGDBRptRecord read GetItems;

  end;

implementation

{ TFGDBRptReader}

constructor TFGDBRptReader.Create(const sfile: string; const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;   
  FLogEvent := LogEvent;
  Open;
end;

destructor TFGDBRptReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TFGDBRptReader.Clear;
var
  i: Integer;
  p: PFGDBRptRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p :=  PFGDBRptRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
          
procedure TFGDBRptReader.GetNumberSet(slNumber: TStringList);
var
  i: Integer;
  p: PFGDBRptRecord;     
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

procedure TFGDBRptReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;  
       
function TFGDBRptReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TFGDBRptReader.GetItems(i: Integer): PFGDBRptRecord;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PFGDBRptRecord(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TFGDBRptReader.Open;
const
  CINumber = 1; //物料代码
  CIName = 2;   //物料名称
  CIStock = 3;  //仓库名称
  CIQty = 4;    //常用单位数量

var
  sSheet: string;
  iSheet: Integer;
  iSheetCount: Integer;     
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  p: PFGDBRptRecord;
begin
     
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
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
        if stitle <> '物料代码物料名称仓库名称常用单位数量' then
        begin
          Log(sSheet +'  不是 成品报表 未出货 格式');
          Continue;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin
          p := New(PFGDBRptRecord);
          
          p^.snumber := ExcelApp.Cells[irow, CINumber].Value;
          p^.sname := ExcelApp.Cells[irow, CIName].Value;
          p^.sstock := ExcelApp.Cells[irow, CIStock].Value;
          p^.dqty := ExcelApp.Cells[irow, CIQty].Value;

          FList.AddObject(snumber, TObject(p));
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;
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
