unit SalePlanWFReader;

interface

uses
   Classes, SysUtils, ComObj, CommUtils, Variants;

type
  TSalePlanWFRecord = packed record
    dt: TDateTime;
    dqty_ib1: Double;
    dqty_ib2: Double;
    dqty_ob1: Double;
    dqty_ob2: Double;
  end;
  PSalePlanWFRecord = ^TSalePlanWFRecord;
  
  TSalePlanWFProj = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PSalePlanWFRecord;
  public
    sname: string;
    snote_ib: string;
    snote_ob: string;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    function GetLastWeekQty(dt1: TDateTime; bIn: Boolean): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSalePlanWFRecord read GetItems;
  end;
  
  TSalePlanWFWeek = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): TSalePlanWFProj;
  public
    sweek1: string;
    sweek2: string; 
    constructor Create;
    destructor Destroy; override;
    procedure Clear;       
    function GetLastWeekQty(const sproj: string; dt1: TDateTime; bIn: Boolean): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSalePlanWFProj read GetItems;
  end;
  
  TSalePlanWFReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FList: TList;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSalePlanWFWeek;
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetLastWeekQty(const sproj: string; dt1: TDateTime; bIn: Boolean): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSalePlanWFWeek read GetItems;
  end;

implementation
 
{ TSalePlanWFProj }

constructor TSalePlanWFProj.Create;
begin
  FList := TList.Create;
end;

destructor TSalePlanWFProj.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSalePlanWFProj.Clear;
var
  i: Integer;
  p: PSalePlanWFRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSalePlanWFRecord(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSalePlanWFProj.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSalePlanWFProj.GetItems(i: Integer): PSalePlanWFRecord;
begin
  Result := PSalePlanWFRecord(FList[i]);
end;

function TSalePlanWFProj.GetLastWeekQty(dt1: TDateTime; bIn: Boolean): Double;
var
  i: Integer;
  p: PSalePlanWFRecord;
  ss: string;
begin
  Result := 0;

  ss := FormatDateTime('yyyy-MM', dt1);
  for i := 0 to FList.Count - 1 do
  begin
    p := PSalePlanWFRecord(FList[i]);
    if FormatDateTime('yyyy-MM', p^.dt) = ss then
    begin
      if bIn then
        Result := p^.dqty_ib2
      else Result := p^.dqty_ob2;
      Break;
    end;
  end;
end;

{ TSalePlanWFWeek }

constructor TSalePlanWFWeek.Create;
begin
  FList := TList.Create;
end;

destructor TSalePlanWFWeek.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSalePlanWFWeek.Clear;
var
  i: Integer;
  aSalePlanWFProj: TSalePlanWFProj;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSalePlanWFProj := TSalePlanWFProj(FList[i]);
    aSalePlanWFProj.Free;
  end;
  FList.Clear;
end;

function TSalePlanWFWeek.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSalePlanWFWeek.GetItems(i: Integer): TSalePlanWFProj;
begin
  Result := TSalePlanWFProj(FList[i]);
end;  

function TSalePlanWFWeek.GetLastWeekQty(const sproj: string; dt1: TDateTime; bIn: Boolean): Double;
var
  i: Integer;
  aSalePlanWFProj: TSalePlanWFProj;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aSalePlanWFProj := TSalePlanWFProj(FList[i]);
    if aSalePlanWFProj.sname = sproj then
    begin
      Result := aSalePlanWFProj.GetLastWeekQty(dt1, bIn);
      Break;
    end;
  end;
end;

{ TSalePlanWFReader }

constructor TSalePlanWFReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;

destructor TSalePlanWFReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSalePlanWFReader.Clear;
var
  i: Integer;
  aSalePlanWFWeek: TSalePlanWFWeek;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSalePlanWFWeek := TSalePlanWFWeek(FList[i]);
    aSalePlanWFWeek.Free;
  end;
  FList.Clear;
end;

function TSalePlanWFReader.GetLastWeekQty(const sproj: string; dt1: TDateTime;
  bIn: Boolean): Double;
var
  i: Integer;
  aSalePlanWFWeek: TSalePlanWFWeek;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aSalePlanWFWeek := TSalePlanWFWeek(FList[i]);
    Result := aSalePlanWFWeek.GetLastWeekQty(sproj, dt1, bIn);

    Break;
  end;
end;

procedure TSalePlanWFReader.Log(const str: string);
begin

end;

function TSalePlanWFReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSalePlanWFReader.GetItems(i: Integer): TSalePlanWFWeek;
begin
  Result := TSalePlanWFWeek(FList[i]);
end;

procedure TSalePlanWFReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  sproj, sbound, sweek: string;
  aSalePlanWFWeek: TSalePlanWFWeek;
  aSalePlanWFProj: TSalePlanWFProj;
  send: string;
  icol: Integer;
  v: Variant;
  irow_t: Integer;
  aSalePlanWFRecordPtr: PSalePlanWFRecord;
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
        stitle := stitle1 + stitle2 + stitle3;
        if stitle <> '项目国内/海外销售计划' then
        begin
          Log(sSheet +'  不是销售计划Waterfall格式');
          Continue;
        end;

        aSalePlanWFWeek := nil;

        irow_t := irow;
        
        irow := irow + 1;
        send := ExcelApp.Cells[irow, 3].Value + ExcelApp.Cells[irow + 1, 3].Value + ExcelApp.Cells[irow + 2, 3].Value; // 连续 3 行为空结束
        while send <> '' do
        begin
          sweek := ExcelApp.Cells[irow, 3].Value;
          if sweek = '' then
          begin
            irow_t := irow;
            aSalePlanWFWeek := nil;
            irow := irow + 1;
            send := ExcelApp.Cells[irow, 3].Value + ExcelApp.Cells[irow + 1, 3].Value + ExcelApp.Cells[irow + 2, 3].Value;
            Continue;
          end;

          send := ExcelApp.Cells[irow, 1].Value + ExcelApp.Cells[irow, 2].Value + ExcelApp.Cells[irow, 3].Value;  // 标题列
          if send = '项目国内/海外销售计划' then
          begin
            irow_t := irow;     
            irow := irow + 1;
            Continue;
          end;

          if aSalePlanWFWeek = nil then
          begin
            aSalePlanWFWeek := TSalePlanWFWeek.Create;
            FList.Add(aSalePlanWFWeek);

            aSalePlanWFWeek.sweek1 := ExcelApp.Cells[irow, 3].Value;
            aSalePlanWFWeek.sweek2 := ExcelApp.Cells[irow + 1, 3].Value;
          end;

          aSalePlanWFProj := TSalePlanWFProj.Create;
          aSalePlanWFProj.sname := ExcelApp.Cells[irow, 1].Value;

          aSalePlanWFWeek.FList.Add(aSalePlanWFProj);

          icol := 4;
          v := ExcelApp.Cells[irow_t, icol].Value;
          while VarIsType(v, varDate) do
          begin
            aSalePlanWFRecordPtr := New(PSalePlanWFRecord);
            aSalePlanWFProj.FList.Add(aSalePlanWFRecordPtr);

            aSalePlanWFRecordPtr^.dt := v;
            aSalePlanWFRecordPtr^.dqty_ib1 := ExcelApp.Cells[irow, icol].Value;
            aSalePlanWFRecordPtr^.dqty_ib2 := ExcelApp.Cells[irow + 1, icol].Value;
            aSalePlanWFRecordPtr^.dqty_ob1 := ExcelApp.Cells[irow + 3, icol].Value;
            aSalePlanWFRecordPtr^.dqty_ob2 := ExcelApp.Cells[irow + 4, icol].Value;

            icol := icol + 1;
            v := ExcelApp.Cells[irow_t, icol].Value;
          end;


          icol := icol + 1;
          aSalePlanWFProj.snote_ib := ExcelApp.Cells[irow, icol].Value;
          aSalePlanWFProj.snote_ob := ExcelApp.Cells[irow + 3, icol].Value;

          irow := irow + 6;
          send := ExcelApp.Cells[irow, 3].Value + ExcelApp.Cells[irow + 1, 3].Value + ExcelApp.Cells[irow + 2, 3].Value;
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
