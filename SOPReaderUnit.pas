unit SOPReaderUnit;

interface

uses
  Windows, Classes, SysUtils, ComObj, Variants, CommUtils, Excel2000, DateUtils;

type
  TSOPProj = class;
  TSOPCol = class;
  TSOPLine = class;

  TSOPReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    
    FHaveArea: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetProjCount: Integer;
    function GetProjs(i: Integer): TSOPProj;
  public
    FProjYear: TStringList;                
    FProjs: TStringList;
    constructor Create(slProjYear: TStringList; const sfile: string;
      aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function GetProj(const sName: string): TSOPProj;
    procedure GetNumberList(slFGNumber: TStringList);
    procedure GetDateList(sldate: TStringList);
    procedure GetMonthList(slMonth: TStringList);
    function GetDemand(const snumber: string; dt1: TDateTime): TSOPCol;
    procedure GetDemands(const snumber: string; dt1, dtMemand: TDateTime;
      lstDemand: TList);
    function GetDemandQty(const snumber: string; dt1: TDateTime): Double;
    function GetDemandSum(dt1: TDateTime; const snumber: string): Double;
    property ProjCount: Integer read GetProjCount;
    property Projs[i: Integer]: TSOPProj read GetProjs;
    property HaveArea: Boolean read FHaveArea;
    property sFile: string read FFile;
  end;

  TSOPProj = class
  private
    function GetLineCount: Integer;
    function GetLines(i: Integer): TSOPLine;
  public
    FName: string;
    FList: TStringList;
    slMonths: TStringList;
    constructor Create(const sproj: string);
    destructor Destroy; override;
    procedure Clear;
    function GetLine(const sVer, sNumber, sColor, sCap: string): TSOPLine;
    procedure GetVerList(sl: TStringList);
    procedure GetCapList(sl: TStringList);
    function GetSumVer(const sver: string; dt: TDateTime): Double;
    function GetSumCap(const scap: string; dt: TDateTime): Double;    
    function GetSumColor(const scolor: string; dt: TDateTime): Double;    
    function GetSumAll(dt: TDateTime): Double;
    function GetNumberQty(const sNumber, sver, scolor, scap: string; dt: TDateTime): Double;

    procedure GetDateList(sl: TStringList);
    procedure GetColorList(sl: TStringList);


    property LineCount: Integer read GetLineCount;
    property Lines[i: Integer]: TSOPLine read GetLines;
    

  end;

  TSOPLine = class
  private                                             
    function GetDateCount: Integer;
    function GetDates(i: Integer): TSOPCol;
  public
    sDate: string;
    sProj: string;
    sFG: string;
    sPkg: string;
    sStdVer: string;
    sVer: string; //制式
    sNumber: string; //	物料编码
    sColor: string; //颜色
    sCap: string; //容量
    sMRPArea: string;
    FList: TStringList;
    FCalc: Boolean;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    procedure Add( const sMonth, sWeek, sDate: string; dt1, dt2: TDateTime;
      iQty: Integer ); overload;
    procedure Add( const sMonth, sWeek, sDate: string; dt1, dt2: TDateTime;
      iQty_sop, iQty_mps: Integer ); overload;   
    procedure Insert(idx: Integer; const sMonth, sWeek, sDate: string;
      dt1, dt2: TDateTime; iQty_sop, iQty_mps: Integer );
    function GetCol(const sDate: string): TSOPCol;
    function DateIdx(const dt1: TDateTime): Integer;
    property DateCount: Integer read GetDateCount;
    property Dates[i: Integer]: TSOPCol read GetDates;
  end;

  TSOPCol = class
  public
    sMonth: string;
    sWeek: string;
    sDate: string;
    dt1: TDateTime;
    dt2: TDateTime;
    iQty: Double;
    iQty_Left: Double; // 历史遗留需求
    iQty_sop: Double;
    iQty_mps: Double;
    iQty_ok: Double;   // 可齐套数量
    iQty_calc: Double; // 齐套分析，已计算数量
    icol: Integer;
    sShortageICItem: string;
    procedure AddShortageICItem(const smsg: string);
    function DemandQty: Double;
  end;

implementation

type
  TSopColHead = packed record
    sdate: string;
    dt1: TDateTime;
    dt2: TDateTime;
    icol: Integer;
  end;
  PSopColHead = ^TSopColHead;

{ TSOPProj }

constructor TSOPProj.Create(const sproj: string);
begin
  FName := sproj;
  FList := TStringList.Create;
  slMonths := TStringList.Create;
end;

destructor TSOPProj.Destroy;
begin
  Clear;
  FList.Free;
  slMonths.Free;
  inherited;
end;

procedure TSOPProj.Clear;
var
  i: Integer;
  iweek: Integer;
  aSOPLine: TSOPLine;
  slweek: TStringList;
  aSopColHeadPtr: PSopColHead;
begin
  for i := 0 to slMonths.Count - 1 do
  begin
    slweek := TStringList(slMonths.Objects[i]);
    for iweek := 0 to slweek.Count - 1 do
    begin
      aSopColHeadPtr := PSopColHead(slweek.Objects[iweek]);
      Dispose(aSopColHeadPtr);
    end;
    slweek.Free;
  end;
  slMonths.Clear;

  for i := 0 to FList.Count - 1 do
  begin
    aSOPLine := TSOPLine(FList.Objects[i]);
    aSOPLine.Free;
  end;
  FList.Clear;
end;

function TSOPProj.GetLineCount: Integer;
begin
  Result := FList.Count;
end;

function TSOPProj.GetLines(i: Integer): TSOPLine;
begin
  Result := TSOPLine(FList.Objects[i]);
end;

function TSOPProj.GetLine(const sVer, sNumber, sColor, sCap: string): TSOPLine;
var
  i: Integer;
  aSOPLine: TSOPLine;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPLine := TSOPLine(FList.Objects[i]);
    if (aSOPLine.sVer = sVer) and (aSOPLine.sNumber = sNumber)
      and (aSOPLine.sColor = sColor) and (aSOPLine.sCap = sCap) then
    begin
      Result := aSOPLine;
      Break;
    end;
  end;
end;

procedure TSOPProj.GetVerList(sl: TStringList);
var
  iLine: Integer;
  aSOPLine: TSOPLine;
begin
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if sl.IndexOf( aSOPLine.sVer ) < 0 then
    begin
      sl.Add(aSOPLine.sVer);
    end;
  end;
end;     

procedure TSOPProj.GetCapList(sl: TStringList);
var
  iLine: Integer;
  aSOPLine: TSOPLine;
begin
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if sl.IndexOf( aSOPLine.sCap ) < 0 then
    begin
      sl.Add(aSOPLine.sCap);
    end;
  end;
end;

procedure TSOPProj.GetColorList(sl: TStringList);
var
  iLine: Integer;
  aSOPLine: TSOPLine;
begin
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if sl.IndexOf( aSOPLine.sColor ) < 0 then
    begin
      sl.Add(aSOPLine.sColor);
    end;
  end;
end;  
   
function TSOPProj.GetSumVer(const sver: string; dt: TDateTime): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if aSOPLine.sVer <> sver then Continue;

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];
      if DoubleE(aSOPCol.dt1, dt) then
      begin
        Result := Result + aSOPCol.iQty;
        Break;
      end;
    end;
  end;
end;
   
function TSOPProj.GetSumCap(const scap: string; dt: TDateTime): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if aSOPLine.sCap <> scap then Continue;

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];
      if DoubleE(aSOPCol.dt1, dt) then
      begin
        Result := Result + aSOPCol.iQty;
        Break;
      end;
    end;
  end;
end;      
   
function TSOPProj.GetSumColor(const scolor: string; dt: TDateTime): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    if aSOPLine.sColor <> scolor then Continue;

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];
      if DoubleE(aSOPCol.dt1, dt) then
      begin
        Result := Result + aSOPCol.iQty;
        Break;
      end;
    end;
  end;
end;         
   
function TSOPProj.GetSumAll(dt: TDateTime): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];
      if DoubleE(aSOPCol.dt1, dt) then
      begin
        Result := Result + aSOPCol.iQty;
        Break;
      end;
    end;
  end;
end;

function TSOPProj.GetNumberQty(const sNumber, sver, scolor, scap: string; dt: TDateTime): Double;
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := 0;
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];
    // 编码为空，比较 制式、颜色、容量
    if (Trim(sNumber) = '') and (Trim(aSOPLine.sNumber) = '') then
    begin
      if (aSOPLine.sVer <> sver) or (aSOPLine.sColor <> scolor)
        or (aSOPLine.sCap <> scap) then Continue; 
    end
    else if (aSOPLine.sNumber <> sNumber) then Continue;

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];
      if DoubleE(aSOPCol.dt1, dt) then
      begin
        Result := Result + aSOPCol.iQty;
        Break;
      end;
    end;
  end;
end;

procedure TSOPProj.GetDateList(sl: TStringList);
var
  iLine: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  aSOPColNew: TSOPCol;
begin
  for iLine := 0 to self.LineCount - 1 do
  begin
    aSOPLine := self.Lines[iLine];

    for idate := 0 to aSOPLine.DateCount - 1 do
    begin
      aSOPCol := aSOPLine.Dates[idate];

      aSOPColNew := TSOPCol.Create;
      aSOPColNew.sMonth := aSOPCol.sMonth;
      aSOPColNew.sWeek := aSOPCol.sWeek;
      aSOPColNew.sDate := aSOPCol.sDate;
      aSOPColNew.dt1 := aSOPCol.dt1;

      sl.AddObject(aSOPColNew.sDate, aSOPColNew);
    end;

    Break;
  end;
end;

{ TSOPLine }
 
constructor TSOPLine.Create;
begin
  FList := TStringList.Create;
  FCalc := False;
end;

destructor TSOPLine.Destroy;
begin
  Clear;
  FList.Free;
end;

procedure TSOPLine.Clear;
var
  i: Integer;
  aSOPCol: TSOPCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSOPCol := TSOPCol(FList.Objects[i]);
    aSOPCol.Free;
  end;
  FList.Clear;
end;

function TSOPLine.GetDateCount: Integer;
begin
  Result := FList.Count;
end;

function TSOPLine.GetDates(i: Integer): TSOPCol;
begin
  Result := TSOPCol(FList.Objects[i]);
end;

procedure TSOPLine.Add( const sMonth, sWeek, sDate: string; dt1, dt2: TDateTime;
  iQty: Integer );
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := TSOPCol.Create;
  aSOPCol.sMonth := sMonth;
  aSOPCol.sWeek := sWeek;
  aSOPCol.sDate := sDate;
  aSOPCol.iQty := iQty;
  aSOPCol.dt1 := dt1;
  aSOPCol.dt2 := dt2;
  FList.AddObject(sDate, aSOPCol);
end;   

procedure TSOPLine.Add( const sMonth, sWeek, sDate: string; dt1, dt2: TDateTime;
  iQty_sop, iQty_mps: Integer );
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := TSOPCol.Create;
  aSOPCol.sMonth := sMonth;
  aSOPCol.sWeek := sWeek;
  aSOPCol.sDate := sDate;
  aSOPCol.dt1 := dt1;
  aSOPCol.dt2 := dt2;
  aSOPCol.iQty_sop := iQty_sop;
  aSOPCol.iQty_mps := iQty_mps;
  FList.AddObject(sDate, aSOPCol);
end;      

procedure TSOPLine.Insert(idx: Integer; const sMonth, sWeek, sDate: string;
  dt1, dt2: TDateTime; iQty_sop, iQty_mps: Integer );
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := TSOPCol.Create;
  aSOPCol.sMonth := sMonth;
  aSOPCol.sWeek := sWeek;
  aSOPCol.sDate := sDate;
  aSOPCol.dt1 := dt1;
  aSOPCol.dt2 := dt2;
  aSOPCol.iQty_sop := iQty_sop;
  aSOPCol.iQty_mps := iQty_mps;
  FList.InsertObject(idx, sDate, aSOPCol);
end;

function TSOPLine.GetCol(const sDate: string): TSOPCol;
var
  i: Integer;
  aSOPCol: TSOPCol;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPCol := TSOPCol(FList.Objects[i]);
    if aSOPCol.sDate = sDate then
    begin
      Result := aSOPCol;
      Break;
    end;
  end;
end;

function TSOPLine.DateIdx(const dt1: TDateTime): Integer;
var
  i: Integer;
  aSOPCol: TSOPCol;
begin
  Result := -1;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPCol := TSOPCol(FList.Objects[i]);
    if FormatDateTime('yyyy-MM-dd', aSOPCol.dt1) = FormatDateTime('yyyy-MM-dd', dt1) then
    begin
      Result := i;
      Break;
    end;
  end;
end;  

{ TSOPReader }

constructor TSOPReader.Create(slProjYear: TStringList; const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FProjYear := slProjYear;
  FProjs := TStringList.Create;

  Open;


end;

destructor TSOPReader.Destroy;
begin
  Clear;
  FProjs.Free;
  inherited;
end;

procedure TSOPReader.Clear;
var
  i: Integer;
  aSOPProj: TSOPProj;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPProj := TSOPProj(FProjs.Objects[i]);
    aSOPProj.Free;
  end;
  FProjs.Clear;
end;

function TSOPReader.GetProj(const sName: string): TSOPProj;
var
  i: Integer;
  aSOPProj: TSOPProj;
begin
  Result := nil;
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPProj := TSOPProj(FProjs.Objects[i]);
    if aSOPProj.FName = sName then
    begin
      Result := aSOPProj;
      Break;
    end;
  end;
end;

procedure TSOPReader.GetNumberList(slFGNumber: TStringList);
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
begin
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
       if slFGNumber.IndexOf(aSOPLine.sNumber) >= 0 then Continue;
       slFGNumber.Add(aSOPLine.sNumber);
     end;
  end;
end;

procedure TSOPReader.GetDateList(sldate: TStringList);
  function IndexOfDate(dt1: TDateTime): Integer;
  var
    iCount: Integer;
    aSOPCol: TSOPCol;
  begin
    Result := -1;
    for iCount := 0 to sldate.Count - 1 do
    begin
      aSOPCol := TSOPCol(sldate.Objects[iCount]);
      if dt1 = aSOPCol.dt1 then
      begin
        Result := iCount;
        Break;
      end;
    end;
  end;
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  aSOPColNew: TSOPCol;
begin
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);

       for idate := 0 to aSOPLine.FList.Count - 1 do
       begin                                         
         aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
         if IndexOfDate(aSOPCol.dt1) >= 0 then Continue; 

         aSOPColNew := TSOPCol.Create;
         aSOPColNew.sMonth := aSOPCol.sMonth;
         aSOPColNew.sWeek := aSOPCol.sWeek;
         aSOPColNew.sDate := aSOPCol.sDate;
         aSOPColNew.dt1 := aSOPCol.dt1;
         aSOPColNew.dt2 := aSOPCol.dt2;
         sldate.AddObject(aSOPColNew.sDate, aSOPColNew);
       end;

       Break;
     end;
  end;
end;
                               
  function StringListSortCompare_month(List: TStringList; Index1, Index2: Integer): Integer;
  var
    dt1, dt2: TDateTime;
  begin
    dt1 := myStrToDateTime(List[Index1]); 
    dt2 := myStrToDateTime(List[Index2]);
    if DoubleG(dt1, dt2) then
      Result := 1
    else if DoubleE(dt1, dt2) then
      Result := 0
    else Result := -1;
  end;

procedure TSOPReader.GetMonthList(slMonth: TStringList);
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  aSOPColNew: TSOPCol;
  sdt: string;
begin
  slMonth.Clear;
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);

       for idate := 0 to aSOPLine.FList.Count - 1 do
       begin                                         
         aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
         sdt := FormatDateTime('yyyy-MM', aSOPCol.dt1) + '-01';
         if slMonth.IndexOf(sdt) >= 0 then Continue;
         slMonth.Add(sdt);
       end;

       Break;
     end;
  end;
  slMonth.CustomSort(StringListSortCompare_month); 
end;

function TSOPReader.GetDemand(const snumber: string; dt1: TDateTime): TSOPCol;
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
begin
  Result := nil;
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
       if aSOPLine.sNumber <> snumber then Continue;
       
       for idate := 0 to aSOPLine.FList.Count - 1 do
       begin
         aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
         if aSOPCol.dt1 <> dt1 then Continue;

         Result := aSOPCol;

         Break;
       end;

       Break;
     end;
     if Result <>  nil then Break;
  end;
end;

procedure TSOPReader.GetDemands(const snumber: string; dt1, dtMemand: TDateTime;
  lstDemand: TList);
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  aSOPCol_last: TSOPCol;
begin
  lstDemand.Clear;
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
       if aSOPLine.sNumber <> snumber then Continue;
       
       for idate := 0 to aSOPLine.FList.Count - 1 do
       begin
         aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
         if aSOPCol.dt1 <> dt1 then Continue;

         if idate > 0 then
         begin
           aSOPCol_last := TSOPCol(aSOPLine.FList.Objects[idate - 1]);
           if aSOPCol_last.dt1 >= dtMemand then
           begin
             aSOPCol.iQty_Left := aSOPCol_last.DemandQty - aSOPCol_last.iQty_ok; // 上一日期为满足数量
           end;
         end;

         lstDemand.Add(aSOPCol);
         Break;
       end;
     end; 
  end;
end;

function TSOPReader.GetDemandQty(const snumber: string; dt1: TDateTime): Double;
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := GetDemand(snumber, dt1);
  if aSOPCol = nil then
    Result := 0
  else Result := aSOPCol.iQty;
end;

function TSOPReader.GetDemandSum(dt1: TDateTime; const snumber: string): Double;
var
  iproj: Integer;
  aSOPProj: TSOPProj;
  iline: Integer;
  aSOPLine: TSOPLine;
  idate: Integer;
  aSOPCol: TSOPCol;
  bFound: Boolean;
begin
  Result := 0;
  bFound := False;
  for iproj := 0 to self.FProjs.Count - 1 do
  begin
     aSOPProj := TSOPProj(FProjs.Objects[iproj]);
     for iline := 0 to aSOPProj.FList.Count - 1 do
     begin
       aSOPLine := TSOPLine(aSOPProj.FList.Objects[iline]);
       if aSOPLine.sNumber <> snumber then Continue;
       bFound := True;
       
       for idate := 0 to aSOPLine.FList.Count - 1 do
       begin                               
         aSOPCol := TSOPCol(aSOPLine.FList.Objects[idate]);
         if aSOPCol.dt1 < dt1 then Continue;
         Result := Result + aSOPCol.iQty;
       end;
       Break;
     end;
     if bFound then Break;
  end;
end;
    
function TSOPReader.GetProjCount: Integer;
begin
  Result := FProjs.Count;
end;

function TSOPReader.GetProjs(i: Integer): TSOPProj;
begin
  Result := TSOPProj(FProjs.Objects[i]);
end;

procedure TSOPReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;

function CheckSopCol(const sdate: string): Boolean;
var
  s: string;
  idx: Integer;
begin
  Result := False;
  s := sdate;
  idx := Pos('/', s);
  if idx <= 0 then Exit;
  s := Copy(s, idx + 1, Length(s));

  idx := Pos('-', s);
  if idx <= 0 then Exit;
  s := Copy(s, idx + 1, Length(s));

  idx := Pos('/', s);
  if idx <= 0 then Exit;
  s := Copy(s, idx + 1, Length(s));

  if Length(s) > 2 then Exit;

  Result := True;
end;

procedure MoveActOut(ExcelApp: Variant; icolDate1: Integer; const syear: string);
var
  icol: Integer;
  irow: Integer;
  s: string;
  icolActOut: Integer;
  sFirstWeek: string;
  sweek: string;
  icolCut: Integer;
  s1, s2, s3, s4, s5: string;
  dt: TDateTime;
begin
  icolActOut := 0;
  irow := 1;
  for icol := 5 to 1000 do
  begin
    s := ExcelApp.Cells[irow, icol].Value;
    if s = '实际出货' then
    begin
      icolActOut := icol;
      Break;
    end;

    s1 := ExcelApp.Cells[irow, icol].Value;
    s2 := ExcelApp.Cells[irow, icol + 1].Value;
    s3 := ExcelApp.Cells[irow, icol + 2].Value;
    s4 := ExcelApp.Cells[irow, icol + 3].Value;
    s5 := ExcelApp.Cells[irow, icol + 4].Value;
    
    s := s1 + s2 + s3 + s4 + s5;
      
    if s = '' then Break; 
  end;

  if icolActOut = 0 then Exit;

  sFirstWeek := ExcelApp.Cells[2, icolDate1];

  icolCut := 0;
  for icol := icolActOut to icolActOut + 10 do
  begin
    if not IsCellMerged(ExcelApp, 1, icolActOut, 1, icol) then Break;   
    sweek := ExcelApp.Cells[2, icol];
    if sweek = sFirstWeek then
    begin
      icolCut := icol - 1;
      Break;
    end;
  end;

  if icolCut < icolActOut then Exit;

  ExcelApp.Cells[1, icolActOut].UnMerge;

  ExcelApp.Columns[GetRef(icolActOut) + ':' + GetRef(icolCut)].Select;
  ExcelApp.Selection.Cut;
  ExcelApp.Columns[GetRef(icolDate1) + ':' + GetRef(icolDate1)].Select;
  ExcelApp.Selection.Insert (Shift:=xlToRight);

  for icol := icolDate1 to icolDate1 + icolCut - icolActOut do
  begin
    s := ExcelApp.Cells[2, icol].Value;
    s := Copy(s, 1, Pos('-', s) - 1);
    s := syear + '-' + StringReplace(s, '/', '-', [rfReplaceAll]);
    dt := myStrToDateTime(s);
    ExcelApp.Cells[1, icol].Value := 'WK' + IntToStr(WeekOfTheYear(dt));
  end;  
end;

function GetStartYear(ExcelApp: Variant; icolDate1: Integer): string;
var
  icol: Integer;
  irow: Integer;
  s: string;    
  s1, s2, s3, s4, s5: string;
  v: variant;
  dt: TDateTime;
  syear: string;
begin
  syear := '';
          
  irow := 1;
  for icol := icolDate1 to 500 do
  begin
    v := ExcelApp.Cells[irow, icol].Value;
    if VarIsType(v, varDate) then // 标准日期格式
    begin
      dt := v;
      syear := FormatDateTime('yyyy', dt);
      Break;
    end
    else
    begin
      s := v;                     // 格式中有年度
      if Pos('年', s) > 0 then
      begin
        syear := Copy(s, 1, 4);
        Break;
      end;
    end;

    s1 := ExcelApp.Cells[irow, icol].Value;
    s2 := ExcelApp.Cells[irow, icol + 1].Value;
    s3 := ExcelApp.Cells[irow, icol + 2].Value;
    s4 := ExcelApp.Cells[irow, icol + 3].Value;
    s5 := ExcelApp.Cells[irow, icol + 4].Value;
    
    s := s1 + s2 + s3 + s4 + s5;
      
    if s = '' then Break;  
  end;

  Result := syear;
end;

procedure TSOPReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sproj: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9: string;
  stitle4x, stitle8x: string;    
  stitle5x, stitle9x: string;
  irow, icol: Integer;
  icol1: Integer;
  smonth: string;
  sweek: string;
  sdate: string;
  irow1: Integer;
  icolDate1: Integer;
  icolMRPArea: Integer;
  icolVer: Integer; //制式
  icolNumber: Integer; //	物料编码
  icolColor: Integer; //颜色
  icolCap: Integer; //容量

  icolProj: Integer;
  icolFG: Integer;
  icolPkg: Integer;
  icolStdVer: Integer;
			
  sVer: string; //制式
  sNumber: string; //	物料编码
  sColor: string; //颜色
  sCap: string; //容量
  v: Variant;
  iQty: Integer;

  slWeeks: TStringList; 
  iMonth: Integer;
  iWeek: Integer;
  
  aSOPLine: TSOPLine;
  aSOPProj: TSOPProj;
  iProj: Integer;
  aSopColHeadPtr: PSopColHead;

  dt0: TDateTime;
  syear: string;
  sdt1, sdt2: string;
  dt1, dt2: TDateTime;
  idx: Integer;

  s: string;
begin
  Clear;
                 
  icolDate1 := 0;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    FHaveArea := False;
    
    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;

        ExcelApp.Columns[1].ColumnWidth := 2;
        ExcelApp.Columns[2].ColumnWidth := 18;
        ExcelApp.Columns[3].ColumnWidth := 10;
          
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
        stitle4x := stitle1 + stitle2 + stitle3 + stitle4;
        stitle8x := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8;

        stitle5x := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        stitle9x := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8 + stitle9;


        if stitle4x = '制式物料编码颜色容量' then
        begin      
          FHaveArea := False;
          
          icolVer := 1; // Integer; //制式
          icolNumber := 2; // Integer; //	物料编码
          icolColor := 3; // Integer; //颜色
          icolCap := 4; // Integer; //容量        
          icolDate1 := 5;
        end
        else if stitle8x = '项目整机/裸机包装标准制式制式物料编码颜色容量' then
        begin       
          FHaveArea := False;

          icolProj := 1;
          icolFG := 2;
          icolPkg := 3;
          icolStdVer := 4;

          icolVer := 5; // Integer; //制式
          icolNumber := 6; // Integer; //	物料编码
          icolColor := 7; // Integer; //颜色
          icolCap := 8; // Integer; //容量
          icolDate1 := 9;
        end    
        else if (stitle5x = 'MRP区域制式物料编码颜色容量')
          or (stitle5x = 'MRP范围制式物料编码颜色容量') then
        begin
          FHaveArea := True;
          icolMRPArea := 1;
          icolVer := 2; // Integer; //制式
          icolNumber := 3; // Integer; //	物料编码
          icolColor := 4; // Integer; //颜色
          icolCap := 5; // Integer; //容量
          icolDate1 := 6;
        end
        else if (stitle9x = '项目整机/裸机包装标准制式MRP区域制式物料编码颜色容量')
          or (stitle9x = '项目整机/裸机包装标准制式MRP范围制式物料编码颜色容量') then
        begin              
          FHaveArea := True;

          icolProj := 1;
          icolFG := 2;
          icolPkg := 3;
          icolStdVer := 4;
          
          icolMRPArea := 5;
          icolVer := 6; // Integer; //制式
          icolNumber := 7; // Integer; //	物料编码
          icolColor := 8; // Integer; //颜色
          icolCap := 9; // Integer; //容量
          icolDate1 := 10;
        end
        else
        begin
          Log(sSheet + ' 不是SOP格式');
          Continue;
        end;


        sproj := sSheet;
        if Pos(' ', sSheet) > 0 then
        begin
          sproj := Copy(sSheet, 1, Pos(' ', sSheet) - 1);
        end;

        if (FProjYear <> nil) and
          (FProjYear.Count > 0) and
          (FProjYear.IndexOfName(sproj) < 0) then
        begin
          Log(sSheet + ' 没有项目开始年度');
          Continue;
        end;

        syear := GetStartYear(ExcelApp, icolDate1);
        if syear = '' then
        begin
          // 文件格式中获取不到年份， 从配置获取
          if FProjYear = nil then
          begin
            syear := inttostr(yearof(now))
          end
          else
          begin
            syear := FProjYear.Values[sproj];
            syear := Trim(syear);
            if syear = '' then
            begin
              syear := inttostr(yearof(now));
            end;
          end;
        end;

        MoveActOut(ExcelApp, icolDate1, syear);





        
        aSOPProj := TSOPProj.Create(sproj);
        FProjs.AddObject(sproj, aSOPProj);
   

        irow := 1;
        icol := icolDate1;
        sweek := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        icol1 := icol;      

        dt0 := 0;

        slWeeks := TStringList.Create;
        while Trim(sweek + sdate) <> '' do
        begin
          if IsCellMerged(ExcelApp, irow, icol, irow + 1, icol)
            and (icol > icol1) then
          begin                    
            smonth := ExcelApp.Cells[irow, icol].Value;
            if slWeeks.Count > 0 then
            begin
              aSOPProj.slMonths.AddObject(smonth, slWeeks);
              slWeeks := TStringList.Create;   
            end;   
          
            icol := icol + 1;
            sweek := ExcelApp.Cells[irow, icol].Value;
            sdate := ExcelApp.Cells[irow + 1, icol].Value;

//            v := ExcelApp.Cells[irow + 1, icol].Value;
//            if VarIsType(v, varDate) then
//            begin
//              sdate := FormatDateTime('MM/dd-MM/dd', v);
//            end
//            else sdate := v;


            Continue;
          end;

          if not CheckSopCol(sdate) then
          begin 
            icol := icol + 1;
            sweek := ExcelApp.Cells[irow, icol].Value;
            sdate := ExcelApp.Cells[irow + 1, icol].Value;
            Continue;
          end;

          aSopColHeadPtr := New(PSopColHead);
          aSopColHeadPtr^.sdate := sdate;
          aSopColHeadPtr^.icol := icol;

          idx := Pos('-', sdate);
          if idx > 0 then
          begin
            sdt1 := Copy(sdate, 1, idx - 1);
            sdt2 := Copy(sdate, idx + 1, Length(sdate) - idx)
          end
          else
          begin
            sdt1 := sdate;
            sdt2 := sdate;
          end;
          sdt1 := StringReplace(sdt1, '/', '-', [rfReplaceAll]);
          sdt2 := StringReplace(sdt2, '/', '-', [rfReplaceAll]);
          dt1 := myStrToDateTime(syear + '-' + sdt1);
          dt2 := myStrToDateTime(syear + '-' + sdt2);
          if dt0 > dt1 then
          begin
            syear := IntToStr(StrToInt(syear) + 1);
            dt1 := myStrToDateTime(syear + '-' + sdt1);
            dt2 := myStrToDateTime(syear + '-' + sdt2);
          end;
          dt0 := dt1;

          aSopColHeadPtr^.dt1 := dt1;
          aSopColHeadPtr^.dt2 := dt2;

          slWeeks.AddObject(sweek + '=' + sdate, TObject(aSopColHeadPtr));
          
          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
        end;
        if slWeeks.Count > 0 then
        begin
          aSOPProj.slMonths.AddObject(smonth, slWeeks);
        end
        else slWeeks.Free;
        

        irow := 3;
        irow1 := 0;
        while not IsCellMerged(ExcelApp, irow, icolNumber, irow, icolCap) do
        begin
          if (irow1 = 0) or
            not IsCellMerged(ExcelApp, irow1, icolVer, irow, icolVer) then
          begin
            irow1 := irow;       
            sVer := ExcelApp.Cells[irow, icolVer].Value;
          end;   
          sNumber := ExcelApp.Cells[irow, icolNumber].Value;
          sColor := ExcelApp.Cells[irow, icolColor].Value;
          sCap := ExcelApp.Cells[irow, icolCap].Value;

          sNumber := Trim(sNumber);

          if {(sVer = '') and} (sNumber = '') and (sColor = '') and (sCap = '') then
          begin
            Log('结束');
            Break;
          end;


          aSOPLine := TSOPLine.Create;
          aSOPProj.FList.AddObject(sNumber, aSOPLine);

          if (icolVer = 5) or (icolVer = 6) then
          begin
            aSOPLine.sProj := ExcelApp.Cells[irow, 1].Value;
            aSOPLine.sFG := ExcelApp.Cells[irow, 2].Value;
            aSOPLine.sPkg := ExcelApp.Cells[irow, 3].Value;
            aSOPLine.sStdVer := ExcelApp.Cells[irow, 4].Value;
          end;

          aSOPLine.sVer := sVer;
          aSOPLine.sNumber := sNumber;
          aSOPLine.sColor := sColor;
          aSOPLine.sCap := sCap;
          if FHaveArea then
          begin
            aSOPLine.sMRPArea := ExcelApp.Cells[irow, icolMRPArea].Value;
          end;

          for iMonth := 0 to aSOPProj.slMonths.Count - 1 do
          begin
            slWeeks := TStringList(aSOPProj.slMonths.Objects[iMonth]);
            for iWeek := 0 to slWeeks.Count - 1 do
            begin
              aSopColHeadPtr := PSopColHead( slWeeks.Objects[iWeek] );
              v := ExcelApp.Cells[irow, aSopColHeadPtr^.icol].Value;
              if VarIsNumeric(v) then
              begin
                iQty := v;
              end
              else
              begin
                s := v;
                iQty := StrToIntDef(s, 0);
              end;

              aSOPLine.Add(aSOPProj.slMonths[iMonth], slWeeks.Names[iWeek],
                slWeeks.ValueFromIndex[iWeek],
                aSopColHeadPtr^.dt1, aSopColHeadPtr^.dt2, iQty);
            end;
          end;
       
          irow := irow + 1;
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

{ TSOPCol }

procedure TSOPCol.AddShortageICItem(const smsg: string);
begin
  if sShortageICItem <> '' then
    sShortageICItem := sShortageICItem + #13#10'';
  sShortageICItem := sShortageICItem + smsg;
end;

function TSOPCol.DemandQty: Double;
begin
  Result := iQty + iQty_Left;
end;

end.
