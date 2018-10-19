unit SOPSimReader;

interface

uses
  Classes, ComObj, CommUtils, SysUtils
  {$ifndef __NoK3}
  {$ifndef __SAP}
  , CommVars
  {$endif}
  {$endif}
  ;

type

  TSIMColHead = packed record
    sweek: string;
    sdate: string;
    dt1: TDateTime;
    dt2: TDateTime;
    icol: Integer;
  end;
  PSIMColHead = ^TSIMColHead;

  TSOPSimCol = class
  public
    sweek: string;
    sdate: string;
    dt1: TDateTime;
    qty_demand: Double;
    qty_a: Double;
//    qty_sop0: Double;
//    qty_mps0: Double;
  end;

  TSOPSimLine = class
  private
    FList: TList;
    procedure Clear;
    function GetItems(i: Integer): TSOPSimCol;
    function GetCount: Integer; 
  public
    sarea: string;
    snumber: string;
    sver: string;
    scolor: string;
    scap: string;
    constructor Create;
    destructor Destroy; override;
    function GetQty(dt1: TDateTime): Double;
    function GetQtyAvail(dt1: TDateTime): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSOPSimCol read GetItems;
  end;

  TSOPSimProj = class
  private
    slMonths: TStringList;
    FList: TStringList;
    function GetItems(i: Integer): TSOPSimLine;
    function GetCount: Integer;
    function GetMonthCount: Integer;
    function GetMonths(i: Integer): TStringList;
  public
    constructor Create(const sname: string);
    destructor Destroy; override;
    procedure Clear;
    function GetLine(const snumber: string): TSOPSimLine;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSOPSimLine read GetItems;
    property MonthCount: Integer read GetMonthCount;
    property Months[i: Integer]: TStringList read GetMonths;
  end;

  TSOPSimReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FProjYear: TStringList;
    FList: TStringList;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetProjCount: Integer;
    function GetProjs(i: Integer): TSOPSimProj;
    function GetProjByName(const sproj: string): TSOPSimProj; 
  public
    constructor Create(const sfile: string; slProjYear: TStringList;
      const LogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property ProjCount: Integer read GetProjCount;
    property Projs[i: Integer]: TSOPSimProj read GetProjs;
    property ProjByName[const sproj: string]: TSOPSimProj read GetProjByName;
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
  end;

implementation
 
{ TSOPSimLine }

constructor TSOPSimLine.Create;
begin
  FList := TList.Create;
end;

destructor TSOPSimLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TSOPSimLine.GetItems(i: Integer): TSOPSimCol;
begin
  Result := TSOPSimCol(FList[i]);
end;

function TSOPSimLine.GetCount: Integer;
begin
  Result := FList.Count;
end;

procedure TSOPSimLine.Clear;
var
  i: Integer;
  aSOPSimCol: TSOPSimCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimCol := TSOPSimCol(FList[i]);
    aSOPSimCol.Free;
  end;
  FList.Clear;
end;

function TSOPSimLine.GetQty(dt1: TDateTime): Double;
var
  i: Integer;
  aSOPSimCol: TSOPSimCol;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimCol := TSOPSimCol(FList[i]);
    if aSOPSimCol.dt1 = dt1 then
    begin
      Result := aSOPSimCol.qty_demand;
      Break;
    end;
  end;
end;

function TSOPSimLine.GetQtyAvail(dt1: TDateTime): Double;
var
  i: Integer;
  aSOPSimCol: TSOPSimCol;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimCol := TSOPSimCol(FList[i]);
    if aSOPSimCol.dt1 = dt1 then
    begin
      Result := aSOPSimCol.qty_a;
      Break;
    end;
  end;
end;

{ TSOPSimProj }

constructor TSOPSimProj.Create(const sname: string);
begin
  slMonths := TStringList.Create;
  FList := TStringList.Create;
end;

destructor TSOPSimProj.Destroy;
begin
  Clear;
  slMonths.Free;
  FList.Free;
end;

procedure TSOPSimProj.Clear;
var
  i: Integer;
  slWeeks: TStringList;
  aSopColHeadPtr: PSIMColHead;
  iweek: Integer;
  aSOPSimLine: TSOPSimLine;
begin
  for i := 0 to slMonths.Count - 1 do
  begin
    slWeeks := TStringList(slMonths.Objects[i]);
    for iweek := 0 to slWeeks.Count - 1 do
    begin
      aSopColHeadPtr := PSIMColHead(slWeeks.Objects[iweek]);
      Dispose(aSopColHeadPtr);
    end;
    slWeeks.Free;
  end;
  slMonths.Clear;

  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimLine := TSOPSimLine(FList.Objects[i]);
    aSOPSimLine.Free;
  end;
  FList.Clear;
end;

function TSOPSimProj.GetItems(i: Integer): TSOPSimLine;
begin
  Result := TSOPSimLine(FList.Objects[i]);
end;

function TSOPSimProj.GetCount: Integer;
begin
  Result := FList.Count;
end;  

function TSOPSimProj.GetMonthCount;
begin
  Result := slMonths.Count;
end;

function TSOPSimProj.GetMonths(i: Integer): TStringList;
begin
  Result := TStringList(slMonths.Objects[i]);
end;

function TSOPSimProj.GetLine(const snumber: string): TSOPSimLine;
var
  i: Integer;
  aSOPSimLine: TSOPSimLine;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimLine := TSOPSimLine(FList.Objects[i]);
    if aSOPSimLine.snumber = snumber then
    begin
      Result := aSOPSimLine;
      Break;
    end;
  end;
end;

  
{ TSOPSimReader }

constructor TSOPSimReader.Create(const sfile: string; slProjYear: TStringList;
  const LogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FList := TStringList.Create;
  FProjYear := slProjYear;
  FLogEvent := LogEvent;
  Open;
end;

destructor TSOPSimReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSOPSimReader.Clear;
var
  i: Integer;
  aSOPSimProj: TSOPSimProj;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSOPSimProj := TSOPSimProj(FList.Objects[i]);
    aSOPSimProj.Free;
  end;
  FList.Clear;
end;
     
procedure TSOPSimReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
    FLogEvent(str);
end;  

procedure TSOPSimReader.Open;    
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sproj: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  icol1: Integer;

  sweek: string;
  sdate: string;
  dt0: TDateTime;
  syear: string;
  smonth: string;

  slWeeks: TStringList;

  snumber: string;

  aSOPSimProj: TSOPSimProj;
  aSopColHeadPtr: PSIMColHead;
  idx: Integer;
  sdt1, sdt2: string;
  dt1, dt2: TDateTime;
  aSOPSimLine: TSOPSimLine;
  aSOPSimCol: TSOPSimCol;
  iweek: Integer;
  idate: Integer;
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

        if sSheet = 'M1891' then
        begin
          Sleep(1);
        end;

        irow := 1;                                     
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;      
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7;
        if stitle <> 'MRP区域项目物料编码标准制式颜色容量内容项' then
        begin      
          Log(sSheet +'  不是简易SOP模拟格式  项目物料编码标准制式颜色容量内容项');
          Continue;
        end;



        sproj := sSheet;
        if Pos(' ', sproj) > 0 then
        begin
          sproj := Copy(sproj, 1, Pos(' ', sproj) - 1);
        end;

        if (FProjYear.Count > 0) and (FProjYear.IndexOfName(sproj) < 0) then Continue;

        aSOPSimProj := TSOPSimProj.Create(sproj);
        FList.AddObject(sproj, aSOPSimProj);

        
        irow := 1;
        icol := 8;
        sweek := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        icol1 := icol;

        dt0 := 0;     
        syear := FProjYear.Values[sproj];
        if syear = '' then
        begin
          syear := '2017';
        end;
        
        slWeeks := TStringList.Create;
        while Trim(sweek + sdate) <> '' do
        begin
          if IsCellMerged(ExcelApp, irow, icol, irow + 1, icol)
            and (icol > icol1) then
          begin                    
            smonth := ExcelApp.Cells[irow, icol].Value;
            if slWeeks.Count > 0 then
            begin
              aSOPSimProj.slMonths.AddObject(smonth, slWeeks);
              slWeeks := TStringList.Create;
            end;   
          
            icol := icol + 1;
            sweek := ExcelApp.Cells[irow, icol].Value;
            sdate := ExcelApp.Cells[irow + 1, icol].Value;
            Continue;
          end;                                          

          aSopColHeadPtr := New(PSIMColHead);  
          aSopColHeadPtr^.sweek := sweek;
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
          log('sdate: ' + sdate + '   sdt1: ' + sdt1 + '   irow: ' + IntToStr(irow) + '  icol: ' + IntToStr(icol));
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
        slWeeks.Free;
        


        
        irow := 3;
        snumber := ExcelApp.Cells[irow, 3].Value;
        while snumber <> '' do
        begin
          if snumber = '03.60.36041002CN' then
          begin
            Sleep(1);
          end;
          aSOPSimLine := TSOPSimLine.Create;
          aSOPSimProj.FList.AddObject(snumber, aSOPSimLine);

          aSOPSimLine.snumber := snumber;
          aSOPSimLine.sarea := '';  
          aSOPSimLine.sver := ExcelApp.Cells[irow, 4].Value;
          aSOPSimLine.scolor := ExcelApp.Cells[irow, 5].Value;
          aSOPSimLine.scap := ExcelApp.Cells[irow, 6].Value;

          for iweek := 0 to aSOPSimProj.slMonths.Count - 1 do
          begin
            slWeeks := TStringList(aSOPSimProj.slMonths.Objects[iweek]);
            for idate := 0 to slWeeks.Count - 1 do
            begin
              aSopColHeadPtr := PSIMColHead(slWeeks.Objects[idate]);

              aSOPSimCol := TSOPSimCol.Create;
              aSOPSimLine.FList.Add(aSOPSimCol);

              aSOPSimCol.sweek := aSopColHeadPtr^.sweek;
              aSOPSimCol.sdate := aSopColHeadPtr^.sdate;
              aSOPSimCol.dt1 := aSopColHeadPtr^.dt1;
              aSOPSimCol.qty_demand := ExcelApp.Cells[irow, aSopColHeadPtr^.icol].Value;
              aSOPSimCol.qty_a := ExcelApp.Cells[irow + 1, aSopColHeadPtr^.icol].Value;
//              aSOPSimCol.qty_sop0 := ExcelApp.Cells[irow + 3, aSopColHeadPtr^.icol].Value;
//              aSOPSimCol.qty_mps0 := ExcelApp.Cells[irow + 4, aSopColHeadPtr^.icol].Value;

            end;
          end;
        
          irow := irow + 4;
          snumber := ExcelApp.Cells[irow, 3].Value;
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

function TSOPSimReader.GetProjCount: Integer;
begin
  Result := FList.Count ;
end;

function TSOPSimReader.GetProjs(i: Integer): TSOPSimProj;
begin
  if (i >= 0) and (i < ProjCount) then
  begin
    Result := TSOPSimProj(FList.Objects[i]);
  end
  else Result := nil;
end;

function TSOPSimReader.GetProjByName(const sproj: string): TSOPSimProj;
var
  idx: Integer;
begin
  Result := nil;
  idx := FList.IndexOf(sproj);
  if idx < 0 then Exit;
  Result := TSOPSimProj(FList.Objects[idx]);
end;

end.
