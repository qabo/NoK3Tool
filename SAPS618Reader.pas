unit SAPS618Reader;

interface

uses
  Classes, ComObj, DateUtils, SysUtils, CommUtils, SOPReaderUnit;

type
  TSAPS618Col = packed record
    snumber: string;
    sname: string;
    sweek: string;
    dt1, dt2: TDateTime;
    dqty: Double;

    dQty_Left: Double;
    dDemandQty: Double;
    dQty_ok: Double;
    dQty_calc: Double;
    sShortageICItem: string;
  end;
  PSAPS618Col = ^TSAPS618Col;

  TSAPS618 = class
  private    
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPS618Col;
  public
    FNumber: string; 
    sname: string;
    sgroup: string;
    sgroupname: string;
    sfac: string;
    sDemandType: string;
    sDemandVer: string;
    sAct: string;
    sPlanNo: string;
    FMrpArea: string;
    sMrper: string;
    sUnit: string;

    Tag: TObject;

    constructor Create(const snumber, smrparea: string);
    destructor Destroy; override;
    procedure Clear;
    procedure Add(const snumber, sname, sweek: string; dt1, dt2: TDateTime;
      dqty: Double); overload;
    procedure Add(p: PSAPS618Col); overload;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPS618Col read GetItems;
    function GetSum: Double;
  end; 
  
  TSAPS618Reader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FList: TStringList;
    FPlan: string;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPS618;
  public
    slWeek: TStringList;
    constructor Create(const sfile, splan: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function GetItem(const snumber, smrparea: string): TSAPS618; // GetItem里是新Create的，调用者要负责释放掉      
    procedure GetDemands(const snumber: string; dt1, dtMemand: TDateTime;
      lstDemand: TList);
    procedure GetDateList(sldate: TStringList);
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPS618 read GetItems;
  end;

  TSAPPIRReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FList: TStringList;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPS618;
  public
    slWeek: TStringList;
    FNumbers: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function GetItem(const snumber, smrparea: string): TSAPS618; // GetItem里是新Create的，调用者要负责释放掉
    procedure SubNumber(const snumber: string; var iQty: Integer);
    property LogEvent: TLogEvent read FLogEvent write FLogEvent;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPS618 read GetItems;
  end;

implementation

{ TSAPS618 }

constructor TSAPS618.Create(const snumber, smrparea: string);
begin
  FNumber := snumber;
  FMrpArea := smrparea;
  FList := TList.Create;
end;

destructor TSAPS618.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPS618.Clear;
var
  i: Integer;
  aSAPS618ColPtr: PSAPS618Col;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSAPS618ColPtr := PSAPS618Col(FList[i]);
    Dispose(aSAPS618ColPtr);
  end;
  FList.Clear;
end;

function TSAPS618.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPS618.GetItems(i: Integer): PSAPS618Col;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PSAPS618Col(FList[i]);
  end
  else Result := nil;
end;

procedure TSAPS618.Add(const snumber, sname, sweek: string; dt1, dt2: TDateTime;
  dqty: Double);
var
  aSAPS618ColPtr: PSAPS618Col;
begin
  aSAPS618ColPtr := New(PSAPS618Col);
  aSAPS618ColPtr^.snumber := snumber;
  aSAPS618ColPtr^.sname := sname;
  aSAPS618ColPtr^.sweek := sweek;
  aSAPS618ColPtr^.dt1 := dt1;
  aSAPS618ColPtr^.dt2 := dt2;
  aSAPS618ColPtr^.dqty := dQty;

  aSAPS618ColPtr^.dQty_Left := 0;
  aSAPS618ColPtr^.dDemandQty := 0;
  aSAPS618ColPtr^.dQty_ok := 0;
  aSAPS618ColPtr^.dQty_calc := 0;
  aSAPS618ColPtr^.sShortageICItem := '';

  FList.Add(aSAPS618ColPtr);
end;

procedure TSAPS618.Add(p: PSAPS618Col);
var
  aSAPS618ColPtr: PSAPS618Col;
begin
  aSAPS618ColPtr := New(PSAPS618Col);
  aSAPS618ColPtr^ := p^;
  FList.Add(aSAPS618ColPtr);
end;

function TSAPS618.GetSum: Double;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    Result := Result + Items[i]^.dqty;
  end;
end;  

{ TSAPS618Reader }

constructor TSAPS618Reader.Create(const sfile, splan: string;
  aLogEvent: TLogEvent = nil);
begin
  FLogEvent := aLogEvent;
  FPlan := splan;
  FFile := sfile;
  FList := TStringList.Create;
  slWeek := TStringList.Create;
  Open;
end;

destructor TSAPS618Reader.Destroy;
begin
  Clear;
  FList.Free;   
  slWeek.Free;
  inherited;
end;

procedure TSAPS618Reader.Clear;
var
  i: Integer;
  aSAPS618: TSAPS618;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSAPS618 := TSAPS618(FList.Objects[i]);
    aSAPS618.Free;
  end;
  FList.Clear;

  slWeek.Clear;
end;

function TSAPS618Reader.GetItem(const snumber, smrparea: string): TSAPS618;  // GetItem里是新Create的，调用者要负责释放掉
var
  i: Integer;
  j: Integer;
  aSAPS618: TSAPS618;
  lst: TList;
  aSAPS618New: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  aSAPS618ColPtrNew: PSAPS618Col;
begin
  Result := nil;

  lst := TList.Create;
  try
    for i := 0 to FList.Count - 1 do
    begin
      aSAPS618 := TSAPS618(FList.Objects[i]);
      if smrparea = '' then
      begin
        if aSAPS618.FNumber = snumber then
        begin
          lst.Add(Pointer(aSAPS618));
        end;
      end
      else
      begin
        if (aSAPS618.FNumber = snumber) and (aSAPS618.FMrpArea = smrparea) then
        begin
          lst.Add(Pointer(aSAPS618));
        end;
      end;
    end;

    if lst.Count > 0 then
    begin      
      aSAPS618 := TSAPS618(lst[0]);
      aSAPS618New := TSAPS618.Create(aSAPS618.FNumber, aSAPS618.FMrpArea);
      for i := 0 to aSAPS618.Count - 1 do
      begin
        aSAPS618New.Add(aSAPS618.Items[i]);
      end;
            
      for i := 1 to lst.Count - 1 do
      begin           
        aSAPS618 := TSAPS618(lst[i]);
        for j := 0 to aSAPS618.Count - 1 do
        begin
          aSAPS618ColPtr := aSAPS618.Items[j];
          aSAPS618ColPtrNew := aSAPS618New.Items[j];
          aSAPS618ColPtrNew^.dqty := aSAPS618ColPtrNew^.dqty + aSAPS618ColPtr^.dqty;
        end;

        lst.Clear;
      end;

      Result := aSAPS618New;
    end;  
  finally
    lst.Free;
  end;
end;
     
procedure TSAPS618Reader.GetDemands(const snumber: string; dt1, dtMemand: TDateTime;
  lstDemand: TList);
var
  iline: Integer;
  aSAPS618: TSAPS618;
  idate: Integer;
  aSAPS618ColPtr: PSAPS618Col;
  aSAPS618ColPtr_last: PSAPS618Col;
begin
  lstDemand.Clear;


   for iline := 0 to Self.Count - 1 do
   begin
     aSAPS618 := Self.Items[iline];
     if aSAPS618.FNumber <> snumber then Continue;


     for idate := 0 to aSAPS618.Count - 1 do
     begin
       aSAPS618ColPtr := aSAPS618.Items[idate];
       if aSAPS618ColPtr^.dt1 <> dt1 then Continue;

       if idate > 0 then
       begin
         aSAPS618ColPtr_last := aSAPS618.Items[idate - 1];
         if aSAPS618ColPtr_last^.dt1 >= dtMemand then
         begin
           aSAPS618ColPtr_last^.dQty_Left := aSAPS618ColPtr_last^.dDemandQty - aSAPS618ColPtr_last^.dQty_ok; // 上一日期为满足数量
         end;
       end;

       lstDemand.Add(aSAPS618ColPtr);
       Break;
     end;
   end; 

end;  

procedure TSAPS618Reader.GetDateList(sldate: TStringList);  
  function IndexOfDate(dt1: TDateTime): Integer;
  var
    iCount: Integer;
    a: PSAPS618Col;
  begin
    Result := -1;
    for iCount := 0 to sldate.Count - 1 do
    begin
      a := PSAPS618Col(sldate.Objects[iCount]);
      if dt1 = a^.dt1 then
      begin
        Result := iCount;
        Break;
      end;
    end;
  end;
var
  i: Integer;
  aSAPS618: TSAPS618;    
  j: Integer;
  aSAPS618ColPtr: PSAPS618Col;     
  aSAPS618ColPtr_New: PSAPS618Col; 
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSAPS618 := TSAPS618(FList.Objects[i]);

    for j := 0 to aSAPS618.Count - 1 do
    begin
      aSAPS618ColPtr := PSAPS618Col(aSAPS618.Items[j]);

      if IndexOfDate(aSAPS618ColPtr^.dt1) >= 0 then Continue;

      aSAPS618ColPtr_New := New(PSAPS618Col);
      aSAPS618ColPtr_New^ := aSAPS618ColPtr^;
      sldate.AddObject(aSAPS618ColPtr_New^.sweek, TObject(aSAPS618ColPtr_New));

    
    end;

  end;
end;

procedure TSAPS618Reader.Log(const str: string);
begin

end;

function TSAPS618Reader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPS618Reader.GetItems(i: Integer): TSAPS618;
begin
  Result := TSAPS618(FList.Objects[i]);
end;

procedure TSAPS618Reader.Open;
var
  irow: Integer;
  stitle1, stitle2, stitle3, stitle4,
  stitle5, stitle6, stitle7, stitle8: string;
  stitle: string;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  snumber: string;
  sname: string;
  smrparea: string;
  sletter: string;
  aSAPS618: TSAPS618;
  sweek: string;
  icol: Integer;
  iweek: Integer;
  dqty: Double;
  dt1, dt2: TDateTime;
  sdt1, sdt2: string;
  swk: string;

  dt0: TDateTime;
  iyear: Integer;
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

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;  // 物料
        stitle2 := ExcelApp.Cells[irow, 2].Value;  //工厂
        stitle3 := ExcelApp.Cells[irow, 3].Value;  //MRP 范围
        stitle4 := ExcelApp.Cells[irow, 4].Value;  //物料描述
        stitle5 := ExcelApp.Cells[irow, 5].Value;  //计划类型
        stitle6 := ExcelApp.Cells[irow, 6].Value;  //字符
        stitle7 := ExcelApp.Cells[irow, 7].Value;  //订单测量单位
        stitle8 := ExcelApp.Cells[irow, 8].Value;  //合计

        stitle := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6 + stitle7 + stitle8;
        if stitle <> '物料工厂MRP 范围物料描述计划类型字符订单测量单位合计' then
        begin
          Log('sheet ' + sSheet + '  格式不符');
          Continue;
        end;

        icol := 10;
        sweek := ExcelApp.Cells[irow, icol].Value;
        while sweek <> '' do
        begin
          swk := Copy(sweek, 6, 2);
          if Pos('(', sweek) > 0 then
          begin
            sweek := Copy(sweek, Pos('(', sweek) + 1, Length(sweek));
          end;
          if Pos(')', sweek) > 0 then
          begin
            sweek := Copy(sweek, 1, Pos(')', sweek) - 1);
          end;

          slWeek.AddObject(sweek + '=' + swk, TObject(icol));

          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          if snumber = '83.68.36810002CN' then
          begin
            Sleep(1);
          end;
          
          sletter := ExcelApp.Cells[irow, 6].Value;
          if sletter <> FPlan then
          begin    
            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 1].Value;
            Continue;
          end;
          sname := ExcelApp.Cells[irow, 4].Value;
          smrparea := ExcelApp.Cells[irow, 3].Value;
          aSAPS618 := TSAPS618.Create(snumber, smrparea);
          aSAPS618.sDemandType := ExcelApp.Cells[irow, 6].Value;
          FList.AddObject(snumber, aSAPS618);

          //aSAPS618.FNumber := snumber;

          dt0 := 0;
          iyear := YearOf(Now);
          for iweek := 0 to slWeek.Count - 1 do
          begin
            sweek := slWeek.Names[iweek];
            icol := Integer(slWeek.Objects[iweek]);
            dqty := ExcelApp.Cells[irow, icol].Value;
            
            sdt1 := Copy(sweek, 1, Pos('-', sweek) - 1);
            sdt2 := Copy(sweek, Pos('-', sweek) + 1, Length(sweek));
            sdt1 := IntToStr( iyear ) + '-' + Copy(sdt1, 1, 2) + '-' + Copy(sdt1, 3, 2);
            sdt2 := IntToStr( iyear ) + '-' + Copy(sdt2, 1, 2) + '-' + Copy(sdt2, 3, 2);
            dt1 := myStrToDateTime(sdt1);
            dt2 := myStrToDateTime(sdt2);

            if dt1 < dt0 then
            begin
              iyear := iyear + 1;         
              sdt1 := Copy(sweek, 1, Pos('-', sweek) - 1);
              sdt2 := Copy(sweek, Pos('-', sweek) + 1, Length(sweek));
              sdt1 := IntToStr( iyear ) + '-' + Copy(sdt1, 1, 2) + '-' + Copy(sdt1, 3, 2);
              sdt2 := IntToStr( iyear ) + '-' + Copy(sdt2, 1, 2) + '-' + Copy(sdt2, 3, 2);
              dt1 := myStrToDateTime(sdt1);
              dt2 := myStrToDateTime(sdt2);
            end;  

            aSAPS618.Add(snumber, sname, sweek, dt1, dt2, dqty);
            dt0 := dt1;
          end;

          irow := irow + 1;
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

{ TSAPPIRReader }

constructor TSAPPIRReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FLogEvent := aLogEvent;
  FFile := sfile;
  FList := TStringList.Create;
  slWeek := TStringList.Create;
  FNumbers := TStringList.Create;
  Open;
end;

destructor TSAPPIRReader.Destroy;
begin
  Clear;
  FList.Free;   
  slWeek.Free;
  FNumbers.Free;
  inherited;
end;

procedure TSAPPIRReader.Clear;
var
  i: Integer;
  aSAPS618: TSAPS618;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSAPS618 := TSAPS618(FList.Objects[i]);
    aSAPS618.Free;
  end;
  FList.Clear;

  slWeek.Clear;
  FNumbers.Clear;
end;

function TSAPPIRReader.GetItem(const snumber, smrparea: string): TSAPS618;  // GetItem里是新Create的，调用者要负责释放掉
var
  i: Integer;
  j: Integer;
  aSAPS618: TSAPS618;
  lst: TList;
  aSAPS618New: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  aSAPS618ColPtrNew: PSAPS618Col;
begin
  Result := nil;

  lst := TList.Create;
  try
    for i := 0 to FList.Count - 1 do
    begin
      aSAPS618 := TSAPS618(FList.Objects[i]);
      if smrparea = '' then
      begin
        if aSAPS618.FNumber = snumber then
        begin
          lst.Add(Pointer(aSAPS618));
        end;
      end
      else
      begin
        if (aSAPS618.FNumber = snumber) and (aSAPS618.FMrpArea = smrparea) then
        begin
          lst.Add(Pointer(aSAPS618));
        end;
      end;
    end;

    if lst.Count > 0 then
    begin      
      aSAPS618 := TSAPS618(lst[0]);
      aSAPS618New := TSAPS618.Create(aSAPS618.FNumber, aSAPS618.FMrpArea);
      for i := 0 to aSAPS618.Count - 1 do
      begin
        aSAPS618New.Add(aSAPS618.Items[i]);
      end;
            
      for i := 1 to lst.Count - 1 do
      begin           
        aSAPS618 := TSAPS618(lst[i]);
        for j := 0 to aSAPS618.Count - 1 do
        begin
          aSAPS618ColPtr := aSAPS618.Items[j];
          aSAPS618ColPtrNew := aSAPS618New.Items[j];
          aSAPS618ColPtrNew^.dqty := aSAPS618ColPtrNew^.dqty + aSAPS618ColPtr^.dqty;
        end;

        lst.Clear;
      end;

      Result := aSAPS618New;
    end;  
  finally
    lst.Free;
  end;
end;

procedure TSAPPIRReader.SubNumber(const snumber: string; var iQty: Integer);
var
  i: Integer;
  aSAPS618: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  icol: Integer;
begin     
  if iQty > 0 then  // 减少需求
  begin
    for i := 0 to FList.Count - 1 do
    begin
      aSAPS618 := TSAPS618(FList.Objects[i]);
      if aSAPS618.FNumber <> snumber then Continue;

      for icol := aSAPS618.FList.Count - 1 downto 0 do
      begin
        aSAPS618ColPtr := PSAPS618Col(aSAPS618.FList[icol]);
        if aSAPS618ColPtr^.dqty >= iQty then
        begin
          aSAPS618ColPtr^.dqty := aSAPS618ColPtr^.dqty - iQty;
          iQty := 0;
          Break;
        end
        else
        begin
          iQty := iQty - Round(aSAPS618ColPtr^.dqty);
          aSAPS618ColPtr^.dqty := 0;
        end;
      end;

      if iQty = 0 then Break;
    end;
  end
  else if iQty < 0 then // 增加需求
  begin            
    for i := 0 to FList.Count - 1 do
    begin
      aSAPS618 := TSAPS618(FList.Objects[i]);  
      if aSAPS618.FNumber <> snumber then Continue;

      if aSAPS618.FList.Count > 0 then
      begin
        icol := aSAPS618.FList.Count - 1;
        aSAPS618ColPtr := PSAPS618Col(aSAPS618.FList[icol]);
        aSAPS618ColPtr^.dqty := aSAPS618ColPtr^.dqty - iQty;
        iQty := 0;
      end;    

      Break;
    end;
  end;
end;  

procedure TSAPPIRReader.Log(const str: string);
begin

end;

function TSAPPIRReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPPIRReader.GetItems(i: Integer): TSAPS618;
begin
  Result := TSAPS618(FList.Objects[i]);
end;

procedure TSAPPIRReader.Open;
var
  irow: Integer;
  stitle1, stitle2, stitle3, stitle4,
  stitle5, stitle6, stitle7, stitle8: string;
  stitle: string;
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  snumber: string;
  sname: string;
  smrparea: string;
  sletter: string;
  aSAPS618: TSAPS618;
  sweek: string;
  icol: Integer;
  iweek: Integer;
  dqty: Double;
  dt1, dt2: TDateTime; 
  swk: string;
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

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle8 := ExcelApp.Cells[irow, 8].Value;

        stitle := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6 + stitle7 + stitle8;
        if stitle <> '物料物料描述物料组物料组描述工厂需求类型版本Act' then
        begin
          Log('sheet ' + sSheet + '  格式不符');
          Continue;
        end;

        icol := 13;
        sweek := ExcelApp.Cells[irow, icol].Value;
        while sweek <> '' do
        begin
          sletter := Copy(sweek, 1, 1);
          sletter := UpperCase(sletter);
          dt1 := 0;
          if sletter = 'W' then
          begin
            dt1 := EncodeDateWeek( StrToInt( Copy(sweek, 2, 4) ), StrToInt( Copy(sweek, 6, 2) ));
          end
          else if sletter = 'D' then
          begin
            dt1 := EncodeDate(StrToInt( Copy(sweek, 2, 4) ), StrToInt( Copy(sweek, 6, 2) ), StrToInt( Copy(sweek, 8, 2) ));
          end      
          else if sletter = 'M' then
          begin
            dt1 := EncodeDate(StrToInt( Copy(sweek, 2, 4) ), StrToInt( Copy(sweek, 6, 2) ), 1);
          end;
          swk := FormatDateTime('yyyy-MM-dd', dt1);
          slWeek.AddObject(sweek + '=' + swk, TObject(icol));

          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          if FNumbers.IndexOf(snumber) < 0 then
          begin
            FNumbers.Add(snumber);
          end;
          
          sname := ExcelApp.Cells[irow, 2].Value;
          smrparea := ExcelApp.Cells[irow, 10].Value;
          aSAPS618 := TSAPS618.Create(snumber, smrparea);
          aSAPS618.sname := sname;

          aSAPS618.sgroup := ExcelApp.Cells[irow, 3].Value;
          aSAPS618.sgroupname := ExcelApp.Cells[irow, 4].Value;
          aSAPS618.sfac := ExcelApp.Cells[irow, 5].Value;
          aSAPS618.sDemandType := ExcelApp.Cells[irow, 6].Value;
          aSAPS618.sDemandVer := ExcelApp.Cells[irow, 7].Value;
          aSAPS618.sAct := ExcelApp.Cells[irow, 8].Value;
          aSAPS618.sPlanNo := ExcelApp.Cells[irow, 9].Value;
          aSAPS618.FMrpArea := ExcelApp.Cells[irow, 10].Value;
          aSAPS618.sMrper := ExcelApp.Cells[irow, 11].Value;
          aSAPS618.sUnit := ExcelApp.Cells[irow, 12].Value;

          FList.AddObject(snumber, aSAPS618);
 
          for iweek := 0 to slWeek.Count - 1 do
          begin
            sweek := slWeek.Names[iweek];
            icol := Integer(slWeek.Objects[iweek]);
            dqty := ExcelApp.Cells[irow, icol].Value;
            dt1 := myStrToDateTime( slWeek.ValueFromIndex[iweek] );
            dt2 := myStrToDateTime( slWeek.ValueFromIndex[iweek] );
            aSAPS618.Add(snumber, sname, sweek, dt1, dt2, dqty);
          end;

          irow := irow + 1;
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
