unit CPInAndStockReader;

interface

uses
  Classes, SysUtils, ComObj, CommUtils, ADODB, SAPMB51Reader2, SAPStockReader2,
  SAPMrpAreaStockReader;

type                       
  TCPInAndStorkColHeader = packed record
    icol: Integer;
    sml: string;
    sfox: string;
    sdate: string;
  end;
  PCPInAndStorkColHeader = ^TCPInAndStorkColHeader;

  TCPInAndStorkCol = packed record
    dQtyML: Double;
    dQtyFox: Double;
  end;
  PCPInAndStorkCol = ^TCPInAndStorkCol;

  TCPInAndStorkLine = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PCPInAndStorkCol;
  public                 
    snumber: string;
    sname: string;  
    dCPInAcct: Double;
    dCPInYesterday: Double;      
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    procedure AddTodayStock(aSAPStockReader2: TSAPStockReader2;
      aSAPMrpAreaStockReader: TSAPMrpAreaStockReader );
    property Count: Integer read GetCount;
    property Items[i: Integer]: PCPInAndStorkCol read GetItems;
  end;

  TCPInAndStorkProj = class
  private
    FList: TStringList;
    function GetCount: Integer;
    function GetItems(i: Integer): TCPInAndStorkLine;
  public
    sname: string;
    sno: string;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    procedure AddCPIn(aSAPMB51RecordPtr: PSAPMB51Record; lstDate: TList);
    procedure AddTodayStock(aSAPStockReader2: TSAPStockReader2;
      aSAPMrpAreaStockReader: TSAPMrpAreaStockReader; lstDate: TList );
    property Count: Integer read GetCount;
    property Items[i: Integer]: TCPInAndStorkLine read GetItems;
  end;

  TCPInAndStockReader = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent; 
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TCPInAndStorkProj;
  public
    lstDate: TList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    procedure AddCPIn(aSAPMB51RecordPtr: PSAPMB51Record; const sproj: string);
    procedure AddTodayStock( const dt: TDateTime;
      aSAPStockReader2: TSAPStockReader2;
      aSAPMrpAreaStockReader: TSAPMrpAreaStockReader );
    property Count: Integer read GetCount;
    property Items[i: Integer]: TCPInAndStorkProj read GetItems;
  end;
 

implementation
                
{ TCPInAndStorkLine }

constructor TCPInAndStorkLine.Create;
begin
  FList := TList.Create;
  
end;

destructor TCPInAndStorkLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TCPInAndStorkLine.Clear;
var
  i: Integer;
  p: PCPInAndStorkCol;
begin
  for i := 0 to FList.Count -1 do
  begin
    p := PCPInAndStorkCol(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

procedure TCPInAndStorkLine.AddTodayStock(aSAPStockReader2: TSAPStockReader2;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader );
var
  p: PCPInAndStorkCol;
  aSAPMrpAREA: TSAPMrpAREA;
  sl: TStringList;
  i: Integer;
begin
  p := New(PCPInAndStorkCol);
  FList.Add(p);

  if self.snumber = '03.43.3430030' then
  begin
    Sleep(1);
  end;

  for i := 0 to aSAPMrpAreaStockReader.Count - 1 do
  begin
    aSAPMrpAREA := aSAPMrpAreaStockReader.Items[i];

    if aSAPMrpAREA.sAreaName = '魅力MRP区域' then
    begin
      p^.dQtyML := aSAPStockReader2.GetStocks(snumber, aSAPMrpAREA.FList);
    end;

    if aSAPMrpAREA.sAreaName = '富士康MRP区域' then
    begin
      p^.dQtyFox := aSAPStockReader2.GetStocks(snumber, aSAPMrpAREA.FList);
    end;
  end;

end;  

function TCPInAndStorkLine.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TCPInAndStorkLine.GetItems(i: Integer): PCPInAndStorkCol;
begin
  Result := PCPInAndStorkCol(FList[i]);
end;  
 
{ TCPInAndStorkProj }

constructor TCPInAndStorkProj.Create;
begin
  FList := TStringList.Create;
end;

destructor TCPInAndStorkProj.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TCPInAndStorkProj.Clear;
var
  i: Integer;
  a: TCPInAndStorkLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    a := TCPInAndStorkLine(FList.Objects[i]);
    a.Free;
  end;
  FList.Clear;  
end;

procedure TCPInAndStorkProj.AddCPIn(aSAPMB51RecordPtr: PSAPMB51Record;
  lstDate: TList);
var
  idx: Integer;
  aCPInAndStorkLine: TCPInAndStorkLine;
  i: Integer;
  p: PCPInAndStorkColHeader;
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader;
begin
  idx := FList.IndexOf(aSAPMB51RecordPtr^.snumber);
  if idx < 0 then
  begin
    aCPInAndStorkLine := TCPInAndStorkLine.Create;
    aCPInAndStorkLine.snumber := aSAPMB51RecordPtr^.snumber;
    aCPInAndStorkLine.sname := aSAPMB51RecordPtr^.sname;
    aCPInAndStorkLine.dCPInAcct := 0;
    aCPInAndStorkLine.dCPInYesterday := aSAPMB51RecordPtr^.dqty;

    for i := 0 to lstDate.Count - 1 do
    begin
      aCPInAndStorkColHeaderPtr := PCPInAndStorkColHeader(lstDate[i]);
      p := New(PCPInAndStorkColHeader);
      p^ := aCPInAndStorkColHeaderPtr^;
      aCPInAndStorkLine.FList.Add(p);
    end;
  end
  else
  begin
    aCPInAndStorkLine := Items[idx];
    aCPInAndStorkLine.dCPInYesterday :=
      aCPInAndStorkLine.dCPInYesterday + aSAPMB51RecordPtr^.dqty;
  end;
end;

procedure TCPInAndStorkProj.AddTodayStock(aSAPStockReader2: TSAPStockReader2;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader; lstDate: TList);
var
  i: Integer;
  idate: Integer;
  aCPInAndStorkLine: TCPInAndStorkLine;
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader;
  aSAPStockRecordPtr: PSAPStockRecord;
  aCPInAndStorkColPtr: PCPInAndStorkCol; 
begin

  for i := 0 to aSAPStockReader2.Count - 1 do
  begin
    aSAPStockRecordPtr := aSAPStockReader2.Items[i];
    if Copy(aSAPStockRecordPtr^.snumber, 1, 5) = self.sno then
    begin
      if FList.IndexOf(aSAPStockRecordPtr^.snumber) < 0 then // 不存在
      begin
        aCPInAndStorkLine := TCPInAndStorkLine.Create;
        FList.AddObject(aSAPStockRecordPtr^.snumber, aCPInAndStorkLine);

        aCPInAndStorkLine.snumber := aSAPStockRecordPtr^.snumber;
        aCPInAndStorkLine.sname := aSAPStockRecordPtr^.sname;

        for idate := 0 to lstDate.Count - 1 do 
        begin
          aCPInAndStorkColHeaderPtr := PCPInAndStorkColHeader(lstDate[idate]);

          aCPInAndStorkColPtr := New(PCPInAndStorkCol);
          aCPInAndStorkColPtr^.dQtyML := 0;
          aCPInAndStorkColPtr^.dQtyFox := 0;
          aCPInAndStorkLine.FList.Add(aCPInAndStorkColPtr);
        end;
      end;  
    end;
  end;

  for i := 0 to FList.Count - 1 do
  begin
    aCPInAndStorkLine := TCPInAndStorkLine(FList.Objects[i]);
    aCPInAndStorkLine.AddTodayStock( aSAPStockReader2, aSAPMrpAreaStockReader );
  end;

end;  

function TCPInAndStorkProj.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TCPInAndStorkProj.GetItems(i: Integer): TCPInAndStorkLine;
begin
  Result := TCPInAndStorkLine(FList.Objects[i]);
end;  
 
{ TCPInAndStockReader }

constructor TCPInAndStockReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  
  lstDate := TList.Create;

  Open;
end;

destructor TCPInAndStockReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TCPInAndStockReader.Clear;
var
  i: Integer;
  a: TCPInAndStorkProj;
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader;
begin
  for i := 0 to FList.Count - 1 do
  begin
    a := TCPInAndStorkProj(FList.Objects[i]);
    a.Free;
  end;
  FList.Clear;

          for i := 0 to lstDate.Count -1  do
          begin
            aCPInAndStorkColHeaderPtr := PCPInAndStorkColHeader(lstDate[i]);
            Dispose(aCPInAndStorkColHeaderPtr);
          end;
          lstDate.Clear;  
end;

procedure TCPInAndStockReader.AddCPIn(aSAPMB51RecordPtr: PSAPMB51Record;
  const sproj: string);
var
  idx: Integer;
  aCPInAndStorkProj: TCPInAndStorkProj;
begin
  idx := FList.IndexOf(sproj);
  if idx < 0 then
  begin
    aCPInAndStorkProj := TCPInAndStorkProj.Create;
  end
  else
  begin
    aCPInAndStorkProj := Items[idx];
  end;

  aCPInAndStorkProj.AddCPIn(aSAPMB51RecordPtr, lstDate);
end;

procedure TCPInAndStockReader.AddTodayStock(const dt: TDateTime;
  aSAPStockReader2: TSAPStockReader2;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader );
var
  i: Integer;
  a: TCPInAndStorkProj;       
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader; 
begin
  for i := 0 to FList.Count - 1 do
  begin
    a := TCPInAndStorkProj(FList.Objects[i]);
    a.AddTodayStock(aSAPStockReader2, aSAPMrpAreaStockReader, lstDate);
  end;

  // 这个要放在后面 
  aCPInAndStorkColHeaderPtr := New(PCPInAndStorkColHeader);
  aCPInAndStorkColHeaderPtr^.icol := -1; 
  aCPInAndStorkColHeaderPtr^.sml := FormatDateTime('MM.dd', dt) + '魅力';
  aCPInAndStorkColHeaderPtr^.sfox := FormatDateTime('MM.dd', dt) + '富士康';
  aCPInAndStorkColHeaderPtr^.sdate := FormatDateTime('MM月dd日', dt);
  lstDate.Add(aCPInAndStorkColHeaderPtr);

end;

function TCPInAndStockReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TCPInAndStockReader.GetItems(i: Integer): TCPInAndStorkProj;
begin
  Result := TCPInAndStorkProj(FList.Objects[i]);
end;

procedure TCPInAndStockReader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TCPInAndStockReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  snumber: string;   

  sdate: string;  
  aCPInAndStorkColHeaderPtr: PCPInAndStorkColHeader;
  aCPInAndStorkColPtr: PCPInAndStorkCol;

  sproj: string;
  aCPInAndStorkProj: TCPInAndStorkProj;
  aCPInAndStorkLine: TCPInAndStorkLine;

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

        irow := 2;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4;
        if stitle <> '机型料号描述当月累计入库数' then
        begin
          Log(sSheet +'  不是 成品入库与库存格式 格式(机型料号描述当月累计入库数)');
          Continue;
        end;


          icol := 6;
          sdate := ExcelApp.Cells[irow, icol].Value;
          while sdate <> '' do
          begin
            aCPInAndStorkColHeaderPtr := New(PCPInAndStorkColHeader);
            aCPInAndStorkColHeaderPtr^.sml := ExcelApp.Cells[irow, icol].Value;
            aCPInAndStorkColHeaderPtr^.sfox := ExcelApp.Cells[irow, icol + 1].Value;
            aCPInAndStorkColHeaderPtr^.sdate := ExcelApp.Cells[irow, icol + 2].Value;
            aCPInAndStorkColHeaderPtr^.icol := icol;
            lstDate.Add(aCPInAndStorkColHeaderPtr);
            icol := icol + 3;
            sdate := ExcelApp.Cells[irow, icol].Value;
          end;

          irow := 3;
          sproj := ExcelApp.Cells[irow, 1].Value;
          while sproj <> '' do
          begin
            aCPInAndStorkProj := TCPInAndStorkProj.Create;
            aCPInAndStorkProj.sname := sproj;
            FList.AddObject(sproj, aCPInAndStorkProj);

            snumber := ExcelApp.Cells[irow, 2].Value;
            while snumber<> '' do
            begin
              aCPInAndStorkProj.sno := Copy(snumber, 1, 5);
              
              aCPInAndStorkLine := TCPInAndStorkLine.Create;
              aCPInAndStorkProj.FList.AddObject(snumber, aCPInAndStorkLine);
              aCPInAndStorkLine.snumber := snumber;
              aCPInAndStorkLine.sname := ExcelApp.Cells[irow, 3].Value;

              for idate := 0 to lstDate.Count - 1 do
              begin
                aCPInAndStorkColHeaderPtr := PCPInAndStorkColHeader(lstDate[idate]);

                aCPInAndStorkColPtr := New(PCPInAndStorkCol);
                aCPInAndStorkColPtr^.dQtyML := ExcelApp.Cells[irow, aCPInAndStorkColHeaderPtr^.icol].Value;         
                aCPInAndStorkColPtr^.dQtyFox := ExcelApp.Cells[irow, aCPInAndStorkColHeaderPtr^.icol + 1].Value;
                aCPInAndStorkLine.FList.Add(aCPInAndStorkColPtr);
              end;
                                                           
              irow := irow + 1;
              snumber := ExcelApp.Cells[irow, 2].Value;

              if IsCellMerged(ExcelApp, irow, 2, irow, 3) then Break;
            end;  
            
            
            irow := irow + 1;
            sproj := ExcelApp.Cells[irow, 1].Value;
          end;


        Break; //读一页就好
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
