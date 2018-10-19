unit SAPMB51Reader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TSAPMB51Record = packed record
    sbillno: string; //物料凭证
    sentryid: string; //物料凭证项目
    fbilldate: TDateTime; //凭证日期
    fstockno: string; //库存地点
    fstockname: string; //仓储地点的描述
    fnote: string; //凭证抬头文本   用于存储 代工厂单号
    smovingtype: string; //移动类型
    smovingtypeText: string; //移动类型文本
    snumber: string; //物料
    sname: string; //物料描述
    dqty: Double; //以录入单位表示的数量
    fdate: TDateTime; //过账日期
    finputdate: TDateTime; //输入日期
    finputtime: TDateTime; //输入时间
    spo: string; //订单
    sbillno_po: string; //采购订单
    snote_entry: string; // 文本

    bCalc: Boolean;
    sMatchType: string;
  end;
  PSAPMB51Record= ^TSAPMB51Record;

  TSAPMB51Reader2 = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FWinBCount: Integer;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPMB51Record;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function GetMB51Qty101(aSAPMB51RecordPtr: PSAPMB51Record): Double;
    procedure SetCalcFlag(aSAPMB51RecordPtr: PSAPMB51Record; const sMatchType: string);
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPMB51Record read GetItems;
    property WinBCount: Integer read FWinBCount;
  end;

  TSAPMB51RecordFox = packed record
    sbillno: string; //物料凭证      
    snumber: string; //物料
    sname: string; //物料描述      
    dqty: Double; //以录入单位表示的数量     
    smovingtype: string; //移动类型
    splnt: string;
    fstockno: string; //库存地点
    fstockname: string; //仓储地点的描述
    stext: string; //凭证抬头文本   用于存储 代工厂单号    
    sorder: string; //订单
    sref: string;   
    fdate: TDateTime; //过账日期
    finputdate: TDateTime; //输入日期
    finputtime: TDateTime; //输入时间
    sheadtext: string;
    susername: string;
    splnt_name: string; 
    smovingtypeText: string; //移动类型文本
    sitem: string;
     
    bCalc: Boolean;
    sMatchType: string;
  end;
  PSAPMB51RecordFox = ^TSAPMB51RecordFox;

  TSAPMB51ReaderFox = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FWinBCount: Integer;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPMB51RecordFox;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPMB51RecordFox read GetItems;
    property WinBCount: Integer read FWinBCount;
  end;

implementation

{ TSAPMB51Reader2 }

constructor TSAPMB51Reader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPMB51Reader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPMB51Reader2.Clear;
var
  i: Integer;
  p: PSAPMB51Record;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPMB51Record(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPMB51Reader2.GetMB51Qty101(aSAPMB51RecordPtr: PSAPMB51Record): Double;
var
  i: Integer;
  p: PSAPMB51Record;
begin
  Result := 0;
  for i := 0 to self.Count - 1 do
  begin
    p := self.Items[i];

    if p^.bCalc then Continue;

    if (p^.fnote = aSAPMB51RecordPtr^.fnote) and
      (p^.snumber = aSAPMB51RecordPtr^.snumber) and
      (p^.sbillno_po = aSAPMB51RecordPtr^.sbillno_po) then
    begin
      Result := Result + p^.dqty;
    end;
  end;
end;

procedure TSAPMB51Reader2.SetCalcFlag(aSAPMB51RecordPtr: PSAPMB51Record;
  const sMatchType: string);
var
  i: Integer;
  p: PSAPMB51Record;
begin 
  for i := 0 to self.Count - 1 do
  begin
    p := self.Items[i];
    if (p^.fnote = aSAPMB51RecordPtr^.fnote) and
      (p^.snumber = aSAPMB51RecordPtr^.snumber) and
      (p^.sbillno_po = aSAPMB51RecordPtr^.sbillno_po) then
    begin
      p^.bCalc := True;
      p^.sMatchType := sMatchType;
    end;
  end;
end;  

function TSAPMB51Reader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMB51Reader2.GetItems(i: Integer): PSAPMB51Record;
begin
  Result := PSAPMB51Record(FList.Objects[i]);
end;

procedure TSAPMB51Reader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TSAPMB51Reader2.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PSAPMB51Record;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'Sheet1'+'$]';

    ADOTabXLS.Active:=true;

    FWinBCount := 0;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      snumber := ADOTabXLS.FieldByName('物料').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;

      aSAPOPOAllocPtr := New(PSAPMB51Record);
      aSAPOPOAllocPtr^.bCalc := False;


      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('物料凭证').AsString;
      //aSAPOPOAllocPtr^.sentryid := ADOTabXLS.FieldByName('物料凭证项目').AsString;
      aSAPOPOAllocPtr^.fbilldate := ADOTabXLS.FieldByName('凭证日期').AsDateTime;
      if ADOTabXLS.FindField('库位') <> nil then
      begin
        aSAPOPOAllocPtr^.fstockno := ADOTabXLS.FieldByName('库位').AsString;
      end
      else
      begin
        aSAPOPOAllocPtr^.fstockno := ADOTabXLS.FieldByName('库存地点').AsString;
      end;  
      aSAPOPOAllocPtr^.fstockname := ADOTabXLS.FieldByName('仓储地点的描述').AsString;
      aSAPOPOAllocPtr^.fnote := ADOTabXLS.FieldByName('凭证抬头文本').AsString;  //   用于存储 代工厂单号
      if ADOTabXLS.FindField('MvT') <> nil then
      begin
        aSAPOPOAllocPtr^.smovingtype := ADOTabXLS.FieldByName('MvT').AsString;
      end
      else
      begin
        aSAPOPOAllocPtr^.smovingtype := ADOTabXLS.FieldByName('移动类型').AsString;
      end;

//      aSAPOPOAllocPtr^.smovingtypeText := ADOTabXLS.FieldByName('移动类型文本').AsString;
      aSAPOPOAllocPtr^.snumber := snumber;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('物料描述').AsString;
      if ADOTabXLS.FindField('数量') <> nil then
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量').AsFloat;
      end
      else if ADOTabXLS.FindField('数量(录入单位)') <> nil then
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('数量(录入单位)').AsFloat;
      end
      else
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('以录入单位表示的数量').AsFloat;
      end;

      aSAPOPOAllocPtr^.fdate := ADOTabXLS.FieldByName('过账日期').AsDateTime;
      aSAPOPOAllocPtr^.finputdate := ADOTabXLS.FieldByName('输入日期').AsDateTime;
      aSAPOPOAllocPtr^.finputtime := ADOTabXLS.FieldByName('输入时间').AsDateTime;
      aSAPOPOAllocPtr^.spo := ADOTabXLS.FieldByName('订单').AsString;
      aSAPOPOAllocPtr^.sbillno_po := ADOTabXLS.FieldByName('采购订单').AsString;
      aSAPOPOAllocPtr^.snote_entry := ADOTabXLS.FieldByName('文本').AsString;

      if aSAPOPOAllocPtr^.smovingtype = '101' then
      begin
        FWinBCount := FWinBCount + 1;
      end;
                


      FList.AddObject(snumber, TObject(aSAPOPOAllocPtr));

      
      irow := irow + 1;
      snumber := ADOTabXLS.FieldByName('物料').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;

          
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;

////////////////////////////////////////////////////////////////////////////////



{ TSAPMB51ReaderFox }

constructor TSAPMB51ReaderFox.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPMB51ReaderFox.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPMB51ReaderFox.Clear;
var
  i: Integer;
  p: PSAPMB51RecordFox;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPMB51RecordFox(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPMB51ReaderFox.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMB51ReaderFox.GetItems(i: Integer): PSAPMB51RecordFox;
begin
  Result := PSAPMB51RecordFox(FList.Objects[i]);
end;

procedure TSAPMB51ReaderFox.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TSAPMB51ReaderFox.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer; 
  aSAPOPOAllocPtr: PSAPMB51RecordFox;
   
  Conn: TADOConnection;
  ADOTabXLS: TADOTable;

  snumber: string;
  sdt: string;
begin
  Clear;

  if not FileExists(FFile) then Exit;


  ADOTabXLS := TADOTable.Create(nil);
  Conn:=TADOConnection.Create(nil);

  Conn.ConnectionString:='Provider=Microsoft.ACE.OLEDB.12.0;Data Source="' + FFile + '";Extended Properties=excel 8.0;Persist Security Info=False';

  Conn.LoginPrompt:=false;

  try

    Conn.Connected:=true;

    ADOTabXLS.Connection:=Conn;

    ADOTabXLS.TableName:='['+'Sheet1'+'$]';

    ADOTabXLS.Active:=true;

    FWinBCount := 0;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      snumber := ADOTabXLS.FieldByName('Material').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then 
      begin
        Continue;
      end;

      aSAPOPOAllocPtr := New(PSAPMB51RecordFox);
      aSAPOPOAllocPtr^.bCalc := False;
      FList.AddObject(snumber, TObject(aSAPOPOAllocPtr));

      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('Mat. doc.').AsString;
      aSAPOPOAllocPtr^.snumber := ADOTabXLS.FieldByName('Material').AsString;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('Material description').AsString;
      aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('   Quantity').AsFloat;
      aSAPOPOAllocPtr^.smovingtype := ADOTabXLS.FieldByName('MvT').AsString;
      aSAPOPOAllocPtr^.splnt := ADOTabXLS.FieldByName('Plnt').AsString;
      aSAPOPOAllocPtr^.fstockno := ADOTabXLS.FieldByName('SLoc').AsString;
      aSAPOPOAllocPtr^.fstockname := ADOTabXLS.FieldByName('富士康仓位').AsString;
      aSAPOPOAllocPtr^.stext := ADOTabXLS.FieldByName('Text').AsString;
      aSAPOPOAllocPtr^.sorder := ADOTabXLS.FieldByName('Order').AsString;
      aSAPOPOAllocPtr^.sref := ADOTabXLS.FieldByName('Reference').AsString;
      sdt := ADOTabXLS.FieldByName('Entry date').AsString;
      aSAPOPOAllocPtr^.fdate := myStrToDateTime(sdt);     
      sdt := ADOTabXLS.FieldByName('Pstg date').AsString;
      aSAPOPOAllocPtr^.finputdate := myStrToDateTime(sdt);
      aSAPOPOAllocPtr^.finputtime := ADOTabXLS.FieldByName('Time').AsDateTime;
      aSAPOPOAllocPtr^.sheadtext := ADOTabXLS.FieldByName('HeadText').AsString;
      aSAPOPOAllocPtr^.susername := ADOTabXLS.FieldByName('User name').AsString;
      aSAPOPOAllocPtr^.splnt_name := ADOTabXLS.FieldByName('Name 1').AsString;
      aSAPOPOAllocPtr^.smovingtypeText := ADOTabXLS.FieldByName('MvtTypeTxt').AsString;
      aSAPOPOAllocPtr^.sitem := ADOTabXLS.FieldByName('Item').AsString;
   
          
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;


end.
