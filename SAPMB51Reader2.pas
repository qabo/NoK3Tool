unit SAPMB51Reader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TSAPMB51Record = packed record
    sbillno: string; //����ƾ֤
    sentryid: string; //����ƾ֤��Ŀ
    fbilldate: TDateTime; //ƾ֤����
    fstockno: string; //���ص�
    fstockname: string; //�ִ��ص������
    fnote: string; //ƾ̧֤ͷ�ı�   ���ڴ洢 ����������
    smovingtype: string; //�ƶ�����
    smovingtypeText: string; //�ƶ������ı�
    snumber: string; //����
    sname: string; //��������
    dqty: Double; //��¼�뵥λ��ʾ������
    fdate: TDateTime; //��������
    finputdate: TDateTime; //��������
    finputtime: TDateTime; //����ʱ��
    spo: string; //����
    sbillno_po: string; //�ɹ�����
    snote_entry: string; // �ı�

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
    sbillno: string; //����ƾ֤      
    snumber: string; //����
    sname: string; //��������      
    dqty: Double; //��¼�뵥λ��ʾ������     
    smovingtype: string; //�ƶ�����
    splnt: string;
    fstockno: string; //���ص�
    fstockname: string; //�ִ��ص������
    stext: string; //ƾ̧֤ͷ�ı�   ���ڴ洢 ����������    
    sorder: string; //����
    sref: string;   
    fdate: TDateTime; //��������
    finputdate: TDateTime; //��������
    finputtime: TDateTime; //����ʱ��
    sheadtext: string;
    susername: string;
    splnt_name: string; 
    smovingtypeText: string; //�ƶ������ı�
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

      snumber := ADOTabXLS.FieldByName('����').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;

      aSAPOPOAllocPtr := New(PSAPMB51Record);
      aSAPOPOAllocPtr^.bCalc := False;


      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('����ƾ֤').AsString;
      //aSAPOPOAllocPtr^.sentryid := ADOTabXLS.FieldByName('����ƾ֤��Ŀ').AsString;
      aSAPOPOAllocPtr^.fbilldate := ADOTabXLS.FieldByName('ƾ֤����').AsDateTime;
      if ADOTabXLS.FindField('��λ') <> nil then
      begin
        aSAPOPOAllocPtr^.fstockno := ADOTabXLS.FieldByName('��λ').AsString;
      end
      else
      begin
        aSAPOPOAllocPtr^.fstockno := ADOTabXLS.FieldByName('���ص�').AsString;
      end;  
      aSAPOPOAllocPtr^.fstockname := ADOTabXLS.FieldByName('�ִ��ص������').AsString;
      aSAPOPOAllocPtr^.fnote := ADOTabXLS.FieldByName('ƾ̧֤ͷ�ı�').AsString;  //   ���ڴ洢 ����������
      if ADOTabXLS.FindField('MvT') <> nil then
      begin
        aSAPOPOAllocPtr^.smovingtype := ADOTabXLS.FieldByName('MvT').AsString;
      end
      else
      begin
        aSAPOPOAllocPtr^.smovingtype := ADOTabXLS.FieldByName('�ƶ�����').AsString;
      end;

//      aSAPOPOAllocPtr^.smovingtypeText := ADOTabXLS.FieldByName('�ƶ������ı�').AsString;
      aSAPOPOAllocPtr^.snumber := snumber;
      aSAPOPOAllocPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      if ADOTabXLS.FindField('����') <> nil then
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����').AsFloat;
      end
      else if ADOTabXLS.FindField('����(¼�뵥λ)') <> nil then
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('����(¼�뵥λ)').AsFloat;
      end
      else
      begin
        aSAPOPOAllocPtr^.dqty := ADOTabXLS.FieldByName('��¼�뵥λ��ʾ������').AsFloat;
      end;

      aSAPOPOAllocPtr^.fdate := ADOTabXLS.FieldByName('��������').AsDateTime;
      aSAPOPOAllocPtr^.finputdate := ADOTabXLS.FieldByName('��������').AsDateTime;
      aSAPOPOAllocPtr^.finputtime := ADOTabXLS.FieldByName('����ʱ��').AsDateTime;
      aSAPOPOAllocPtr^.spo := ADOTabXLS.FieldByName('����').AsString;
      aSAPOPOAllocPtr^.sbillno_po := ADOTabXLS.FieldByName('�ɹ�����').AsString;
      aSAPOPOAllocPtr^.snote_entry := ADOTabXLS.FieldByName('�ı�').AsString;

      if aSAPOPOAllocPtr^.smovingtype = '101' then
      begin
        FWinBCount := FWinBCount + 1;
      end;
                


      FList.AddObject(snumber, TObject(aSAPOPOAllocPtr));

      
      irow := irow + 1;
      snumber := ADOTabXLS.FieldByName('����').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;

          
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
      aSAPOPOAllocPtr^.fstockname := ADOTabXLS.FieldByName('��ʿ����λ').AsString;
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
