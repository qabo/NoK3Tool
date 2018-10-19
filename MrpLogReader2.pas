unit MrpLogReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB, SAPMaterialReader;

type
  TMrpLogRecord = packed record
    id: Integer; //ID
    pid: Integer; //��ID
    snumber: string; //����
    sname: string; //��������
    dt: TDateTime; //��������
    dtReq: TDateTime; //�����µ�����
    dqty: Double; //��������
    dqtyStock: Double; //���ÿ��
    dqtyOPO: Double; //OPO
    dqtyNet: Double; //������
    sGroup: string; //�����
    sMrp: string; //MRP������
    sBuyer: string; //�ɹ�Ա
    sArea: string; //MRP����
    spnumber: string; //	�ϲ��Ϻ�
    srnumber: string; //	���Ϻ�
    slt: string; //	L/T
    bCalc: Boolean;
  end;
  PMrpLogRecord = ^TMrpLogRecord;



  TMrpLogReader2 = class
  private         
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string); 
    function GetCount: Integer;
    function GetItems(i: Integer): PMrpLogRecord;
  public          
    FList: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear; 
    property Count: Integer read GetCount;
    property Items[i: Integer]: PMrpLogRecord read GetItems;
    function GetSAPMaterialRecord(const snumber: string): PMrpLogRecord;
  end;

implementation
       
{ TMrpLogReader2 }

constructor TMrpLogReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TMrpLogReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TMrpLogReader2.Clear;
var
  i: Integer;
  p: PMrpLogRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PMrpLogRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TMrpLogReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TMrpLogReader2.GetItems(i: Integer): PMrpLogRecord;
begin
  Result := PMrpLogRecord(FList.Objects[i]);
end;

function TMrpLogReader2.GetSAPMaterialRecord(const snumber: string): PMrpLogRecord;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PMrpLogRecord(FList.Objects[idx]);
  end
  else Result := nil;
end;
 
procedure TMrpLogReader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function IndexOfCol(ExcelApp: Variant; irow: Integer; const scol: string): Integer;
var
  i: Integer;
  s: string;
begin
  Result := -1;
  for i := 1 to 50 do
  begin
    s := ExcelApp.Cells[irow, i].Value;
    if s = scol then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
var
  item1, item2: PMrpLogRecord;
begin
  item1 := PMrpLogRecord(List.Objects[Index1]);
  item2 := PMrpLogRecord(List.Objects[Index2]);
  if item1.id > item2.id then
    Result := 1
  else if item1.id < item2.id then
    Result := -1
  else Result := 0;
end;

procedure TMrpLogReader2.Open;
const
  CSNumber = '���ϱ���';
  CSName = '��������';
  CSRecvTime = '�ջ�����ʱ��';
  CSMRPType = 'MRP����';
  CSSPQ = '����ֵ';
  CSLT_PD = '�ƻ�����ʱ��';
  CSLT_M0 = '��������ʱ��';
  
  CSLT_MOQ = '��С������С';
  CSMRPGroup = 'MRP��';
  CSPType = '�ɹ�����';
  
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  aMrpLogRecordPtr: PMrpLogRecord;
  iColNumber: Integer;
  iColName: Integer;
  iColRecvTime: Integer;
  iColMRPType: Integer;
  iColSPQ: Integer;
  iColLT_PD: Integer;
  iColLT_M0: Integer;
  
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

    ADOTabXLS.TableName:='[Mrp Log$]';

    ADOTabXLS.Active:=true;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin 
      aMrpLogRecordPtr := New(PMrpLogRecord);
      aMrpLogRecordPtr.bCalc := False;

      aMrpLogRecordPtr^.id := ADOTabXLS.FieldByName('ID').AsInteger;
      aMrpLogRecordPtr^.pid := ADOTabXLS.FieldByName('��ID').AsInteger;
      aMrpLogRecordPtr^.snumber := ADOTabXLS.FieldByName('����').AsString;
      aMrpLogRecordPtr^.sname := ADOTabXLS.FieldByName('��������').AsString;
      aMrpLogRecordPtr^.dt := ADOTabXLS.FieldByName('��������').AsDateTime;
      aMrpLogRecordPtr^.dtReq := ADOTabXLS.FieldByName('�����µ�����').AsDateTime;
      aMrpLogRecordPtr^.dqty := ADOTabXLS.FieldByName('��������').AsFloat;
      aMrpLogRecordPtr^.dqtyStock := ADOTabXLS.FieldByName('���ÿ��').AsFloat;
      aMrpLogRecordPtr^.dqtyOPO := ADOTabXLS.FieldByName('OPO').AsFloat;
      aMrpLogRecordPtr^.dqtyNet := ADOTabXLS.FieldByName('������').AsFloat;
      aMrpLogRecordPtr^.sGroup := ADOTabXLS.FieldByName('�����').AsString;
      aMrpLogRecordPtr^.sMrp := ADOTabXLS.FieldByName('MRP������').AsString;
      aMrpLogRecordPtr^.sBuyer := ADOTabXLS.FieldByName('�ɹ�Ա').AsString;
      aMrpLogRecordPtr^.sArea := ADOTabXLS.FieldByName('MRP����').AsString;
      aMrpLogRecordPtr^.spnumber := ADOTabXLS.FieldByName('�ϲ��Ϻ�').AsString;
      aMrpLogRecordPtr^.srnumber := ADOTabXLS.FieldByName('���Ϻ�').AsString;
      aMrpLogRecordPtr^.slt := ADOTabXLS.FieldByName('L/T').AsString;
                                             
      FList.AddObject(aMrpLogRecordPtr^.snumber, TObject(aMrpLogRecordPtr));
      
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;

    FList.CustomSort( StringListSortCompare );
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end; 
end;
 
end.
