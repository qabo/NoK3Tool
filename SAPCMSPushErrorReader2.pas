unit SAPCMSPushErrorReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TCMSPushError = packed record
    sbillno: string; //���ݱ��
    sbilltype: string; //��������
    dt: TDateTime; //����
    dtCreate: TDateTime; //��������
    sfac: string; //����������
    serrtype: string; //��������
    serreason: string; //����ԭ��
    smsg1: string; //��Ϣ1
    smsg2: string; //��Ϣ2
    smsg3: string; //��Ϣ3
    smsg4: string; //��Ϣ4
    scount: string; //����

  end;
  PCMSPushError = ^TCMSPushError;

  TSAPCMSPushErrorReader2 = class
  private         
    FList: TList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PCMSPushError;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PCMSPushError read GetItems;
  end;
 

implementation

{ TSAPCMSPushErrorReader2 }

constructor TSAPCMSPushErrorReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TSAPCMSPushErrorReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPCMSPushErrorReader2.Clear;
var
  i: Integer;
  p: PCMSPushError;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PCMSPushError(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPCMSPushErrorReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPCMSPushErrorReader2.GetItems(i: Integer): PCMSPushError;
begin
  Result := PCMSPushError(FList[i]);
end;

procedure TSAPCMSPushErrorReader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TSAPCMSPushErrorReader2.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPOPOAllocPtr: PCMSPushError;
   
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

    ADOTabXLS.TableName:='['+'���ʹ����б�'+'$]';

    ADOTabXLS.Active:=true;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PCMSPushError);
      FList.Add(aSAPOPOAllocPtr);
      
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('���ݱ��').AsString;
      aSAPOPOAllocPtr^.sbilltype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.dt := ADOTabXLS.FieldByName('����').AsDateTime;
      aSAPOPOAllocPtr^.dtCreate := ADOTabXLS.FieldByName('��������').AsDateTime;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('����������').AsString;
      aSAPOPOAllocPtr^.serrtype := ADOTabXLS.FieldByName('��������').AsString;
      aSAPOPOAllocPtr^.serreason := ADOTabXLS.FieldByName('����ԭ��').AsString;
      aSAPOPOAllocPtr^.smsg1 := ADOTabXLS.FieldByName('��Ϣ1').AsString;
      aSAPOPOAllocPtr^.smsg2 := ADOTabXLS.FieldByName('��Ϣ2').AsString;
      aSAPOPOAllocPtr^.smsg3 := ADOTabXLS.FieldByName('��Ϣ3').AsString;
      aSAPOPOAllocPtr^.smsg4 := ADOTabXLS.FieldByName('��Ϣ4').AsString;
      aSAPOPOAllocPtr^.scount := ADOTabXLS.FieldByName('����').AsString;

 
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
