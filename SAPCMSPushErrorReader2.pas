unit SAPCMSPushErrorReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TCMSPushError = packed record
    sbillno: string; //单据编号
    sbilltype: string; //单据类型
    dt: TDateTime; //日期
    dtCreate: TDateTime; //创建日期
    sfac: string; //代工厂名称
    serrtype: string; //错误类型
    serreason: string; //错误原因
    smsg1: string; //信息1
    smsg2: string; //信息2
    smsg3: string; //信息3
    smsg4: string; //信息4
    scount: string; //次数

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

    ADOTabXLS.TableName:='['+'推送错误列表'+'$]';

    ADOTabXLS.Active:=true;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      aSAPOPOAllocPtr := New(PCMSPushError);
      FList.Add(aSAPOPOAllocPtr);
      
      aSAPOPOAllocPtr^.sbillno := ADOTabXLS.FieldByName('单据编号').AsString;
      aSAPOPOAllocPtr^.sbilltype := ADOTabXLS.FieldByName('单据类型').AsString;
      aSAPOPOAllocPtr^.dt := ADOTabXLS.FieldByName('日期').AsDateTime;
      aSAPOPOAllocPtr^.dtCreate := ADOTabXLS.FieldByName('创建日期').AsDateTime;
      aSAPOPOAllocPtr^.sfac := ADOTabXLS.FieldByName('代工厂名称').AsString;
      aSAPOPOAllocPtr^.serrtype := ADOTabXLS.FieldByName('错误类型').AsString;
      aSAPOPOAllocPtr^.serreason := ADOTabXLS.FieldByName('错误原因').AsString;
      aSAPOPOAllocPtr^.smsg1 := ADOTabXLS.FieldByName('信息1').AsString;
      aSAPOPOAllocPtr^.smsg2 := ADOTabXLS.FieldByName('信息2').AsString;
      aSAPOPOAllocPtr^.smsg3 := ADOTabXLS.FieldByName('信息3').AsString;
      aSAPOPOAllocPtr^.smsg4 := ADOTabXLS.FieldByName('信息4').AsString;
      aSAPOPOAllocPtr^.scount := ADOTabXLS.FieldByName('次数').AsString;

 
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
