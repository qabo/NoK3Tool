unit SAPWhereUseReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB, SAPMaterialReader;

type
  TSAPWhereUseRecord = packed record
    snumber: string;
    swhereuse: string;
  end;
  PSAPWhereUseRecord = ^TSAPWhereUseRecord;

  TSAPWhereUseReader = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string); 
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPWhereUseRecord;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear; 
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPWhereUseRecord read GetItems;
    function FindItem(const snumber: string): PSAPWhereUseRecord;
    function GetWhereUse(const snumber: string): string;
  end;

implementation
       
{ TSAPWhereUseReader }

constructor TSAPWhereUseReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPWhereUseReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPWhereUseReader.Clear;
var
  i: Integer;
  p: PSAPWhereUseRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPWhereUseRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPWhereUseReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPWhereUseReader.GetItems(i: Integer): PSAPWhereUseRecord;
begin
  Result := PSAPWhereUseRecord(FList.Objects[i]);
end;

function TSAPWhereUseReader.FindItem(const snumber: string): PSAPWhereUseRecord;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PSAPWhereUseRecord(FList.Objects[idx]);
  end
  else Result := nil;
end;

function TSAPWhereUseReader.GetWhereUse(const snumber: string): string;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PSAPWhereUseRecord(FList.Objects[idx])^.swhereuse;
  end
  else Result := '';
end;

procedure TSAPWhereUseReader.Log(const str: string);
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

procedure TSAPWhereUseReader.Open;
const
  CSNumber = '物料编码';
  CSName = '物料描述';
  CSRecvTime = '收货处理时间';
  CSMRPType = 'MRP类型';
  CSSPQ = '舍入值';
  CSLT_PD = '计划交货时间';
  CSLT_M0 = '自制生产时间';
  
  CSLT_MOQ = '最小批量大小';
  CSMRPGroup = 'MRP组';
  CSPType = '采购类型';
  
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  aSAPWhereUseRecordPtr: PSAPWhereUseRecord;
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

    ADOTabXLS.TableName:='['+'Sheet1'+'$]';

    ADOTabXLS.Active:=true;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      snumber := ADOTabXLS.FieldByName('物料').AsString;  // ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;
                
      aSAPWhereUseRecordPtr := New(PSAPWhereUseRecord);
      FList.AddObject(snumber, TObject(aSAPWhereUseRecordPtr));

      aSAPWhereUseRecordPtr^.sNumber := snumber;
      aSAPWhereUseRecordPtr^.swhereuse := ADOTabXLS.FieldByName('在产项目').AsString;  //   ExcelApp.Cells[irow, iColName].Value; 
 
             
      ADOTabXLS.Next;
    end;


    ADOTabXLS.Close;

    Conn.Connected := False;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;

  FList.Sort;
end;
 
end.
