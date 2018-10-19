unit SAPMaterialReader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB, SAPMaterialReader;

type
  (*
  TSAPMaterialRecord = packed record
    sNumber: string;
    sName: string;
    dRecvTime: Double;
    sMRPType: string; // PD mrp, 外购  M0 mps 自制半成品 
    dSPQ: Double; // = '舍入值';
    dLT_PD: Double; // = '计划交货时间';
    dLT_M0: Double; // = '自制生产时间';
    iLowestCode: Integer; // 低位码
    sMRPer: string;
    sBuyer: string;
  end;
  PSAPMaterialRecord = ^TSAPMaterialRecord;
  *) 
  TSAPMaterialReader2 = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string); 
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPMaterialRecord;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear; 
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPMaterialRecord read GetItems;
    function GetSAPMaterialRecord(const snumber: string): PSAPMaterialRecord;
    function GetMrpType( const snumber: string ): string;
  end;

implementation
       
{ TSAPMaterialReader2 }

constructor TSAPMaterialReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPMaterialReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPMaterialReader2.Clear;
var
  i: Integer;
  p: PSAPMaterialRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPMaterialRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
function TSAPMaterialReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPMaterialReader2.GetItems(i: Integer): PSAPMaterialRecord;
begin
  Result := PSAPMaterialRecord(FList.Objects[i]);
end;

function TSAPMaterialReader2.GetSAPMaterialRecord(const snumber: string): PSAPMaterialRecord;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PSAPMaterialRecord(FList.Objects[idx]);
  end
  else Result := nil;
end;

function TSAPMaterialReader2.GetMrpType( const snumber: string ): string;
var
  idx: Integer;
begin
  idx := FList.IndexOf(snumber);
  if idx >= 0 then
  begin
    Result := PSAPMaterialRecord(FList.Objects[idx])^.sMRPType;
  end
  else Result := '';
end;  

procedure TSAPMaterialReader2.Log(const str: string);
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

procedure TSAPMaterialReader2.Open;
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
  aSAPMaterialRecordPtr: PSAPMaterialRecord;
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

      snumber := ADOTabXLS.FieldByName('物料编码').AsString;  // ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;
                
      aSAPMaterialRecordPtr := New(PSAPMaterialRecord);
      FList.AddObject(snumber, TObject(aSAPMaterialRecordPtr));

      aSAPMaterialRecordPtr^.sNumber := snumber;
      aSAPMaterialRecordPtr^.sName := ADOTabXLS.FieldByName('物料描述').AsString;  //   ExcelApp.Cells[irow, iColName].Value;
      aSAPMaterialRecordPtr^.dRecvTime := ADOTabXLS.FieldByName('收货处理时间').AsFloat;  //   ExcelApp.Cells[irow, iColRecvTime].Value;
      aSAPMaterialRecordPtr^.sMRPType := ADOTabXLS.FieldByName('MRP类型').AsString;  //  ExcelApp.Cells[irow, iColMRPType].Value;
      aSAPMaterialRecordPtr^.dSPQ := ADOTabXLS.FieldByName('舍入值').AsFloat;  //  ExcelApp.Cells[irow, iColSPQ].Value;
      aSAPMaterialRecordPtr^.dLT_PD := ADOTabXLS.FieldByName('计划交货时间').AsFloat;  //  ExcelApp.Cells[irow, iColLT_PD].Value;
      aSAPMaterialRecordPtr^.dLT_M0 := ADOTabXLS.FieldByName('自制生产时间').AsFloat;  //  ExcelApp.Cells[irow, iColLT_M0].Value;
      aSAPMaterialRecordPtr^.iLowestCode := 0; // 低位码，默认为1
      aSAPMaterialRecordPtr^.sMRPer := ADOTabXLS.FieldByName('MRP控制者').AsString;
      aSAPMaterialRecordPtr^.sMRPerDesc := ADOTabXLS.FieldByName('MRP控制者描述').AsString;
      aSAPMaterialRecordPtr^.sBuyer := ADOTabXLS.FieldByName('采购组描述').AsString;

      aSAPMaterialRecordPtr^.dMOQ := ADOTabXLS.FieldByName('最小批量大小').AsFloat;
      aSAPMaterialRecordPtr^.sMRPGroup := ADOTabXLS.FieldByName('MRP组').AsString;
      aSAPMaterialRecordPtr^.sPType := ADOTabXLS.FieldByName('采购类型').AsString;
      aSAPMaterialRecordPtr^.sAbc := ADOTabXLS.FieldByName('ABC标识').AsString;
      aSAPMaterialRecordPtr^.sGroupName := ADOTabXLS.FieldByName('物料组描述').AsString;
      aSAPMaterialRecordPtr^.sPlanNumber := ADOTabXLS.FieldByName('计划物料').AsString;
      aSAPMaterialRecordPtr^.sMMTypeDesc := ADOTabXLS.FieldByName('物料类型描述').AsString;
 

      irow := irow + 1;
      snumber := ADOTabXLS.FieldByName('物料编码').AsString;  //  ExcelApp.Cells[irow, iColNumber].Value;

             
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
