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
    sMRPType: string; // PD mrp, �⹺  M0 mps ���ư��Ʒ 
    dSPQ: Double; // = '����ֵ';
    dLT_PD: Double; // = '�ƻ�����ʱ��';
    dLT_M0: Double; // = '��������ʱ��';
    iLowestCode: Integer; // ��λ��
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

      snumber := ADOTabXLS.FieldByName('���ϱ���').AsString;  // ExcelApp.Cells[irow, iColNumber].Value;
      if snumber = '' then Break;
                
      aSAPMaterialRecordPtr := New(PSAPMaterialRecord);
      FList.AddObject(snumber, TObject(aSAPMaterialRecordPtr));

      aSAPMaterialRecordPtr^.sNumber := snumber;
      aSAPMaterialRecordPtr^.sName := ADOTabXLS.FieldByName('��������').AsString;  //   ExcelApp.Cells[irow, iColName].Value;
      aSAPMaterialRecordPtr^.dRecvTime := ADOTabXLS.FieldByName('�ջ�����ʱ��').AsFloat;  //   ExcelApp.Cells[irow, iColRecvTime].Value;
      aSAPMaterialRecordPtr^.sMRPType := ADOTabXLS.FieldByName('MRP����').AsString;  //  ExcelApp.Cells[irow, iColMRPType].Value;
      aSAPMaterialRecordPtr^.dSPQ := ADOTabXLS.FieldByName('����ֵ').AsFloat;  //  ExcelApp.Cells[irow, iColSPQ].Value;
      aSAPMaterialRecordPtr^.dLT_PD := ADOTabXLS.FieldByName('�ƻ�����ʱ��').AsFloat;  //  ExcelApp.Cells[irow, iColLT_PD].Value;
      aSAPMaterialRecordPtr^.dLT_M0 := ADOTabXLS.FieldByName('��������ʱ��').AsFloat;  //  ExcelApp.Cells[irow, iColLT_M0].Value;
      aSAPMaterialRecordPtr^.iLowestCode := 0; // ��λ�룬Ĭ��Ϊ1
      aSAPMaterialRecordPtr^.sMRPer := ADOTabXLS.FieldByName('MRP������').AsString;
      aSAPMaterialRecordPtr^.sMRPerDesc := ADOTabXLS.FieldByName('MRP����������').AsString;
      aSAPMaterialRecordPtr^.sBuyer := ADOTabXLS.FieldByName('�ɹ�������').AsString;

      aSAPMaterialRecordPtr^.dMOQ := ADOTabXLS.FieldByName('��С������С').AsFloat;
      aSAPMaterialRecordPtr^.sMRPGroup := ADOTabXLS.FieldByName('MRP��').AsString;
      aSAPMaterialRecordPtr^.sPType := ADOTabXLS.FieldByName('�ɹ�����').AsString;
      aSAPMaterialRecordPtr^.sAbc := ADOTabXLS.FieldByName('ABC��ʶ').AsString;
      aSAPMaterialRecordPtr^.sGroupName := ADOTabXLS.FieldByName('����������').AsString;
      aSAPMaterialRecordPtr^.sPlanNumber := ADOTabXLS.FieldByName('�ƻ�����').AsString;
      aSAPMaterialRecordPtr^.sMMTypeDesc := ADOTabXLS.FieldByName('������������').AsString;
 

      irow := irow + 1;
      snumber := ADOTabXLS.FieldByName('���ϱ���').AsString;  //  ExcelApp.Cells[irow, iColNumber].Value;

             
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
