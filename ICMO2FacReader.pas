unit ICMO2FacReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ADODB;

type
  TICMO2FacRecord = packed record
    sbillno_mz: string;
    sbillno_fac: string;
  end;
  PICMO2FacRecord = ^TICMO2FacRecord;

  TICMO2FacReader2 = class
  private         
    FList: TList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PICMO2FacRecord;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PICMO2FacRecord read GetItems;
    function ICMOFac2MZ(const sicmo_fac: string): string;
  end;
 

implementation

{ TICMO2FacReader2 }

constructor TICMO2FacReader2.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TList.Create;
  Open;
end;

destructor TICMO2FacReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TICMO2FacReader2.Clear;
var
  i: Integer;
  p: PICMO2FacRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PICMO2FacRecord(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TICMO2FacReader2.ICMOFac2MZ(const sicmo_fac: string): string;
var
  i: Integer;
  p: PICMO2FacRecord;
  sbillno: string;
begin
  Result := '';
  for i := 0 to FList.Count - 1 do
  begin
    p := PICMO2FacRecord(FList[i]);

    sbillno := p^.sbillno_fac;
    if Copy(sbillno, 1, 3) = 'NWT' then
    begin
      sbillno := Copy(sbillno, 4, Length(sbillno) - 3);
    end;
    if Copy(sbillno, 1, 2) = 'WT' then
    begin
      sbillno := Copy(sbillno, 3, Length(sbillno) - 2);
    end;

    if sbillno = sicmo_fac then
    begin
      Result := p^.sbillno_mz;
      Break;
    end;
  end;
end;

function TICMO2FacReader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TICMO2FacReader2.GetItems(i: Integer): PICMO2FacRecord;
begin
  Result := PICMO2FacRecord(FList[i]);
end;

procedure TICMO2FacReader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

procedure TICMO2FacReader2.Open;
var
  iSheetCount, iSheet: Integer; 
  stitle: string;
  irow: Integer;
  snumber: string;
  aICMO2FacRecordPtr: PICMO2FacRecord;
   
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

    try
      Conn.Connected:=true;

      ADOTabXLS.Connection:=Conn;

      ADOTabXLS.TableName:='[Sheet1$]';

      ADOTabXLS.Active:=true;

      ADOTabXLS.First;
      while not ADOTabXLS.Eof do
      begin
        aICMO2FacRecordPtr := New(PICMO2FacRecord);
        FList.Add(aICMO2FacRecordPtr);

        aICMO2FacRecordPtr^.sbillno_mz := ADOTabXLS.FieldByName('订单').AsString;
        aICMO2FacRecordPtr^.sbillno_fac := ADOTabXLS.FieldByName('代工厂工单').AsString;
 
        ADOTabXLS.Next;
      end;


      ADOTabXLS.Close;

      Conn.Connected := False;
    except
      on e: Exception do
      begin
        raise Exception.Create(FFile + #13#10 + e.Message);
      end;
    end;
  finally
    FreeAndNil(Conn);
    FreeAndNil(ADOTabXLS);
  end;
end;
 
end.
