unit BomD2Reader2;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ProjNameReader, BomD2Reader, ADODB;

type 
  TBomD2Reader2 = class
  private
    FProjNameReader: TProjNameReader;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TBomD2;
  public
    FNumbers: TStringList;
    FList: TStringList;
    constructor Create(const sfile: string; aProjNameReader: TProjNameReader;
      aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    function GetWhereUse(const snumber_child: string): string;
    function BomByNumber(const snumber: string): TBomD2;
    property ReadOk: Boolean read FReadOk;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TBomD2 read GetItems;
  end;

implementation

{ TBomD2Reader2 }

constructor TBomD2Reader2.Create(const sfile: string;
  aProjNameReader: TProjNameReader; aLogEvent: TLogEvent = nil);
begin
  FProjNameReader := aProjNameReader;
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FNumbers := TStringList.Create;
  Open;
end;

destructor TBomD2Reader2.Destroy;
begin
  Clear;
  FList.Free;
  FNumbers.Free;
  inherited;
end;

procedure TBomD2Reader2.Clear;
var
  i: Integer;
  aBomD2: TBomD2;
begin
  FNumbers.Clear;  //要放在前面，因为引用了BOM的Item指针
  
  for i := 0 to FList.Count - 1 do
  begin
    aBomD2 := TBomD2(FList.Objects[i]);
    aBomD2.Free;
  end;
  FList.Clear;
end;

function TBomD2Reader2.GetWhereUse(const snumber_child: string): string;
var
  i: Integer;
  aBomD2: TBomD2;
  sprojs: string;
begin 
  for i := 0 to FList.Count - 1 do
  begin
    aBomD2 := TBomD2(FList.Objects[i]);
    if aBomD2.ChildExists(snumber_child) then
    begin
      if Pos(aBomD2.fproj, sprojs) <= 0 then
      begin
        if sprojs = '' then
        begin
          sprojs := aBomD2.fproj;
        end
        else
        begin
          sprojs := sprojs + ',' + aBomD2.fproj;
        end;
      end;
    end;
  end;
  Result := sprojs;
end;

function TBomD2Reader2.BomByNumber(const snumber: string): TBomD2;
var
  i: Integer;
  aBomD2: TBomD2; 
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aBomD2 := TBomD2(FList.Objects[i]);
    if aBomD2.fnumber = snumber then
    begin
      Result := aBomD2;
      Break;
    end;
  end; 
end;

procedure TBomD2Reader2.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TBomD2Reader2.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TBomD2Reader2.GetItems(i: Integer): TBomD2;
begin
  Result := TBomD2(FList[i]);
end;

procedure TBomD2Reader2.Open;
var
  snumber: string;
  snumber_child: string;
  p: PBomD2Item;
  aBomD2: TBomD2;
  sproj: string;
   
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

    aBomD2 := nil;

    ADOTabXLS.First;
    while not ADOTabXLS.Eof do
    begin

      snumber_child := ADOTabXLS.FieldByName('子件物料编码').AsString;  //   ExcelApp.Cells[irow, iColNumber].Value;
      if snumber_child = '' then Break;


      snumber := ADOTabXLS.FieldByName('母件物料编码').AsString;
      snumber := Trim(snumber);

      if snumber <> '' then
      begin
        sproj := FProjNameReader.ProjOfNumber(snumber);
        aBomD2 := TBomD2.Create(snumber, sproj);
        FList.AddObject(snumber, aBomD2); 
      end;
          
      p := New(PBomD2Item);

      p^.snumber := snumber;
      p^.sname := ADOTabXLS.FieldByName('母件物料描述').AsString; 
      p^.snumber_child := snumber_child;
      p^.sname_child := ADOTabXLS.FieldByName('子件物料描述').AsString; 
      p^.snumber_p := ADOTabXLS.FieldByName('层级').AsString;     
      p^.dusage := ADOTabXLS.FieldByName('子件用量').AsFloat;   
      p^.sabc := ADOTabXLS.FieldByName('ABC标识').AsString;
      p^.sptype := ADOTabXLS.FieldByName('采购类型').AsString;
      p^.slt := ADOTabXLS.FieldByName('L/T').AsFloat;
      p^.dper := ADOTabXLS.FieldByName('使用可能性').AsFloat;
      p^.sgroup := ADOTabXLS.FieldByName('替代项目组').AsString;
      p^.sparent := ADOTabXLS.FieldByName('上层物料编码').AsString;

      aBomD2.FList.Add(p);

      if FNumbers.IndexOf(snumber_child) < 0 then
      begin
        FNumbers.AddObject(snumber_child, TObject(p));
      end;
                         
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
