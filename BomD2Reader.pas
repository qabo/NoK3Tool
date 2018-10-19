unit BomD2Reader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, ProjNameReader;

type 
  TBomD2Item = packed record
    snumber: string;
    sname: string;
    snumber_child: string;
    sname_child: string;
    snumber_p: string;
    dusage: Double;
    sabc: string;
    sptype: string;
    slt: Double;
    dper: Double;
    sgroup: string; //替代组
    sparent: string;
  end;
  PBomD2Item = ^TBomD2Item;

  TBomD2 = class
  private
  public 
    FList: TList;
    fnumber: string;
    fproj: string;
    constructor Create(const snumber,sproj: string);
    destructor Destroy; override;
    procedure Clear;
    function ChildExists(const snumber_child: string): Boolean;
    function ChildByNumber(const snumber_child: string): PBomD2Item;
    function CheckAlloc(ptrBomD2Item_child: PBomD2Item): Boolean;
  end;
 
  TBomD2Reader = class
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
     
{ TBomD2 }

constructor TBomD2.Create(const snumber, sproj: string);
begin
  fnumber := snumber;
  fproj := sproj;
  FList := TList.Create;
end;

destructor TBomD2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TBomD2.Clear;
var
  i: Integer;
  p: PBomD2Item;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomD2Item(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TBomD2.ChildExists(const snumber_child: string): Boolean;
var
  i: Integer;
  p: PBomD2Item;
begin
  Result := False;
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomD2Item(FList[i]);
    if p^.snumber_child = snumber_child then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function TBomD2.ChildByNumber(const snumber_child: string): PBomD2Item;
var
  i: Integer;
  p: PBomD2Item;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomD2Item(FList[i]);
    if p^.snumber_child = snumber_child then
    begin
      Result := p;
      Break;
    end;
  end;
end;

function RightPos(const SubStr, Str: string): Integer;
var
   i, j, k, LenSub, LenS: Integer;
begin
   Result:= 0;
   LenSub:= Length(SubStr);
   LenS:= Length(Str);
   if (LenSub = 0) or (LenS = 0) or (LenSub > LenS) then Exit;
   for i:= LenS downto 1 do
   begin
      if Str[i] = SubStr[LenSub] then
      begin
         k:= i-1;
         for j:= LenSub - 1 downto 1 do
         begin
            if Str[k] = SubStr[j] then
               Dec(k)
            else
               Break;
         end;
      end;
      if i-k = LenSub then
      begin
         Result:= k+1;
         Exit;
      end;
   end;    
end;

function TBomD2.CheckAlloc(ptrBomD2Item_child: PBomD2Item): Boolean;
var
  i: Integer;
  p: PBomD2Item;
  dPer100: Double;
  s1, s2: string;
begin
  Result := True;
     
  if Trim( ptrBomD2Item_child^.sgroup ) = '' then Exit;

  s1 := Copy(ptrBomD2Item_child^.snumber_p, 1, RightPos('.', ptrBomD2Item_child^.snumber_p));

  dPer100 := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomD2Item(FList[i]);
    s2 := Copy(p^.snumber_p, 1, RightPos('.', p^.snumber_p));
    if (s1 = s2) and (p^.sgroup = ptrBomD2Item_child^.sgroup) then
    begin
      dPer100 := dPer100 + p^.dper;
    end;
  end;
  Result := DoubleE(dPer100, 100);
end;  

{ TBomD2Reader }

constructor TBomD2Reader.Create(const sfile: string;
  aProjNameReader: TProjNameReader; aLogEvent: TLogEvent = nil);
begin
  FProjNameReader := aProjNameReader;
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FNumbers := TStringList.Create;
  Open;
end;

destructor TBomD2Reader.Destroy;
begin
  Clear;
  FList.Free;
  FNumbers.Free;
  inherited;
end;

procedure TBomD2Reader.Clear;
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

function TBomD2Reader.GetWhereUse(const snumber_child: string): string;
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

function TBomD2Reader.BomByNumber(const snumber: string): TBomD2;
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

procedure TBomD2Reader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TBomD2Reader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TBomD2Reader.GetItems(i: Integer): TBomD2;
begin
  Result := TBomD2(FList[i]);
end;

procedure TBomD2Reader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  snumber_child: string;   
  p: PBomD2Item;
  aBomD2: TBomD2;
  sproj: string;

begin
  Clear;
          
  FReadOk := False;

  if not FileExists(FFile) then Exit;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(FFile);   

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;
        Log(sSheet);

        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;  
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        if stitle <> '母件物料编码母件物料描述工厂用途代工厂' then
        begin        
          Log(sSheet +'  不是  BOM  格式  母件物料编码母件物料描述工厂用途代工厂');
          Continue;
        end;
    
        FReadOk := True;

        aBomD2 := nil;
        irow := 2;
        snumber_child := ExcelApp.Cells[irow, 13].Value;
        while snumber_child <> '' do
        begin
          snumber := ExcelApp.Cells[irow, 1].Value;

          if snumber <> '' then
          begin
            sproj := FProjNameReader.ProjOfNumber(snumber);
            aBomD2 := TBomD2.Create(snumber, sproj);
            FList.AddObject(snumber, aBomD2); 
          end;
          
          p := New(PBomD2Item);

          p^.snumber := snumber;
          p^.sname := ExcelApp.Cells[irow, 2].Value;
          p^.snumber_child := snumber_child;
          p^.sname_child := ExcelApp.Cells[irow, 14].Value;
          p^.snumber_p := ExcelApp.Cells[irow, 9].Value;
          p^.dusage := ExcelApp.Cells[irow, 19].Value;
          p^.sabc := ExcelApp.Cells[irow, 16].Value;
          p^.sptype := ExcelApp.Cells[irow, 15].Value;
          p^.slt := ExcelApp.Cells[irow, 18].Value;
          p^.dper := ExcelApp.Cells[irow, 23].Value;
          p^.sgroup := ExcelApp.Cells[irow, 21].Value;
          p^.sparent := ExcelApp.Cells[irow, 9].Value;

          aBomD2.FList.Add(p);

          if FNumbers.IndexOf(snumber_child) < 0 then
          begin
            FNumbers.AddObject(snumber_child, TObject(p));
          end;

          irow := irow + 1;
          snumber_child := ExcelApp.Cells[irow, 13].Value;
        end;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;  
end;
 
end.
