unit SAPOPOReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TSAPOPOAlloc = packed record
    dQty: Double;
    dt: TDateTime;
  end;
  PSAPOPOAlloc= ^TSAPOPOAlloc;
  
  TSAPOPOLine = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPOPOAlloc;    
  public
    FNumber: string;
    FName: string;
    FDate: TDateTime;
    FStock: string; // ���ص�
    FQty: Double;
    FQtyAlloc: Double;
    FBillNo: string;
    FLine: string;
    FPlanLine: string;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;              
    function Alloc(dt: TDateTime; var dQty: Double): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPOPOAlloc read GetItems;
  end;
  
  TSAPOPOReader = class
  private         
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPOPOLine;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                                                           
    procedure GetOPOs(slNumber: TStringList; lst: TList);
    function Alloc(slNumber: TStringList; dt: TDateTime; dQty: Double): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPOPOLine read GetItems;
  end;

    function ListSortCompare_DateTime_PO(Item1, Item2: Pointer): Integer;

implementation
      
{ TSAPOPOLine }

constructor TSAPOPOLine.Create;
begin
  FList := TList.Create;
  FQtyAlloc := 0;
end;

destructor TSAPOPOLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPOPOLine.Clear;
var
  p: PSAPOPOAlloc;
  i: Integer;
begin
  for i := 0 to FList.Count -1 do
  begin
    p := PSAPOPOAlloc(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPOPOLine.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPOPOLine.GetItems(i: Integer): PSAPOPOAlloc;
begin
  Result := PSAPOPOAlloc(FList[i]);
end;

function TSAPOPOLine.Alloc(dt: TDateTime; var dQty: Double): Double;
var                      
  p: PSAPOPOAlloc;
begin
  Result := 0;
  if DoubleG( FQty, FQtyAlloc ) then //�пɷ���
  begin
    p := New(PSAPOPOAlloc);
    FList.Add(p);
    if DoubleGE( FQty - FQtyAlloc, dQty ) then //����������
    begin                             
      Result := dQty;
      FQtyAlloc := FQtyAlloc + dQty;

      dQty := 0;
    end
    else  // ������������
    begin                         
      Result := FQty - FQtyAlloc;
      FQtyAlloc := FQty;
      
      dQty := dQty - Result;
    end;                            

    p^.dQty := Result;
    p^.dt := dt;
  end;
end;  

{ TSAPOPOReader }

constructor TSAPOPOReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TSAPOPOReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPOPOReader.Clear;
var
  i: Integer;
  p: TSAPOPOLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := TSAPOPOLine(FList.Objects[i]);
    p.Free;
  end;
  FList.Clear;
end;

procedure TSAPOPOReader.GetOPOs(slNumber: TStringList; lst: TList);
var
  i: Integer;
begin
  for i := 0 to FList.Count - 1 do
  begin
    if slNumber.IndexOf(FList[i]) >= 0 then
    begin
      lst.Add( FList.Objects[i] );
    end;
  end;
end;

function TSAPOPOReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPOPOReader.GetItems(i: Integer): TSAPOPOLine;
begin
  Result := TSAPOPOLine(FList.Objects[i]);
end;

              
    function ListSortCompare_DateTime_PO(Item1, Item2: Pointer): Integer;
    var
      p1, p2: TSAPOPOLine;
    begin
      p1 := TSAPOPOLine(Item1);
      p2 := TSAPOPOLine(Item2);
      
      if DoubleG(p1.FDate, p2.FDate) then
        Result := 1
      else if DoubleL(p1.FDate, p2.FDate) then
        Result := -1
      else Result := 0;
    end;

function TSAPOPOReader.Alloc(slNumber: TStringList; dt: TDateTime;
  dQty: Double): Double;
var
  lst: TList;
  i: Integer;    
  p: TSAPOPOLine;
begin
  Result := 0;
  
  lst := TList.Create;
  GetOPOs(slNumber, lst); // �ҵ���������ϵĿ��òɹ�����

  lst.Sort(ListSortCompare_DateTime_PO);  // ����������
  for i := 0 to lst.Count - 1 do
  begin
    p := TSAPOPOLine(lst[i]);
    Result := Result + p.Alloc(dt, dQty);
    if DoubleE( dQty, 0) then // �������������
    begin
      Break;
    end;
  end;

  lst.Free;
end;

procedure TSAPOPOReader.Log(const str: string);
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

procedure TSAPOPOReader.Open;
const
  CSNumber = '����';
  CSName = '���ı�';   
  CSColDate = 'ƾ֤����';
  CSStock = '���ص�';
  CSQty = '�ƻ�����';
  CSBillNo = '�ɹ�ƾ֤';
  CSLine = '��Ŀ';
  CSPlanLine = '�ƻ���';
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  aSAPOPOLine: TSAPOPOLine;
  iColNumber: Integer;
  iColName: Integer;
  iColDate: Integer;
  iColStock: Integer;
  iColQty: Integer;
  iColBillNo: Integer;
  iColLine: Integer;
  iColPlanLine: Integer;
begin
  Clear;

  if not FileExists(FFile) then Exit;

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
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
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
        if stitle <> '�ɹ�ƾ֤��Ŀ�ƻ��вɹ�ƾ֤���Ͳɹ�ƾ֤���ɹ���' then
        begin

          Log(sSheet +'  ����  SAP����OPO  ��ʽ');
          Continue;
        end;

        iColNumber := IndexOfCol(ExcelApp, irow, CSNumber);
        iColName := IndexOfCol(ExcelApp, irow, CSName);
        iColDate := IndexOfCol(ExcelApp, irow, CSColDate);
        iColStock := IndexOfCol(ExcelApp, irow, CSStock);
        iColQty := IndexOfCol(ExcelApp, irow, CSQty);
        iColBillNo := IndexOfCol(ExcelApp, irow, CSBillNo);
        iColLine := IndexOfCol(ExcelApp, irow, CSLine);
        iColPlanLine := IndexOfCol(ExcelApp, irow, CSPlanLine);

        if (iColNumber = -1) or (iColName = -1) or (iColDate = -1)
          or (iColStock = -1) or (iColQty = -1) or (iColBillno = -1)
          or (iColLine = -1) or (iColPlanLine = -1)
          then
        begin
          Log(sSheet +'  ����  SAP����OPO  ��ʽ');
          Continue;
        end;
                
        irow := 2;
        snumber := ExcelApp.Cells[irow, iColNumber].Value;
        while snumber <> '' do
        begin                                
          aSAPOPOLine := TSAPOPOLine.Create;
          FList.AddObject(snumber, aSAPOPOLine);

          aSAPOPOLine.FNumber := snumber;
          aSAPOPOLine.FName := ExcelApp.Cells[irow, iColName].Value;
          aSAPOPOLine.FDate := ExcelApp.Cells[irow, iColDate].Value;
          aSAPOPOLine.FStock := ExcelApp.Cells[irow, iColStock].Value;   
          aSAPOPOLine.FQty := ExcelApp.Cells[irow, iColQty].Value;
          aSAPOPOLine.FBillNo := ExcelApp.Cells[irow, iColBillNo].Value;
          aSAPOPOLine.FLine := ExcelApp.Cells[irow, iColLine].Value;
          aSAPOPOLine.FPlanLine := ExcelApp.Cells[irow, iColPlanLine].Value;

          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, iColNumber].Value;
        end;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end;  
end;
 
end.
