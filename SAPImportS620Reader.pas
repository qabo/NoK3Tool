unit SAPImportS620Reader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TSAPImportS620Col = packed record
    sdt: string;
    dQty: Double;
  end;
  PSAPImportS620Col= ^TSAPImportS620Col;
  
  TSAPImportS620Line = class
  private
    FList: TList;
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPImportS620Col;
  public
    sMATNR: string;
    sBERID: string;

    snumber: string; //产品编码
    sver: string; //版本
    scolor: string; //	颜色
    scap: string; //	容量
    sproj: string; //	项目


    constructor Create;
    destructor Destroy; override;
    function AddDateQty(const aSAPImportS620Col: TSAPImportS620Col): Integer;
    procedure Clear;               
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPImportS620Col read GetItems;
  end;
  
  TSAPImportS620Reader = class
  private     
    FList: TStringList;
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): TSAPImportS620Line;
  public                 
    FSum: Double;    
    FDates: TStringList;   
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;                                                            
    property Count: Integer read GetCount;
    property Items[i: Integer]: TSAPImportS620Line read GetItems;
  end;
 
implementation
      
{ TSAPImportS620Line }

constructor TSAPImportS620Line.Create;
begin
  FList := TList.Create; 
end;

destructor TSAPImportS620Line.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSAPImportS620Line.Clear;
var
  p: PSAPImportS620Col;
  i: Integer;
begin
  for i := 0 to FList.Count -1 do
  begin
    p := PSAPImportS620Col(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TSAPImportS620Line.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPImportS620Line.GetItems(i: Integer): PSAPImportS620Col;
begin
  Result := PSAPImportS620Col(FList[i]);
end;

function TSAPImportS620Line.AddDateQty(const aSAPImportS620Col: TSAPImportS620Col): Integer;
var
  p: PSAPImportS620Col;
begin
  p := New(PSAPImportS620Col);
  p^ := aSAPImportS620Col;
  Result := FList.Add(p);
end;

{ TSAPOPOReader }

constructor TSAPImportS620Reader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  FDates := TStringList.Create;
  Open;
end;

destructor TSAPImportS620Reader.Destroy;
begin
  Clear;
  FList.Free;
  FDates.Free;
  inherited;
end;

procedure TSAPImportS620Reader.Clear;
var
  i: Integer;
  p: TSAPImportS620Line;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := TSAPImportS620Line(FList.Objects[i]);
    p.Free;
  end;
  FList.Clear;

  FDates.Clear;

  FSum := 0;
end;
 
function TSAPImportS620Reader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPImportS620Reader.GetItems(i: Integer): TSAPImportS620Line;
begin
  Result := TSAPImportS620Line(FList.Objects[i]);
end;
 
procedure TSAPImportS620Reader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;
 
procedure TSAPImportS620Reader.Open; 
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  snumber: string;   
  aSAPOPOLine: TSAPImportS620Line; 
  iColPlanLine: Integer;
  aSAPImportS620Col: TSAPImportS620Col;
  idate: Integer;
  icolExtra: Integer;
begin
  Clear;

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
        stitle := stitle1 + stitle2;
        if stitle <> 'MATNRBERID' then
        begin

          Log(sSheet +'  不是  SAP导入S620  格式');
          Continue;
        end;

        irow := 1;
        icol := 3;
        aSAPImportS620Col.sdt := ExcelApp.Cells[irow, icol].Value;
        while aSAPImportS620Col.sdt <> '' do
        begin
          FDates.AddObject(aSAPImportS620Col.sdt, TObject(icol));
          
          icol := icol + 1;
          aSAPImportS620Col.sdt := ExcelApp.Cells[irow, icol].Value;
        end;

        icolExtra := icol + 3;
     
        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin                                
          aSAPOPOLine := TSAPImportS620Line.Create;
          aSAPOPOLine.sMATNR := snumber;
          aSAPOPOLine.sBERID := ExcelApp.Cells[irow, 2].Value;

          aSAPOPOLine.snumber := ExcelApp.Cells[irow, icolExtra].Value;
          aSAPOPOLine.sver := ExcelApp.Cells[irow, icolExtra + 1].Value;
          aSAPOPOLine.scolor := ExcelApp.Cells[irow, icolExtra + 2].Value;
          aSAPOPOLine.scap := ExcelApp.Cells[irow, icolExtra + 3].Value;
          aSAPOPOLine.sproj := ExcelApp.Cells[irow, icolExtra + 4].Value;

          FList.AddObject(snumber, aSAPOPOLine);

          icol := 3;
          for idate := 0 to FDates.Count - 1 do
          begin
            icol := Integer(FDates.Objects[idate]);

            aSAPImportS620Col.sdt := FDates[idate];
            aSAPImportS620Col.dQty := ExcelApp.Cells[irow, icol].Value;

            aSAPOPOLine.AddDateQty(aSAPImportS620Col);

            FSum := FSum + aSAPImportS620Col.dQty;
          end; 
          
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 1].Value;
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
