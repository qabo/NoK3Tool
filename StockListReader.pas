unit StockListReader;

interface
           
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TStockInfoRecord = packed record
    sfac: string;
    snumber: string;
    sname: string;
    sMrpArea: string;
    sIsMrp: string;
  end;
  PStockInfoRecord = ^TStockInfoRecord;

  TStockListReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    function GetItems(i: Integer): PStockInfoRecord;
    function GetCount: Integer;
  public
    FList: TList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PStockInfoRecord read GetItems;
  end;

implementation

{ TStockListReader }

constructor TStockListReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;

destructor TStockListReader.Destroy;
begin
  Clear;
  FList.Free;
end;

procedure TStockListReader.Clear;
begin
  FList.Clear;
end;

procedure TStockListReader.Log(const str: string);
var
  i: Integer;
  p: PStockInfoRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PStockInfoRecord(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TStockListReader.GetItems(i: Integer): PStockInfoRecord;
begin
  Result := PStockInfoRecord(FList[i]);
end;

function TStockListReader.GetCount: Integer;
begin
  Result := FList.Count;
end;  

procedure TStockListReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  snumber: string; 
  p: PStockInfoRecord;
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
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle := stitle1 + stitle2 + stitle3;
        if stitle <> '工厂库位仓储地点描述' then
        begin
          Log(sSheet +'  不是 仓库列表 格式');
          Continue;
        end;
 
        irow := 2;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin                                
          p := New(PStockInfoRecord);
          FList.Add(p);

          p^.snumber := snumber;
          p^.sfac := ExcelApp.Cells[irow, 1].Value;   
          p^.sname := ExcelApp.Cells[irow, 3].Value;

          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, 2].Value;
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
