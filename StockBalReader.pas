unit StockBalReader;

interface
          
uses
  Classes, ComObj, CommUtils;

type
  TStockBalRecord = packed record
    snumber: string;
    qty: Double;
  end;
  PStockBalRecord = ^TStockBalRecord;

  TStockBalReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FList: TStringList;
    procedure Open;
    procedure Log(const str: string);
    procedure Clear;
    function GetCount: Integer;
    function GetItems(i: Integer): PStockBalRecord;
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    function GetNumberBal(const sNumber: string): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PStockBalRecord read GetItems;
  end;

implementation

{ TStockBalReader }
               
constructor TStockBalReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TStockBalReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TStockBalReader.Clear;
var
  i: Integer;
  p: PStockBalRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PStockBalRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TStockBalReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TStockBalReader.GetItems(i: Integer): PStockBalRecord;
begin
  Result := PStockBalRecord(FList.Objects[i]);
end;

procedure TStockBalReader.Log(const str: string);
begin

end;

function TStockBalReader.GetNumberBal(const sNumber: string): Double;
var
  i: Integer;
  p: PStockBalRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PStockBalRecord(FList.Objects[i]);
    if p^.snumber = sNumber then
    begin
      Result := p^.qty;
      Break;
    end;
  end;
end;

procedure TStockBalReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PStockBalRecord;
begin
  Clear;


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
        if stitle <> '物料编码期末库存' then
        begin
          Log(sSheet +'  不是期末库存格式');
          Continue;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin                                
          p := New(PStockBalRecord);
          FList.AddObject(snumber, TObject(p));

          p^.snumber := snumber;
          p^.qty := ExcelApp.Cells[irow, 2].Value; 

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
