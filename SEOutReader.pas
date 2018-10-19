unit SEOutReader;

interface
        
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TSEOutRecord = packed record
    dt: TDateTime;
    snumber: string;
    qty: Double;
  end;
  PSEOutRecord = ^TSEOutRecord;

  TSEOutReader = class
  private
    FFile: string;    
    FList: TStringList;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    procedure Clear;                                       
    function GetCount: Integer;
    function GetItems(i: Integer): PSEOutRecord;
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    function GetQty(const snumber: string; dt1, dt2: TDateTime): Double;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSEOutRecord read GetItems;
  end;

implementation
      
{ TSEOutReader }
               
constructor TSEOutReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TSEOutReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TSEOutReader.GetQty(const snumber: string; dt1, dt2: TDateTime): Double;
var
  i: Integer;
  p: PSEOutRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PSEOutRecord(FList.Objects[i]);
    if p^.snumber <> snumber then Continue;
    if p^.dt < dt1 then Continue;
    if p^.dt > dt2 then Continue;
    Result := Result + p^.qty;
  end;
end;

procedure TSEOutReader.Clear;
var
  i: Integer;
  p: PSEOutRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSEOutRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
         
function TSEOutReader.GetCount: Integer;
begin
  Result := FList.Count ;
end;

function TSEOutReader.GetItems(i: Integer): PSEOutRecord;
begin
  Result := PSEOutRecord(FList.Objects[i]);
end;

procedure TSEOutReader.Log(const str: string);
begin

end;

procedure TSEOutReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PSEOutRecord;
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
        if stitle <> '日期物料编码数量' then
        begin
          Log(sSheet +'  不是期末库存格式');
          Continue;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin                                
          p := New(PSEOutRecord);
          FList.AddObject(snumber, TObject(p));

          p^.dt := ExcelApp.Cells[irow, 1].Value;
          p^.snumber := snumber;
          p^.qty := ExcelApp.Cells[irow, 3].Value;

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
