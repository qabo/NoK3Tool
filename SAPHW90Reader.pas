unit SAPHW90Reader;

interface

uses
  Classes, ComObj, ActiveX, SysUtils, Windows, CommUtils, SAPStockReader,
  SAPMaterialReader;

type
  TSAPHW90Record = packed record
    snumber: string;
    sname: string;
    snumber_luoji: string;
    snumber90: string;
    sproj: string;
  end;
  PSAPHW90Record = ^TSAPHW90Record;

  TSAPHW90Reader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
                      
    FList: TStringList;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string);
    function GetCount: Integer;
    function GetItems(i: Integer): PSAPHW90Record;
  public
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;
    property ReadOk: Boolean read FReadOk;
    property Count: Integer read GetCount;
    property Items[i: Integer]: PSAPHW90Record read GetItems;
  end;

implementation
 
{ TSAPHW90Reader }

constructor TSAPHW90Reader.Create(const sfile: string; aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create; 
  Open;
end;

destructor TSAPHW90Reader.Destroy;
begin
  Clear;
  FList.Free; 
  inherited;
end;

procedure TSAPHW90Reader.Clear;
var
  i: Integer;
  p: PSAPHW90Record;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PSAPHW90Record(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear; 
end;

procedure TSAPHW90Reader.Log(const str: string);
begin

end;
 
procedure TSAPHW90Reader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  irow: Integer;
  p: PSAPHW90Record;
  snumber: string;
begin
  Clear;

  FReadOk := False;

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
        if stitle <> '代码名称裸机计划物料编码机型' then
        begin
          Log(sSheet +'  不是 海外计划物料对照表 格式');
          Continue;
        end;

        FReadOk := True;
 
        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          p := New(PSAPHW90Record);

          p^.snumber := snumber;
          p^.sname := ExcelApp.Cells[irow, 2].Value;
          p^.snumber_luoji := ExcelApp.Cells[irow, 3].Value;
          p^.snumber90 := ExcelApp.Cells[irow, 4].Value;
          p^.sproj := ExcelApp.Cells[irow, 5].Value;

          FList.AddObject(snumber, TObject(p));
          
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

function TSAPHW90Reader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TSAPHW90Reader.GetItems(i: Integer): PSAPHW90Record;
begin
  Result := PSAPHW90Record(FList.Objects[i]);
end;
 
end.
