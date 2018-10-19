unit MRPWinReader;

interface
              
uses
  SysUtils, Classes, ComObj, CommUtils;

type
  TWinRecord = packed record
    snumber: string;
    sname: string;
    dqty: Double;
  end;
  PWinRecord = ^TWinRecord;

  TRPWinReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetQty(const snumber: string): Double;
  end;

implementation
      
{ TRPWinReader }

constructor TRPWinReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TRPWinReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TRPWinReader.Clear;
var
  i: Integer;
  p: PWinRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PWinRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;

function TRPWinReader.GetQty(const snumber: string): Double;
var
  i: Integer;
  p: PWinRecord;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    p := PWinRecord(FList.Objects[i]);
    if p^.snumber = snumber then
    begin
      Result := p^.dqty;
      Break;
    end;
  end;
end;

procedure TRPWinReader.Log(const str: string);
begin

end;

procedure TRPWinReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PWinRecord;
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
        if stitle <> '物料编码物料名称数量' then
        begin
          Log(sSheet +'  外购入库数据');
          Continue;
        end;


        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin                                
          p := New(PWinRecord);
          FList.AddObject(snumber, TObject(p));

          p^.snumber := snumber;
          p^.sname := ExcelApp.Cells[irow, 2].Value;
          p^.dqty := ExcelApp.Cells[irow, 3].Value; 

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
