unit ProjNameReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, Variants;

type 
  TProjNameRecord = packed record
    snumber: string;
    sname: string;
    sproj: string;
  end;
  PProjNameRecord = ^TProjNameRecord;
 
  TProjNameReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FReadOk: Boolean;
    procedure Open;
    procedure Log(const str: string); 
  public
    FList: TStringList;
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    function ProjOfNumber(const snumber: string): string;
    procedure Clear; 
    property ReadOk: Boolean read FReadOk;
  end;

implementation

{ TProjNameReader }

constructor TProjNameReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FList := TStringList.Create;
  Open;
end;

destructor TProjNameReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TProjNameReader.ProjOfNumber(const snumber: string): string;
var
  i: Integer;
  p: PProjNameRecord;
begin
  Result := '';
  for i := 0 to FList.Count - 1 do
  begin
    p := PProjNameRecord(FList.Objects[i]);
    if Copy(p^.snumber, 1, 5) = Copy(snumber, 1, 5) then
    begin
      Result := p^.sproj;
      Break;
    end;
  end;
end;

procedure TProjNameReader.Clear;
var
  i: Integer;
  p: PProjNameRecord;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PProjNameRecord(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;
end;
 
procedure TProjNameReader.Log(const str: string);
begin
  savelogtoexe(str);
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;
 
procedure TProjNameReader.Open; 
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  snumber: string;   
  p: PProjNameRecord;
  v: Variant;
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
        stitle := stitle1 + stitle2 + stitle3;
        if stitle <> '物料长代码物料名称机种' then
        begin

          Log(sSheet +'  不是  机种名称  格式  物料长代码物料名称机种');
          Continue;
        end;
    
        FReadOk := True;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin                                
          p := New(PProjNameRecord);
          FList.AddObject(snumber, TObject(p));

          p^.snumber := snumber;
          v := ExcelApp.Cells[irow, 2].Value;
          if VarIsType(v, varError) then
          begin
            p^.sname := '';
          end
          else
          begin
            p^.sname := v;
          end;
          p^.sproj := ExcelApp.Cells[irow, 3].Value;
  
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
