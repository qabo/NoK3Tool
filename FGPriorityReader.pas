unit FGPriorityReader;

interface
           
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TFGPriorityReader = class
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
  end;

implementation

{ TFGPriorityReader }

constructor TFGPriorityReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TFGPriorityReader.Destroy;
begin
  Clear;
  FList.Free;
end;

procedure TFGPriorityReader.Clear;
begin
  FList.Clear;
end;

procedure TFGPriorityReader.Log(const str: string);
begin

end;

procedure TFGPriorityReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2: string;
  stitle: string;
  irow: Integer;
  snumber: string;
  iPriority: Integer; 
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
        if stitle <> '产品编码优先级' then
        begin
          Log(sSheet +'  不是新旧SKU格式');
          Continue;
        end;
 
        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin                                
          iPriority := ExcelApp.Cells[irow, 2].Value;

          FList.AddObject(snumber, TObject(iPriority));

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
