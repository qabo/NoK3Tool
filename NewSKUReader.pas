unit NewSKUReader;

interface
        
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TNewSKUReader = class 
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

{ TNewSKUReader }

constructor TNewSKUReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TNewSKUReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TNewSKUReader.Clear;
begin
  FList.Clear;
end;

procedure TNewSKUReader.Log(const str: string);
begin

end;

procedure TNewSKUReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4: string;
  stitle: string;
  irow: Integer;
  snumber: string;
//  sproj: string;
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

        irow := 2;
        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle := stitle1 + stitle2;
        if stitle <> 'ProjectSKU No' then
        begin
          Log(sSheet +'  不是新旧SKU格式');
          Continue;
        end;

        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle + stitle3 + stitle4;
        if stitle = 'ProjectSKU NoDOS目标' then
        begin
          Log(sSheet +'  不是新旧SKU格式');
          Continue;
        end;  

        irow := 3;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin
//          sproj := ExcelApp.Cells[irow, 1].Value;

          FList.Add(snumber);

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
