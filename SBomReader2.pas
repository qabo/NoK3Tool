unit SBomReader2;

interface

uses
  Classes, ComObj, CommUtils;

type

  TSBomChild2 = class
  public
    FNumber99: string;
    FName: string;
    FLT: Integer;
    Obj: TObject;
  end;
  
  TSBom2 = class
  private
    procedure Clear;
  public
    FNumber: string;
    FName: string;
    FList: TStringList;
    constructor Create;
    destructor Destroy; override;
  end;

  TSBomReader2 = class
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
    function GetBom(const sNumber: string): TSBom2;
  end;

implementation

{ TSBom2 }

constructor TSBom2.Create;
begin
  FList := TStringList.Create;
end;

destructor TSBom2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSBom2.Clear;
var
  i: Integer;
  abc: TSBomChild2;
begin
  for i := 0 to FList.Count - 1 do
  begin
    abc := TSBomChild2(FList.Objects[i]);
    abc.Free;
  end;
  FList.Clear;
end;

{ TSBomReader2 }

constructor TSBomReader2.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TSBomReader2.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSBomReader2.Clear;
var
  i: Integer;
  aSBom: TSBom2;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSBom := TSBom2(FList.Objects[i]);
    aSBom.Free;
  end;
  FList.Clear;
end;

procedure TSBomReader2.Log(const str: string);
begin

end;

function TSBomReader2.GetBom(const sNumber: string): TSBom2;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    if FList[i] = sNumber then
    begin
      Result := TSBom2(FList.Objects[i]);
      Break;
    end;
  end;
end;

procedure TSBomReader2.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  irow: Integer;
  irow1: Integer;
  snumber: string;
  aSBom: TSBom2;
  aSBomChild: TSBomChild2;
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
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
        if stitle <> '产品编码产品名称物料编码物料名称提前期' then
        begin
          Log(sSheet +'  不是简易BOM格式');
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 1].Value;
        while snumber <> '' do
        begin
          aSBom := TSBom2.Create;
          FList.AddObject(snumber, aSBom);

          aSBom.FNumber := snumber;
          aSBom.FName := ExcelApp.Cells[irow, 2].Value;
          
          irow1 := irow;
          while IsCellMerged(ExcelApp, irow1, 1, irow, 1) do
          begin                             
            aSBomChild := TSBomChild2.Create;
            aSBomChild.FNumber99 := ExcelApp.Cells[irow, 3].Value;
            aSBomChild.FName := ExcelApp.Cells[irow, 4].Value;
            aSBomChild.FLT := ExcelApp.Cells[irow, 5].Value;
            aSBom.FList.AddObject(aSBomChild.FNumber99, aSBomChild);
            irow := irow + 1;
          end;   
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
