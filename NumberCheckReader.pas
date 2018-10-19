unit NumberCheckReader;

interface

uses
  Classes, ComObj, SysUtils, CommUtils;

type
  TBomLocLine = packed record
    snumber: string;
    sname: string;
    svendor: string;
    ssourcing: string;
    sbuyer: string;
    slt: string;
    scheck: string;
    schecknote: string;
    smcnote: string;
    sriskqty: string;
    smc: string;
    sver: string;
    scap: string;
    scol: string;
  end;
  PBomLocLine = ^TBomLocLine;

  TBomLoc = class
  public
    scategory: string;
    susage: string;
    sloc: string;

    FList: TStringList;    
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
  end;
  
  TNumberCheckReader = class
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

{ TBomLoc }
    
constructor TBomLoc.Create;
begin
  FList := TStringList.Create;
end;

destructor TBomLoc.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TBomLoc.Clear;
var
  i: Integer;
  p: PBomLocLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PBomLocLine(FList.Objects[i]);
    Dispose(p);
  end;
  FList.Clear;  
end;  

{ TNumberCheckReader }

constructor TNumberCheckReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create; 
  Open;
end;

destructor TNumberCheckReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TNumberCheckReader.Clear;
var
  i: Integer;
  aBomLoc: TBomLoc;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aBomLoc := TBomLoc(FList.Objects[i]);
    aBomLoc.Free;
  end;  
  FList.Clear; 
end;

procedure TNumberCheckReader.Log(const str: string);
begin

end;

procedure TNumberCheckReader.Open;
var
  iSheet: Integer;
  iSheetCount: Integer;
  sSheet: string;
  stitle: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  irow: Integer;
  irow_title: Integer; 
  icol: Integer;
  aBomLoc: TBomLoc;
  snumber: string;  
  aBomLocLinePtr: PBomLocLine; 
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

        irow_title := -1;
        for irow := 1 to 10 do
        begin
          stitle1 := ExcelApp.Cells[irow, 1].Value;
          stitle2 := ExcelApp.Cells[irow, 2].Value;
          stitle3 := ExcelApp.Cells[irow, 3].Value;
          stitle4 := ExcelApp.Cells[irow, 4].Value;
          stitle5 := ExcelApp.Cells[irow, 5].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
          if (stitle = '物料类型物料编码物料名称数量厂家')
            or (stitle = '物料类型物料编码物料名称用量厂家') then
          begin
            irow_title := irow;
            Break;
          end;
        end;

        if irow_title = -1 then
        begin
          Log(sSheet + ' 不是 BOM 签核格式');
          Continue;
        end;

        aBomLoc := nil;

        irow := irow_title + 1;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin
          if not IsCellMerged(ExcelApp, irow, 4, irow - 1, 4) then // 第一行物料，跟标题不可能是合并单元格
          begin
            aBomLoc := TBomLoc.Create;
            aBomLoc.scategory := ExcelApp.Cells[irow, 1].Value;
            aBomLoc.susage := ExcelApp.Cells[irow, 4].Value;
            FList.AddObject(IntToStr(irow), aBomLoc); 
          end;

          aBomLocLinePtr := New(PBomLocLine);
          aBomLocLinePtr^.snumber := snumber;
          aBomLocLinePtr^.sname := ExcelApp.Cells[irow, 3].Value;
          aBomLocLinePtr^.svendor := ExcelApp.Cells[irow, 5].Value;
          aBomLocLinePtr^.ssourcing := ExcelApp.Cells[irow, 6].Value;
          aBomLocLinePtr^.sbuyer := ExcelApp.Cells[irow, 7].Value;
          aBomLocLinePtr^.slt := ExcelApp.Cells[irow, 8].Value;
          aBomLocLinePtr^.scheck := ExcelApp.Cells[irow, 9].Value;
          aBomLocLinePtr^.schecknote := ExcelApp.Cells[irow, 10].Value;
          aBomLocLinePtr^.smcnote := ExcelApp.Cells[irow, 11].Value;
          aBomLocLinePtr^.sriskqty := ExcelApp.Cells[irow, 12].Value;
          aBomLocLinePtr^.smc := ExcelApp.Cells[irow, 13].Value;
          aBomLoc.FList.AddObject(aBomLocLinePtr^.snumber, TObject(aBomLocLinePtr));
 
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
