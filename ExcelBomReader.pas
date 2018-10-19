unit ExcelBomReader;

interface

uses
  Windows, Classes, ComObj, SysUtils, NumberCheckReader, CommUtils;

type
  TExcelBomReader = class
  private             
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    procedure ReadSMT(irow_title: Integer);
    procedure ReadAsm(irow_title: Integer);
  public
    sProj: string;
    FList: TStringList;    
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

implementation
 
{ TExcelBomReader }

constructor TExcelBomReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create; 
  Open;
end;

destructor TExcelBomReader.Destroy;
begin
  Clear;
  FList.Free; 
  inherited;
end;

procedure TExcelBomReader.Clear;
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

procedure TExcelBomReader.Log(const str: string);
begin

end;

procedure TExcelBomReader.ReadSMT(irow_title: Integer);
var
  irow: Integer;
  aBomLoc: TBomLoc;
  snumber: string;
  aBomLocLinePtr: PBomLocLine;
  aBomLocLine: TBomLocLine;
  stitle: string;
begin

  aBomLoc := nil;

  irow := irow_title + 1;
  snumber := ExcelApp.Cells[irow, 2].Value +
    ExcelApp.Cells[irow + 1, 2].Value +
    ExcelApp.Cells[irow + 2, 2].Value +
    ExcelApp.Cells[irow + 3, 2].Value +
    ExcelApp.Cells[irow + 4, 2].Value;
  while (snumber <> '')  or IsCellMerged(ExcelApp, irow, 2, irow - 1, 2) do
  begin
    if (aBomLoc = nil) or not IsCellMerged(ExcelApp, irow, 3, irow - 1, 3) then // 第一行物料，跟标题不可能是合并单元格
    begin
      aBomLoc := TBomLoc.Create;
      aBomLoc.scategory := ExcelApp.Cells[irow, 1].Value;
      aBomLoc.susage := ExcelApp.Cells[irow, 3].Value;
      aBomLoc.sloc := ExcelApp.Cells[irow, 4].Value;
      FList.AddObject(IntToStr(irow), aBomLoc); 
    end;

    ZeroMemory(@aBomLocLine, SizeOf(aBomLocLine));
    aBomLocLine.snumber := ExcelApp.Cells[irow, 2].Value;
    aBomLocLine.sname := ExcelApp.Cells[irow, 5].Value;
    aBomLocLine.svendor := ExcelApp.Cells[irow, 6].Value;
    aBomLocLine.sver := ExcelApp.Cells[irow, 9].Value;   
    aBomLocLine.scap := ExcelApp.Cells[irow, 10].Value;
    aBomLocLine.scol := ExcelApp.Cells[irow, 11].Value;

    stitle := string(ExcelApp.Cells[irow, 1].Value) +
      string(ExcelApp.Cells[irow, 2].Value) +
      string(ExcelApp.Cells[irow, 3].Value) +
      string(ExcelApp.Cells[irow, 4].Value) +
      string(ExcelApp.Cells[irow, 5].Value) +
      string(ExcelApp.Cells[irow, 6].Value);
      
    if
//      (aBomLocLine.snumber = '') or
//      (aBomLocLine.sname = '') or
//      (aBomLocLine.sver = '') or
      (stitle = '物料类型物料编码数量位号物料名称厂家') then
    begin       
      aBomLoc := nil;
      irow := irow + 1;   
      snumber := ExcelApp.Cells[irow, 2].Value +
        ExcelApp.Cells[irow + 1, 2].Value +
        ExcelApp.Cells[irow + 2, 2].Value +
        ExcelApp.Cells[irow + 3, 2].Value +
        ExcelApp.Cells[irow + 4, 2].Value;
      Continue;
    end;  

    aBomLocLinePtr := New(PBomLocLine);
    aBomLocLinePtr^ := aBomLocLine; 
    aBomLoc.FList.AddObject(aBomLocLinePtr^.snumber, TObject(aBomLocLinePtr));
 
    irow := irow + 1;
    snumber := ExcelApp.Cells[irow, 2].Value +
      ExcelApp.Cells[irow + 1, 2].Value +
      ExcelApp.Cells[irow + 2, 2].Value +
      ExcelApp.Cells[irow + 3, 2].Value +
      ExcelApp.Cells[irow + 4, 2].Value;
  end; 
end;

procedure TExcelBomReader.ReadAsm(irow_title: Integer);
var
  irow: Integer;
  aBomLoc: TBomLoc;
  snumber: string;
  aBomLocLinePtr: PBomLocLine;
  aBomLocLine: TBomLocLine;
//  stitle: string;
begin

  aBomLoc := nil;

  irow := irow_title + 1;
  snumber := ExcelApp.Cells[irow, 6].Value +
    ExcelApp.Cells[irow + 1, 6].Value +
    ExcelApp.Cells[irow + 2, 6].Value +
    ExcelApp.Cells[irow + 3, 6].Value +
    ExcelApp.Cells[irow + 4, 6].Value;
  while (snumber <> '') or IsCellMerged(ExcelApp, irow, 6, irow - 1, 6) do
  begin

    if (aBomLoc = nil) or not IsCellMerged(ExcelApp, irow, 5, irow - 1, 5) then // 第一行物料，跟标题不可能是合并单元格
    begin
      aBomLoc := TBomLoc.Create;
      aBomLoc.scategory := ExcelApp.Cells[irow, 1].Value;
      aBomLoc.susage := ExcelApp.Cells[irow, 5].Value;
      FList.AddObject(IntToStr(irow), aBomLoc); 
    end;

    ZeroMemory(@aBomLocLine, SizeOf(aBomLocLine));
    aBomLocLine.snumber := ExcelApp.Cells[irow, 3].Value;
    aBomLocLine.sname := ExcelApp.Cells[irow, 4].Value;
    aBomLocLine.sver := ExcelApp.Cells[irow, 9].Value;
    aBomLocLine.scap := ExcelApp.Cells[irow, 10].Value;
    aBomLocLine.scol := ExcelApp.Cells[irow, 11].Value;

    if (aBomLocLine.snumber = '') or (aBomLocLine.sname = '') then
    begin
      aBomLoc := nil;
      irow := irow + 1;   
      snumber := ExcelApp.Cells[irow, 6].Value +
        ExcelApp.Cells[irow + 1, 6].Value +
        ExcelApp.Cells[irow + 2, 6].Value +
        ExcelApp.Cells[irow + 3, 6].Value +
        ExcelApp.Cells[irow + 4, 6].Value;
      Continue;
    end;

    aBomLocLinePtr := New(PBomLocLine);
    aBomLocLinePtr^ := aBomLocLine; 
    aBomLoc.FList.AddObject(aBomLocLinePtr^.snumber, TObject(aBomLocLinePtr));
 
    irow := irow + 1;
    snumber := ExcelApp.Cells[irow, 6].Value +
      ExcelApp.Cells[irow + 1, 6].Value +
      ExcelApp.Cells[irow + 2, 6].Value +
      ExcelApp.Cells[irow + 3, 6].Value +
      ExcelApp.Cells[irow + 4, 6].Value;
  end; 
end;

procedure TExcelBomReader.Open;
var
  iSheet: Integer;
  iSheetCount: Integer;
  sSheet: string;
  stitle: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6,
    stitle7: string;
//  stitle8, stitle9, stitle10, stitle11: string;
  irow: Integer;
  icol: Integer;
  irow_title: Integer;
begin
  Clear;


  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try

    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    sProj := ChangeFileExt( ExtractFileName(FFile), '' );

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
          icol := 1;
          stitle1 := ExcelApp.Cells[irow, icol].Value;
          stitle2 := ExcelApp.Cells[irow, icol + 1].Value;
          stitle3 := ExcelApp.Cells[irow, icol + 2].Value;
          stitle4 := ExcelApp.Cells[irow, icol + 3].Value;
          stitle5 := ExcelApp.Cells[irow, icol + 4].Value;  
          stitle6 := ExcelApp.Cells[irow, icol + 5].Value;
          stitle7 := ExcelApp.Cells[irow, icol + 6].Value;
//          stitle8 := ExcelApp.Cells[irow, icol + 7].Value;
//          stitle9 := ExcelApp.Cells[irow, icol + 8].Value;
//          stitle10 := ExcelApp.Cells[irow, icol + 9].Value;
//          stitle11 := ExcelApp.Cells[irow, icol + 10].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7; // + stitle8 + stitle9 + stitle10 + stitle11;
          if (stitle = '物料类型物料编码数量位号物料名称厂家变更项') or
            (stitle ='物料类型物料编码数量位号物料描述厂家变更项') then
          begin
            irow_title := irow;
            ReadSMT(irow_title);
            Break;
          end;
                                                                                                             
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;
          if stitle = '名称序号物料编码物料名称描述用量仓位' then
          begin
            irow_title := irow;
            ReadASM(irow_title);
            Break;
          end; 
        end;
                       
        if irow_title = -1 then
        begin
          Log(sSheet + ' 不是 Excel Bom 格式');
          Continue;
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
