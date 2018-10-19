unit ManMrpReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils;

type
  TManMrpLine = class
  private
    FDates: TStringList;
  public
    sver: string;
    snumber: string;
    sname: string;
    scolor: string;
    scap: string;
    sfg: string;
    spkg: string;

    dSum: Integer;

    constructor Create;
    destructor Destroy; override;
  end;

  TManMrpReader = class
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
    procedure GetSubs(const snumber: string; slSubs: TStringList);
  end;

implementation
 
constructor TManMrpLine.Create;
begin
  FDates := TStringList.Create;
end;

destructor TManMrpLine.Destroy;
begin
  FDates.Free;
end;

{ TManMrpReader }

constructor TManMrpReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TManMrpReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TManMrpReader.Clear;
var
  i: Integer;
  p: TManMrpLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := TManMrpLine(FList.Objects[i]);
    p.Free;
  end;
  FList.Clear;
end;

procedure TManMrpReader.GetSubs(const snumber: string; slSubs: TStringList);
var
  idx: Integer;
  i: Integer;
  p: TManMrpLine;    
  p0: TManMrpLine;
begin
  slSubs.Add(snumber);
  idx := FList.IndexOf(snumber);
  if idx < 0 then Exit;

  p0 := TManMrpLine(FList.Objects[idx]);

  for i := 0 to FList.Count - 1 do
  begin                         
    p := TManMrpLine(FList.Objects[i]);

    // 自己在前面已经第一个加了
    if i = idx then Continue;

    // 同项目
    if Copy(snumber, 1, 5) = Copy(p.snumber, 1, 5) then Continue;

    // 同项目 计划物料                                                             
    if '90.' + Copy(snumber, 3, 3) = Copy(p.snumber, 1, 5) then Continue;

    if (p0.sver = p.sver) and (p0.scolor = p.scolor) and (p0.scap = p.scap) then
    begin
      slSubs.Add(p.snumber);
    end;
  end;      

  for i := 0 to FList.Count - 1 do
  begin                         
    p := TManMrpLine(FList.Objects[i]);

    // 自己在前面已经第一个加了
    if i = idx then Continue;

    // 同项目
    if Copy(snumber, 1, 5) = Copy(p.snumber, 1, 5) then Continue;

    // 同项目 计划物料                                                             
    if '90.' + Copy(snumber, 3, 3) = Copy(p.snumber, 1, 5) then Continue;

    if (p0.scolor = p.scolor) and (p0.scap = p.scap) then
    begin
      if slSubs.IndexOf(p.snumber) >= 0 then Continue;
      slSubs.Add(p.snumber);
    end;
  end;
end;

procedure TManMrpReader.Log(const str: string);
begin

end;

procedure TManMrpReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  snumber: string;   
  aManMrpLine: TManMrpLine;
  slDate: TStringList;
  s: string;
  iDate: Integer;
  dQty: Integer;
begin
  Clear;
      
  if not FileExists(FFile) then Exit;


  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';

  slDate := TStringList.Create;
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
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7;
        if stitle <> '制式物料长代码物料名称颜色容量整机/裸机豪华装' then
        begin
          Log(sSheet +'  不是  物料需求进度表  格式  制式物料长代码物料名称颜色容量整机/裸机豪华装');
          Continue;
        end;

        slDate.Clear;
        
        icol := 8;
        s := ExcelApp.Cells[irow, icol].Value;
        while s <> '' do
        begin
          if Pos('合计', s) = 0 then
          begin
            slDate.AddObject(s, TObject(icol));
          end;
          icol := icol + 1;
          s := ExcelApp.Cells[irow, icol].Value;
        end;

        irow := 2;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin                                
          aManMrpLine := TManMrpLine.Create;
          FList.AddObject(snumber, aManMrpLine);

          aManMrpLine.sver := ExcelApp.Cells[irow, 1].Value;
          aManMrpLine.snumber := ExcelApp.Cells[irow, 2].Value;
          aManMrpLine.sname := ExcelApp.Cells[irow, 3].Value;
          aManMrpLine.scolor := ExcelApp.Cells[irow, 4].Value;
          aManMrpLine.scap := ExcelApp.Cells[irow, 5].Value;
          aManMrpLine.sfg := ExcelApp.Cells[irow, 6].Value;
          aManMrpLine.spkg := ExcelApp.Cells[irow, 7].Value;

          aManMrpLine.dSum := 0;
          for iDate := 0 to slDate.Count - 1 do
          begin
            s := slDate[iDate];
            icol := Integer(slDate.Objects[iDate]);
            dQty := ExcelApp.Cells[irow, icol].Value;
            aManMrpLine.FDates.AddObject(s, TObject(dQty));
            aManMrpLine.dSum := aManMrpLine.dSum + dQty;
          end;
               
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

    slDate.Free;
  end;  
end;


end.                              
