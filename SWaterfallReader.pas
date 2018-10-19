unit SWaterfallReader;

interface

uses
  Classes, ComObj, Variants, SysUtils, CommUtils;

type
  TSWFLine = class;
  TSWFCol = class;

  TSWFProj = class
  private
    FList: TStringList;
    function GetLineCount: Integer;
    function GetLines(i: Integer): TSWFLine;
  public
    FName: string;
    constructor Create(const sname: string);
    destructor Destroy; override;
    procedure Clear;
    property LineCount: Integer read GetLineCount;
    property Lines[i: Integer]: TSWFLine read GetLines;
  end;

  TSWFLine = class
  private
    function GetDateCount: Integer;
    function GetDates(i: Integer): TSWFCol;
  public
    FName: string;
    FDate: TDateTime;
    FList: TList;
    FRow: Integer;
    FPayCol: Integer;
    constructor Create(const sname: string);
    destructor Destroy; override;
    procedure Clear;
    property DateCount: Integer read GetDateCount;
    property Dates[i: Integer]: TSWFCol read GetDates;
  end;

  TSWFCol = class
  public
    FWeek: string;
    FDate: TDateTime;
    FQty: Double;
  end;

  TSWaterfallReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
    function GetProjCount: Integer;
    function GetProjs(i: Integer): TSWFProj;
  public
    FList: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    property ProjCount: Integer read GetProjCount;
    property Projs[i: Integer]: TSWFProj read GetProjs;
    function GetProj(const sname: string): TSWFProj;
  end;

implementation

{ TSWFProj }

constructor TSWFProj.Create(const sname: string);
begin
  FName := sname;
  FList := TStringList.Create;
end;

destructor TSWFProj.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSWFProj.Clear;
var
  i: Integer;
  aSWFLine: TSWFLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSWFLine := TSWFLine(FList.Objects[i]);
    aSWFLine.Free;
  end;
  FList.Clear;
end;

function TSWFProj.GetLineCount: Integer;
begin
  Result := FList.Count;
end;

function TSWFProj.GetLines(i: Integer): TSWFLine;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := TSWFLine(FList.Objects[i]);
  end
  else Result := nil;
end;

{ TSWFLine }
 
constructor TSWFLine.Create(const sname: string);
begin
  FName := sname;
  FList := TList.Create;
end;

destructor TSWFLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSWFLine.Clear;
var
  i: Integer;
  aSWFCol: TSWFCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSWFCol := TSWFCol(FList[i]);
    aSWFCol.Free;
  end;
  FList.Clear;
end;

function TSWFLine.GetDateCount: Integer;
begin
  Result := FList.Count;
end;

function TSWFLine.GetDates(i: Integer): TSWFCol;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := TSWFCol(FList[i]);
  end
  else Result := nil;
end;

{ TSWaterfallReader }
                   
constructor TSWaterfallReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TSWaterfallReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TSWaterfallReader.Clear;
var
  i: Integer;
  aSWFProj: TSWFProj;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSWFProj := TSWFProj(FList.Objects[i]);
    aSWFProj.Free;
  end;
  FList.Clear;
end;

function TSWaterfallReader.GetProj(const sname: string): TSWFProj;
var
  i: Integer;
  aSWFProj: TSWFProj;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSWFProj := TSWFProj(FList.Objects[i]);
    if aSWFProj.FName = sname then
    begin
      Result := aSWFProj;
      Break;
    end;
  end;
end;

procedure TSWaterfallReader.Log(const str: string);
begin

end;

function TSWaterfallReader.GetProjCount: Integer;
begin
  Result := FList.Count;
end;

function TSWaterfallReader.GetProjs(i: Integer): TSWFProj;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := TSWFProj(FList.Objects[i]);
  end
  else Result := nil;
end;

procedure TSWaterfallReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3: string;
  stitle: string;
  irow: Integer;
  icol: Integer;
  slWeek: TStringList;
  iWeek: Integer;
  sweek: string;
  dt: TDateTime;
  sdate: string;
  v: Variant;
  sname: string;   
  aSWFProj: TSWFProj;
  aSWFLine: TSWFLine;
  aSWFCol: TSWFCol;
begin
  Clear;

  if not FileExists(FFile) then Exit;

  slWeek := TStringList.Create;

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

        irow := 3;
        stitle1 := ExcelApp.Cells[irow, 2].Value;
        stitle2 := ExcelApp.Cells[irow, 3].Value;
        stitle3 := ExcelApp.Cells[irow, 4].Value;
        stitle := stitle1 + stitle2 + stitle3;
        if stitle <> 'MPS版本需求日期责任总量' then
        begin
          Log('sheet ' + sSheet + ' 不是Simple waterfall 格式');
          Continue;
        end;

        irow := 3;
        icol := 5;
        v := ExcelApp.Cells[irow, icol].Value;
        while VarIsType(v, varDate) do
        begin
          sweek := ExcelApp.Cells[irow - 1, icol].Value;
          dt := ExcelApp.Cells[irow, icol].Value;
          sdate := FormatDateTime('yyyy-MM-dd', dt);
          slWeek.Add(sweek + '=' + sdate);
          
          icol := icol + 1;
          v := ExcelApp.Cells[irow, icol].Value;
        end;

        if slWeek.Count = 0 then
        begin
          Log('sheet ' + sSheet + ' 没有week数据');
          Continue;
        end;

        aSWFProj := TSWFProj.Create(sSheet);
        FList.AddObject(sSheet, aSWFProj);

        irow := 4;
        sname := ExcelApp.Cells[irow, 2].Value;
        while sname <> '' do
        begin                                
          aSWFLine := TSWFLine.Create(sname);
          aSWFProj.FList.AddObject(sname, aSWFLine);

          aSWFLine.FDate := ExcelApp.Cells[irow, 3].Value;

          for iWeek := 0 to slWeek.Count - 1 do
          begin
            aSWFCol := TSWFCol.Create;
            aSWFCol.FWeek := slWeek.Names[iWeek];
            aSWFCol.FDate := myStrToDateTime(slWeek.ValueFromIndex[iWeek]);
            aSWFCol.FQty := ExcelApp.Cells[irow, iWeek + 5].Value;
            aSWFLine.FList.Add(aSWFCol);
          end;  
 
          irow := irow + 1;
          sname := ExcelApp.Cells[irow, 2].Value;
        end;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;

    slWeek.Free;
  end;  
end;

end.
