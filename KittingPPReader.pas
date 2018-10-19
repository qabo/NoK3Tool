unit KittingPPReader;

interface
          
uses
  Classes, SysUtils, ComObj, CommUtils, KittingKeyNumberReader;

type
  TKittingPPSheet = class;

  TKittingChild = class
  public
    FNumber: string;
    FLT: Double;
    FUsage: Double;
  end;

  TKittingNumber = class
  private
    FRows: TStringList;
    FChilds: TStringList;
    function GetRom: string;
    function GetRam: string;
    function GetName: string;
    procedure FindPPLines(aKittingPPSheet: TKittingPPSheet);
    procedure Clear;
  public
    FProj: string;
    FColor: string;
    FCap: string;
    FVer: string;
    constructor Create;
    destructor Destroy; override;
    procedure AddChildIfOk(aKittingKeyNumber: TKittingKeyNumber);
    property Rom: string read GetRom;
    property Ram: string read GetRam;
    property Name: string read GetName;
  end;
  
  TKittingPPCol = class
  public
    sdate: string;
    FQty: Double;
  end;

  TKittingPPRow = class
  private
    FCols: TStringList;
    function GetColCount: Integer;
    function GetCols(i: Integer): TKittingPPCol;
    procedure Clear;
  public
    FVer: string;
    FNumber: string;
    FColor: string;
    FCap: string;
    FQtyMade: Double;
    FQtyToMake: Double;         
    constructor Create;
    destructor Destroy; override;
    
    property ColCount: Integer read GetColCount;
    property Cols[i: Integer]: TKittingPPCol read GetCols;
  end;

  TKittingPPSheet = class
  private
    FVers: TStringList;
    FColors: TStringList;
    FCaps: TStringList;
    FDates: TStringList;
    FRows: TStringList;
    function GetDateCount: Integer;
    function GetDates(i: Integer): string;

    function GetRowCount: Integer;
    function GetRows(i: Integer): TKittingPPRow;
    function GetProjName: string;
    procedure Clear;
  public
    FName: string;
    constructor Create;
    destructor Destroy; override;

    procedure GenCombos(sl: TStringList);

    property DateCount: Integer read GetDateCount;
    property Dates[i: Integer]: string read GetDates;
    property RowCount: Integer read GetRowCount;
    property Rows[i: Integer]: TKittingPPRow read GetRows;
    property ProjName: string read GetProjName;
  end;

  TKittingPPReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FLogEvent: TLogEvent;
    FSheets: TStringList;

    procedure Open;
    procedure Log(const str: string);

    function GetCount: Integer;
    function GetSheets(i: Integer): TKittingPPSheet;
  public 
    constructor Create(const sfile: string; aLogEvent: TLogEvent = nil);
    destructor Destroy; override;
    procedure Clear;

    property Count: Integer read GetCount;
    property Sheets[i: Integer]: TKittingPPSheet read GetSheets;
  end;

implementation

{ TKittingNumber }

constructor TKittingNumber.Create;
begin
  FRows := TStringList.Create;
  FChilds := TStringList.Create;
end;

destructor TKittingNumber.Destroy;
begin
  Clear;
  FRows.Free;
  FChilds.Free;
  inherited;
end;

procedure TKittingNumber.Clear;
var
  i: Integer;
  aKittingChild: TKittingChild;
begin
  FRows.Clear;

  for i := 0 to FChilds.Count - 1 do
  begin
    aKittingChild := TKittingChild(FChilds.Objects[i]);
    aKittingChild.Free;
  end;
  FChilds.Clear;
end;

function TKittingNumber.GetRom: string;
var
  ipos: Integer;
begin
  ipos := Pos('+', FCap);
  if ipos > 0 then
  begin
    Result := Copy(FCap, ipos + 1, Length(FCap));
  end
  else
  begin
    Result := '';
  end;
end;

function TKittingNumber.GetRam: string;
var
  ipos: Integer;
begin
  ipos := Pos('+', FCap);
  if ipos > 0 then
  begin
    Result := Copy(FCap, 1, ipos - 1);
  end
  else
  begin
    Result := '';
  end;
end;

function TKittingNumber.GetName: string;
begin
  Result := FProj + FVer + FColor + FCap;
end;

procedure TKittingNumber.FindPPLines(aKittingPPSheet: TKittingPPSheet);
var
  irow: Integer;
  aKittingPPRow: TKittingPPRow;
begin
  for irow := 0 to aKittingPPSheet.RowCount - 1 do
  begin
    aKittingPPRow := aKittingPPSheet.Rows[irow];
    if (aKittingPPRow.FColor = FColor)
      and (aKittingPPRow.FCap = FCap)
      and (aKittingPPRow.FVer = FVer) then
    begin
      FRows.AddObject(aKittingPPRow.FNumber, aKittingPPRow);
    end;
  end
end;

procedure TKittingNumber.AddChildIfOk(aKittingKeyNumber: TKittingKeyNumber);
var
  aKittingChild: TKittingChild;
begin
  if aKittingKeyNumber.sproj <> FProj then Exit;

  if (aKittingKeyNumber.sver <> '通用') and
    (aKittingKeyNumber.sver <> FVer) then Exit;

  if (aKittingKeyNumber.scolor <> '通用') and
    (aKittingKeyNumber.scolor <> FColor) then Exit;

  if (aKittingKeyNumber.scap <> '通用') and
    (aKittingKeyNumber.scap <> FCap) then Exit;

  aKittingChild := TKittingChild.Create;
  FChilds.AddObject(aKittingKeyNumber.name, aKittingChild);
end;

{ TKittingPPCol } 

{ TKittingPPRow }

constructor TKittingPPRow.Create;
begin
  FCols := TStringList.Create;
end;

destructor TKittingPPRow.Destroy;
begin
  Clear;
  FCols.Free;
end;

procedure TKittingPPRow.Clear;
var
  i: Integer;
  aKittingPPCol: TKittingPPCol;
begin
  for i := 0 to FCols.Count - 1 do
  begin
    aKittingPPCol := TKittingPPCol(FCols.Objects[i]);
    aKittingPPCol.Free;
  end;
  FCols.Clear;
end;

function TKittingPPRow.GetColCount: Integer;
begin
  Result := FCols.Count;
end;

function TKittingPPRow.GetCols(i: Integer): TKittingPPCol;
begin
  Result := TKittingPPCol(FCols.Objects[i]);
end;  

{ TKittingPPSheet }

constructor TKittingPPSheet.Create;
begin
  FVers := TStringList.Create;
  FColors := TStringList.Create;
  FCaps := TStringList.Create;
  FDates := TStringList.Create;
  FRows := TStringList.Create;
end;

destructor TKittingPPSheet.Destroy;
begin
  Clear;
  FVers.Free;
  FColors.Free;
  FCaps.Free;
  FDates.Free;
  FRows.Free;
end;

procedure TKittingPPSheet.Clear;
var
  i: Integer;
  aKittingPPRow: TKittingPPRow;
begin
  FVers.Clear;
  FColors.Clear;
  FCaps.Clear;
  FDates.Clear;
  for i := 0 to FRows.Count - 1  do
  begin
    aKittingPPRow := TKittingPPRow(FRows.Objects[i]);
    aKittingPPRow.Free;
  end;
  FRows.Clear;
end;

function TKittingPPSheet.GetDateCount: Integer;
begin
  Result := FDates.Count;
end;

function TKittingPPSheet.GetDates(i: Integer): string;
begin
  Result := FDates[i];
end;

function TKittingPPSheet.GetRowCount: Integer;
begin
  Result := FRows.Count;
end;

function TKittingPPSheet.GetRows(i: Integer): TKittingPPRow;
begin
  Result := TKittingPPRow(FRows.Objects[i]);
end;

function TKittingPPSheet.GetProjName: string;
var
  ipos: Integer;
begin
  ipos := Pos('_', FName);
  if ipos = 0 then
  begin
    Result := FName;
  end
  else
  begin
    Result := Copy(FName, 1, ipos - 1);
  end;
end;

procedure TKittingPPSheet.GenCombos(sl: TStringList);
var
  iver: Integer;
  icolor: Integer;
  icap: Integer;
  inumber: Integer;
  aKittingNumber: TKittingNumber;
begin
//  for inumber := 0 to sl.Count - 1 do
//  begin
//    aKittingNumber := TKittingNumber(sl.Objects[inumber]);
//    aKittingNumber.Free;
//  end;  
//  sl.Clear;

  for iver := 0 to FVers.Count - 1 do
  begin
    for icolor := 0 to FColors.Count - 1 do
    begin
      for icap := 0 to FCaps.Count - 1 do
      begin
        aKittingNumber := TKittingNumber.Create;
        aKittingNumber.FProj := self.ProjName;
        aKittingNumber.FVer := FVers[iVer];
        aKittingNumber.FColor := FColors[icolor];
        aKittingNumber.FCap := FCaps[icap];

        aKittingNumber.FindPPLines(Self);

        sl.AddObject(aKittingNumber.Name, aKittingNumber);
      end;
    end;
  end;
end;  

{ TKittingPPReader }

constructor TKittingPPReader.Create(const sfile: string;
  aLogEvent: TLogEvent = nil);
begin
  FFile := sfile;
  FLogEvent := aLogEvent;
  FSheets := TStringList.Create;
  Open;
end;

destructor TKittingPPReader.Destroy;
begin
  Clear;
  FSheets.Free;
  inherited;
end;

procedure TKittingPPReader.Clear;
var
  i: Integer;
  aKittingPPSheet: TKittingPPSheet;
begin
  for i := 0 to FSheets.Count - 1 do
  begin
    aKittingPPSheet := TKittingPPSheet(FSheets.Objects[i]);
    aKittingPPSheet.Free;
  end;
  FSheets.Clear;
end;

procedure TKittingPPReader.Log(const str: string);
begin
  if Assigned(FLogEvent) then
  begin
    FLogEvent(str);
  end;
end;

function TKittingPPReader.GetCount: Integer;
begin
  Result := FSheets.Count;
end;

function TKittingPPReader.GetSheets(i: Integer): TKittingPPSheet;
begin
  Result := TKittingPPSheet(FSheets.Objects[i]);
end;  

procedure TKittingPPReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4,
  stitle5, stitle6, stitle7: string;
  stitle: string;
  irow: Integer;
  icol: Integer;  
  sdate: string;
  snumber: string;    
  sver: string;
  scolor: string;
  scap: string;
  irow1: Integer;
  i: Integer;

  aKittingPPSheet: TKittingPPSheet;
  aKittingPPRow: TKittingPPRow;
  aKittingPPCol: TKittingPPCol;
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
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;   
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6 + stitle7;
        if stitle <> '标准制式物料编码颜色容量模拟分类累计历史产出未来计划产出' then
        begin
          Log(sSheet +'  不是   生产计划模拟');
          Continue;
        end;
                
        aKittingPPSheet := TKittingPPSheet.Create;
        aKittingPPSheet.FName := sSheet;
        FSheets.AddObject(sSheet, aKittingPPSheet);

        icol := 8;
        sdate := ExcelApp.Cells[irow, icol].value;
        while sdate <> '' do
        begin
          aKittingPPSheet.FDates.AddObject(sdate, TObject(icol));
          icol := icol + 1;    
          sdate := ExcelApp.Cells[irow, icol].value;
        end;  

        irow1 := 2;
        irow := 3;
        snumber := ExcelApp.Cells[irow, 2].Value;
        while snumber <> '' do
        begin
          aKittingPPRow := TKittingPPRow.Create;
          aKittingPPSheet.FRows.AddObject(snumber, aKittingPPRow);

          if not IsCellMerged(ExcelApp, irow1, 1, irow, 1) then
          begin
            sver := ExcelApp.Cells[irow, 1].Value;
            irow1 := irow;
          end;

          scolor := ExcelApp.Cells[irow, 3].Value;
          scap := ExcelApp.Cells[irow, 4].Value;

          aKittingPPRow.FVer := sver;
          aKittingPPRow.FNumber := snumber;
          aKittingPPRow.FColor := scolor;
          aKittingPPRow.FCap := scap;
          aKittingPPRow.FQtyMade := ExcelApp.Cells[irow, 6].Value;
          aKittingPPRow.FQtyToMake := ExcelApp.Cells[irow, 7].Value;

          if aKittingPPSheet.FVers.IndexOf(sver) < 0 then
          begin
            aKittingPPSheet.FVers.Add(sver);
          end;
          if aKittingPPSheet.FColors.IndexOf(scolor) < 0 then
          begin
            aKittingPPSheet.FColors.Add(scolor);
          end;
          if aKittingPPSheet.FCaps.IndexOf(scap) < 0 then
          begin
            aKittingPPSheet.FCaps.Add(scap);
          end;

          for i := 0 to aKittingPPSheet.FDates.Count - 1 do
          begin
            sdate := aKittingPPSheet.FDates[i];
            icol := Integer(aKittingPPSheet.FDates.Objects[i]);
            aKittingPPCol := TKittingPPCol.Create;
            aKittingPPRow.FCols.AddObject(sdate, aKittingPPCol);
            aKittingPPCol.sdate := sdate;
            try
              aKittingPPCol.FQty := ExcelApp.Cells[irow, icol].Value;
            except
              on e: Exception do
              begin
                Log(Format('sheet: %s, irow: %d, icol: %d  不是有效 数值 格式', [sSheet, irow, icol]));
                raise e;
              end;
            end;
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
  end;  
end;


end.                              
