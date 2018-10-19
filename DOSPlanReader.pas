unit DOSPlanReader;

interface
             
uses
  Classes, SysUtils, ComObj, CommUtils;

type
       
  TProjDOSCol = class
  public
    FWeek: string;
    FDate: string;
    fdos: Double;
  end;

  TProjDOSLine = class
  private
    FList: TStringList;
  public
    FNumber: string;
    FProj: string;
    FVer: string;
    FColor: string;
    FCap: string;

    constructor Create(const snumber: string);
    destructor Destroy; override;
    procedure Clear;
    function GetDosPlan(const sdate: string): Double;
  end;

  TDOSPlanProj = class
  private
    FList: TStringList;
  public
    FName: string;
    constructor Create(const sname: string);
    destructor Destroy; override;
    procedure Clear;
    function GetDosPlan(const sNumber: string; const sdate: string): Double;
  end;

  TDOSPlanReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;     
    FList: TStringList;
    procedure Open;
    procedure Log(const str: string);
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetDosPlan(const sproj, sNumber, sdate: string): Double; 
  end;

implementation
        
{ TProjDOSLine }

constructor TProjDOSLine.Create(const snumber: string);
begin
  FNumber := snumber;
  FList := TStringList.Create;
end;

destructor TProjDOSLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TProjDOSLine.Clear;
var
  i: Integer;
  aProjDOSCol: TProjDOSCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOSCol := TProjDOSCol(FList.Objects[i]);
    aProjDOSCol.Free;
  end;
  FList.Clear;
end;

function TProjDOSLine.GetDosPlan(const sdate: string): Double;
var
  i: Integer;
  aProjDOSCol: TProjDOSCol;
begin
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOSCol := TProjDOSCol(FList.Objects[i]);
    savelogtoexe('aProjDOSCol.FDate: ' + aProjDOSCol.FDate + '   sdate: ' + sdate);
    if aProjDOSCol.FDate = sdate then
    begin
      Result := aProjDOSCol.fdos;
      Break;
    end;
  end;
end;

{ TDOSPlanProj }

constructor TDOSPlanProj.Create(const sname: string);
begin
  FName := sname;
  FList := TStringList.Create;
end;

destructor TDOSPlanProj.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TDOSPlanProj.Clear;
var
  i: Integer;
  aProjDOSLine: TProjDOSLine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOSLine := TProjDOSLine(FList.Objects[i]);
    aProjDOSLine.Free;
  end;
  FList.Clear;
end;

function TDOSPlanProj.GetDosPlan(const sNumber: string; const sdate: string): Double;
var
  i: Integer;
  aProjDOSLine: TProjDOSLine;
begin 
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOSLine := TProjDOSLine(FList.Objects[i]);
    if aProjDOSLine.FNumber = sNumber then
    begin
      Result := aProjDOSLine.GetDosPlan(sdate);
      Break;
    end;  
  end;
end;

{ TDOSPlanReader }

constructor TDOSPlanReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  Open;
end;

destructor TDOSPlanReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TDOSPlanReader.Clear;
var
  i: Integer;
  aProjDOS: TDOSPlanProj;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOS := TDOSPlanProj(FList.Objects[i]);
    aProjDOS.Free;
  end;
  FList.Clear;
end;

procedure TDOSPlanReader.Log(const str: string);
begin
  savelogtoexe(str);
end;

function TDOSPlanReader.GetDosPlan(const sproj, sNumber, sdate: string): Double;
var
  i: Integer;
  aProjDOS: TDOSPlanProj;
begin
  savelogtoexe('GetDosPlan  ' + sNumber);
  Result := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aProjDOS := TDOSPlanProj(FList.Objects[i]);   
    savelogtoexe('GetDosPlan  aProjDOS.FName: ' + aProjDOS.FName + '   sproj: ' + sproj);
    if aProjDOS.FName = sproj then
    begin
      Result := aProjDOS.GetDosPlan(sNumber, sdate);
      Break;
    end;
  end;
end;

procedure TDOSPlanReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  irow: Integer; 
  icol: Integer;
  snumber: string;
  aDOSPlanProj: TDOSPlanProj;
  sldate: TStringList;
  sdate: string;
  sweek: string;
  idate: Integer;
  aProjDOSLine: TProjDOSLine;
  aProjDOSCol: TProjDOSCol;
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
        if stitle <> '项目制式物料编码颜色容量' then
        begin            
          Log(sSheet +'  不是DOS计划格式');
          Continue;
        end;

        aDOSPlanProj := TDOSPlanProj.Create(sSheet);
        FList.AddObject(sSheet, aDOSPlanProj);
             
        sldate := TStringList.Create;
        try         
          irow := 1;
          icol := 6;   
          sweek := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
          while Trim(sdate + sweek) <> '' do
          begin
            if not IsCellMerged(ExcelApp, irow, icol, irow + 1, icol) then
            begin
              sldate.AddObject(sdate + '=' + sweek, TObject(icol));
            end;
            icol := icol + 1;             
            sweek := ExcelApp.Cells[irow, icol].Value;
            sdate := ExcelApp.Cells[irow + 1, icol].Value;
          end;


          irow := 3;
          snumber := ExcelApp.Cells[irow, 3].Value;
          while snumber <> '' do
          begin
            savelogtoexe(snumber);
            aProjDOSLine := TProjDOSLine.Create(snumber);
            aDOSPlanProj.FList.AddObject(snumber, aProjDOSLine);
            
            aProjDOSLine.FProj := ExcelApp.Cells[irow, 1].Value;
            aProjDOSLine.FVer := ExcelApp.Cells[irow, 2].Value;    
            aProjDOSLine.FColor := ExcelApp.Cells[irow, 4].Value;  
            aProjDOSLine.FCap := ExcelApp.Cells[irow, 5].Value;

            for idate := 0 to sldate.Count - 1 do
            begin
              icol := Integer(sldate.Objects[idate]);

              aProjDOSCol := TProjDOSCol.Create;        
              aProjDOSLine.FList.AddObject(sldate[idate], aProjDOSCol);

              aProjDOSCol.FWeek := sldate.ValueFromIndex[idate];
              aProjDOSCol.FDate := sldate.Names[idate];
              aProjDOSCol.fdos := ExcelApp.Cells[irow, icol].Value;
            end;

            irow := irow + 1;
            snumber := ExcelApp.Cells[irow, 3].Value;
          end; 
        finally
          sldate.Free;
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
