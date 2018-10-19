unit SOPReaderUnit;

interface

uses
  Windows, Classes, SysUtils, ComObj, Variants, CommUtils;

type
  TSOPProj = class;
  TSOPCol = class;
  TSOPLine = class;

  
  TSOPReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string);
  public                        
    FProjs: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetProj(const sName: string): TSOPProj;
  end;

  TSOPProj = class
  private
  public
    FName: string;
    FList: TStringList;
    slMonths: TStringList;
    constructor Create(const sproj: string);
    destructor Destroy; override;
    procedure Clear;
    function GetLine(const sVer, sNumber, sColor, sCap: string): TSOPLine;
  end;

  TSOPLine = class
  private
  public
    sVer: string; //制式
    sNumber: string; //	物料编码
    sColor: string; //颜色
    sCap: string; //容量
    FList: TStringList;
    FCalc: Boolean;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    procedure Add( const sMonth, sWeek, sDate: string; iQty: Integer ); overload; 
    procedure Add( const sMonth, sWeek, sDate: string; iQty_sop, iQty_mps: Integer ); overload;
    function GetCol(const sDate: string): TSOPCol;
  end;

  TSOPCol = class
  public
    sMonth: string;
    sWeek: string;
    sDate: string;
    iQty: Integer;
    iQty_sop: Integer;
    iQty_mps: Integer;
  end;

implementation
 
{ TSOPProj }

constructor TSOPProj.Create(const sproj: string);
begin
  FName := sproj;
  FList := TStringList.Create;
  slMonths := TStringList.Create;
end;

destructor TSOPProj.Destroy;
begin
  Clear;
  FList.Free;
  slMonths.Free;
  inherited;
end;

procedure TSOPProj.Clear;
var
  i: Integer;
  aSOPLine: TSOPLine;
begin
  for i := 0 to slMonths.Count - 1 do
  begin
    TStringList(slMonths.Objects[i]).Free;
  end;
  slMonths.Clear;

  for i := 0 to FList.Count - 1 do
  begin
    aSOPLine := TSOPLine(FList.Objects[i]);
    aSOPLine.Free;
  end;
  FList.Clear;
end;

function TSOPProj.GetLine(const sVer, sNumber, sColor, sCap: string): TSOPLine;     
var
  i: Integer;
  aSOPLine: TSOPLine;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPLine := TSOPLine(FList[i]);
    if (aSOPLine.sVer = sVer) and (aSOPLine.sNumber = sNumber)
      and (aSOPLine.sColor = sColor) and (aSOPLine.sCap = sCap) then
    begin
      Result := aSOPLine;
      Break;
    end;
  end;
end;

{ TSOPLine }
 
constructor TSOPLine.Create;
begin
  FList := TStringList.Create;
  FCalc := False;
end;

destructor TSOPLine.Destroy;
begin
  Clear;
  FList.Free;
end;

procedure TSOPLine.Clear;
var
  i: Integer;
  aSOPCol: TSOPCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aSOPCol := TSOPCol(FList.Objects[i]);
    aSOPCol.Free;
  end;
  FList.Clear;
end;

procedure TSOPLine.Add( const sMonth, sWeek, sDate: string; iQty: Integer );
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := TSOPCol.Create;
  aSOPCol.sMonth := sMonth;
  aSOPCol.sWeek := sWeek;
  aSOPCol.sDate := sDate;
  aSOPCol.iQty := iQty;
  FList.AddObject(sDate, aSOPCol);
end;   

procedure TSOPLine.Add( const sMonth, sWeek, sDate: string; iQty_sop, iQty_mps: Integer );
var
  aSOPCol: TSOPCol;
begin
  aSOPCol := TSOPCol.Create;
  aSOPCol.sMonth := sMonth;
  aSOPCol.sWeek := sWeek;
  aSOPCol.sDate := sDate;
  aSOPCol.iQty_sop := iQty_sop;
  aSOPCol.iQty_mps := iQty_mps;
  FList.AddObject(sDate, aSOPCol);
end;

function TSOPLine.GetCol(const sDate: string): TSOPCol;
var
  i: Integer;
  aSOPCol: TSOPCol;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aSOPCol := TSOPCol(FList[i]);
    if aSOPCol.sDate = sDate then
    begin
      Result := aSOPCol;
      Break;
    end;
  end;
end;

{ TSOPReader }

constructor TSOPReader.Create(const sfile: string);
begin
  FFile := sfile;

  FProjs := TStringList.Create;

  Open;


end;

destructor TSOPReader.Destroy;
begin
  Clear;
  FProjs.Free;
end;

procedure TSOPReader.Clear;
var
  i: Integer;
  aSOPProj: TSOPProj;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPProj := TSOPProj(FProjs.Objects[i]);
    aSOPProj.Free;
  end;
  FProjs.Clear;
end;

function TSOPReader.GetProj(const sName: string): TSOPProj;
var
  i: Integer;
  aSOPProj: TSOPProj;
begin
  Result := nil;
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPProj := TSOPProj(FProjs.Objects[i]);
    if aSOPProj.FName = sName then
    begin
      Result := aSOPProj;
      Break;
    end;
  end;
end;

procedure TSOPReader.Log(const str: string);
begin

end;

procedure TSOPReader.Open;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sproj: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8: string;
  stitle4x, stitle8x: string;
  irow, icol: Integer;
  icol1: Integer;
  smonth: string;
  sweek: string;
  sdate: string;
  irow1: Integer;
  icolDate1: Integer;
  icolVer: Integer; //制式
  icolNumber: Integer; //	物料编码
  icolColor: Integer; //颜色
  icolCap: Integer; //容量
			
  sVer: string; //制式
  sNumber: string; //	物料编码
  sColor: string; //颜色
  sCap: string; //容量
  v: Variant;
  iQty: Integer;

  slWeeks: TStringList; 
  iMonth: Integer;
  iWeek: Integer;
  
  aSOPLine: TSOPLine;
  aSOPProj: TSOPProj;
  iProj: Integer;
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


        irow := 1;
        stitle1 := ExcelApp.Cells[irow, 1].Value;  
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;  
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle8 := ExcelApp.Cells[irow, 8].Value;
        stitle4x := stitle1 + stitle2 + stitle3 + stitle4;
        stitle8x := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6 + stitle7 + stitle8; 
        if stitle4x = '制式物料编码颜色容量' then
        begin
          icolVer := 1; // Integer; //制式
          icolNumber := 2; // Integer; //	物料编码
          icolColor := 3; // Integer; //颜色
          icolCap := 4; // Integer; //容量        
          icolDate1 := 5;
        end
        else if stitle8x = '项目整机/裸机包装标准制式制式物料编码颜色容量' then
        begin
          icolVer := 5; // Integer; //制式
          icolNumber := 6; // Integer; //	物料编码
          icolColor := 7; // Integer; //颜色
          icolCap := 8; // Integer; //容量
          icolDate1 := 9;
        end
        else
        begin     
          icolDate1 := 0;
          Log(sSheet + ' 不是SOP格式');
          Continue;
        end;  

        sproj := sSheet;
        if Pos(' ', sSheet) > 0 then
        begin
          sproj := Copy(sSheet, 1, Pos(' ', sSheet) - 1);
        end;
        aSOPProj := TSOPProj.Create(sproj);
        FProjs.AddObject(sproj, aSOPProj);
   
        
        irow := 1;
        icol := icolDate1;
        sweek := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        icol1 := icol;      

        slWeeks := TStringList.Create;
        while Trim(sweek + sdate) <> '' do
        begin
          if IsCellMerged(ExcelApp, irow, icol, irow + 1, icol) then
          begin
            if icol > icol1 then
            begin
              smonth := ExcelApp.Cells[irow, icol].Value;
              if slWeeks.Count > 0 then
              begin
                aSOPProj.slMonths.AddObject(smonth, slWeeks);
                slWeeks := TStringList.Create;
              end;
            end;
            
            icol := icol + 1;
            sweek := ExcelApp.Cells[irow, icol].Value;
            sdate := ExcelApp.Cells[irow + 1, icol].Value;
            Continue;
          end;                                          

          slWeeks.AddObject(sweek + '=' + sdate, TObject(icol));
          
          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
        end;
        slWeeks.Free;
        

        irow := 3;
        irow1 := 0;
        while not IsCellMerged(ExcelApp, irow, icolNumber, irow, icolCap) do
        begin
          if (irow1 = 0) or
            not IsCellMerged(ExcelApp, irow1, icolVer, irow, icolVer) then
          begin
            irow1 := irow;       
            sVer := ExcelApp.Cells[irow, icolVer].Value;
          end;   
          sNumber := ExcelApp.Cells[irow, icolNumber].Value;
          sColor := ExcelApp.Cells[irow, icolColor].Value;
          sCap := ExcelApp.Cells[irow, icolCap].Value;


          aSOPLine := TSOPLine.Create;
          aSOPProj.FList.AddObject(sNumber, aSOPLine);

          aSOPLine.sVer := sVer;
          aSOPLine.sNumber := sNumber;
          aSOPLine.sColor := sColor;
          aSOPLine.sCap := sCap;

          for iMonth := 0 to aSOPProj.slMonths.Count - 1 do
          begin
            slWeeks := TStringList(aSOPProj.slMonths.Objects[iMonth]);
            for iWeek := 0 to slWeeks.Count - 1 do
            begin
              icol := Integer( slWeeks.Objects[iWeek] );
              v := ExcelApp.Cells[irow, icol].Value;
              if VarIsNumeric(v) then
              begin
                iQty := v;
              end
              else iQty := 0;

              aSOPLine.Add(aSOPProj.slMonths[iMonth], slWeeks.Names[iWeek], slWeeks.ValueFromIndex[iWeek], iQty);
            end;
          end;
       
          irow := irow + 1;
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
