unit SOPVSActReaderUnit;

interface
  
uses
  Windows, Classes, SysUtils, ComObj, Variants, CommUtils, SOPReaderUnit;

type
  TSOPVSActProj = class;
  
  TSOPVSActReader = class
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FProjs: TStringList;
    procedure Open;
    procedure Log(const str: string);
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
    function GetProj(const sName: string): TSOPVSActProj;
  end;

  TSOPVSActProj = class
  public
    slWeeks: TStringList;
    FList: TList;
    FName: string;
    constructor Create(const sproj: string);
    destructor Destroy; override;
    procedure Clear;
  end;

implementation
                 
{ TSOPVSActProj }

constructor TSOPVSActProj.Create(const sproj: string);
begin
  FName := sproj;
  FList := TList.Create;  
  slWeeks := TStringList.Create;
end;

destructor TSOPVSActProj.Destroy;
begin
  Clear;
  FList.Free;
  slWeeks.Free;
end;

procedure TSOPVSActProj.Clear;
var
  i: Integer;
  aSOPLine: TSOPLine;
begin
  slWeeks.Clear;
  
  for i := 0 to FList.Count - 1 do
  begin
    aSOPLine := TSOPLine(FList[i]);
    aSOPLine.Free;
  end;
  FList.Clear;
end;

{ TSOPVSActReader }

constructor TSOPVSActReader.Create(const sfile: string);
begin
  FFile := sfile;

  FProjs := TStringList.Create;

  Open;
end;

destructor TSOPVSActReader.Destroy;
begin
  Clear;
  FProjs.Free;
end;

procedure TSOPVSActReader.Clear;
var
  i: Integer;
  aSOPVSActProj: TSOPVSActProj;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPVSActProj := TSOPVSActProj( FProjs.Objects[i] );
    aSOPVSActProj.Free;
  end;
  FProjs.Clear;
end;

function TSOPVSActReader.GetProj(const sName: string): TSOPVSActProj;
var
  i: Integer;
  aSOPVSActProj: TSOPVSActProj;
begin
  Result := nil;
  for i := 0 to FProjs.Count - 1 do
  begin
    aSOPVSActProj := TSOPVSActProj(FProjs.Objects[i]);
    if aSOPVSActProj.FName = sName then
    begin
      Result := aSOPVSActProj;
      Break;
    end;
  end;
end;

procedure TSOPVSActReader.Log(const str: string);
begin

end;

procedure TSOPVSActReader.Open;
var
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  aSOPVSActProj: TSOPVSActProj;
  icol: Integer;
  irow: Integer;
  sweek: string;
  sdate: string;
  sver: string; //制式
  snumber: string; //	物料编码
  scolor: string; //	颜色
  scap: string;//	容量
  aSOPLine: TSOPLine;
  iWeek: Integer;
  v: Variant;
  iqty_sop, iqty_mps: Integer;
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

        aSOPVSActProj := TSOPVSActProj.Create(sSheet);
        FProjs.AddObject(sSheet, aSOPVSActProj);

        irow := 1;
        icol := 7;

        sweek := ExcelApp.Cells[irow, icol].Value;
        sdate := ExcelApp.Cells[irow + 1, icol].Value;
        while sweek <> '' do
        begin
          aSOPVSActProj.slWeeks.AddObject(sweek + '=' + sdate, TObject(icol));
          icol := icol + 1;
          sweek := ExcelApp.Cells[irow, icol].Value;
          sdate := ExcelApp.Cells[irow + 1, icol].Value;
        end;

        irow := 3;                                 
        sdate := ExcelApp.Cells[irow, 1].Value;
        sver := ExcelApp.Cells[irow, 2].Value;
        snumber := ExcelApp.Cells[irow, 3].Value;
        scolor := ExcelApp.Cells[irow, 4].Value;
        scap := ExcelApp.Cells[irow, 5].Value;
        while sver <> '' do
        begin
          aSOPLine := TSOPLine.Create;
          aSOPLine.sVer := sver;
          aSOPLine.sNumber := snumber;
          aSOPLine.sColor := scolor;
          aSOPLine.sCap := scap;
          aSOPVSActProj.FList.Add(aSOPLine);

          for iWeek := 0 to aSOPVSActProj.slWeeks.Count - 1 do
          begin
            icol := Integer(aSOPVSActProj.slWeeks.Objects[iWeek]);
            iqty_sop := 0;
            iqty_mps := 0;
            
            v := ExcelApp.Cells[irow, icol].Value;
            if VarIsNumeric(v) then
            begin
              iqty_sop := v;
            end;

            v := ExcelApp.Cells[irow + 1, icol].Value;
            if VarIsNumeric(v) then
            begin
              iqty_mps := v;
            end;

            aSOPLine.Add('', aSOPVSActProj.slWeeks.Names[iWeek],
              aSOPVSActProj.slWeeks.ValueFromIndex[iWeek], iqty_sop, 0, 0,
              iqty_mps);
          end;

          irow := irow + 3;
          sdate := ExcelApp.Cells[irow, 1].Value;
          sver := ExcelApp.Cells[irow, 2].Value;
          snumber := ExcelApp.Cells[irow, 3].Value;
          scolor := ExcelApp.Cells[irow, 4].Value;
          scap := ExcelApp.Cells[irow, 5].Value;
        end;

      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;

  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit; 
  end
end;

end.
