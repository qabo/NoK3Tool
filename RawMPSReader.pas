unit RawMPSReader;

interface

uses
  Classes, ComObj, SysUtils;

type
  TRawMPSCol = packed record
    iQty: Integer;
  end;
  PRawMPSCol = ^TRawMPSCol;

  TRawMPSLine = Class
  public
    snumber: string; //产品编码
    sarea: string; 
    sver: string; //版本
    scol: string; //颜色
    scap: string; //容量
    sproj: string; //项目
    sbom: string;
    FList: TList;     
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
  end;

  TRawBom = packed record
    sbom: string;
    sver: string;
    scap: string;
    scol: string;
  end;
  PRawBom = ^TRawBom;
  
  TRawMPSReader = class
  private             
    FFile: string;
    ExcelApp, WorkBook: Variant;
    procedure Open;
    procedure Log(const str: string); 
  public
    FList: TStringList;   
    slWeek: TStringList;
    slBomNumber: TStringList;
    constructor Create(const sfile: string);
    destructor Destroy; override;
    procedure Clear;
  end;

implementation

{ TRawMPSLine }

constructor TRawMPSLine.Create;
begin  
  FList := TList.Create;
end;

destructor TRawMPSLine.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TRawMPSLine.Clear;
var
  i: Integer;
  p: PRawMPSCol;
begin
  for i := 0 to FList.Count - 1 do
  begin
    p := PRawMPSCol(FList[i]);
    Dispose(p);
  end;
  FList.Clear;
end;  
  
{ TRawMPSReader }

constructor TRawMPSReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TStringList.Create;
  slWeek := TStringList.Create;
  slBomNumber := TStringList.Create;
  Open;
end;

destructor TRawMPSReader.Destroy;
begin
  Clear;
  FList.Free;
  slWeek.Free;
  slBomNumber.Free;
  inherited;
end;

procedure TRawMPSReader.Clear;
var
  i: Integer;  
  aRawMPSLine: TRawMPSLine;
  aRawBomPtr: PRawBom;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aRawMPSLine := TRawMPSLine(FList.Objects[i]);
    aRawMPSLine.Free;
  end;
  FList.Clear;
  slWeek.Clear;

  for i := 0 to slBomNumber.Count - 1 do
  begin
    aRawBomPtr := PRawBom(slBomNumber.Objects[i]);
    Dispose(aRawBomPtr);
  end;
  slBomNumber.Clear;
end;

procedure TRawMPSReader.Log(const str: string);
begin

end;

procedure TRawMPSReader.Open;
var
  iSheet: Integer;
  iSheetCount: Integer;
  sSheet: string;
  stitle: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  irow: Integer;
  icol: Integer;
  icol_ver: Integer;
  aRawMPSLine: TRawMPSLine;
  sver: string;
  bColEnd: Boolean;
  iweek: Integer;
  aRawMPSColPtr: PRawMPSCol;
  aRawBomPtr: PRawBom;
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
        stitle := stitle1 + stitle2;
        if stitle <> 'MATNRBERID' then
        begin
          Log(sSheet + ' 不是 原始MPS格式');
          Continue;
        end;

        icol_ver := -1;
        bColEnd := False;
        
        for icol := 3 to 200 do
        begin                                          
          stitle1 := ExcelApp.Cells[irow, icol].Value;
          stitle2 := ExcelApp.Cells[irow, icol + 1].Value;
          stitle3 := ExcelApp.Cells[irow, icol + 2].Value;
          stitle4 := ExcelApp.Cells[irow, icol + 3].Value;
          stitle5 := ExcelApp.Cells[irow, icol + 4].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
          if stitle = '产品编码版本颜色容量项目' then
          begin
            icol_ver := icol + 1;
            Break;
          end;

          if stitle1 = '' then
          begin
            bColEnd := True;
          end;

          if not bColEnd then
          begin
            slWeek.AddObject(stitle1, TObject(icol));
          end;

        end;

        if icol_ver = -1 then
        begin
          Log(sSheet + ' 不是 原始MPS格式 找不到制式列');
          Continue;
        end;

        irow := irow + 1;
        sver := ExcelApp.Cells[irow, icol_ver].Value;
        while sver <> '' do
        begin
          aRawMPSLine := TRawMPSLine.Create;
          FList.AddObject(IntToStr(irow), aRawMPSLine);

          aRawMPSLine.snumber := ExcelApp.Cells[irow, 1].Value;
          aRawMPSLine.sarea := ExcelApp.Cells[irow, 2].Value;
          aRawMPSLine.sver := sver;
          aRawMPSLine.scol := ExcelApp.Cells[irow, icol_ver + 1].Value;
          aRawMPSLine.scap := ExcelApp.Cells[irow, icol_ver + 2].Value;
          aRawMPSLine.sproj := ExcelApp.Cells[irow, icol_ver + 3].Value;

          for iweek := 0 to slWeek.Count - 1 do
          begin
            icol := Integer(slWeek.Objects[iweek]);
            aRawMPSColPtr := New(PRawMPSCol);
            aRawMPSColPtr^.iQty := ExcelApp.Cells[irow, icol].Value;
            aRawMPSLine.FList.Add(aRawMPSColPtr);
          end;

          aRawMPSLine.sbom := aRawMPSLine.sver + aRawMPSLine.scol + aRawMPSLine.scap + aRawMPSLine.sproj;
          if slBomNumber.IndexOf(aRawMPSLine.sbom) < 0 then
          begin
            aRawBomPtr := New(PRawBom);
            aRawBomPtr^.sbom := aRawMPSLine.sbom;
            aRawBomPtr^.sver := aRawMPSLine.sver;
            aRawBomPtr^.scap := aRawMPSLine.scap;
            aRawBomPtr^.scol := aRawMPSLine.scol;
            slBomNumber.AddObject(aRawMPSLine.sbom, TObject(aRawBomPtr));
          end;  

          irow := irow + 1;
          sver := ExcelApp.Cells[irow, icol_ver].Value;
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
