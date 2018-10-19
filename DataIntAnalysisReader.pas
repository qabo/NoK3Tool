unit DataIntAnalysisReader;

interface

uses
  Classes, ComObj, CommUtils;

type
  TDataIntAnalysisCol = packed record
    icol: Integer;
    sweek: string;
    qty1: Double;
    qty2: Double;
    fcalc: Boolean;
  end;
  PDataIntAnalysisCol = ^TDataIntAnalysisCol;

  TDataIntAnalysisLine = class
  private
    procedure Clear;
  public
    smode: string;
    sproj: string; //项目
    sweek: string; //	week
    snumber: string; //	物料编码
    scolor: string; //	颜色
    scap: string; //	容量
    sver: string; //	制式
    splan: string; //	计划
    irow: Integer;
    slweeks: TStringList;
    fcalc: Boolean;
    constructor Create;
    destructor Destroy; override;
    function FindWeek(const sweek: string): PDataIntAnalysisCol;
  end;

  TDataIntAnalysisReader = class
  private
    FFile: string;
    FWeeks: TStringList;
    procedure Open;
    procedure Clear;
  public          
    FListSOPvsDemand: TList;   
    FListACTvsDemand: TList;
    FListACTvsSch: TList;
    constructor Create(const sfile: string; slWeeks: TStrings);
    destructor Destroy; override;
    function FindLine(lst: TList;
      aDataIntAnalysisLine1: TDataIntAnalysisLine): TDataIntAnalysisLine;
  end;

implementation

{ TDataIntAnalysisLine }



constructor TDataIntAnalysisLine.Create;
begin
  fcalc := False;
  slweeks := TStringList.Create;
end;

destructor TDataIntAnalysisLine.Destroy;
begin
  Clear;
  slweeks.Free;
  inherited;
end;

procedure TDataIntAnalysisLine.Clear;
var
  i: Integer;
  p: PDataIntAnalysisCol;
begin
  for i := 0 to slweeks.Count - 1 do
  begin
    p := PDataIntAnalysisCol(slWeeks.Objects[i]);
    Dispose(p);
  end;
  slweeks.Clear;
end;

function TDataIntAnalysisLine.FindWeek(const sweek: string): PDataIntAnalysisCol;
var
  i: Integer;
  p: PDataIntAnalysisCol;
begin
  Result := nil;
  for i := 0 to slweeks.Count - 1 do
  begin
    p := PDataIntAnalysisCol(slWeeks.Objects[i]);
    if p^.sweek = sweek then
    begin
      Result := p;
      Break;
    end;
  end;
end;

{ TDataIntAnalysisReader }

constructor TDataIntAnalysisReader.Create(const sfile: string;
  slWeeks: TStrings);
begin
  FFile := sfile;
  FWeeks := TStringList.Create;
  FWeeks.Text := slWeeks.Text;
  FListSOPvsDemand := TList.Create;
  FListACTvsDemand := TList.Create;
  FListACTvsSch := TList.Create;
  Open;
end;

destructor TDataIntAnalysisReader.Destroy;
begin
  Clear;
  FWeeks.Free;
  FListSOPvsDemand.Free;
  FListACTvsDemand.Free;
  FListACTvsSch.Free;
  inherited;
end;

procedure TDataIntAnalysisReader.Clear;
var
  i: Integer;
  aDataIntAnalysisLine: TDataIntAnalysisLine;
begin
  for i := 0 to FListSOPvsDemand.Count - 1 do
  begin
    aDataIntAnalysisLine := TDataIntAnalysisLine(FListSOPvsDemand[i]);
    aDataIntAnalysisLine.Free;
  end;
  FListSOPvsDemand.Clear;

  for i := 0 to FListACTvsDemand.Count - 1 do
  begin
    aDataIntAnalysisLine := TDataIntAnalysisLine(FListACTvsDemand[i]);
    aDataIntAnalysisLine.Free;
  end;
  FListACTvsDemand.Clear;

  for i := 0 to FListACTvsSch.Count - 1 do
  begin
    aDataIntAnalysisLine := TDataIntAnalysisLine(FListACTvsSch[i]);
    aDataIntAnalysisLine.Free;
  end;
  FListACTvsSch.Clear;

  FWeeks.Clear;
end;

function TDataIntAnalysisReader.FindLine(lst: TList;
  aDataIntAnalysisLine1: TDataIntAnalysisLine): TDataIntAnalysisLine;
var
  i: Integer;
  aline: TDataIntAnalysisLine;
begin
  Result := nil;
  for i := 0 to lst.Count - 1 do
  begin
    aline := TDataIntAnalysisLine(lst[i]);
    if (aline.smode = aDataIntAnalysisLine1.smode)
      and (aline.sproj = aDataIntAnalysisLine1.sproj) 
      and (aline.snumber = aDataIntAnalysisLine1.snumber)
      and (aline.scolor = aDataIntAnalysisLine1.scolor) 
      and (aline.scap = aDataIntAnalysisLine1.scap) 
      and (aline.sver = aDataIntAnalysisLine1.sver) then
    begin
      Result := aline;
      Break;
    end;
  end;
end;

procedure TDataIntAnalysisReader.Open;
var
  ExcelApp, WorkBook: Variant;
  iSheet, iSheetCount: Integer;
  sSheet: string;
  slWeeks: TStringList;
  iweek: Integer;
  icol: Integer;
  sweek: string;
  snumber: string;
  scolor: string;
  irow: Integer;
  aDataIntAnalysisLine: TDataIntAnalysisLine;
  aDataIntAnalysisColPtr: PDataIntAnalysisCol;
begin

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';

  ExcelApp.ScreenUpdating := False; 

  try
    WorkBook := ExcelApp.WorkBooks.Open(FFile);

    try
      iSheetCount := ExcelApp.Sheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        if not ExcelApp.Sheets[iSheet].Visible then Continue;

        ExcelApp.Sheets[iSheet].Activate;

        sSheet := ExcelApp.Sheets[iSheet].Name;

        if sSheet = 'KPI分析-S&OP供应计划 VS 销售计划' then
        begin
          slWeeks := TStringList.Create;
          try
            icol := 9;
            sweek := ExcelApp.Cells[1, icol].Value;
            while sweek <> '' do
            begin
              slWeeks.AddObject(sweek, TObject(icol));
              icol := icol + 1;       
              sweek := ExcelApp.Cells[1, icol].Value;
            end;

            if FWeeks.Count > 0 then
            begin
              for iweek := slWeeks.Count - 1 downto 0 do
              begin
                if FWeeks.IndexOf(slWeeks[iweek]) < 0 then
                begin
                  slWeeks.Delete(iweek);
                end;
              end;
            end;

            irow := 2;
            snumber := ExcelApp.Cells[irow, 4].Value;
            scolor := ExcelApp.Cells[irow, 5].Value;
            while (snumber <> '') or (scolor <> '') do
            begin
              aDataIntAnalysisLine := TDataIntAnalysisLine.Create;
              aDataIntAnalysisLine.irow := irow;
              aDataIntAnalysisLine.smode := ExcelApp.Cells[irow, 1].Value;
              aDataIntAnalysisLine.sproj := ExcelApp.Cells[irow, 2].Value;   
              aDataIntAnalysisLine.sweek := ExcelApp.Cells[irow, 3].Value;
              aDataIntAnalysisLine.snumber := snumber;
              aDataIntAnalysisLine.scolor := scolor;
              aDataIntAnalysisLine.scap := ExcelApp.Cells[irow, 6].Value;
              aDataIntAnalysisLine.sver := ExcelApp.Cells[irow, 7].Value;
              aDataIntAnalysisLine.splan := ExcelApp.Cells[irow, 8].Value;

              for iweek := 0 to slWeeks.Count - 1 do
              begin
                aDataIntAnalysisColPtr := New(PDataIntAnalysisCol);
                aDataIntAnalysisColPtr^.fcalc := False; 
                aDataIntAnalysisColPtr^.icol := Integer(slWeeks.Objects[iweek]);  
                aDataIntAnalysisColPtr^.sweek := slWeeks[iweek];
                aDataIntAnalysisColPtr^.qty1 := ExcelApp.Cells[irow, Integer(slWeeks.Objects[iweek])].Value;
                aDataIntAnalysisColPtr^.qty2 := ExcelApp.Cells[irow + 1, Integer(slWeeks.Objects[iweek])].Value;       
                aDataIntAnalysisLine.slweeks.AddObject(slWeeks[iweek], TObject(aDataIntAnalysisColPtr));
              end;

              FListSOPvsDemand.Add(aDataIntAnalysisLine);

              if aDataIntAnalysisLine.smode = 'OEM' then
              begin
                irow := irow + Length(CSOEMSOPvsDemand_OEM) + 4;
              end
              else
              begin
                irow := irow + Length(CSOEMSOPvsDemand_OEM) + 4;
              end;


              snumber := ExcelApp.Cells[irow, 4].Value;     
              scolor := ExcelApp.Cells[irow, 5].Value;
            end;

          finally
            slWeeks.Free;
          end;
        end;

        if sSheet = 'KPI分析-实际产出 VS S&OP供应计划' then
        begin
          slWeeks := TStringList.Create;
          try
            icol := 9;
            sweek := ExcelApp.Cells[1, icol].Value;
            while sweek <> '' do
            begin
              slWeeks.AddObject(sweek, TObject(icol));
              icol := icol + 1;  
              sweek := ExcelApp.Cells[1, icol].Value;
            end;

            if FWeeks.Count > 0 then
            begin
              for iweek := slWeeks.Count - 1 downto 0 do
              begin
                if FWeeks.IndexOf(slWeeks[iweek]) < 0 then
                begin
                  slWeeks.Delete(iweek);
                end;
              end;
            end;

            
            irow := 2;
            snumber := ExcelApp.Cells[irow, 4].Value;
            scolor := ExcelApp.Cells[irow, 5].Value;
            while (snumber <> '') or (scolor <> '') do
            begin
              aDataIntAnalysisLine := TDataIntAnalysisLine.Create;
              aDataIntAnalysisLine.irow := irow;
              aDataIntAnalysisLine.smode := ExcelApp.Cells[irow, 1].Value;
              aDataIntAnalysisLine.sproj := ExcelApp.Cells[irow, 2].Value;   
              aDataIntAnalysisLine.sweek := ExcelApp.Cells[irow, 3].Value;
              aDataIntAnalysisLine.snumber := snumber;
              aDataIntAnalysisLine.scolor := scolor;
              aDataIntAnalysisLine.scap := ExcelApp.Cells[irow, 6].Value;
              aDataIntAnalysisLine.sver := ExcelApp.Cells[irow, 7].Value;
              aDataIntAnalysisLine.splan := ExcelApp.Cells[irow, 8].Value;

              for iweek := 0 to slWeeks.Count - 1 do
              begin
                aDataIntAnalysisColPtr := New(PDataIntAnalysisCol);
                aDataIntAnalysisColPtr^.fcalc := False;
                aDataIntAnalysisColPtr^.icol := Integer(slWeeks.Objects[iweek]);     
                aDataIntAnalysisColPtr^.sweek := slWeeks[iweek];
                aDataIntAnalysisColPtr^.qty1 := ExcelApp.Cells[irow, Integer(slWeeks.Objects[iweek])].Value;
                aDataIntAnalysisColPtr^.qty2 := ExcelApp.Cells[irow + 1, Integer(slWeeks.Objects[iweek])].Value;       
                aDataIntAnalysisLine.slweeks.AddObject(slWeeks[iweek], TObject(aDataIntAnalysisColPtr));
              end;

              FListACTvsDemand.Add(aDataIntAnalysisLine);

              if aDataIntAnalysisLine.smode = 'OEM' then
              begin
                irow := irow + Length(CSOEMACTvsDemand_OEM) + 3;
              end
              else
              begin
                irow := irow + Length(CSOEMACTvsDemand_ODM) + 3;
              end;
              snumber := ExcelApp.Cells[irow, 4].Value;   
              scolor := ExcelApp.Cells[irow, 5].Value;
            end;

          finally
            slWeeks.Free;
          end;
        end;

        if sSheet = 'KPI分析-实际产出 VS 排产计划' then
        begin
          slWeeks := TStringList.Create;
          try
            icol := 9;
            sweek := ExcelApp.Cells[1, icol].Value;
            while sweek <> '' do
            begin
              slWeeks.AddObject(sweek, TObject(icol));
              icol := icol + 1;    
              sweek := ExcelApp.Cells[1, icol].Value;
            end;

            if FWeeks.Count > 0 then
            begin
              for iweek := slWeeks.Count - 1 downto 0 do
              begin
                if FWeeks.IndexOf(slWeeks[iweek]) < 0 then
                begin
                  slWeeks.Delete(iweek);
                end;
              end;
            end;  

            
            irow := 2;
            snumber := ExcelApp.Cells[irow, 4].Value;
            scolor := ExcelApp.Cells[irow, 5].Value;
            while (snumber <> '') or (scolor <> '') do
            begin
              aDataIntAnalysisLine := TDataIntAnalysisLine.Create;
              aDataIntAnalysisLine.irow := irow;
              aDataIntAnalysisLine.smode := ExcelApp.Cells[irow, 1].Value;
              aDataIntAnalysisLine.sproj := ExcelApp.Cells[irow, 2].Value;   
              aDataIntAnalysisLine.sweek := ExcelApp.Cells[irow, 3].Value;
              aDataIntAnalysisLine.snumber := snumber;
              aDataIntAnalysisLine.scolor := scolor;
              aDataIntAnalysisLine.scap := ExcelApp.Cells[irow, 6].Value;
              aDataIntAnalysisLine.sver := ExcelApp.Cells[irow, 7].Value;
              aDataIntAnalysisLine.splan := ExcelApp.Cells[irow, 8].Value;

              for iweek := 0 to slWeeks.Count - 1 do
              begin
                aDataIntAnalysisColPtr := New(PDataIntAnalysisCol);
                aDataIntAnalysisColPtr^.fcalc := False; 
                aDataIntAnalysisColPtr^.icol := Integer(slWeeks.Objects[iweek]);    
                aDataIntAnalysisColPtr^.sweek := slWeeks[iweek];
                aDataIntAnalysisColPtr^.qty1 := ExcelApp.Cells[irow, Integer(slWeeks.Objects[iweek])].Value;
                aDataIntAnalysisColPtr^.qty2 := ExcelApp.Cells[irow + 1, Integer(slWeeks.Objects[iweek])].Value;       
                aDataIntAnalysisLine.slweeks.AddObject(slWeeks[iweek], TObject(aDataIntAnalysisColPtr));
              end;

              FListACTvsSch.Add(aDataIntAnalysisLine);

              if aDataIntAnalysisLine.smode = 'OEM' then
              begin
                irow := irow + Length(CSOEMACTvsSch_OEM) + 3;
              end
              else
              begin
                irow := irow + Length(CSOEMACTvsSch_ODM) + 3;
              end;
              snumber := ExcelApp.Cells[irow, 4].Value;   
              scolor := ExcelApp.Cells[irow, 5].Value;
            end;

            
          finally
            slWeeks.Free;
          end;
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
