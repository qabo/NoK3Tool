unit DailyPlanVsActReader;

interface

uses
  SysUtils, Classes, ComObj, CommUtils, Variants;

type
  TDatePlanAct = packed record
    dt: TDateTime;
    dQty: Double;
    dQtyAct: Double;
    scomment: string;
  end;
  PDatePlanAct= ^TDatePlanAct;
  
  TDPVALine = class
  public
    sver: string;
    snumber: string;
    scolor: string;
    scap: string;
    FList: TList;
    constructor Create;
    destructor Destroy; override;
    procedure Clear;
    procedure GetQty(dt1, dt2: TDateTime; var dSOPQtyPlan, dSOPQtyAct: Double);
  end;

  TDailyPlanVsAcsSheet = class
  public
    FName: string;
    FList: TList;
    constructor Create(const sname: string);
    destructor Destroy; override;     
    procedure Clear;
    function GetLine(const sNumber: string): TDPVALine;
  end;

  TDailyPlanVsActReader = class    
  private
    FFile: string;
    ExcelApp, WorkBook: Variant;
    FList: TList;
    procedure Open;
    procedure Log(const str: string);
    procedure Clear;
    function GetCount: Integer;
    function GetItems(i: Integer): TDailyPlanVsAcsSheet;
    function GetSheetByName(const sName: string): TDailyPlanVsAcsSheet;
  public
    constructor Create(const sfile: string);
    destructor Destroy; override;
    property Count: Integer read GetCount;
    property Items[i: Integer]: TDailyPlanVsAcsSheet read GetItems;
    property SheetByName[const sName: string]: TDailyPlanVsAcsSheet read GetSheetByName;
  end;

implementation

{ TDPVALine }

constructor TDPVALine.Create;
begin
  FList := TList.Create;
end;

destructor TDPVALine.Destroy;
begin
  Clear;
  FList.Clear;
  inherited;
end;

procedure TDPVALine.Clear;
var
  i: Integer;
  aDatePlanActPtr: PDatePlanAct;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aDatePlanActPtr := PDatePlanAct(FList[i]);
    Dispose(aDatePlanActPtr);
  end;
  FList.Clear;
end;

procedure TDPVALine.GetQty(dt1, dt2: TDateTime; var dSOPQtyPlan, dSOPQtyAct: Double);
var
  i: Integer;
  aDatePlanActPtr: PDatePlanAct;
begin
  dSOPQtyPlan := 0;
  dSOPQtyAct := 0;
  for i := 0 to FList.Count - 1 do
  begin
    aDatePlanActPtr := PDatePlanAct(FList[i]);
    if aDatePlanActPtr^.dt < dt1 then Continue;
    if aDatePlanActPtr^.dt > dt2 then Break;
    dSOPQtyPlan := dSOPQtyPlan + aDatePlanActPtr^.dQtyAct;     
    dSOPQtyAct := dSOPQtyAct + aDatePlanActPtr^.dQtyAct;
  end;
end;


{ TDailyPlanVsAcsSheet }

constructor TDailyPlanVsAcsSheet.Create(const sname: string);
begin
  FName := sname;
  FList := TList.Create;
end;
destructor TDailyPlanVsAcsSheet.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;
procedure TDailyPlanVsAcsSheet.Clear;
var
  i: Integer;
  aDPVALine: TDPVALine;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aDPVALine := TDPVALine(FList[i]);
    aDPVALine.Free;
  end;
  FList.Clear;
end;

function TDailyPlanVsAcsSheet.GetLine(const sNumber: string): TDPVALine;
var
  i: Integer;
  aDPVALine: TDPVALine;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aDPVALine := TDPVALine(FList[i]);
    if aDPVALine.snumber = sNumber then
    begin
      Result := aDPVALine;
      Break;
    end;
  end;
end;

{ TDailyPlanVsActReader }

procedure TDailyPlanVsActReader.Open;
var
  iSheetCount: Integer;
  iSheet: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5: string;
  stitle: string;
  sproj: string;
  irow: Integer;
  icol: Integer;
  v: Variant;
  dt: TDateTime;
  sldates: TStringList; // 字符串是日期， object 是 列号， 整数
  sitem1, sitem2: string;
  sitem: string;
  aDPVALine: TDPVALine;
  iDate: Integer;
  aDatePlanActPtr: PDatePlanAct;
  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
  vComment: Variant;
begin
  Clear;

  sldates := TStringList.Create;
  try
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

          if UpperCase(Copy(sSheet, 1, 3)) <> 'CTB' then
          begin
            Log(sSheet + ' is not CTB');
            Continue;
          end;

//          sproj := Copy(sSheet, Pos('CTB-', UpperCase(sSheet)) + 4, Length(sSheet));
//          if Pos('-', sproj) > 0 then
//          begin
//            sproj := Copy(sproj, 1, Pos('-', sproj) - 1);
//          end;

          sproj := ExtractFileName(FFile);
          if Pos('-', sproj) > 0 then
          begin
            sproj := Copy(sproj, 1, Pos('-', sproj) - 1);
          end;

          irow := 2;
          stitle1 := ExcelApp.Cells[irow, 1].Value;
          stitle2 := ExcelApp.Cells[irow, 2].Value;
          stitle3 := ExcelApp.Cells[irow, 3].Value;
          stitle4 := ExcelApp.Cells[irow, 4].Value;
          stitle5 := ExcelApp.Cells[irow, 5].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5;
          if stitle <> '机器型号物料编码颜色容量项目' then
          begin
            Log(sSheet + ' is not 日生产计划跟进');
            Continue;
          end;

          aDailyPlanVsAcsSheet := TDailyPlanVsAcsSheet.Create(sproj);
          FList.Add(aDailyPlanVsAcsSheet);

          sldates.Clear;
				              
          irow := 3;
          icol := 7;
          v := ExcelApp.Cells[irow, icol].Value;
          while VarIsType(v, varDate) do
          begin
            dt := v;
            sldates.AddObject(FormatDateTime('yyyy-MM-dd', dt), TObject(icol));
            icol := icol + 1;
            v := ExcelApp.Cells[irow, icol].Value;
          end;

          Log('日期列表');
          Log(sldates.Text);

          irow := 4;
          while not IsCellMerged(ExcelApp, irow, 2, irow, 3) do
          begin
            sitem1 := ExcelApp.Cells[irow, 5].Value;
            sitem2 := ExcelApp.Cells[irow + 1, 5].Value;
            sitem := sitem1 + sitem2;
            if sitem = '计划实际' then
            begin
              aDPVALine := TDPVALine.Create;
              aDailyPlanVsAcsSheet.FList.Add(aDPVALine);

              aDPVALine.sver := ExcelApp.Cells[irow, 1].Value;
              aDPVALine.snumber := ExcelApp.Cells[irow, 2].Value;
              aDPVALine.scolor := ExcelApp.Cells[irow, 3].Value;
              aDPVALine.scap := ExcelApp.Cells[irow, 4].Value;

              for iDate := 0 to sldates.Count - 1 do
              begin
                icol := Integer(sldates.Objects[iDate]);
                aDatePlanActPtr := New(PDatePlanAct);
                aDPVALine.FList.Add(aDatePlanActPtr);

                aDatePlanActPtr^.dt := myStrToDateTime(sldates[iDate]);
                aDatePlanActPtr^.dQty := ExcelApp.Cells[irow, icol].Value;
                aDatePlanActPtr^.dQtyAct := ExcelApp.Cells[irow + 1, icol].Value;

                vComment := ExcelApp.Cells[irow + 1, icol].Comment;

                if FindVarData(vComment)^.VDispatch <> nil then
                begin
                  aDatePlanActPtr^.scomment := vComment.Text;
                end
                else
                begin
                  aDatePlanActPtr^.scomment := '';
                end;

              end;
            end;
            
            irow := irow + 2;
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
  finally
    sldates.Free;
  end;
end;

procedure TDailyPlanVsActReader.Log(const str: string);
begin

end;

constructor TDailyPlanVsActReader.Create(const sfile: string);
begin
  FFile := sfile;
  FList := TList.Create;
  Open;
end;

destructor TDailyPlanVsActReader.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TDailyPlanVsActReader.Clear;
var
  i: Integer;
  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aDailyPlanVsAcsSheet := TDailyPlanVsAcsSheet(FList[i]);
    aDailyPlanVsAcsSheet.Free;
  end;
  FList.Clear;
end;

function TDailyPlanVsActReader.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TDailyPlanVsActReader.GetItems(i: Integer): TDailyPlanVsAcsSheet; 
begin
  Result := TDailyPlanVsAcsSheet(FList[i]);
end;

function TDailyPlanVsActReader.GetSheetByName(const sName: string): TDailyPlanVsAcsSheet;
var
  i: Integer;
  aDailyPlanVsAcsSheet: TDailyPlanVsAcsSheet;
begin
  Result := nil;
  for i := 0 to FList.Count - 1 do
  begin
    aDailyPlanVsAcsSheet := TDailyPlanVsAcsSheet(FList[i]);
    if aDailyPlanVsAcsSheet.FName = sName then
    begin
      Result := aDailyPlanVsAcsSheet;
      Break;
    end;
  end;
end;

end.
