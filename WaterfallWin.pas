unit WaterfallWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ImgList, ToolWin, StdCtrls, ExtCtrls, ComObj, DateUtils,
  IniFiles, CommUtils;

const
  xlBetween = 1;
  xlEqual = 3;
  xlGreater = 5;
  xlGreaterEqual = 7;
  xlLess = 6;
  xlLessEqual = 8;
  xlNotBetween = 2;
  xlNotEqual = 4;
  xlCellValue = 1;
  xlExpression = 2;
  xlButtonControl = 0;
  xlCheckBox = 1;
  xlDropDown = 2;
  xlEditBox = 3;
  xlGroupBox = 4;
  xlLabel = 5;
  xlListBox = 6;
  xlOptionButton = 7;
  xlScrollBar = 8;
  xlSpinner = 9;
  xlColumnLabels = 2;
  xlMixedLabels = 3;
  xlNoLabels = -4142;
  xlRowLabels = 1;
  xlHAlignCenter = -4108;
  xlHAlignCenterAcrossSelection = 7;
  xlHAlignDistributed = -4117;
  xlHAlignFill = 5;
  xlHAlignGeneral = 1;
  xlHAlignJustify = -4130;
  xlHAlignLeft = -4131;
  xlHAlignRight = -4152;
  xlHebrewFullScript = 0;
  xlHebrewMixedAuthorizedScript = 3;
  xlHebrewMixedScript = 2;
  xlHebrewPartialScript = 1;
  xlAllChanges = 2;
  xlNotYetReviewed = 3;
  xlSinceMyLastSave = 1;
  xlHtmlCalc = 1;
  xlHtmlChart = 3;
  xlHtmlList = 2;
  xlHtmlStatic = 0;
  xlIMEModeAlpha = 8;
  xlIMEModeAlphaFull = 7;
  xlIMEModeDisable = 3;
  xlIMEModeHangul = 10;
  xlIMEModeHangulFull = 9;

type
  TReadExcelProc = procedure (ExcelApp, WorkBook: Variant) of object;
  TWriteExcelProc = function (ExcelApp, WorkBook: Variant): Boolean of object;

  TDateQty = packed record
    dt: TDateTime;
    dQty: Double;
  end;
  PDateQty = ^TDateQty;

  TLineData = class
  private
    FName: string;
    FList: TList;
    FRow: Integer;
    FMonthList: TList;
    procedure Clear;
    function GetDateQtyCount: Integer;
    function GetDateQtys(i: Integer): PDateQty;
    function GetMonthQtyCount: Integer;
    function GetMonthQtys(i: Integer): PDateQty;
  public
    constructor Create(const sname: string);
    destructor Destroy; override;
    procedure Add(dt: TDateTime; dQty: Double);
    procedure AddMonth(dt: TDateTime; dQty: Double);
    procedure Insert(idx: Integer; dt: TDateTime; dQty: Double);
    function IndexOfDate(dt: TDateTime): Integer;
    property DateQtyCount: Integer read GetDateQtyCount;
    property DateQtys[i: Integer]: PDateQty read GetDateQtys;
    property MonthQtyCount: Integer read GetMonthQtyCount;
    property MonthQtys[i: Integer]: PDateQty read GetMonthQtys;
  end;
  
  TWeekData = class
  private
    FName: string;
    FColors: TList;
    FVers: TList;
    FCaps: TList;
    FFGs: TList;
    FPkgs: TList;
    FRow: Integer;
    procedure Clear;
    function GetColorCount: Integer;
    function GetVerCount: Integer;
    function GetCapCount: Integer;
    function GetFGCount: Integer;
    function GetPkgCount: Integer;
    function GetColors(i: Integer): TLineData;
    function GetVers(i: Integer): TLineData;
    function GetCaps(i: Integer): TLineData;
    function GetFGs(i: Integer): TLineData;
    function GetPkgs(i: Integer): TLineData;
  public
    constructor Create(const sWeek: string);
    destructor Destroy; override;

//    function GetColorLineData(const sKey: string): TLineData;
//    function GetVerLineData(const sKey: string): TLineData;
//    function GetCapLineData(const sKey: string): TLineData;
//    function GetFGLineData(const sKey: string): TLineData;
//    function GetPkgLineData(const sKey: string): TLineData;   
//    function GetLineData(lst: TList; const skey: string): TLineData;
    
    property ColorCount: Integer read GetColorCount;
    property VerCount: Integer read GetVerCount;
    property CapCount: Integer read GetCapCount;
    property FGCount: Integer read GetFGCount;
    property PkgCount: Integer read GetPkgCount;
    property Colors[i: Integer]: TLineData read GetColors;
    property Vers[i: Integer]: TLineData read GetVers;
    property Caps[i: Integer]: TLineData read GetCaps;
    property FGs[i: Integer]: TLineData read GetFGs;
    property Pkgs[i: Integer]: TLineData read GetPkgs;
  end;
  
  TProjData = class
  private
    FName: string;
    FList: TList;
    FCurrWeekData: TWeekData;
    procedure Clear;
    function GetWeekCount: Integer;
    function GetWeekds(i: Integer): TWeekData;
  public
    constructor Create(const sName: string);
    destructor Destroy; override;
    function AddWeek(const sWeek: string): Integer;
    procedure SortDate;
    procedure SumByMonth;
    procedure InsertDateQty(aWeekData: TWeekData; dt: TDateTime);
    property WeekCount: Integer read GetWeekCount;
    property Weeks[i: Integer]: TWeekData read GetWeekds;
  end;
    
  TLineDatasGetter = function(aWeekData: TWeekData): TList;

  TfrmWaterfall = class(TForm)
    ToolBar1: TToolBar;
    ImageList1: TImageList;
    tbSave: TToolButton;
    leWaterFall: TLabeledEdit;
    leSOPSum: TLabeledEdit;
    btnWaterFall: TButton;
    btnSOPSum: TButton;
    Memo1: TMemo;
    procedure btnWaterFallClick(Sender: TObject);
    procedure btnSOPSumClick(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FProjs: TList;
    procedure ReadExcelFile(const sfile: string; aProc: TReadExcelProc);
    procedure WriteExcelFile(const sfile: string; aProc: TWriteExcelProc);
    procedure Clear;
    procedure OpenWF(ExcelApp, WorkBook: Variant);
    procedure OpenSOP(ExcelApp, WorkBook: Variant);
    procedure OpenWFDetail(ExcelApp, WorkBook: Variant; aProjData: TProjData); 
    procedure OpenSOPSheet(ExcelApp: Variant; aProjData: TProjData;
      const sWeek: string);
    function FindFirstNeedDateCellDetail(ExcelApp: Variant; var row, col: Integer): Boolean;
    function GetDateColRangeDetail(ExcelApp: Variant; r, c: Integer; var col1, col2: Integer): Boolean;
    procedure GetWeekCountDetail(ExcelApp: Variant; irow, icol: Integer; lstRowOfWeeks: TList);
    function ReadDatesQtyDetail(ExcelApp: Variant; iRow: Integer;
      irow1, irow2: Integer; icol1, icol2: Integer; lst: TList; iKeyCol: Integer): string;
    procedure ReadWeekDetail(ExcelApp: Variant; irow, icol: Integer;
      dtcol1, dtcol2: Integer; aProjData: TProjData);
    function FindProj(const sProj: string): TProjData;
    procedure SortDate;
    procedure SumByMonth;
    function WriteWF(ExcelApp, WorkBook: Variant): Boolean;
    procedure WriteWF11(ExcelApp: Variant; aProj: TProjData; bSOP: Boolean);
    procedure WriteWFKeys(ExcelApp: Variant; aWeekData: TWeekData; const sTitle: string;
      iDateQtyCount: Integer; aLst: TList; aKeys: TStringList; var irow, irow1, irow2: Integer);
    procedure WriteWF01(ExcelApp: Variant; aProj: TProjData);
    procedure WriteWF11Compares(ExcelApp: Variant; aProj: TProjData;
      aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
      var irow: Integer; aLineDatasGetter: TLineDatasGetter; bSOP: Boolean);
    procedure WriteWF11ComparesTTL(ExcelApp: Variant; aProj: TProjData;
      aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
      var irow: Integer; aLineDatasGetter: TLineDatasGetter; bSOP: Boolean);
    procedure WriteWF01Compares(ExcelApp: Variant; aProj: TProjData;
      aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
      var irow: Integer; aLineDatasGetter: TLineDatasGetter);


    procedure WriteWF11Sum(ExcelApp: Variant; aProj: TProjData; bSOP: Boolean);
    procedure WriteWF01Sum(ExcelApp: Variant; aProj: TProjData);
    procedure WriteWFKeysSum(ExcelApp: Variant; aWeekData: TWeekData; const sTitle: string;
      iMonthQtyCount: Integer; aLst: TList; aKeys: TStringList; var irow, irow1, irow2: Integer);
    procedure WriteWF11ComparesSum(ExcelApp: Variant; aProj: TProjData;
      aColors: TStringList; const sKeyType: string; iMonthQtyCount: Integer;
      var irow: Integer; aLineDatasGetter: TLineDatasGetter; wd1, wd2: TWeekData);
  public
    { Public declarations }
    class procedure ShowForm;
  end;


implementation

{$R *.dfm}

var
  gFormatSettings: TFormatSettings;

class procedure TfrmWaterfall.ShowForm;
var
  frmWaterfall: TfrmWaterfall;
begin
  frmWaterfall := TfrmWaterfall.Create(nil);
  try
    frmWaterfall.ShowModal;
  finally
    frmWaterfall.Free;
  end;
end;
                
function GetRef(const X:Integer):string;
var
  token,I,R:Integer;
begin
  Result:='';
  token:=X;
  repeat
    I := token div 26;
    R := token mod 26;
    if R <> 0 then
    begin
      Result:=Char(R + 64) + Result;
    end
    else if I > 0 then
    begin
      Result := 'Z' + Result ;
      Dec(I);
    end;
    token := I;
  until I = 0;
end;

function ExcelOpenDialog(var sfile: string): Boolean;
begin
  with TOpenDialog.Create(nil) do
  try
    Filter := 'Excel Files|*.xls;*.xlsx';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options - [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := FileName;
    end;
  finally
    Free;
  end;
end;
    
function ExcelSaveDialog(var sfile: string): Boolean;
begin
  with TSaveDialog.Create(nil) do
  try
    FileName := sfile;
    Filter := 'Excel Files|*.xls;*.xlsx';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options - [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := FileName;
    end;
  finally
    Free;
  end;
end;
    
procedure GetAllKeys(lst: TList; aColors: TStringList);
var
  iKey: Integer;
  aLineData: TLineData;
begin
  for iKey := 0 to lst.Count - 1 do
  begin
    aLineData := TLineData(lst[iKey]);
    if aColors.IndexOf(aLineData.FName) < 0 then
    begin
      aColors.Add(aLineData.FName);
    end;
  end;
end;

function LineDatasGetter4Color(aWeekData: TWeekData): TList;
begin
  Result := aWeekData.FColors;
end;
         
function LineDatasGetter4Ver(aWeekData: TWeekData): TList;
begin
  Result := aWeekData.FVers;
end;

function LineDatasGetter4Cap(aWeekData: TWeekData): TList;
begin
  Result := aWeekData.FCaps;
end;

function LineDatasGetter4FG(aWeekData: TWeekData): TList;
begin
  Result := aWeekData.FFGs;
end;

function LineDatasGetter4Pkg(aWeekData: TWeekData): TList;
begin
  Result := aWeekData.FPkgs;
end;

procedure TfrmWaterfall.ReadExcelFile(const sfile: string; aProc: TReadExcelProc);
var
  ExcelApp, WorkBook: Variant;
begin
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '错误', 0);
      Exit;
    end;
  end;

  try
    try
      WorkBook := ExcelApp.WorkBooks.Open(sFile);
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '错误', 0);
        Exit;
      end;
    end;

    try
      aProc(ExcelApp, WorkBook);
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;
  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;
  end;
end;
        
procedure TfrmWaterfall.WriteExcelFile(const sfile: string; aProc: TWriteExcelProc);
var
  ExcelApp, WorkBook: Variant;
begin
  try
    ExcelApp := CreateOleObject('Excel.Application' );
//    ExcelApp.Visible := False;                           
    ExcelApp.Visible := True;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(0, PChar(e.Message), '错误', 0);
      Exit;
    end;
  end;

  try
    try
      WorkBook := ExcelApp.WorkBooks.Add;
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.Message), '错误', 0);
        Exit;
      end;
    end;

    try
      if aProc(ExcelApp, WorkBook) then
      begin
        WorkBook.SaveAs(sfile);
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
    
function CellMerged(ExcelApp: Variant; irow1, irow2: Integer; icol: Integer): Boolean;
var
  vma1, vma2: Variant;
begin
  vma1 := ExcelApp.Cells[irow1, icol].MergeArea;
  vma2 := ExcelApp.Cells[irow2, icol].MergeArea;
  Result := vma1.Address = vma2.Address;
end;

procedure GetRowMergedRange(ExcelApp: Variant; irow, icol: Integer;
  var irow1, irow2: Integer);
begin
  irow1 := irow;
  irow2 := irow1 + 1;
  while CellMerged(ExcelApp, irow1, irow2, icol) do
  begin
    irow2 := irow2 + 1;
  end;
  irow2 := irow2 - 1;
end;

////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

{ TLineData }

constructor TLineData.Create(const sname: string);
begin
  FName := sname;
  FList := TList.Create;
  FMonthList := TList.Create;
end;

destructor TLineData.Destroy;
begin
  Clear;
  FList.Free;
  FMonthList.Free;
  inherited;
end;

procedure TLineData.Add(dt: TDateTime; dQty: Double);
var
  p: PDateQty;
begin
  p := New(PDateQty);
  p^.dt := dt;
  p^.dQty := dQty;
  FList.Add(p);
end;

procedure TLineData.AddMonth(dt: TDateTime; dQty: Double);
var
  p: PDateQty;
begin
  p := New(PDateQty);
  p^.dt := dt;
  p^.dQty := dQty;
  FMonthList.Add(p);
end;

procedure TLineData.Insert(idx: Integer; dt: TDateTime; dQty: Double);
var
  p: PDateQty;
begin
  p := New(PDateQty);
  p^.dt := dt;
  p^.dQty := dQty;
  FList.Insert(idx, p);
end;

function TLineData.IndexOfDate(dt: TDateTime): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to DateQtyCount - 1 do
  begin
    if DateQtys[i].dt = dt then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TLineData.Clear;
var
  i: Integer;
  p: PDateQty;
begin
  for i := 0 to DateQtyCount - 1 do
  begin
    p := DateQtys[i];
    Dispose(p);
  end;
  FList.Clear;


  for i := 0 to FMonthList.Count - 1 do
  begin
    p := FMonthList[i];
    Dispose(p);
  end;
  FMonthList.Clear;
end;

function TLineData.GetDateQtyCount: Integer;
begin
  Result := FList.Count;
end;

function TLineData.GetDateQtys(i: Integer): PDateQty;
begin
  Result := nil;
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := PDateQty(FList[i]);
  end;
end;

function TLineData.GetMonthQtyCount: Integer;
begin
  Result := FMonthList.Count;
end;

function TLineData.GetMonthQtys(i: Integer): PDateQty;
begin
  Result := nil;
  if (i >= 0) and (i < FMonthList.Count) then
  begin
    Result := PDateQty(FMonthList[i]);
  end;
end;

{ TWeekData }
           
constructor TWeekData.Create(const sWeek: string);
begin
  FName := sWeek;
  FColors := TList.Create;
  FVers := TList.Create;
  FCaps := TList.Create;
  FFGs := TList.Create;
  FPkgs := TList.Create;
end;

destructor TWeekData.Destroy;
begin
  Clear;
  FColors.Free;
  FVers.Free;
  FCaps.Free;
  FFGs.Free;
  FPkgs.Free;
  inherited;
end;

function GetLineData(lst: TList; const skey: string): TLineData;
var
  i: Integer;
  ld: TLineData;
begin
  Result := nil;
  for i := 0 to lst.Count -1  do
  begin
    ld := TLineData(lst[i]);
    if ld.FName = skey then
    begin
      Result := ld;
      Break;
    end;
  end;
end;

procedure TWeekData.Clear;
  procedure ClearList(lst: TList);    
  var
    i: Integer;
    ld: TLineData;
  begin
    for i := 0 to lst.Count - 1 do
    begin
      ld := lst[i];
      ld.Free;
    end;
    lst.Clear;
  end;
begin
  ClearList(FColors);   
  ClearList(FVers);
  ClearList(FCaps);
  ClearList(FFGs);
  ClearList(FPkgs);
end;

function TWeekData.GetColorCount: Integer;
begin
  Result := FColors.Count;
end;

function TWeekData.GetVerCount: Integer;
begin
  Result := FVers.Count;
end;

function TWeekData.GetCapCount: Integer;
begin
  Result := FCaps.Count;
end;

function TWeekData.GetFGCount: Integer;
begin
  Result := FFGs.Count;
end;

function TWeekData.GetPkgCount: Integer;
begin
  Result := FPkgs.Count;
end;

function TWeekData.GetColors(i: Integer): TLineData;
begin
  if (i >= 0) and (i < ColorCount) then
  begin
    Result := TLineData(FColors[i]);
  end
  else Result := nil;
end;

function TWeekData.GetVers(i: Integer): TLineData;
begin
  if (i >= 0) and (i < VerCount) then
  begin
    Result := TLineData(FVers[i]);
  end
  else Result := nil;
end;

function TWeekData.GetCaps(i: Integer): TLineData;
begin
  if (i >= 0) and (i < CapCount) then
  begin
    Result := TLineData(FCaps[i]);
  end
  else Result := nil;
end;

function TWeekData.GetFGs(i: Integer): TLineData;
begin
  if (i >= 0) and (i < FGCount) then
  begin
    Result := TLineData(FFGs[i]);
  end
  else Result := nil;
end;

function TWeekData.GetPkgs(i: Integer): TLineData;
begin
  if (i >= 0) and (i < PkgCount) then
  begin
    Result := TLineData(FPkgs[i]);
  end
  else Result := nil;
end;
  
{ TProjData }

constructor TProjData.Create(const sName: string);
begin
  FName := sname;
  FList := TList.Create;
  FCurrWeekData := TWeekData.Create('');
end;

destructor TProjData.Destroy;
begin
  Clear;
  FList.Free;
  FCurrWeekData.Free;
  inherited;
end;

function TProjData.AddWeek(const sWeek: string): Integer;
var
  wd: TWeekData;
begin
  wd := TWeekData.Create(sWeek);
  Result := FList.Add(wd);
end;

procedure TProjData.SortDate;
var
  wd: TWeekData;
  i: Integer;
  idx: Integer;
  ld: TLineData;
  iWeek: Integer;
begin
  if WeekCount = 0 then Exit;

  if FCurrWeekData.ColorCount = 0 then Exit;

  wd := Weeks[0];
  ld := wd.Colors[0];

  for i := 0 to ld.DateQtyCount - 1 do
  begin
    idx := FCurrWeekData.Colors[0].IndexOfDate(ld.DateQtys[i].dt);
    if idx < 0 then
    begin
      InsertDateQty(FCurrWeekData, ld.DateQtys[i].dt);
    end;
  end;

  for i := 0 to FCurrWeekData.Colors[0].DateQtyCount - 1 do
  begin
    idx := ld.IndexOfDate(FCurrWeekData.Colors[0].DateQtys[i].dt);
    if idx < 0 then
    begin
      for iWeek := 0 to WeekCount -1 do
      begin
        InsertDateQty(Weeks[iWeek], FCurrWeekData.Colors[0].DateQtys[i].dt);
      end;
    end;
  end; 
end;

procedure TProjData.SumByMonth;
var
  iWeek: Integer;
  aWeekData: TWeekData;
  aLineData: TLineData;
  dtMonth0: TDateTime;
  dtMonth: TDateTime;
  iKey: Integer;
  iDate: Integer;
  aDateQtyPtr: PDateQty;
begin
  for iWeek := 0 to WeekCount - 1 do
  begin
    aWeekData := Weeks[iWeek];
    
    for iKey := 0 to aWeekData.ColorCount - 1 do
    begin
      dtMonth0 := 0;
      
      aLineData := aWeekData.Colors[iKey];
      for iDate := 0 to aLineData.DateQtyCount - 1 do
      begin
        dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
        if (iDate = 0) or (dtMonth0 <> dtMonth) then
        begin
          dtMonth0 := dtMonth;
          aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
        end
        else
        begin
          aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
          aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
        end;
      end;
    end;
    
    for iKey := 0 to aWeekData.VerCount - 1 do
    begin
      dtMonth0 := 0;
      
      aLineData := aWeekData.Vers[iKey];
      for iDate := 0 to aLineData.DateQtyCount - 1 do
      begin
        dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
        if (iDate = 0) or (dtMonth0 <> dtMonth) then
        begin
          dtMonth0 := dtMonth;
          aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
        end
        else
        begin
          aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
          aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
        end;
      end;
    end;  
    
    for iKey := 0 to aWeekData.CapCount - 1 do
    begin
      dtMonth0 := 0;
      
      aLineData := aWeekData.Caps[iKey];
      for iDate := 0 to aLineData.DateQtyCount - 1 do
      begin
        dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
        if (iDate = 0) or (dtMonth0 <> dtMonth) then
        begin
          dtMonth0 := dtMonth;
          aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
        end
        else
        begin
          aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
          aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
        end;
      end;
    end;   
    
    for iKey := 0 to aWeekData.FGCount - 1 do
    begin
      dtMonth0 := 0;
      
      aLineData := aWeekData.FGs[iKey];
      for iDate := 0 to aLineData.DateQtyCount - 1 do
      begin
        dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
        if (iDate = 0) or (dtMonth0 <> dtMonth) then
        begin
          dtMonth0 := dtMonth;
          aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
        end
        else
        begin
          aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
          aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
        end;
      end;
    end;   
    
    for iKey := 0 to aWeekData.PkgCount - 1 do
    begin
      dtMonth0 := 0;
      
      aLineData := aWeekData.Pkgs[iKey];
      for iDate := 0 to aLineData.DateQtyCount - 1 do
      begin
        dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
        if (iDate = 0) or (dtMonth0 <> dtMonth) then
        begin
          dtMonth0 := dtMonth;
          aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
        end
        else
        begin
          aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
          aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
        end;
      end;
    end;
  end;

  aWeekData := FCurrWeekData;  ///////////////////////////////////////////////////////////////////////

  for iKey := 0 to aWeekData.ColorCount - 1 do
  begin
    dtMonth0 := 0;
    aLineData := aWeekData.Colors[iKey];
    for iDate := 0 to aLineData.DateQtyCount - 1 do
    begin
      dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
      if (iDate = 0) or (dtMonth0 <> dtMonth) then
      begin
        dtMonth0 := dtMonth;
        aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
      end
      else
      begin
        aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
        aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
      end;
    end;
  end;   

  for iKey := 0 to aWeekData.VerCount - 1 do
  begin
    dtMonth0 := 0;
    aLineData := aWeekData.Vers[iKey];
    for iDate := 0 to aLineData.DateQtyCount - 1 do
    begin
      dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
      if (iDate = 0) or (dtMonth0 <> dtMonth) then
      begin
        dtMonth0 := dtMonth;
        aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
      end
      else
      begin
        aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
        aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
      end;
    end;
  end;

  for iKey := 0 to aWeekData.CapCount - 1 do
  begin
    dtMonth0 := 0;
    aLineData := aWeekData.Caps[iKey];
    for iDate := 0 to aLineData.DateQtyCount - 1 do
    begin
      dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
      if (iDate = 0) or (dtMonth0 <> dtMonth) then
      begin
        dtMonth0 := dtMonth;
        aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
      end
      else
      begin
        aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
        aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
      end;
    end;
  end;

  for iKey := 0 to aWeekData.FGCount - 1 do
  begin
    dtMonth0 := 0;
    aLineData := aWeekData.FGs[iKey];
    for iDate := 0 to aLineData.DateQtyCount - 1 do
    begin
      dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
      if (iDate = 0) or (dtMonth0 <> dtMonth) then
      begin
        dtMonth0 := dtMonth;
        aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
      end
      else
      begin
        aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
        aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
      end;
    end;
  end;

  for iKey := 0 to aWeekData.PkgCount - 1 do
  begin
    dtMonth0 := 0;
    aLineData := aWeekData.Pkgs[iKey];
    for iDate := 0 to aLineData.DateQtyCount - 1 do
    begin
      dtMonth := EncodeDate( YearOf(aLineData.DateQtys[iDate]^.dt), MonthOf(aLineData.DateQtys[iDate]^.dt), 1);
      if (iDate = 0) or (dtMonth0 <> dtMonth) then
      begin
        dtMonth0 := dtMonth;
        aLineData.AddMonth(dtMonth, aLineData.DateQtys[iDate]^.dQty);
      end
      else
      begin
        aDateQtyPtr := aLineData.MonthQtys[aLineData.MonthQtyCount - 1];
        aDateQtyPtr^.dQty := aDateQtyPtr^.dQty + aLineData.DateQtys[iDate]^.dQty;
      end;
    end;
  end;
end;

procedure TProjData.InsertDateQty(aWeekData: TWeekData; dt: TDateTime);
var
  wd: TWeekData;
  ld: TLineData;
  i: Integer;
  iInsertIdx: Integer;
  iKey: Integer;
  p: PDateQty;
begin
  p := New(PDateQty);
  p^.dt := dt;
  p^.dQty := 0;

  wd := aWeekData;
  ld := wd.Colors[0];

  iInsertIdx := -1;

  for i := 0 to ld.DateQtyCount - 1 do
  begin
    if ld.DateQtys[i].dt > dt then
    begin
      iInsertIdx := i;
      Break;
    end;
  end;

  if iInsertIdx >= 0 then
  begin
    for iKey := 0 to aWeekData.ColorCount - 1 do
    begin
      aWeekData.Colors[iKey].Insert(iInsertIdx, p^.dt, 0); 
    end;    
    for iKey := 0 to aWeekData.VerCount - 1 do
    begin
      aWeekData.Vers[iKey].Insert(iInsertIdx, p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.CapCount - 1 do
    begin
      aWeekData.Caps[iKey].Insert(iInsertIdx, p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.FGCount - 1 do
    begin
      aWeekData.FGs[iKey].Insert(iInsertIdx, p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.PkgCount - 1 do
    begin
      aWeekData.Pkgs[iKey].Insert(iInsertIdx, p^.dt, 0);
    end;
  end
  else
  begin
    for iKey := 0 to aWeekData.ColorCount - 1 do
    begin
      aWeekData.Colors[iKey].Add(p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.VerCount - 1 do
    begin
      aWeekData.Vers[iKey].Add(p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.CapCount - 1 do
    begin
      aWeekData.Caps[iKey].Add(p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.FGCount - 1 do
    begin
      aWeekData.FGs[iKey].Add(p^.dt, 0);
    end;
    for iKey := 0 to aWeekData.PkgCount - 1 do
    begin
      aWeekData.Pkgs[iKey].Add(p^.dt, 0);
    end;
  end;
end;

procedure TProjData.Clear;
var
  i: Integer;
  wd: TWeekData;
begin
  for i := 0 to WeekCount -1  do
  begin
    wd := Weeks[i];
    wd.Free;
  end;
  FList.Clear;
end;

function TProjData.GetWeekCount;
begin
  Result := FList.Count;
end;

function TProjData.GetWeekds(i: Integer): TWeekData;
begin
  if (i >= 0) and (i < FList.Count) then
  begin
    Result := TWeekData(FList[i]);
  end
  else Result := nil;
end;
  
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

procedure TfrmWaterfall.btnWaterFallClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;

  leWaterFall.Text := sfile;
end;

procedure TfrmWaterfall.btnSOPSumClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;

  leSOPSum.Text := sfile;
end;

procedure TfrmWaterfall.tbSaveClick(Sender: TObject);
var 
  sfile: string;
  dwtick: DWORD;
begin
  sfile := 'wf.xlsx';
  if not ExcelSaveDialog(sfile) then Exit;
 // DeleteFile(sfile);

//  sfile := 'C:\Users\qiujinbo\Desktop\WF\wf.xlsx';

  FProjs := TList.Create;
  try
    dwtick := GetTickCount;

    if FileExists(leWaterFall.Text) then
    begin
      ReadExcelFile(leWaterFall.Text, OpenWF);
    end;

    dwtick := GetTickCount - dwtick;
    Memo1.Lines.Add(Format('读取Waterfall耗时：%0.2f 秒', [dwtick / 1000]));
  
    ReadExcelFile(leSOPSum.Text, OpenSOP);
    SortDate;        
    SumByMonth;
    WriteExcelFile(sfile, WriteWF);
    MessageBox(Handle, 'ok', 'ok', 0);
  finally
    Clear;
    FProjs.Free;
  end;
end;

function TfrmWaterfall.FindFirstNeedDateCellDetail(ExcelApp: Variant; var row, col: Integer): Boolean;
var
  r, c: Integer;
  s: string;
  s1, s2: string;
begin
  Result := False;
  for r := 1 to 5 do
  begin
    for c := 1 to 5 do
    begin
      s1 := ExcelApp.Cells[r, c].Value;
      s2 := ExcelApp.Cells[r, c + 1].Value;
      s := s1 + s2;
      if s = '需求日期合计' then
      begin
        Result := True;
        row := r;
        col := c;
        Exit;
      end;
    end;
  end;
end;

function TfrmWaterfall.GetDateColRangeDetail(ExcelApp: Variant; r, c: Integer; var col1, col2: Integer): Boolean;
var
  v: Variant;
begin
  col1 := c;
  col2 := col1;
  v := ExcelApp.Cells[r, col1].Value;
  while VarIsType(v, varDate) do
  begin
    col2 := col2 + 1;
    v := ExcelApp.Cells[r, col2].Value;
  end;
  col2 := col2 - 1;
  Result := col1 <= col2;
end;

procedure TfrmWaterfall.GetWeekCountDetail(ExcelApp: Variant; irow, icol: Integer; lstRowOfWeeks: TList);
var
  r: Integer;
  s: string;
  s2: string;
begin
  r := irow;
  s := ExcelApp.Cells[r, icol].Value;

  lstRowOfWeeks.Add(Pointer(r));

  while s <> '' do
  begin
    r := r + 1;
    s := ExcelApp.Cells[r, icol].Value;
    if s = '' then
    begin
      s2 := ExcelApp.Cells[r + 1, icol].Value;
      if s2 = '' then    // 连续两行空白，认为结束
      begin
        Break;
      end
      else
      begin
        s := s2;
        r := r + 1;
        lstRowOfWeeks.Add(Pointer(r));
      end;
    end;
  end;
end;

function TfrmWaterfall.ReadDatesQtyDetail(ExcelApp: Variant; irow: Integer;
  irow1, irow2: Integer; icol1, icol2: Integer; lst: TList; iKeyCol: Integer): string;
var
  r, c: Integer;
  dQty: Double;
  dt: TDateTime;
  ld: TLineData;
  s: string;
  rx: Integer;
  cx: Integer;
begin
  rx := 0;
  cx := 0;
  try
    for r := irow1 to irow2 do
    begin
      rx := r;
      s :=  ExcelApp.Cells[r, iKeyCol].Value;
      ld := TLineData.Create(s);
      lst.Add(ld);
      for c := icol1 to icol2 do
      begin
        cx := c;
        dt := ExcelApp.Cells[irow, c].Value;
        dQty := ExcelApp.Cells[r, c].Value;
        ld.Add(dt, dQty);
      end;
    end;
  except
    on e: Exception do
    begin
      Memo1.Lines.Add(e.Message + 'irow: ' + IntToStr(irow) + '    c: ' + IntToStr(cx) + '   r: ' + IntToStr(rx) + '  iKeyCol: ' + IntToStr(iKeyCol));
      raise e;
    end;
  end;
end;

procedure TfrmWaterfall.ReadWeekDetail(ExcelApp: Variant; irow, icol: Integer;
  dtcol1, dtcol2: Integer; aProjData: TProjData);
var
  sWeek: string;
  irow1, irow2: Integer;
  idx: Integer;
  wd: TWeekData;
  irowDate: Integer;
begin

  sWeek := ExcelApp.Cells[irow, icol - 1].Value;
  idx := aProjData.AddWeek(sWeek);
  wd := aProjData.Weeks[idx];

  irowDate := irow;
  irow := irow + 1;

  // 颜色 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, icol - 1, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, wd.FColors, dtcol1 - 2);
  irow := irow2 + 1;
  // 制式 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, icol - 1, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, wd.FVers, dtcol1 - 2);
  irow := irow2 + 1;
  // 容量 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, icol - 1, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, wd.FCaps, dtcol1 - 2);
  irow := irow2 + 1;
  // 整机/裸机 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, icol - 1, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, wd.FFGs, dtcol1 - 2);
  irow := irow2 + 1;
  // 礼盒装 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, icol - 1, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, wd.FPkgs, dtcol1 - 2);
end;

procedure TfrmWaterfall.OpenWFDetail(ExcelApp, WorkBook: Variant; aProjData: TProjData);
var
  irow, icol: Integer;
  icolDate1, icolDate2: Integer;
  lstRowOfWeeks: TList;
  iWeek: Integer;
begin
  if not FindFirstNeedDateCellDetail(ExcelApp, irow, icol) then Exit;

  GetDateColRangeDetail(ExcelApp, irow, icol + 2, icolDate1, icolDate2);

  lstRowOfWeeks := TList.Create;

  try
    GetWeekCountDetail(ExcelApp, irow, icol, lstRowOfWeeks);

    for iWeek := 0 to lstRowOfWeeks.Count - 1 do
    begin
      irow := Integer(lstRowOfWeeks[iWeek]);
      ReadWeekDetail(ExcelApp, irow, icol, icolDate1, icolDate2, aProjData);
    end;

  finally
    lstRowOfWeeks.Free;
  end;
  
end;
 
procedure TfrmWaterfall.OpenSOPSheet(ExcelApp: Variant; aProjData: TProjData;
  const sWeek: string);
var
  irow: Integer;
  irow1, irow2: Integer;
  s: string;
  dtcol1, dtcol2: Integer;
  irowDate: Integer;
begin 
  aProjData.FCurrWeekData.FName := sWeek;
   
  irow := 1;
  s := ExcelApp.Cells[irow, 1].Value;
  while s <> '' do
  begin
    irow := irow + 1;
    s := ExcelApp.Cells[irow, 1].Value;
  end;
  irow := irow + 1;

  irowDate := irow;
  GetDateColRangeDetail(ExcelApp, irowDate, 8, dtcol1, dtcol2);

  irow := irow + 1;  /////////////////////////////////////////////////////////
  // 颜色 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, 6, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, aProjData.FCurrWeekData.FColors, dtcol1 - 1);
  irow := irow2 + 1;
  // 制式 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, 6, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, aProjData.FCurrWeekData.FVers, dtcol1 - 1);
  irow := irow2 + 1;
  // 容量 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, 6, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, aProjData.FCurrWeekData.FCaps, dtcol1 - 1);
  irow := irow2 + 1;
  // 整机/裸机 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, 6, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, aProjData.FCurrWeekData.FFGs, dtcol1 - 1);
  irow := irow2 + 1;
  // 礼盒装 ------------------------------------------
  GetRowMergedRange(ExcelApp, irow, 6, irow1, irow2);
  ReadDatesQtyDetail(ExcelApp, irowDate, irow1, irow2, dtcol1, dtcol2, aProjData.FCurrWeekData.FPkgs, dtcol1 - 1);
end;

procedure TfrmWaterfall.OpenWF(ExcelApp, WorkBook: Variant);
  function ExtractProj(const sSheet: string): string;
  begin
    Result := Copy(sSheet, 1, Pos('-', sSheet) - 1);
  end;
  function IsDetail(const sSheet: string): Boolean;
  begin
    Result := UpperCase(Copy(sSheet, Pos('-', sSheet) + 1, Length(sSheet) - Pos('-', sSheet))) = 'DETAIL';
  end;  
  function IsSummary(const sSheet: string): Boolean;
  begin
    Result := UpperCase(Copy(sSheet, Pos('-', sSheet) + 1, Length(sSheet) - Pos('-', sSheet))) = 'SUMMARY';
  end;
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sProj: string;
  aProjData: TProjData;
begin
  iSheetCount := ExcelApp.WorkSheets.Count;
  for iSheet := 1 to iSheetCount do
  begin
    ExcelApp.WorkSheets[iSheet].Activate;       
    if not ExcelApp.Sheets[iSheet].Visible then Continue;
    sSheet := ExcelApp.WorkSheets[iSheet].Name;  
    Memo1.Lines.Add(sSheet);

    sProj := ExtractProj(sSheet);
    if sProj = '' then
    begin
      Memo1.Lines.Add('Waterfall Sheet名称不对' + sProj);
      Continue;
    end;

    aProjData := FindProj(sProj);
    if aProjData = nil then
    begin
      aProjData := TProjData.Create(sProj);
      FProjs.Add(aProjData);
    end;

    if IsDetail(sSheet) then
    begin
      OpenWFDetail(ExcelApp, WorkBook, aProjData);
    end; 
  end;
end;
    
procedure TfrmWaterfall.OpenSOP(ExcelApp, WorkBook: Variant);
var
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sProj: string;
  aProjData: TProjData;
  sWeek: string;
  s: string;
begin
  sWeek := leSOPSum.Text;
  sWeek := ChangeFileExt(sWeek, '');
  // xxx(week).xlsx
  sWeek := Copy(sWeek, Pos('(', sWeek) + 1, Pos(')', sWeek) - Pos('(', sWeek) - 1);
  if sWeek = '' then
  begin
    Memo1.Lines.Add('SOP 文件名格式不对');
    Exit;
  end;
  
  iSheetCount := ExcelApp.WorkSheets.Count;
  for iSheet := 1 to iSheetCount do
  begin
    ExcelApp.WorkSheets[iSheet].Activate;
    sSheet := ExcelApp.WorkSheets[iSheet].Name;

    sProj := sSheet;

    s := ExcelApp.Cells[1, 1].Value + ExcelApp.Cells[1, 2].Value;
    if s <> '制式物料长代码' then
    begin
      Memo1.Lines.Add(sProj + '  Sheet 格式不符合SOP' );
      Continue;
    end;

    aProjData := FindProj(sProj);
    if aProjData = nil then
    begin
      aProjData := TProjData.Create(sProj);
      FProjs.Add(aProjData);
    end;
    
    OpenSOPSheet(ExcelApp, aProjData, sWeek);
  end;
end;

procedure TfrmWaterfall.Clear;
var
  i: Integer;
  aProjData: TProjData;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs[i]);
    aProjData.Free;
  end;
  FProjs.Clear;
end;

function TfrmWaterfall.FindProj(const sProj: string): TProjData;
var
  i: Integer;
  aProjData: TProjData;
begin
  Result := nil;
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs[i]);
    if aProjData.FName = sProj then
    begin
      Result := aProjData;
      Break;
    end;
  end;
end;

procedure TfrmWaterfall.SortDate;
var
  i: Integer;
  aProjData: TProjData;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs[i]);
    aProjData.SortDate;
  end;
end;

procedure TfrmWaterfall.SumByMonth;
var
  i: Integer;
  aProjData: TProjData;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs[i]);
    aProjData.SumByMonth;
  end;
end;  

procedure TfrmWaterfall.WriteWFKeys(ExcelApp: Variant; aWeekData: TWeekData; const sTitle: string;
  iDateQtyCount: Integer; aLst: TList; aKeys: TStringList; var irow, irow1, irow2: Integer);
var
  icol: Integer;
  iDate: Integer;
  iKey: Integer;
  aLineData: TLineData;
begin
  irow1 := irow;
  ExcelApp.Cells[irow, 2].Value := sTitle;
  for iKey := 0 to aKeys.Count - 1 do
  begin
    //if aKeys[iKey] = '2+16G' then asm int 3 end;
    
    aLineData := GetLineData(aLst, aKeys[iKey]);
    if aLineData = nil then Continue;
    aLineData.FRow := irow;
    ExcelApp.Cells[irow, 3].Value := aLineData.FName;    
    ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      if iDate >= aLineData.DateQtyCount then asm int 3 end;
      ExcelApp.Cells[irow, icol].Value := aLineData.DateQtys[iDate]^.dQty
    end;     
    irow := irow + 1;
  end;  
  irow2 := irow - 1;
  ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, 2]].MergeCells := True;
  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
  ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
end;
           
procedure TfrmWaterfall.WriteWFKeysSum(ExcelApp: Variant; aWeekData: TWeekData; const sTitle: string;
  iMonthQtyCount: Integer; aLst: TList; aKeys: TStringList; var irow, irow1, irow2: Integer);
var
  icol: Integer;
  iMonth: Integer;
  iKey: Integer;
  aLineData: TLineData;
begin
  irow1 := irow;
  for iKey := 0 to aKeys.Count - 1 do
  begin
    aLineData := GetLineData(aLst, aKeys[iKey]);
    if aLineData = nil then Continue;
    aLineData.FRow := irow;
    ExcelApp.Cells[irow, 3].Value := aLineData.FName;    
    ExcelApp.Cells[irow, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irow) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irow) + ')';
    for iMonth := 0 to iMonthQtyCount - 1 do
    begin
      icol := iMonth + 4;
      try
        ExcelApp.Cells[irow, icol].Value := aLineData.MonthQtys[iMonth]^.dQty
      except
        asm int 3 end;
      end;
    end;     
    irow := irow + 1;
  end;  
  irow2 := irow - 1;
//  ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, 2]].MergeCells := True;
//  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
//  ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
end;

procedure TfrmWaterfall.WriteWF11Compares(ExcelApp: Variant; aProj: TProjData;
  aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
  var irow: Integer; aLineDatasGetter: TLineDatasGetter; bSOP: Boolean);
var
  iKey: Integer;
  iDate: Integer;
  icol: Integer;
  aLineData: TLineData;
  aLineData0: TLineData;
  irow1, irow2: Integer;
  iWeek: Integer;
  aWeekData: TWeekData;
  aWeekData0: TWeekData;
  aKeys: TList;
  sWeekDate: string;
  dtWeekDate: TDateTime;
  dt: TDateTime;
begin
  try
    // 颜色 /////////////////
    for iKey := 0 to aColors.Count - 1 do
    begin
      ExcelApp.Cells[irow, 2].Value := sKeyType + '变化(' + aColors[iKey] + ')'; 
      ExcelApp.Cells[irow, 3].Value := '需求日期';
      ExcelApp.Cells[irow, 4].Value := '合计'; 

      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value := aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt;
      end;

      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

      irow := irow + 1;

      aLineData0 := nil;
      irow1 := irow;
      for iWeek := 0 to aProj.WeekCount - 1 do
      begin
        aWeekData := aProj.Weeks[iWeek];
        aKeys := aLineDatasGetter(aWeekData);
        aLineData := GetLineData(aKeys, aColors[iKey]);
     
        sWeekDate := aWeekData.FName;
        sWeekDate := Copy(sWeekDate, 1, Pos(' ', sWeekDate) - 1);
        dtWeekDate := myStrToDateTime(sWeekDate);

        if iWeek > 0 then
        begin
          aWeekData0 := aProj.Weeks[iWeek - 1];    
          aKeys := aLineDatasGetter(aWeekData0);
          aLineData0 := GetLineData(aKeys, aColors[iKey]);
        end;

        ExcelApp.Cells[irow, 2].Value := aWeekData.FName;
        ExcelApp.Cells[irow, 3].Value := aColors[iKey];
        for iDate := 0 to iDateQtyCount - 1 do
        begin
          icol := iDate + 5;
          if aLineData = nil then
          begin
            if aLineData0 = nil then
            begin
              ExcelApp.Cells[irow, icol].Value :=  '';
            end
            else
            begin
              ExcelApp.Cells[irow, icol].Value :=  '=0-' + GetRef(icol) + IntToStr(aLineData0.FRow);
            end;
          end
          else
          begin
            if aLineData0 = nil then
            begin
              ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData.FRow);
            end
            else
            begin
              ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData0.FRow);
            end;    

            dt := 0;
            if aLineData <> nil then
            begin
              dt := aLineData.DateQtys[iDate].dt;
            end
            else if aLineData0 <> nil then
            begin
              dt := aLineData0.DateQtys[iDate].dt;
            end;
            if (dt <> 0) and (dt <= dtWeekDate) then
            begin
              ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
            end;

          end;
        end;
        ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
        irow := irow + 1;
      end;

      if bSOP then
      begin
        // 写当周 //////////////////////////////////////////////////////////////////
        aWeekData := aProj.FCurrWeekData;
        aKeys := aLineDatasGetter(aWeekData);
        aLineData := GetLineData(aKeys, aColors[iKey]);

        aWeekData0 := aProj.Weeks[aProj.WeekCount - 1];
        aKeys := aLineDatasGetter(aWeekData0);
        aLineData0 := GetLineData(aKeys, aColors[iKey]);

        sWeekDate := aWeekData.FName;
        sWeekDate := Copy(sWeekDate, 1, Pos(' ', sWeekDate) - 1);
        dtWeekDate := myStrToDateTime(sWeekDate);

        ExcelApp.Cells[irow, 2].Value := aWeekData.FName;
        ExcelApp.Cells[irow, 3].Value := aColors[iKey];
        for iDate := 0 to iDateQtyCount - 1 do
        begin
          try
            icol := iDate + 5;
            if aLineData = nil then
            begin
              if aLineData0 = nil then
              begin
                ExcelApp.Cells[irow, icol].Value :=  '';
              end
              else
              begin
                ExcelApp.Cells[irow, icol].Value :=  '=0-' + GetRef(icol) + IntToStr(aLineData0.FRow);
              end;
            end
            else
            begin
              if aLineData0 = nil then
              begin
                ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData.FRow);
              end
              else
              begin
                ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData0.FRow);
              end;
              if aLineData.DateQtys[iDate].dt <= dtWeekDate then
              begin
                ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
              end;
            end;
          
            dt := 0;
            if aLineData <> nil then
            begin
              dt := aLineData.DateQtys[iDate].dt;
            end
            else if aLineData0 <> nil then
            begin
              dt := aLineData0.DateQtys[iDate].dt;
            end;
            if (dt <> 0) and (dt <= dtWeekDate) then
            begin
              ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
            end;
          except
            asm int 3 end;
          end;
        end;
        ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
        irow := irow + 1;
      end;
      /////////////////////////////////////////////////////////////////

      irow2 := irow - 1;
      //ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Interior.Color := $F3EEDA;
      ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
      ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

      ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

      ExcelApp.Cells[irow, 3].Value := aColors[iKey] + '累计变化';
      ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';    
      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
      end;                   
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $ECDFE4;
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框         
      ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';    

      ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
      ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

      irow := irow + 2;
    end;
  except
    on e: Exception do
    begin
      Memo1.Lines.Add(e.Message + '    irow: ' + IntToStr(irow) + '  icol: ' + IntToStr(icol));
      raise e;
    end;
  end;
end;

procedure TfrmWaterfall.WriteWF11ComparesTTL(ExcelApp: Variant; aProj: TProjData;
  aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
  var irow: Integer; aLineDatasGetter: TLineDatasGetter; bSOP: Boolean);
var
  iKey: Integer;
  iDate: Integer;
  icol: Integer; 
  irow1, irow2: Integer;
  iWeek: Integer;
  aWeekData: TWeekData;
  aWeekData0: TWeekData; 
  sWeekDate: string;
  dtWeekDate: TDateTime;
  dt: TDateTime;
  s: string;
begin
  ExcelApp.Cells[irow, 2].Value := sKeyType + '变化';
  ExcelApp.Cells[irow, 3].Value := '需求日期';
  ExcelApp.Cells[irow, 4].Value := '合计'; 

  for iDate := 0 to iDateQtyCount - 1 do
  begin
    icol := iDate + 5;
    ExcelApp.Cells[irow, icol].Value := aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt;
  end;

  ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
  ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

  irow := irow + 1;

  try
    aWeekData0 := nil;
    irow1 := irow;
    for iWeek := 0 to aProj.WeekCount - 1 do
    begin
      aWeekData := aProj.Weeks[iWeek]; 
     
      sWeekDate := aWeekData.FName;
      sWeekDate := Copy(sWeekDate, 1, Pos(' ', sWeekDate) - 1);
      dtWeekDate := myStrToDateTime(sWeekDate);
       
      if iWeek > 0 then
      begin
        aWeekData0 := aProj.Weeks[iWeek - 1];     
      end;

      ExcelApp.Cells[irow, 2].Value := aWeekData.FName;
      ExcelApp.Cells[irow, 3].Value := sKeyType;
      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        if aWeekData0 = nil then
        begin
          s := '=0';
          for iKey := 0 to aWeekData.ColorCount - 1 do
          begin
            s := s + '+' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
          end;    
          for iKey := 0 to aWeekData.ColorCount - 1 do
          begin
            s := s + '-' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
          end; 
          ExcelApp.Cells[irow, icol].Value := s;
        end
        else
        begin
          s := '=0';   
          for iKey := 0 to aWeekData.ColorCount - 1 do
          begin
            s := s + '+' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
          end; 
          for iKey := 0 to aWeekData0.ColorCount - 1 do
          begin
            s := s + '-' + GetRef(icol) + IntToStr(aWeekData0.Colors[iKey].FRow);
          end;
          ExcelApp.Cells[irow, icol].Value := s;
        end;
 
        dt := aWeekData.Colors[0].DateQtys[iDate].dt;
      
        if (dt <> 0) and (dt <= dtWeekDate) then
        begin
          ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
        end;
      end;
      ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
      irow := irow + 1;
    end;
  except
    on e: Exception do
    begin
      memo1.Lines.Add(e.Message + '   irow: ' + IntToStr(irow) + '   icol: ' + IntToStr(icol));
      raise e;
    end;
  end;

  if bSOP then
  begin
    // 写当周 //////////////////////////////////////////////////////////////////
    aWeekData := aProj.FCurrWeekData; 

    aWeekData0 := aProj.Weeks[aProj.WeekCount - 1]; 

    sWeekDate := aWeekData.FName;
    sWeekDate := Copy(sWeekDate, 1, Pos(' ', sWeekDate) - 1);
    dtWeekDate := myStrToDateTime(sWeekDate);

    ExcelApp.Cells[irow, 2].Value := aWeekData.FName;
    ExcelApp.Cells[irow, 3].Value := sKeyType;
    for iDate := 0 to iDateQtyCount - 1 do
    begin

      icol := iDate + 5;
      if aWeekData0 = nil then
      begin
        s := '=0';
        for iKey := 0 to aWeekData.ColorCount - 1 do
        begin
          s := s + '+' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
        end;    
        for iKey := 0 to aWeekData.ColorCount - 1 do
        begin
          s := s + '-' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
        end; 
        ExcelApp.Cells[irow, icol].Value := s;
      end
      else
      begin
        s := '=0';   
        for iKey := 0 to aWeekData.ColorCount - 1 do
        begin
          s := s + '+' + GetRef(icol) + IntToStr(aWeekData.Colors[iKey].FRow);
        end;
        for iKey := 0 to aWeekData0.ColorCount - 1 do
        begin
          s := s + '-' + GetRef(icol) + IntToStr(aWeekData0.Colors[iKey].FRow);
        end;
        ExcelApp.Cells[irow, icol].Value := s;
      end;
 
      dt := aWeekData.Colors[0].DateQtys[iDate].dt;
      
      if (dt <> 0) and (dt <= dtWeekDate) then
      begin
        ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
      end; 
    end;
    ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
    irow := irow + 1;
  end;
  /////////////////////////////////////////////////////////////////

  irow2 := irow - 1;
  //ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Interior.Color := $F3EEDA;
  ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

  ExcelApp.Cells[irow, 3].Value := sKeyType + '累计变化';
  ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
  for iDate := 0 to iDateQtyCount - 1 do
  begin
    icol := iDate + 5;
    ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
  end;                   
  ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $ECDFE4;
  ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框         
  ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';    

  ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

  irow := irow + 2; 
end;

procedure TfrmWaterfall.WriteWF11ComparesSum(ExcelApp: Variant; aProj: TProjData;
  aColors: TStringList; const sKeyType: string; iMonthQtyCount: Integer;
  var irow: Integer; aLineDatasGetter: TLineDatasGetter; wd1, wd2: TWeekData);
var
  iKey: Integer;
  iMonth: Integer;
  icol: Integer;
  aLineData: TLineData;
  aLineData0: TLineData;  
  aKeys: TList;
  irow1, irow2: Integer;
begin
  irow1 := irow;
  irow2 := irow;
  for iKey := 0 to aColors.Count - 1 do
  begin
    aKeys := aLineDatasGetter(wd1);
    aLineData := GetLineData(aKeys, aColors[iKey]);

    aKeys := aLineDatasGetter(wd2);
    aLineData0 := GetLineData(aKeys, aColors[iKey]);

    ExcelApp.Cells[irow, 3].Value :=  aColors[iKey];

    for iMonth := 0 to iMonthQtyCount - 1 do
    begin
      icol := iMonth + 4;
      if aLineData = nil then
      begin
        if aLineData0 = nil then
        begin

        end
        else
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData0.FRow) + '-' + GetRef(icol) + IntToStr(aLineData0.FRow);
        end;
      end
      else
      begin
        if aLineData0 = nil then
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-0';
        end
        else
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData0.FRow) + '-' + GetRef(icol) + IntToStr(aLineData.FRow);
        end;
      end;
    end;
    ExcelApp.Cells[irow, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irow) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irow) + ')';

    irow2 := irow;
    irow := irow + 1; 
  end;

  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
  ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

end;

procedure TfrmWaterfall.WriteWF01Compares(ExcelApp: Variant; aProj: TProjData;
  aColors: TStringList; const sKeyType: string; iDateQtyCount: Integer;
  var irow: Integer; aLineDatasGetter: TLineDatasGetter);
var
  iKey: Integer;
  iDate: Integer;
  icol: Integer;
  aLineData: TLineData;
  aLineData0: TLineData;
  irow1, irow2: Integer; 
  aWeekData: TWeekData; 
  aKeys: TList;      
  sWeekDate: string;
  dtWeekDate: TDateTime;
begin
  // 颜色 /////////////////
  for iKey := 0 to aColors.Count - 1 do
  begin
    ExcelApp.Cells[irow, 2].Value := sKeyType + '变化(' + aColors[iKey] + ')'; 
    ExcelApp.Cells[irow, 3].Value := '需求日期';
    ExcelApp.Cells[irow, 4].Value := '合计'; 

    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      ExcelApp.Cells[irow, icol].Value := aProj.FCurrWeekData.Colors[0].DateQtys[iDate]^.dt;
    end;

    ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
    ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

    irow := irow + 1;
 
    irow1 := irow;

    // 写当周 //////////////////////////////////////////////////////////////////
    aWeekData := aProj.FCurrWeekData;
    aKeys := aLineDatasGetter(aWeekData);
    aLineData := GetLineData(aKeys, aColors[iKey]);

    aLineData0 := nil;
                
    sWeekDate := aWeekData.FName;
    sWeekDate := Copy(sWeekDate, 1, Pos(' ', sWeekDate) - 1);
    dtWeekDate := myStrToDateTime(sWeekDate);

    ExcelApp.Cells[irow, 2].Value := aWeekData.FName;
    ExcelApp.Cells[irow, 3].Value := aColors[iKey];
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      if aLineData = nil then
      begin
        if aLineData0 = nil then
        begin
          ExcelApp.Cells[irow, icol].Value :=  '';
        end
        else
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=0-' + GetRef(icol) + IntToStr(aLineData0.FRow);
        end;
      end
      else
      begin
        if aLineData0 = nil then
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData.FRow);
        end
        else
        begin
          ExcelApp.Cells[irow, icol].Value :=  '=' + GetRef(icol) + IntToStr(aLineData.FRow) + '-' + GetRef(icol) + IntToStr(aLineData0.FRow);
        end; 
        if aLineData.DateQtys[iDate].dt <= dtWeekDate then
        begin
          ExcelApp.Cells[irow, icol].Interior.Color := $F3EEDA;
        end;
      end; 
    end;
    ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
    irow := irow + 1;

    /////////////////////////////////////////////////////////////////

    irow2 := irow - 1;
    //ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Interior.Color := $F3EEDA;
    ExcelApp.Range[ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
    ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

    ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ExcelApp.Range[ExcelApp.Cells[irow1, 4], ExcelApp.Cells[irow2, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

    ExcelApp.Cells[irow, 3].Value := aColors[iKey] + '累计变化';
    ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
    end;                   
    ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $ECDFE4;
    ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框         
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';    

    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions.Add(xlCellValue, xlLess, '=0', EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].FormatConditions[1].Font.Color := clRed;

    irow := irow + 2;
  end;
end;
  
procedure TfrmWaterfall.WriteWF11Sum(ExcelApp: Variant; aProj: TProjData; bSOP: Boolean);
var
  irow: Integer;
  icol: Integer;
  irow1, irow2: Integer;
  irowVer1, irowVer2: Integer;
  iWeek: Integer;
  iMonth: Integer; 
  iMonthQtyCount: Integer;
  aColors: TStringList;
  aVers: TStringList;
  aCaps: TStringList;
  aFGs: TStringList;
  aPkgs: TStringList;
  aWeekData: TWeekData;
  irowMonth, irowSum: Integer;
  wd1, wd2: TWeekData;
  sweek1, sweek2: string;
begin
  ExcelApp.Columns[1].ColumnWidth := 1;  
  ExcelApp.Columns[2].ColumnWidth := 18; 
  ExcelApp.Columns[3].ColumnWidth := 14;


  aColors := TStringList.Create;
  aVers := TStringList.Create;
  aCaps := TStringList.Create;
  aFGs := TStringList.Create;
  aPkgs := TStringList.Create;

  for iWeek := 0 to aProj.WeekCount - 1 do
  begin
    aWeekData := aProj.Weeks[iWeek];
    GetAllKeys(aWeekData.FColors, aColors);
    GetAllKeys(aWeekData.FVers, aVers);
    GetAllKeys(aWeekData.FCaps, aCaps);
    GetAllKeys(aWeekData.FFGs, aFGs);
    GetAllKeys(aWeekData.FPkgs, aPkgs);
  end;

  if bSOP then
  begin
    GetAllKeys(aProj.FCurrWeekData.FColors, aColors);
    GetAllKeys(aProj.FCurrWeekData.FVers, aVers);
    GetAllKeys(aProj.FCurrWeekData.FCaps, aCaps);
    GetAllKeys(aProj.FCurrWeekData.FFGs, aFGs);
    GetAllKeys(aProj.FCurrWeekData.FPkgs, aPkgs);
  end;

  try
    iMonthQtyCount := aProj.Weeks[0].Colors[0].MonthQtyCount;
    irow := 2;
    for iWeek := 0 to aProj.WeekCount - 1 do
    begin
      irowMonth := irow;
      irowSum := irow + 1;
      
      ExcelApp.Cells[irowMonth, 2].Value := 'S&OP日期';
      ExcelApp.Cells[irowMonth, 3].Value := '月份';
      ExcelApp.Cells[irowSum, 3].Value := aProj.FName + ' S&OP';
                                                                        
      ExcelApp.Range[ ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, 2] ].MergeCells := True;

      irow := irow + 2;
      ExcelApp.Cells[irow, 2].Value := aProj.Weeks[iWeek].FName;

      irow1 := irow;
      irow2 := irow;
                          
      irowVer1 := irow;                                                                              
      WriteWFKeysSum(ExcelApp, aProj.Weeks[iWeek], '制式', iMonthQtyCount, aProj.Weeks[iWeek].FVers, aVers, irow, irow1, irow2);
      irowVer2 := irow - 1;

      WriteWFKeysSum(ExcelApp, aProj.Weeks[iWeek], '容量', iMonthQtyCount, aProj.Weeks[iWeek].FCaps, aCaps, irow, irow1, irow2);
      WriteWFKeysSum(ExcelApp, aProj.Weeks[iWeek], '颜色', iMonthQtyCount, aProj.Weeks[iWeek].FColors, aColors, irow, irow1, irow2);
      WriteWFKeysSum(ExcelApp, aProj.Weeks[iWeek], '整机/裸机', iMonthQtyCount, aProj.Weeks[iWeek].FFGs, aFGs, irow, irow1, irow2);

      irow1 := irowVer1;

      
      for iMonth := 0 to iMonthQtyCount - 1 do
      begin
        icol := iMonth + 4;
        ExcelApp.Cells[irowMonth, icol].Value := IntToStr( MonthOf( aProj.Weeks[0].Colors[0].MonthQtys[iMonth]^.dt ) ) + '月';
        ExcelApp.Cells[irowSum, icol].Value := '=IF(SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irow2) + ')/4=SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),"err")';
      end;
      ExcelApp.Cells[irowMonth, iMonthQtyCount + 4].Value := 'TTL S&OP';             
      ExcelApp.Cells[irowSum, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irowSum) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irowSum) + ')';
                 
      aProj.Weeks[iWeek].FRow := irowSum;
      
      ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $E8DEB7;
      ExcelApp.Range[ExcelApp.Cells[irowSum, 3], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $B7B8E6;
      ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].Borders.LineStyle := 1; //加边框
      ExcelApp.Range[ExcelApp.Cells[irowSum, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
                                                                    
      ExcelApp.Range[ ExcelApp.Cells[irowVer1, 2], ExcelApp.Cells[irow2, 2] ].MergeCells := True;
      
      irow := irow + 1;
    end;

    if bSOP then
    begin
      irowMonth := irow;
      irowSum := irow + 1;
      
      ExcelApp.Cells[irowMonth, 2].Value := 'S&OP日期';
      ExcelApp.Cells[irowMonth, 3].Value := '月份';
      ExcelApp.Cells[irowSum, 3].Value := aProj.FName + ' S&OP';
                                                                           
      ExcelApp.Range[ ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, 2] ].MergeCells := True;

      irow := irow + 2;
      ExcelApp.Cells[irow, 2].Value := aProj.FCurrWeekData.FName;

      irow1 := irow;
      irow2 := irow;
                          
      irowVer1 := irow;                                                                              
      WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '制式', iMonthQtyCount, aProj.FCurrWeekData.FVers, aVers, irow, irow1, irow2);
      irowVer2 := irow - 1;

      WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '容量', iMonthQtyCount, aProj.FCurrWeekData.FCaps, aCaps, irow, irow1, irow2);
      WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '颜色', iMonthQtyCount, aProj.FCurrWeekData.FColors, aColors, irow, irow1, irow2);
      WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '整机/裸机', iMonthQtyCount, aProj.FCurrWeekData.FFGs, aFGs, irow, irow1, irow2);

      irow1 := irowVer1;
      
      for iMonth := 0 to iMonthQtyCount - 1 do
      begin
        icol := iMonth + 4;
        ExcelApp.Cells[irowMonth, icol].Value := IntToStr( MonthOf( aProj.FCurrWeekData.Colors[0].MonthQtys[iMonth]^.dt ) ) + '月';
        ExcelApp.Cells[irowSum, icol].Value := '=IF(SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irow2) + ')/4=SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),"err")';
      end;
      ExcelApp.Cells[irowMonth, iMonthQtyCount + 4].Value := 'TTL S&OP';   
      ExcelApp.Cells[irowSum, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irowSum) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irowSum) + ')';
      aProj.FCurrWeekData.FRow := irowSum;

      ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $E8DEB7;
      ExcelApp.Range[ExcelApp.Cells[irowSum, 3], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $B7B8E6;
      ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].Borders.LineStyle := 1; //加边框
      ExcelApp.Range[ExcelApp.Cells[irowMonth, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
                          
      ExcelApp.Range[ ExcelApp.Cells[irowVer1, 2], ExcelApp.Cells[irow2, 2] ].MergeCells := True;
      
      irow := irow + 1;
    end;

    ///  下面是比较  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    if iMonthQtyCount = 0 then Exit;

    if (iMonthQtyCount = 1) and not bSOP then Exit;

    if bSOP then
    begin
      wd1 := aProj.Weeks[aProj.WeekCount - 1];
      wd2 := aProj.FCurrWeekData;
    end
    else
    begin
      wd1 := aProj.Weeks[aProj.WeekCount - 2];
      wd2 := aProj.Weeks[aProj.WeekCount - 1];
    end;

    irowMonth := irow;
    irowSum := irow + 1;
    irow := irow + 2;

    irow1 := irow;
    WriteWF11ComparesSum(ExcelApp, aProj, aVers, '制式', iMonthQtyCount, irow, LineDatasGetter4Ver, wd1, wd2);
    WriteWF11ComparesSum(ExcelApp, aProj, aCaps, '容量', iMonthQtyCount, irow, LineDatasGetter4Cap, wd1, wd2);
    WriteWF11ComparesSum(ExcelApp, aProj, aColors, '颜色', iMonthQtyCount, irow, LineDatasGetter4Color, wd1, wd2);
    WriteWF11ComparesSum(ExcelApp, aProj, aFGs, '整机/裸机', iMonthQtyCount, irow, LineDatasGetter4FG, wd1, wd2);
    irow2 := irow - 1;
    
    ExcelApp.Cells[irowMonth, 2].Value := 'S&OP变化';
    ExcelApp.Cells[irowMonth, 3].Value := '月份';
    ExcelApp.Cells[irowSum, 3].Value := aProj.FName + ' S&OP变化';
                                                                   
    ExcelApp.Range[ ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, 2] ].MergeCells := True;

    for iMonth := 0 to iMonthQtyCount - 1 do
    begin
      ExcelApp.Cells[irowMonth, iMonth + 4].Value := IntToStr( MonthOf( wd1.Colors[0].MonthQtys[iMonth]^.dt )) + '月';
      ExcelApp.Cells[irowSum, iMonth + 4].Value := '=' + GetRef(iMonth + 4) + IntToStr(wd2.FRow) + '-' + GetRef(iMonth + 4) + IntToStr(wd1.FRow);
    end;

    ExcelApp.Cells[irowMonth, iMonthQtyCount + 4].Value := 'TTL S&OP';
    ExcelApp.Cells[irowSum, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irowSum) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irowSum) + ')';

    sweek1 := Copy(wd1.FName, Pos(' ', wd1.FName) + 1, Length(wd1.FName));
    sweek2 := Copy(wd2.FName, Pos(' ', wd2.FName) + 1, Length(wd2.FName));
    ExcelApp.Cells[irowSum + 1, 2].Value := sweek2 + ' vs ' + sweek1;
    ExcelApp.Range[ ExcelApp.Cells[irow1, 2], ExcelApp.Cells[irow2, 2] ].MergeCells := True;

    ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $E8DEB7;
    ExcelApp.Range[ExcelApp.Cells[irowSum, 3], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $00FFFF;
    ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].Borders.LineStyle := 1; //加边框
    ExcelApp.Range[ExcelApp.Cells[irowMonth, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
       
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iMonthQtyCount + 4]].Font.Name := 'Arial';
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iMonthQtyCount + 4]].Font.Size := 9;

  finally
    aColors.Free;
    aVers.Free;
    aCaps.Free;
    aFGs.Free;
    aPkgs.Free;
  end;
end;

procedure TfrmWaterfall.WriteWF01Sum(ExcelApp: Variant; aProj: TProjData);
var
  irow: Integer;
  icol: Integer;
  irow1, irow2: Integer; 
  iMonth: Integer;
  iMonthQtyCount: Integer;
  aColors: TStringList;
  aVers: TStringList;
  aCaps: TStringList;
  aFGs: TStringList;
  aPkgs: TStringList;
  irowMonth, irowSum: Integer;
  irowVer1, irowVer2: Integer;
begin
  ExcelApp.Columns[1].ColumnWidth := 1;  
  ExcelApp.Columns[2].ColumnWidth := 18; 
  ExcelApp.Columns[3].ColumnWidth := 14;
 
  aColors := TStringList.Create;
  aVers := TStringList.Create;
  aCaps := TStringList.Create;
  aFGs := TStringList.Create;
  aPkgs := TStringList.Create;

  GetAllKeys(aProj.FCurrWeekData.FColors, aColors);
  GetAllKeys(aProj.FCurrWeekData.FVers, aVers);
  GetAllKeys(aProj.FCurrWeekData.FCaps, aCaps);
  GetAllKeys(aProj.FCurrWeekData.FFGs, aFGs);
  GetAllKeys(aProj.FCurrWeekData.FPkgs, aPkgs);

  try
                               
    iMonthQtyCount := aProj.FCurrWeekData.Colors[0].MonthQtyCount;
    irow := 2;

    irowMonth := irow;
    irowSum := irow + 1;
      
    ExcelApp.Cells[irowMonth, 2].Value := 'S&OP日期';
    ExcelApp.Cells[irowMonth, 3].Value := '月份';
    ExcelApp.Cells[irowSum, 3].Value := aProj.FName + ' S&OP';
                                                                           
    ExcelApp.Range[ ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, 2] ].MergeCells := True;

    irow := irow + 2;
    ExcelApp.Cells[irow, 2].Value := aProj.FCurrWeekData.FName;

    irow1 := irow;
    irow2 := irow;
                          
    irowVer1 := irow;                                                                              
    WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '制式', iMonthQtyCount, aProj.FCurrWeekData.FVers, aVers, irow, irow1, irow2);
    irowVer2 := irow - 1;

    WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '容量', iMonthQtyCount, aProj.FCurrWeekData.FCaps, aCaps, irow, irow1, irow2);
    WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '颜色', iMonthQtyCount, aProj.FCurrWeekData.FColors, aColors, irow, irow1, irow2);
    WriteWFKeysSum(ExcelApp, aProj.FCurrWeekData, '整机/裸机', iMonthQtyCount, aProj.FCurrWeekData.FFGs, aFGs, irow, irow1, irow2);

    irow1 := irowVer1;
      
    for iMonth := 0 to iMonthQtyCount - 1 do
    begin
      icol := iMonth + 4;
      ExcelApp.Cells[irowMonth, icol].Value := IntToStr( MonthOf( aProj.FCurrWeekData.Colors[0].MonthQtys[iMonth]^.dt ) ) + '月';
      ExcelApp.Cells[irowSum, icol].Value := '=IF(SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irow2) + ')/4=SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),SUM(' + GetRef(icol) + IntToStr(irowVer1) + ':' + GetRef(icol) + IntToStr(irowVer2) + '),"err")';
    end;
    ExcelApp.Cells[irowMonth, iMonthQtyCount + 4].Value := 'TTL S&OP';   
    ExcelApp.Cells[irowSum, iMonthQtyCount + 4].Value := '=SUM(' + GetRef(4) + IntToStr(irowSum) + ':' + GetRef(iMonthQtyCount + 3) + IntToStr(irowSum) + ')';
    aProj.FCurrWeekData.FRow := irowSum;

    ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $E8DEB7;
    ExcelApp.Range[ExcelApp.Cells[irowSum, 3], ExcelApp.Cells[irowSum, iMonthQtyCount + 4]].Interior.Color := $B7B8E6;
    ExcelApp.Range[ExcelApp.Cells[irowMonth, 2], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].Borders.LineStyle := 1; //加边框    
    ExcelApp.Range[ExcelApp.Cells[irowMonth, 4], ExcelApp.Cells[irow2, iMonthQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';
     
    ExcelApp.Range[ ExcelApp.Cells[irowVer1, 2], ExcelApp.Cells[irow2, 2] ].MergeCells := True;
              
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iMonthQtyCount + 4]].Font.Name := 'Arial';
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iMonthQtyCount + 4]].Font.Size := 9;

    irow := irow + 1;

  finally
    aColors.Free;
    aVers.Free;
    aCaps.Free;
    aFGs.Free;
    aPkgs.Free;
  end;
end;

procedure TfrmWaterfall.WriteWF11(ExcelApp: Variant; aProj: TProjData; bSOP: Boolean);
var
  irow: Integer;
  icol: Integer;
  irow1, irow2: Integer;
  iWeek: Integer;
  iDate: Integer; 
  iDateQtyCount: Integer;
  aColors: TStringList;
  aVers: TStringList;
  aCaps: TStringList;
  aFGs: TStringList;
  aPkgs: TStringList;
  aWeekData: TWeekData;  
begin
  ExcelApp.Columns[1].ColumnWidth := 1;  
  ExcelApp.Columns[2].ColumnWidth := 18; 
  ExcelApp.Columns[3].ColumnWidth := 16;   
  ExcelApp.Columns[4].ColumnWidth := 11;


  aColors := TStringList.Create;
  aVers := TStringList.Create;
  aCaps := TStringList.Create;
  aFGs := TStringList.Create;
  aPkgs := TStringList.Create;

  for iWeek := 0 to aProj.WeekCount - 1 do
  begin
    aWeekData := aProj.Weeks[iWeek];
    GetAllKeys(aWeekData.FColors, aColors);
    GetAllKeys(aWeekData.FVers, aVers);
    GetAllKeys(aWeekData.FCaps, aCaps);
    GetAllKeys(aWeekData.FFGs, aFGs);
    GetAllKeys(aWeekData.FPkgs, aPkgs);
  end;

  if bSOP then
  begin
    GetAllKeys(aProj.FCurrWeekData.FColors, aColors);
    GetAllKeys(aProj.FCurrWeekData.FVers, aVers);
    GetAllKeys(aProj.FCurrWeekData.FCaps, aCaps);
    GetAllKeys(aProj.FCurrWeekData.FFGs, aFGs);
    GetAllKeys(aProj.FCurrWeekData.FPkgs, aPkgs);
  end;

  try

    iDateQtyCount := aProj.Weeks[0].Colors[0].DateQtyCount;
    irow := 2;
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      ExcelApp.Cells[irow, iDate + 5].Value := 'WK' + IntToStr(WeekOf(aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt));
      ExcelApp.Columns[iDate + 5].ColumnWidth := 11;
    end;
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框    
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;

    irow := irow + 1;
    for iWeek := 0 to aProj.WeekCount - 1 do
    begin
      ExcelApp.Cells[irow, 2].Value := aProj.Weeks[iWeek].FName;
      ExcelApp.Cells[irow, 3].Value := '需求日期';
      ExcelApp.Cells[irow, 4].Value := '合计';

      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value := aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt;
      end;

      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

      irow := irow + 1;

      WriteWFKeys(ExcelApp, aProj.Weeks[iWeek], '颜色', iDateQtyCount, aProj.Weeks[iWeek].FColors, aColors, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.Weeks[iWeek], '制式', iDateQtyCount, aProj.Weeks[iWeek].FVers, aVers, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.Weeks[iWeek], '容量', iDateQtyCount, aProj.Weeks[iWeek].FCaps, aCaps, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.Weeks[iWeek], '整机/裸机', iDateQtyCount, aProj.Weeks[iWeek].FFGs, aFGs, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.Weeks[iWeek], '包装', iDateQtyCount, aProj.Weeks[iWeek].FPkgs, aPkgs, irow, irow1, irow2);
                         
      ExcelApp.Cells[irow, 3].Value := 'MPS';
      ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
      end;
      ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';           
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $F2F2F2;
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
      irow := irow + 2;
    end;

    if bSOP then
    begin
      // 写当周  //////////////////////////////////////////////////////////////////////////
      ExcelApp.Cells[irow, 2].Value := aProj.FCurrWeekData.FName;
      ExcelApp.Cells[irow, 3].Value := '需求日期';
      ExcelApp.Cells[irow, 4].Value := '合计';

      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value := aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt;
      end;

      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
      ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

      irow := irow + 1;

      WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '颜色', iDateQtyCount, aProj.FCurrWeekData.FColors, aColors, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '制式', iDateQtyCount, aProj.FCurrWeekData.FVers, aVers, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '容量', iDateQtyCount, aProj.FCurrWeekData.FCaps, aCaps, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '整机/裸机', iDateQtyCount, aProj.FCurrWeekData.FFGs, aFGs, irow, irow1, irow2);
      WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '包装', iDateQtyCount, aProj.FCurrWeekData.FPkgs, aPkgs, irow, irow1, irow2);
                         
      ExcelApp.Cells[irow, 3].Value := 'MPS';
      ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
      for iDate := 0 to iDateQtyCount - 1 do
      begin
        icol := iDate + 5;
        ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
      end;
      ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';           
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $F2F2F2;
      ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
      irow := irow + 2;
    end;

    ///  下面是比较  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    for iDate := 0 to iDateQtyCount - 1 do
    begin
      ExcelApp.Cells[irow, iDate + 5].Value := 'WK' + IntToStr(WeekOf(aProj.Weeks[0].Colors[0].DateQtys[iDate]^.dt));
    end;
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框    
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
  
    irow := irow + 1;

    WriteWF11Compares(ExcelApp, aProj, aColors, '颜色', iDateQtyCount, irow, LineDatasGetter4Color, bSOP);
    WriteWF11Compares(ExcelApp, aProj, aVers, '制式', iDateQtyCount, irow, LineDatasGetter4Ver, bSOP);
    WriteWF11Compares(ExcelApp, aProj, aCaps, '容量', iDateQtyCount, irow, LineDatasGetter4Cap, bSOP);
    WriteWF11Compares(ExcelApp, aProj, aFGs, '整机/裸机', iDateQtyCount, irow, LineDatasGetter4FG, bSOP);
    WriteWF11Compares(ExcelApp, aProj, aPkgs, '包装', iDateQtyCount, irow, LineDatasGetter4Pkg, bSOP);
      
    WriteWF11ComparesTTL(ExcelApp, aProj, aColors, '总量', iDateQtyCount, irow, LineDatasGetter4Color, bSOP);

    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iDateQtyCount + 4]].Font.Name := 'Arial';
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iDateQtyCount + 4]].Font.Size := 9;

  finally
    aColors.Free;
    aVers.Free;
    aCaps.Free;
    aFGs.Free;
    aPkgs.Free;
  end;
end;

procedure TfrmWaterfall.WriteWF01(ExcelApp: Variant; aProj: TProjData);
var
  irow: Integer;
  icol: Integer;
  irow1, irow2: Integer; 
  iDate: Integer; 
  iDateQtyCount: Integer;
  aColors: TStringList;
  aVers: TStringList;
  aCaps: TStringList;
  aFGs: TStringList;
  aPkgs: TStringList; 
begin
  ExcelApp.Columns[1].ColumnWidth := 1;  
  ExcelApp.Columns[2].ColumnWidth := 18; 
  ExcelApp.Columns[3].ColumnWidth := 16;   
  ExcelApp.Columns[4].ColumnWidth := 11;
 
  aColors := TStringList.Create;
  aVers := TStringList.Create;
  aCaps := TStringList.Create;
  aFGs := TStringList.Create;
  aPkgs := TStringList.Create;

  GetAllKeys(aProj.FCurrWeekData.FColors, aColors);
  GetAllKeys(aProj.FCurrWeekData.FVers, aVers);
  GetAllKeys(aProj.FCurrWeekData.FCaps, aCaps);
  GetAllKeys(aProj.FCurrWeekData.FFGs, aFGs);
  GetAllKeys(aProj.FCurrWeekData.FPkgs, aPkgs);

  try

    iDateQtyCount := aProj.FCurrWeekData.Colors[0].DateQtyCount;
    irow := 2;
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      ExcelApp.Cells[irow, iDate + 5].Value := 'WK' + IntToStr(WeekOf(aProj.FCurrWeekData.Colors[0].DateQtys[iDate]^.dt));
      ExcelApp.Columns[iDate + 5].ColumnWidth := 11;
    end;
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框    
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;

    irow := irow + 1;

    // 写当周  //////////////////////////////////////////////////////////////////////////
    ExcelApp.Cells[irow, 2].Value := aProj.FCurrWeekData.FName;
    ExcelApp.Cells[irow, 3].Value := '需求日期';
    ExcelApp.Cells[irow, 4].Value := '合计';

    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      ExcelApp.Cells[irow, icol].Value := aProj.FCurrWeekData.Colors[0].DateQtys[iDate]^.dt;
    end;

    ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
    ExcelApp.Range[ExcelApp.Cells[irow, 2], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框

    irow := irow + 1;

    WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '颜色', iDateQtyCount, aProj.FCurrWeekData.FColors, aColors, irow, irow1, irow2);
    WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '制式', iDateQtyCount, aProj.FCurrWeekData.FVers, aVers, irow, irow1, irow2);
    WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '容量', iDateQtyCount, aProj.FCurrWeekData.FCaps, aCaps, irow, irow1, irow2);
    WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '整机/裸机', iDateQtyCount, aProj.FCurrWeekData.FFGs, aFGs, irow, irow1, irow2);
    WriteWFKeys(ExcelApp, aProj.FCurrWeekData, '包装', iDateQtyCount, aProj.FCurrWeekData.FPkgs, aPkgs, irow, irow1, irow2);
                         
    ExcelApp.Cells[irow, 3].Value := 'MPS';
    ExcelApp.Cells[irow, 4].Value := '=SUM(' + GetRef(5) + IntToStr(irow) + ':' + GetRef(iDateQtyCount + 4) + IntToStr(irow) + ')';
    for iDate := 0 to iDateQtyCount - 1 do
    begin
      icol := iDate + 5;
      ExcelApp.Cells[irow, icol].Value :=  '=SUM(' + GetRef(icol) + IntToStr(irow1) + ':' + GetRef(icol) + IntToStr(irow2) + ')';
    end;
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';           
    ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $F2F2F2;
    ExcelApp.Range[ExcelApp.Cells[irow, 3], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框
    irow := irow + 2;
  

    ///  下面是比较  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    for iDate := 0 to iDateQtyCount - 1 do
    begin
      ExcelApp.Cells[irow, iDate + 5].Value := 'WK' + IntToStr(WeekOf(aProj.FCurrWeekData.Colors[0].DateQtys[iDate]^.dt));
    end;
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Borders.LineStyle := 1; //加边框    
    ExcelApp.Range[ExcelApp.Cells[irow, 5], ExcelApp.Cells[irow, iDateQtyCount + 4]].Interior.Color := $DBDCF2;
  
    irow := irow + 1;

    WriteWF01Compares(ExcelApp, aProj, aColors, '颜色', iDateQtyCount, irow, LineDatasGetter4Color);
    WriteWF01Compares(ExcelApp, aProj, aVers, '制式', iDateQtyCount, irow, LineDatasGetter4Ver);
    WriteWF01Compares(ExcelApp, aProj, aCaps, '容量', iDateQtyCount, irow, LineDatasGetter4Cap);
    WriteWF01Compares(ExcelApp, aProj, aFGs, '整机/裸机', iDateQtyCount, irow, LineDatasGetter4FG);
    WriteWF01Compares(ExcelApp, aProj, aPkgs, '包装', iDateQtyCount, irow, LineDatasGetter4Pkg);
                  
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iDateQtyCount + 4]].Font.Name := 'Arial';
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow, iDateQtyCount + 4]].Font.Size := 9;

  finally
    aColors.Free;
    aVers.Free;
    aCaps.Free;
    aFGs.Free;
    aPkgs.Free;
  end;
end;

function TfrmWaterfall.WriteWF(ExcelApp, WorkBook: Variant): Boolean;
var
  iProj: Integer;
  aProj: TProjData;
  iSheet: Integer;
begin
  iSheet := 0;
  for iProj := 0 to FProjs.Count - 1 do
  begin
    aProj := TProjData(FProjs[iProj]);

    ///  Detail  /////////////////////////////////////////////////////////// 
    iSheet := iSheet + 1;
    if iSheet > WorkBook.Sheets.Count then
    begin
      ///  Summary  //////////////////////////////////////////////////////////
      WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet - 1]);
    end;             
    ExcelApp.Sheets[iSheet].Name := aProj.FName + '-Detail';
    ExcelApp.Sheets[iSheet].Activate;
    if aProj.WeekCount > 0 then   // 有 WF
    begin
      if aProj.FCurrWeekData.ColorCount > 0 then  // 有 SOP
      begin
        WriteWF11(ExcelApp, aProj, True);
      end
      else     //无 SOP
      begin
        WriteWF11(ExcelApp, aProj, False);
      end;
    end    // 无 WF ， 有Proj， 肯定有SOP
    else
    begin
      if aProj.FCurrWeekData.ColorCount > 0 then  // 有 SOP
      begin
        WriteWF01(ExcelApp, aProj);
      end;
    end;

    iSheet := iSheet + 1;
    if iSheet > WorkBook.Sheets.Count then
    begin
      ///  Summary  //////////////////////////////////////////////////////////
      WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet - 1]);
    end;                                 
    ExcelApp.Sheets[iSheet].Name := aProj.FName + '-Summary';
    ExcelApp.Sheets[iSheet].Activate;
    if aProj.WeekCount > 0 then   // 有 WF
    begin
      if aProj.FCurrWeekData.ColorCount > 0 then  // 有 SOP
      begin
        WriteWF11Sum(ExcelApp, aProj, True);
      end
      else     //无 SOP
      begin
        WriteWF11Sum(ExcelApp, aProj, False);
      end;
    end    // 无 WF ， 有Proj， 肯定有SOP
    else
    begin  
      if aProj.FCurrWeekData.ColorCount > 0 then  // 有 SOP
      begin
        WriteWF01Sum(ExcelApp, aProj);
      end;
    end;
  end;
  ExcelApp.Sheets[1].Activate;
  Result := True;
end;

procedure TfrmWaterfall.FormCreate(Sender: TObject);
var
  ini: TIniFile;
begin

//  leWaterFall.Text := 'C:\Users\qiujinbo\Desktop\WF\Waterfall-S&OP 2016-12-25(样本).xlsx';
//  leSOPSum.Text := 'C:\Users\qiujinbo\Desktop\WF\M9系列物料需求进度表(2016-12-23 week52).xlsx';

  GetLocaleFormatSettings(0, gFormatSettings);
  gFormatSettings.DateSeparator := '-';

  ini := TIniFile.Create(AppIni);
  try
    leWaterFall.Text := ini.ReadString(self.ClassName, leWaterFall.Name, '');
    leSOPSum.Text := ini.ReadString(self.ClassName, leSOPSum.Name, '');
  finally
    ini.Free;
  end;
end;

procedure TfrmWaterfall.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmWaterfall.FormDestroy(Sender: TObject);    
var
  ini: TIniFile;
begin
  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, leWaterFall.Name, leWaterFall.Text);
    ini.WriteString(self.ClassName, leSOPSum.Name, leSOPSum.Text);
  finally
    ini.Free;
  end;
end;

end.

