unit SOP2PPOrderWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, ComObj, DateUtils,
  ShellAPI, ShlObj, IniFiles, CommUtils;
type        
  TSOPWriter = class
  private
    ffile: string;
    fh: HWND;
    frow: Integer;
    ExcelApp, WorkBook: Variant;
  public
    constructor Create(h: HWND);
    procedure BeginWrite(const sfile, sSheetName: string);
    procedure SOPWriteNumber(const sProj, sVer, sNumber, sColor, sCap, sFG: string;
      dt1, dt2: TDateTime; iQty: Integer; const sMRPArea: string);
    procedure EndWrite;
  end;
     
  TfrmSOP2PPOrder = class(TForm)
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    OpenDialog1: TOpenDialog;
    ToolBar1: TToolBar;
    tbOEMSum: TToolButton;
    ImageList1: TImageList;
    Memo1: TMemo;
    SaveDialog1: TSaveDialog;
    ToolButton2: TToolButton;
    btnDiv: TToolButton;
    ToolButton4: TToolButton;
    tbOEM: TToolButton;
    tbODMSum: TToolButton;
    ToolButton1: TToolButton;
    ToolButton3: TToolButton;
    ToolButton6: TToolButton;
    dtpStart: TDateTimePicker;
    Label1: TLabel;
    GroupBox1: TGroupBox;
    mmoYearOfProj: TMemo;
    procedure btnSOPClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tbOEMSumClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnDivClick(Sender: TObject);
    procedure tbOEMClick(Sender: TObject);
    procedure tbODMSumClick(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FProjs: TStringList;
    FHaveArea: Boolean;
    procedure Clear;
    procedure OpenMDemand(const sfile: string);

    procedure OpenMDemandSum(const sfile: string);
    procedure OpenWDemandSum(const sfile: string);

    procedure OpenODMDemand4Sum(const sfile: string);  
    procedure MakeODMDemand4Sum(const sfile: string);
     
    procedure OpenOEMDemand(const sfile: string; aSOPWriter: TSOPWriter);

    procedure MakeMDemand(const sfile: string);


    procedure WriteToFile(const sfile:string; lstNumber: TList);

    procedure LoadIni;

    procedure SaveIni;

  public
    { Public declarations }
    class procedure ShowForm;
  end;

var
  gFormatSettings: TFormatSettings;

  aFMonths: TStringList;


implementation

{$R *.dfm}

class procedure TfrmSOP2PPOrder.ShowForm;
var
  frmSOP2PPOrder: TfrmSOP2PPOrder;
begin
  frmSOP2PPOrder := TfrmSOP2PPOrder.Create(nil);
  frmSOP2PPOrder.ShowModal;
end;
             
const
  xlCenter = -4108;
  
type 
  TSOPLine = packed record
    Month: string;
    Ver: string;
    Number: string;
    Color: string;
    Cap: string;
    Qty: Integer;
  end;
  PSOPLine = ^TSOPLine;

  TColRange = packed record
    col1: Integer;
    col2: Integer;         
    FDates1: array of TDateTime;
    FDates2: array of TDateTime;
  end;
  PColRange = ^TColRange;

  TCatData = class
  private
    FMonth: string;
    FCats: TStringList;
    procedure Clear;
  public
    constructor Create(const sMonth: string);
    destructor Destroy; override;
    procedure Add(const skey: string; iQty: Integer);
  end;

  TNumberQty = packed record
    sProj: string;
    sMonth: string;
    sVer: string;
    sCap: string;
    sColor: string;
    sFG: string;
    iQty: Integer;
  end;
  PNumberQty = ^TNumberQty;

  TVerData = class
  private
    FVer: string;
    FSums: TStringList;
    FColors: TList;
    FCaps: TList;
    FFGs: TList;
    procedure Clear;
  public
    constructor Create(const sVer: string; slMonth: TStringList);
    destructor Destroy; override;
    procedure AddNumber(iMonth: Integer; const sColor, sCap, sFG: string; iQty:  Integer);
    function GetCatStr(sl: TStringList): string;
    function GetLocStr(iSum: Integer; sl: TStringList): string;
  end;
  
  TProjData = class
  private
    FVers: TStringList;
    FProj: string;
    procedure Clear;
  public
    constructor Create(const sproj: string);
    destructor Destroy; override;
    function AddVer(const sVer: string; slMonth: TStringList): TVerData;
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

function CellMerged(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer): Boolean;
var
  vma1, vma2: Variant;
  sAddress1, sAddress2: string;
begin
  vma1 := ExcelApp.Cells[irow1, icol1].MergeArea;
  vma2 := ExcelApp.Cells[irow2, icol2].MergeArea;
  sAddress1 := vma1.Address;
  sAddress2 := vma2.Address;
  Result := sAddress1 = sAddress2;
end;

function IndexOfYearOfProj( sl: TStringList; const sProj: string ): Integer;
var
  i: Integer;
  s: string;
begin
  Result := -1;
  for i := 0 to sl.Count - 1 do
  begin
    s := sl.Names[i];
    if Pos(s, sProj) > 0 then
    begin
      Result := i;
      Break;
    end;
  end;
end;

{ TSOPWriter }

constructor TSOPWriter.Create(h: HWND);
begin
  fh := h;
end;

procedure TSOPWriter.BeginWrite(const sfile, sSheetName: string);
begin
  ffile := sfile;
  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(fh, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;

  while WorkBook.Sheets.Count > 1 do
  begin
    WorkBook.Sheets[2].Delete;
  end;

  WorkBook.Sheets[1].Name := sSheetName;
  

  frow := 1;


  ExcelApp.Cells[frow, 1].Value := '编号*';
  ExcelApp.Cells[frow, 2].Value := '需求类型*';
  ExcelApp.Cells[frow, 3].Value := '物料长编码*';
  ExcelApp.Cells[frow, 4].Value := '单位*';
  ExcelApp.Cells[frow, 5].Value := '数量*';
  ExcelApp.Cells[frow, 6].Value := '预测开始日期*';
  ExcelApp.Cells[frow, 7].Value := '预测截止日期*';
                                                    
  ExcelApp.Cells[frow, 8].Value := '均化周期类型*';

  ExcelApp.Cells[frow, 9].Value := '源单类型*';
  ExcelApp.Cells[frow, 10].Value := '源单号*';
  ExcelApp.Cells[frow, 11].Value := '源单行号*';
  ExcelApp.Cells[frow, 12].Value := '备注*';
  ExcelApp.Cells[frow, 13].Value := '备注2';
  ExcelApp.Cells[frow, 14].Value := 'MRP区域';

  ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 12]].HorizontalAlignment := xlCenter; //居中

  frow := frow + 1;

end;

procedure TSOPWriter.SOPWriteNumber(const sProj,  sVer, sNumber, sColor, sCap, sFG: string;
  dt1, dt2: TDateTime; iQty: Integer; const sMRPArea: string);
begin
  if iQty = 0 then Exit;

  ExcelApp.Cells[frow, 1].Value := '';
  if Pos('售后', sVer) > 0 then
  begin
    ExcelApp.Cells[frow, 2].Value := '售后';
  end
  else if Pos('海外', sVer) > 0 then
  begin
    ExcelApp.Cells[frow, 2].Value := '海外量产';
  end
  else
  begin
    ExcelApp.Cells[frow, 2].Value := '国内量产';
  end;

  if sNumber = '' then
  begin
    ExcelApp.Cells[frow, 3].Value := sProj + ',' + sVer + ',' + sColor + ',' + sCap;
  end
  else
  begin
    ExcelApp.Cells[frow, 3].Value := sNumber;  
  end;
  ExcelApp.Cells[frow, 4].Value := 'PCS';
  ExcelApp.Cells[frow, 5].Value := iQty;
  ExcelApp.Cells[frow, 6].Value := dt1;
  ExcelApp.Cells[frow, 7].Value := dt2;
                                       
  ExcelApp.Cells[frow, 8].Value := '不均化';

  ExcelApp.Cells[frow, 9].Value := '';
  ExcelApp.Cells[frow, 10].Value := '';
  ExcelApp.Cells[frow, 11].Value := '';
  ExcelApp.Cells[frow, 12].Value := sProj;
  ExcelApp.Cells[frow, 13].Value := sVer;
  ExcelApp.Cells[frow, 14].Value := sMRPArea;

  ExcelApp.Columns[1].ColumnWidth := 16;
  ExcelApp.Columns[3].ColumnWidth := 16;
  ExcelApp.Columns[6].ColumnWidth := 13;
  ExcelApp.Columns[7].ColumnWidth := 13;
  ExcelApp.Columns[12].ColumnWidth := 16;
  ExcelApp.Columns[13].ColumnWidth := 16;   
  ExcelApp.Columns[14].ColumnWidth := 16;

  frow := frow + 1;
end;

procedure TSOPWriter.EndWrite;
begin
  ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[frow-1, 14]].Borders.LineStyle := 1; //加边框

  WorkBook.SaveAs(ffile);
  ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

  WorkBook.Close;
  ExcelApp.Quit; 
end;


{ TCatData }

constructor TCatData.Create(const sMonth: string);
begin
  FMonth := sMonth;
  FCats := TStringList.Create;
end;

destructor TCatData.Destroy;
begin
  Clear;
  FCats.Free;
  inherited;
end;

procedure TCatData.Clear;
begin
  FCats.Clear;
end;

procedure TCatData.Add(const skey: string; iQty: Integer);
var
  idx: Integer;
  iSum: Integer;
begin
  idx := FCats.IndexOf(skey);
  if idx < 0 then
  begin
    idx := FCats.AddObject(skey, TObject(0));
  end;
  iSum := Integer(FCats.Objects[idx]);
  iSum := iSum + iQty;
  FCats.Objects[idx] := TObject(iSum);
end;

{ TVerData }

constructor TVerData.Create(const sVer: string; slMonth: TStringList);
var
  i: Integer;
begin
  FVer := sVer;
  FSums := TStringList.Create;
  FColors := TList.Create;
  FCaps := TList.Create;
  FFGs := TList.Create;
  
  for i := 0 to slMonth.Count - 1 do
  begin
    FSums.AddObject(slMonth[i], TObject(0)); 
    FColors.Add(TStringList.Create);
    FCaps.Add(TStringList.Create);
    FFGs.Add(TStringList.Create);
  end;
end;

destructor TVerData.Destroy;
begin
  Clear;
  FSums.Free;
  FColors.Free;
  FCaps.Free;
  FFGs.Free;
  inherited;
end;

procedure TVerData.Clear;
var
  i: Integer;
  sl: TStringList;
begin
  FSums.Clear;
  
  for i := 0 to FColors.Count - 1 do
  begin
    sl := TStringList(FColors[i]);
    sl.Free;
  end;
  FColors.Clear;   
  
  for i := 0 to FCaps.Count - 1 do
  begin
    sl := TStringList(FCaps[i]);
    sl.Free;
  end;
  FCaps.Clear;
  
  for i := 0 to FFGs.Count - 1 do
  begin
    sl := TStringList(FFGs[i]);
    sl.Free;
  end;
  FFGs.Clear;
end;

procedure AddNumberToKey(const skey: string; sl: TStringList; iQty: Integer);
var
  idx: Integer;    
  iSum: Integer;
begin
  idx := sl.IndexOf(skey);
  if idx < 0 then
  begin
    idx := sl.AddObject(skey, TObject(0));
  end;
  iSum := Integer(sl.Objects[idx]);
  iSum := iSum + iQty;
  sl.Objects[idx] := TObject(iSum);
end;

procedure TVerData.AddNumber(iMonth: Integer; const sColor, sCap, sFG: string; iQty:  Integer); 
begin
  FSums.Objects[iMonth] := TObject(Integer(FSums.Objects[iMonth]) + iQty);
  AddNumberToKey(sColor, TStringList(FColors[iMonth]), iQty);
  AddNumberToKey(sCap, TStringList(FCaps[iMonth]), iQty);
  AddNumberToKey(sFG, TStringList(FFGs[iMonth]), iQty); 
end;

function TVerData.GetCatStr(sl: TStringList): string;
var
  iCat: Integer; 
begin
  Result := '';
  for iCat := 0 to sl.Count - 1 do
  begin
    if iCat > 0 then
    begin
      Result := Result + ' : ';
    end;
    Result := Result + sl[iCat];
  end;
end;

function TVerData.GetLocStr(iSum: Integer; sl: TStringList): string;
var
  iCat: Integer;
  iValue: Integer;
  s: string;
begin
  Result := '';

  if iSum = 0 then Exit;
  
  for iCat := 0 to sl.Count - 1 do
  begin
    if iCat > 0 then
    begin
      Result := Result + ' : ';
    end;
    iValue := Integer(sl.Objects[iCat]);
    s := FormatFloat('##0.##', iValue * 10 / iSum);
    Result := Result + s;
  end;
end;

{ TProjData }

constructor TProjData.Create(const sproj: string);
begin
  FProj := sproj;
  FVers := TStringList.Create;
end;

destructor TProjData.Destroy;
begin
  Clear;
  FVers.Free;
end;

procedure TProjData.Clear;
var
  i: Integer; 
  aVerData: TVerData;
begin  
  for i := 0 to FVers.Count - 1 do
  begin
    aVerData := TVerData(FVers.Objects[i]);
    aVerData.Free;
  end;
  FVers.Clear;
end;

function TProjData.AddVer(const sVer: string; slMonth: TStringList): TVerData;
var
  idx: Integer;
  aVerData: TVerData;
begin
  idx := FVers.IndexOf(sVer);
  if idx < 0 then
  begin
    aVerData := TVerData.Create(sVer, slMonth);
    idx := FVers.AddObject(sVer, aVerData);
  end;
  Result := TVerData(FVers.Objects[idx]);
end;

function GetRoamingFolder: string;  
var
  lpszPath: array[0..260] of Char; 
begin
  ZeroMemory(@lpszPath, SizeOf(lpszPath));
  SHGetSpecialFolderPath(0, lpszPath, CSIDL_APPDATA, FALSE);
  Result := PChar(@lpszPath) + '\';
end;

function GetIniPath: string;
begin
  Result := GetRoamingFolder + ChangeFileExt( ExtractFileName(Application.ExeName), '.ini');
end;

procedure TfrmSOP2PPOrder.LoadIni;
var
  sAppIni: string;
  ini: TIniFile;
begin
  sAppIni := GetIniPath;
  mmoYearOfProj.Lines.Clear;
  ini := TIniFile.Create(sAppIni);
  ini.ReadSectionValues(self.ClassName, mmoYearOfProj.Lines);
  ini.Free;
end;

procedure TfrmSOP2PPOrder.SaveIni;
var
  sAppIni: string;
  ini: TIniFile;
  i: Integer;
begin
  sAppIni := GetIniPath; 
  ini := TIniFile.Create(sAppIni);
  try
    ini.EraseSection(self.ClassName);
    for i := 0 to mmoYearOfProj.Lines.Count - 1 do
    begin
      ini.WriteString(self.ClassName, mmoYearOfProj.Lines.Names[i], mmoYearOfProj.Lines.ValueFromIndex[i]);
    end;
  finally
    ini.Free;
  end;
end;

procedure TfrmSOP2PPOrder.FormCreate(Sender: TObject);
begin
  GetLocaleFormatSettings(GetThreadLocale, gFormatSettings);
  gFormatSettings.DateSeparator := '-';

  Memo1.Clear;

  FProjs := TStringList.Create;

  dtpStart.Date := Now;

  LoadIni;
end;
      
procedure TfrmSOP2PPOrder.Clear;
var
  i: Integer;
  aProjData: TProjData;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs.Objects[i]);
    aProjData.Free;
  end;
  FProjs.Clear;
end;

procedure TfrmSOP2PPOrder.FormDestroy(Sender: TObject);
begin
  Clear;
  FProjs.Free;

  SaveIni;
end;

procedure TfrmSOP2PPOrder.btnSOPClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  leSOP.Text := OpenDialog1.FileName;
end;

function GetCol(irow: Integer; ExcelApp: Variant; const scol: string): Integer;
var
  icol: Integer;
  svalue: string;
begin
  Result := 0;
  for icol := 1 to 100 do
  begin
    svalue := ExcelApp.Cells[irow, icol].Value;
    if svalue = scol then
    begin
      Result := icol;
      Break;
    end;
  end;
  if Result = 0 then
  begin
    raise Exception.Create('列 "' + scol + '" 不存在');
  end;
end;
 
procedure TfrmSOP2PPOrder.MakeMDemand(const sfile: string);
const
  CINumber = 1;
  CIName = 2;
  CIDate = 3;
  CIQty = 4;
  CIProj = 5;
var
  iProj: Integer;  
  irow: Integer;
  ExcelApp, WorkBook: Variant;

  aProjData: TProjData;
  aVerData: TVerData;
  iMonth: Integer;
  iVer: Integer; 
  icol: Integer; 
  iSum: Integer; 
begin

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;
  WorkBook.Sheets[1].Name := 'OEM';
  
  try
    irow := 1;

    // 填写标题的月份

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '型号';
    ExcelApp.Cells[irow, 2].Value := '制式颜色容量占比';
    icol := 3;

    for iMonth := 0 to aFMonths.Count -1 do
    begin
      ExcelApp.Cells[irow, icol].Value := aFMonths[iMonth];
      ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow, icol + 1]].MergeCells := True;

      ExcelApp.Columns[1].ColumnWidth := 5;
      ExcelApp.Columns[2].ColumnWidth := 38;
      ExcelApp.Columns[icol].ColumnWidth := 10;
      ExcelApp.Columns[icol + 1].ColumnWidth := 30;

      icol := icol + 2;
    end;

    irow := 2;
    for iProj := 0 to  FProjs.Count - 1 do
    begin
      aProjData := TProjData(FProjs.Objects[iProj]);

      //开始填写项目数据
      ExcelApp.Cells[irow, 1].Value := aProjData.FProj;
      ExcelApp.Range[ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow + aProjData.FVers.Count * 4 - 1, 1]].MergeCells := True;
      for iVer := 0 to aProjData.FVers.Count - 1 do
      begin                                      
        aVerData := TVerData(aProjData.FVers.Objects[iVer]);

        ExcelApp.Cells[irow, 2].Value := aVerData.FVer;

        icol := 3;
        for iMonth := 0 to aVerData.FSums.Count -1 do
        begin
          if iMonth = 0 then
          begin
            ExcelApp.Cells[irow + 1, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FCaps[0]));
            ExcelApp.Cells[irow + 2, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FColors[0]));
            ExcelApp.Cells[irow + 3, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FFGs[0]));
          end;

          iSum := Integer(aVerData.FSums.Objects[iMonth]);
          if iSum > 0 then
          begin
            ExcelApp.Cells[irow, icol].Value := iSum;
          end;
          ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow + 3, icol]].MergeCells := True;

          ExcelApp.Cells[irow + 1, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FCaps[iMonth]));
          ExcelApp.Cells[irow + 2, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FColors[iMonth]));
          ExcelApp.Cells[irow + 3, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FFGs[iMonth]));

          icol := icol + 2;
        end;
        for iMonth := aVerData.FSums.Count to aFMonths.Count - 1 do
        begin               
          ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow + 3, icol]].MergeCells := True;
          icol := icol + 2;
        end;
        irow := irow + 4;
      end;
    end;

    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 2 + aFMonths.Count * 2]].Borders.LineStyle := 1; //加边框
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 2 + aFMonths.Count * 2]].HorizontalAlignment := xlCenter; //居中

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存


  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

end;

procedure TfrmSOP2PPOrder.MakeODMDemand4Sum(const sfile: string);
const
  CINumber = 1;
  CIName = 2;
  CIDate = 3;
  CIQty = 4;
  CIProj = 5;
var
  iProj: Integer;  
  irow: Integer;
  ExcelApp, WorkBook: Variant;

  aProjData: TProjData;
  aVerData: TVerData;
  iMonth: Integer;
  iVer: Integer; 
  icol: Integer; 
  iSum: Integer; 
begin

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;
  WorkBook.Sheets[1].Name := 'ODM';
  
  try
    irow := 1;

    // 填写标题的月份

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '型号';
    ExcelApp.Cells[irow, 2].Value := '制式颜色容量占比';
    icol := 3;

    for iMonth := 0 to aFMonths.Count -1 do
    begin
      ExcelApp.Cells[irow, icol].Value := aFMonths[iMonth];
      ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow, icol + 1]].MergeCells := True;

      ExcelApp.Columns[1].ColumnWidth := 5;
      ExcelApp.Columns[2].ColumnWidth := 38;
      ExcelApp.Columns[icol].ColumnWidth := 10;
      ExcelApp.Columns[icol + 1].ColumnWidth := 30;

      icol := icol + 2;
    end;

    irow := 2;
    for iProj := 0 to  FProjs.Count - 1 do
    begin
      aProjData := TProjData(FProjs.Objects[iProj]);

      //开始填写项目数据
      ExcelApp.Cells[irow, 1].Value := aProjData.FProj;
      ExcelApp.Range[ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow + aProjData.FVers.Count * 4 - 1, 1]].MergeCells := True;
      for iVer := 0 to aProjData.FVers.Count - 1 do
      begin                                      
        aVerData := TVerData(aProjData.FVers.Objects[iVer]);

        ExcelApp.Cells[irow, 2].Value := aVerData.FVer;

        icol := 3;
        for iMonth := 0 to aVerData.FSums.Count -1 do
        begin
          if iMonth = 0 then
          begin
            ExcelApp.Cells[irow + 1, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FCaps[0]));
            ExcelApp.Cells[irow + 2, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FColors[0]));
            ExcelApp.Cells[irow + 3, 2].Value := aVerData.GetCatStr(TStringList(aVerData.FFGs[0]));
          end;

          iSum := Integer(aVerData.FSums.Objects[iMonth]);
          if iSum > 0 then
          begin
            ExcelApp.Cells[irow, icol].Value := iSum;
          end;
          ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow + 3, icol]].MergeCells := True;

          ExcelApp.Cells[irow + 1, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FCaps[iMonth]));
          ExcelApp.Cells[irow + 2, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FColors[iMonth]));
          ExcelApp.Cells[irow + 3, icol + 1].Value := aVerData.GetLocStr(iSum, TStringList(aVerData.FFGs[iMonth]));

          icol := icol + 2;
        end;
        for iMonth := aVerData.FSums.Count to aFMonths.Count - 1 do
        begin               
          ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow + 3, icol]].MergeCells := True;
          icol := icol + 2;
        end;
        irow := irow + 4;
      end;
    end;

    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 2 + aFMonths.Count * 2]].Borders.LineStyle := 1; //加边框
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 2 + aFMonths.Count * 2]].HorizontalAlignment := xlCenter; //居中

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存


  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

end;

procedure GetMonthList(ExcelApp: Variant; irow: Integer; slMonth: TStringList);
var
  icol: Integer;
  sMonth: string;
begin
  icol := 3;
  sMonth := ExcelApp.Cells[irow, icol].Value;
  while sMonth <> '' do
  begin
    slMonth.Add(sMonth);
    icol := icol + 2;    
    sMonth := ExcelApp.Cells[irow, icol].Value;
  end;
end;

function GetVerCount(ExcelApp: Variant; irow1, irow2: Integer): Integer;
var 
  irowa, irowb: Integer;
begin
  Result := 0;
  irowa := irow1;
  irowb := irow1 + 1;
  while True do
  begin                
    Result := Result + 1;
    while CellMerged(ExcelApp, irowa, 3, irowb, 3) do
    begin
      irowb := irowb + 1;
    end;
    if irowb > irow2 then Break;
    irowa := irowb;
    irowb := irowb + 1;
  end;
end;


procedure TfrmSOP2PPOrder.WriteToFile(const sfile:string; lstNumber: TList);
var
  iCount: Integer;
  aNumberQtyPtr: PNumberQty;
  irow: Integer;    
  ExcelApp, WorkBook: Variant;
begin

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  except
    on e: Exception do
    begin
      MessageBox(Handle, PChar(e.Message), '金蝶提示', 0);
      Exit;
    end;
  end;

  WorkBook := ExcelApp.WorkBooks.Add;
  WorkBook.Sheets[1].Name := 'OEM';
  
  try
    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '项目';
    ExcelApp.Cells[irow, 2].Value := '月份';
    ExcelApp.Cells[irow, 3].Value := '版本';
    ExcelApp.Cells[irow, 4].Value := '容量';
    ExcelApp.Cells[irow, 5].Value := '颜色';
    ExcelApp.Cells[irow, 6].Value := '整机/裸机';
    ExcelApp.Cells[irow, 7].Value := '数量';
    ExcelApp.Range[ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow, 7]].HorizontalAlignment := xlCenter; //居中

    irow := irow + 1;
    for iCount := 0 to lstNumber.Count - 1 do
    begin
      aNumberQtyPtr := PNumberQty(lstNumber[iCount]);
      ExcelApp.Cells[irow + iCount, 1].Value := aNumberQtyPtr^.sProj;
      ExcelApp.Cells[irow + iCount, 2].Value := aNumberQtyPtr^.sMonth;
      ExcelApp.Cells[irow + iCount, 3].Value := aNumberQtyPtr^.sVer;
      ExcelApp.Cells[irow + iCount, 4].Value := aNumberQtyPtr^.sCap;
      ExcelApp.Cells[irow + iCount, 5].Value := aNumberQtyPtr^.sColor;
      ExcelApp.Cells[irow + iCount, 6].Value := aNumberQtyPtr^.sFG;
      ExcelApp.Cells[irow + iCount, 7].Value := aNumberQtyPtr^.iQty;
    end;
     
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[lstNumber.Count + 1, 7]].Borders.LineStyle := 1; //加边框
    
    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
 
       
  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;
  
end;


procedure TfrmSOP2PPOrder.OpenMDemandSum(const sfile: string);
var
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  irow1, irow2: Integer;
  sProj: string;
  iVerCount: Integer;
  slMonth: TStringList;
  iVer: Integer;
  iMonth: Integer;
  sVer: string;
  sCaps: string;
  sColors: string;
  sFGs: string;
  sVerSum: string;
  iVerSum: Integer;
  sCapRate: string;
  sColorRate: string;
  sFGRate: string;    
  slCap: TStringList;
  slColor: TStringList;
  slFG: TStringList;
  slCapRate: TStringList;
  slColorRate: TStringList;
  slFGRate: TStringList;
  iCap: Integer;
  iColor: Integer;
  iFG: Integer;
  iCapSum: Integer;
  iColorSum: Integer;
  iFGValue: Integer;
  lstNumber: TList;
  iCount: Integer;
  aNumberQtyPtr: PNumberQty;
begin
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'SOPNumber-' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;

  lstNumber := TList.Create;

  slCap := TStringList.Create;
  slColor := TStringList.Create;
  slFG := TStringList.Create;

  slCapRate := TStringList.Create;
  slColorRate := TStringList.Create;
  slFGRate := TStringList.Create;

  slMonth := TStringList.Create;

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    try
      WorkBook := ExcelApp.WorkBooks.Open(sFile);
      try
        irow := 1;
        GetMonthList(ExcelApp, irow, slMonth);
        irow := 2;
        sProj := ExcelApp.Cells[irow, 1].Value;
        while sProj <> '' do
        begin
          irow1 := irow;
      
          irow := irow + 1;
          while CellMerged(ExcelApp, irow1, 1, irow, 1) do
          begin
            irow := irow + 1;
          end;
          irow2 := irow - 1;

          iVerCount := GetVerCount(ExcelApp, irow1, irow2);

          for iVer := 0 to iVerCount - 1 do
          begin
            sVer    := ExcelApp.Cells[irow1 + iVer * 4,     2].Value;

            sCaps   := ExcelApp.Cells[irow1 + iVer * 4 + 1, 2].Value;
            sColors := ExcelApp.Cells[irow1 + iVer * 4 + 2, 2].Value;
            sFGs    := ExcelApp.Cells[irow1 + iVer * 4 + 3, 2].Value;

            slCap.Text := StringReplace(sCaps, ' : ', #13#10, [rfReplaceAll]);   
            slColor.Text := StringReplace(sColors, ' : ', #13#10, [rfReplaceAll]);    
            slFG.Text := StringReplace(sFGs, ' : ', #13#10, [rfReplaceAll]);

            for iMonth := 0 to slMonth.Count - 1 do
            begin
              sVerSum := ExcelApp.Cells[irow1 + iVer * 4, 3 + iMonth * 2].Value;
              if sVerSum <> '' then
              begin
                iVerSum    := ExcelApp.Cells[irow1 + iVer * 4,     3 + iMonth * 2].Value;
                sCapRate   := ExcelApp.Cells[irow1 + iVer * 4 + 1, 4 + iMonth * 2].Value;
                sColorRate := ExcelApp.Cells[irow1 + iVer * 4 + 2, 4 + iMonth * 2].Value;
                sFGRate    := ExcelApp.Cells[irow1 + iVer * 4 + 3, 4 + iMonth * 2].Value;

                slCapRate.Text := StringReplace(sCapRate, ' : ', #13#10, [rfReplaceAll]);
                slColorRate.Text := StringReplace(sColorRate, ' : ', #13#10, [rfReplaceAll]);
                slFGRate.Text := StringReplace(sFGRate, ' : ', #13#10, [rfReplaceAll]);

                for iCap := 0 to slCapRate.Count - 1 do
                begin
                  iCapSum := Round(iVerSum * StrToFloat(slCapRate[iCap]) / 10);
                  for iColor := 0 to slColorRate.Count - 1 do
                  begin
                    iColorSum := Round(iCapSum * StrToFloat(slColorRate[iColor]) / 10);
                    for iFG := 0 to slFGRate.Count - 1 do
                    begin
                      iFGValue := Round(iColorSum * StrToFloat(slFGRate[iFG]) / 10);

                      aNumberQtyPtr := New(PNumberQty);
                      
                      aNumberQtyPtr^.sProj := sProj;
                      aNumberQtyPtr^.sMonth := slMonth[iMonth];
                      aNumberQtyPtr^.sVer := sVer;
                      aNumberQtyPtr^.sCap := slCap[iCap];
                      aNumberQtyPtr^.sColor := slColor[iColor];
                      aNumberQtyPtr^.sFG := slFG[iFG];
                      aNumberQtyPtr^.iQty := iFGValue;

                      lstNumber.Add(aNumberQtyPtr);
                    end;
                  end;
                end;
              end;
            end;
          end;
        
          sProj := ExcelApp.Cells[irow, 1].Value;
        end;
      finally
        ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook.Close;
      end
    finally
      ExcelApp.Visible := True;
      ExcelApp.Quit; 
    end;

    WriteToFile(SaveDialog1.FileName, lstNumber);

  finally
    slCap.Free;
    slColor.Free;
    slFG.Free;

    slCapRate.Free;
    slColorRate.Free;
    slFGRate.Free;

    for iCount := 0 to lstNumber.Count - 1 do
    begin
      aNumberQtyPtr := PNumberQty(lstNumber[iCount]);
      Dispose(aNumberQtyPtr);
    end;
    lstNumber.Free;

    slMonth.Free;
  end;
end;
         
procedure TfrmSOP2PPOrder.OpenWDemandSum(const sfile: string);
var
  ExcelApp, WorkBook: Variant;

  procedure GetStartCell(var pirow, picol: Integer);
  var
    ir, ic: Integer;
    s: string;
  begin
    pirow := 0;
    picol := 0;
    for ir := 1 to 10 do
    begin
      for ic := 1 to 10 do
      begin
        s := ExcelApp.Cells[ir, ic].Value;
        if s = '型号' then
        begin
          pirow := ir;
          picol := ic;
          Exit;
        end;
      end;
    end;
  end;  

var
  irow0, icol0: Integer;
  irow, icol: Integer;
  irow1, irow2: Integer;
  icol1, icol2: Integer;
  sMonth: string;
  iMonth: Integer;
  iWeekCol: Integer;
  sWeek: string;
  sWeekDate: string;
  slMonth: TStringList;
  slWeek: TStringList;
  sProj: string;
begin
//  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
//  SaveDialog1.FilterIndex := 0;
//  SaveDialog1.DefaultExt := '.xlsx';
//  SaveDialog1.FileName := 'SOPNumber-' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
//  if not SaveDialog1.Execute then Exit;

  SaveDialog1.FileName := 'SOPNumber-' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
                           
  slMonth := TStringList.Create;

  try
    ExcelApp := CreateOleObject('Excel.Application' );
    ExcelApp.Visible := False;
    ExcelApp.Caption := '应用程序调用 Microsoft Excel';
    try
      WorkBook := ExcelApp.WorkBooks.Open(sFile);
      try

        GetStartCell(irow0, icol0);
        if irow0 = 0 then
        begin
          Memo1.Lines.Add('格式不符合');
          Exit;
        end;

        irow := irow0;
        icol := icol0 + 2;

        icol1 := icol;
        icol2 := icol1 + 1;
        
        sMonth := ExcelApp.Cells[irow, icol].Value;
        while sMonth <> '' do
        begin
          if CellMerged(ExcelApp, irow, icol1, irow, icol2) then
          begin
            icol2 := icol2 + 1;
          end;
          icol2 := icol2 - 1;

          slWeek := TStringList.Create;
          slMonth.AddObject(sMonth, slWeek);
          
          // 一个月数据
          for iWeekCol := icol1 to icol2 do
          begin
            sWeekDate := ExcelApp.Cells[irow + 1, iWeekCol].Value;
            sWeek := ExcelApp.Cells[irow + 2, iWeekCol].Value;
            slWeek.AddObject(sWeek + '=' + sWeekDate, TObject(iWeekCol));
          end;

          icol := icol2;
          sMonth := ExcelApp.Cells[irow, icol].Value;
        end;

        irow := irow0 + 3;
        icol := icol0;

        irow1 := irow;
        irow2 := irow1 + 1;

        sProj := ExcelApp.Cells[irow, icol].Value;
        while sProj <> '' do
        begin
          if CellMerged(ExcelApp, irow1, icol, irow2, icol) then
          begin
            irow2 := irow2 + 1;
          end;
          irow := irow2;
          irow2 := irow2 - 1;

          
        
          sProj := ExcelApp.Cells[irow, icol].Value;
        end;
        
      finally
        ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook.Close;
      end
    finally
      ExcelApp.Visible := True;
      ExcelApp.Quit; 
    end;
 
  finally
    for iMonth := 0 to slMonth.Count - 1 do
    begin
      slWeek := TStringList(slMonth.Objects[iMonth]);
      slWeek.Free;
    end;
    slMonth.Free;
  end;
end;

procedure TfrmSOP2PPOrder.OpenMDemand(const sfile: string);
const
  CIVer = 1;
  CINumber = 2;
  CIColor = 3;
  CICap = 4;

  CSVer = '制式';
  CSNumber = '物料编码';
  CSColor = '颜色';
  CSCap = '容量';

  CITitleRow = 1;
  CIDateRow = 2;
  
var
  sWeek: string;   
  irow, icol: Integer;
  irow1, irow2: Integer;
  icol1: Integer;
  ExcelApp, WorkBook: Variant; 
  iSheetCount, iSheet: Integer;
  sproj: string; 

  sVer: string;
  sNumber: string;
  sColor: string;
  sCap: string;
 
  slMonth: TStringList;
  iMonth: Integer; 

  aColRangePtr: PColRange;
  aProjData: TProjData; 
 
  iQty: Integer;

  aVerData: TVerData;
begin
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try
      iSheetCount := ExcelApp.WorkSheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        ExcelApp.WorkSheets[iSheet].Activate;     
        sproj := ExcelApp.WorkSheets[iSheet].Name; 

        irow := 1;
        sVer := ExcelApp.Cells[irow, CIVer].Value;
        sNumber := ExcelApp.Cells[irow, CINumber].Value;
        sColor := ExcelApp.Cells[irow, CIColor].Value;
        sCap := ExcelApp.Cells[irow, CICap].Value;

        if (sVer <> CSVer) or (sNumber <> CSNumber) or (sColor <> CSColor) or (sCap <> CSCap) then
        begin
          Continue;
        end;

        aProjData := TProjData.Create(sproj); 
        FProjs.AddObject(sproj, aProjData);
        slMonth := TStringList.Create;
 
        try
          //判断有多少个月
          icol := CICap + 1;
          icol1 := icol;
          sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          while sWeek <> '' do
          begin
            if UpperCase(Copy(sWeek, 1, 1)) <> 'W' then
            begin
              aColRangePtr := New(PColRange);
              slMonth.AddObject(sWeek, TObject(aColRangePtr));
              aColRangePtr^.col1 := icol1;
              aColRangePtr^.col2 := icol - 1;
              icol1 := icol + 1;
            end;
            
            icol := icol + 1;                       
            sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          end;

          if aFMonths.Count < slMonth.Count then
          begin
            aFMonths.Clear;
            aFMonths.Text := slMonth.Text;
          end;


          irow := CITitleRow + 2; 
          while True do
          begin
            sVer := ExcelApp.Cells[irow, CIVer].Value;

            if sVer = '' then Break; //读取完了
            if CellMerged(ExcelApp, irow, CINumber, irow, CIColor) then Break; //读取完了

            aVerData := aProjData.AddVer(sVer, slMonth);

            sNumber := ExcelApp.Cells[irow, CINumber].Value;
            sColor := ExcelApp.Cells[irow, CIColor].Value;
            sCap := ExcelApp.Cells[irow, CICap].Value;

 
            irow1 := irow;
            irow2 := irow1 + 1;

            while CellMerged(ExcelApp, irow1, CIVer, irow2, CIVer) do
            begin
              irow2 := irow2 + 1;
            end;
            irow2 := irow2 - 1;

            for irow := irow1 to irow2 do
            begin
              sNumber := ExcelApp.Cells[irow, CINumber].Value;   
              sColor := ExcelApp.Cells[irow, CIColor].Value;
              sCap := ExcelApp.Cells[irow, CICap].Value;

              //按月读取数据
              for iMonth := 0 to slMonth.Count - 1 do
              begin
                aColRangePtr := PColRange(slMonth.Objects[iMonth]);
                for icol := aColRangePtr^.col1 to aColRangePtr^.col2 do
                begin
                  iQty := ExcelApp.Cells[irow, icol].Value;
                  aVerData.AddNumber(iMonth, sColor, sCap, '整机', iQty);
                end;
              end;
            end;

            irow := irow2 + 1;
          end;
        finally
          for iMonth := 0 to slMonth.Count - 1 do
          begin
            aColRangePtr := PColRange(slMonth.Objects[iMonth]);
            Dispose(aColRangePtr);
          end;
          slMonth.Free;
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
               
procedure TfrmSOP2PPOrder.OpenODMDemand4Sum(const sfile: string);
const
  CIVer = 1;
  CINumber = 2;
  CIColor = 3;
  CICap = 4;

  CSVer = '制式';
  CSNumber = '物料编码';
  CSColor = '颜色';
  CSCap = '容量';

  CITitleRow = 1;
  CIDateRow = 2;
  
var
  sWeek: string;   
  irow, icol: Integer;
  irow1, irow2: Integer; 
  ExcelApp, WorkBook: Variant; 
  iSheetCount, iSheet: Integer;
  sproj: string; 

  sVer: string;
  sNumber: string;
  sColor: string;
  sCap: string;
 
  slMonth: TStringList;
  iMonth: Integer; 

  aColRangePtr: PColRange;
  aProjData: TProjData; 
 
  iQty: Integer;

  aVerData: TVerData;
begin
  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try
      iSheetCount := ExcelApp.WorkSheets.Count;
      for iSheet := 1 to iSheetCount do
      begin
        ExcelApp.WorkSheets[iSheet].Activate;     
        sproj := ExcelApp.WorkSheets[iSheet].Name; 

        irow := 1;
        sVer := ExcelApp.Cells[irow, CIVer].Value;
        sNumber := ExcelApp.Cells[irow, CINumber].Value;
        sColor := ExcelApp.Cells[irow, CIColor].Value;
        sCap := ExcelApp.Cells[irow, CICap].Value;

        if (sVer <> CSVer) or (sNumber <> CSNumber) or (sColor <> CSColor) or (sCap <> CSCap) then
        begin
          Memo1.Lines.Add('Seet ' + sproj + ' 格式不符合');
          Continue;
        end;

        aProjData := TProjData.Create(sproj); 
        FProjs.AddObject(sproj, aProjData);
        slMonth := TStringList.Create;
 
        try
          //判断有多少个月
          icol := CICap + 1; 
          sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          while sWeek <> '' do
          begin
            if UpperCase(Copy(sWeek, 1, 1)) <> 'W' then
            begin
              while sWeek <> '' do
              begin
                if not CellMerged(ExcelApp, 1, icol, 2, icol) then
                begin          
                  Memo1.Lines.Add('读取月结束 ' + sWeek);
                  Break;
                end;
                aColRangePtr := New(PColRange);
                slMonth.AddObject(sWeek, TObject(aColRangePtr));
                aColRangePtr^.col1 := icol;
                aColRangePtr^.col2 := icol;
                icol := icol + 1;
                sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
              end;
              Break;
            end;
            icol := icol + 1;                       
            sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          end;

          if aFMonths.Count < slMonth.Count then
          begin
            aFMonths.Clear;
            aFMonths.Text := slMonth.Text;
          end;


          irow := CITitleRow + 2; 
          while True do
          begin
            sVer := ExcelApp.Cells[irow, CIVer].Value;

            if sVer = '' then Break; //读取完了
            if CellMerged(ExcelApp, irow, CINumber, irow, CIColor) then Break; //读取完了

            aVerData := aProjData.AddVer(sVer, slMonth);

            sNumber := ExcelApp.Cells[irow, CINumber].Value;
            sColor := ExcelApp.Cells[irow, CIColor].Value;
            sCap := ExcelApp.Cells[irow, CICap].Value;

 
            irow1 := irow;
            irow2 := irow1 + 1;

            while CellMerged(ExcelApp, irow1, CIVer, irow2, CIVer) do
            begin
              irow2 := irow2 + 1;
            end;
            irow2 := irow2 - 1;

            for irow := irow1 to irow2 do
            begin
              sNumber := ExcelApp.Cells[irow, CINumber].Value;   
              sColor := ExcelApp.Cells[irow, CIColor].Value;
              sCap := ExcelApp.Cells[irow, CICap].Value;

              //按月读取数据
              for iMonth := 0 to slMonth.Count - 1 do
              begin
                aColRangePtr := PColRange(slMonth.Objects[iMonth]);
                for icol := aColRangePtr^.col1 to aColRangePtr^.col2 do
                begin
                  iQty := ExcelApp.Cells[irow, icol].Value;
                  aVerData.AddNumber(iMonth, sColor, sCap, '整机', iQty);
                end;
              end;
            end;

            irow := irow2 + 1;
          end;
        finally
          for iMonth := 0 to slMonth.Count - 1 do
          begin
            aColRangePtr := PColRange(slMonth.Objects[iMonth]);
            Dispose(aColRangePtr);
          end;
          slMonth.Free;
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
     
procedure TfrmSOP2PPOrder.tbOEMSumClick(Sender: TObject);
var
  sfile: string;
begin
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'OEM-SOP-Sum-' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;
  sfile := SaveDialog1.FileName;

  Clear;


  aFMonths := TStringList.Create;

//  sfile := 'C:\Users\qiujinbo\Desktop\SOP-PC\sop.xlsx';

  OpenMDemand(leSOP.Text);

  MakeMDemand(sfile);

  aFMonths.Free;

  MessageBox(Handle, '完成', '金蝶提示', 0);
end;

procedure TfrmSOP2PPOrder.btnDivClick(Sender: TObject);
begin
  OpenMDemandSum(leSOP.Text);
  MessageBox(Handle, '完成', '金蝶提示', 0);
end;
  
procedure TfrmSOP2PPOrder.ToolButton3Click(Sender: TObject);
begin
  OpenWDemandSum(leSOP.Text);
  MessageBox(Handle, '完成', '金蝶提示', 0);
end;

function IsWeek(ExcelApp: Variant; iRow, icol: Integer; var sWeek: string): Boolean;
var 
  sNumberFormatlocal: string;
begin  
  if UpperCase(Copy(sWeek, 1, 1)) = 'W' then
  begin
    Result := True;
    Exit;
  end;
 
  sNumberFormatlocal := ExcelApp.Cells[iRow, icol].NumberFormatlocal;
  sNumberFormatlocal := StringReplace(sNumberFormatlocal, '"', '', [rfReplaceAll]);
  if Copy(sNumberFormatlocal, 1, 2) = 'WK' then
  begin
    Result := True;
    Exit;
  end;

  Result := False;
end;

procedure TfrmSOP2PPOrder.OpenOEMDemand(const sfile: string; aSOPWriter: TSOPWriter);
const
//  CIVer = 1;
//  CINumber = 2;
//  CIColor = 3;
//  CICap = 4;
//
//  CSVer = '制式';
//  CSNumber = '物料编码';
//  CSColor = '颜色';
//  CSCap = '容量';

  CITitleRow = 1;
  CIDateRow = 2;
  
var
  sWeek: string;   
  irow, icol: Integer;
  irow1, irow2: Integer;
  icol1: Integer;
  ExcelApp, WorkBook: Variant; 
  iSheetCount, iSheet: Integer;
  sSheet: string;
  sproj: string;

  sVer: string;
  sNumber: string;
  sColor: string;
  sCap: string;
 
  slMonth: TStringList;
  iMonth: Integer; 

  aColRangePtr: PColRange;
//  aProjData: TProjData; 

  iQty: Integer;

//  aVerData: TVerData;

  dt1, dt2: TDateTime;
  sDate, sDate1, sDate2: string;
  sNumberFormatlocal: string;
  dtStartxx: TDateTime;

  slYearOfProj: TStringList;
  dt1Prev: TDateTime;
  
  idxYear: Integer;
  v: Variant;
  s: string;

  stitle4x, stitle5x, stitle8x, stitle9x: string;
  stitle1, stitle2, stitle3, stitle4, stitle5,
    stitle6, stitle7, stitle8, stitle9: string;

  icolMRPArea: Integer;
  iVer: Integer;
  iNumber: Integer;
  iColor: Integer;
  iCap: Integer;
  sMRPArea: string;
begin       
  slYearOfProj := TStringList.Create;

  dtStartxx := myStrToDateTime(FormatDateTime('yyyy-MM-dd', dtpStart.DateTime));

  ExcelApp := CreateOleObject('Excel.Application' );
  ExcelApp.Visible := False;
  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  try
    WorkBook := ExcelApp.WorkBooks.Open(sFile);
    try
      iSheetCount := ExcelApp.WorkSheets.Count;
      for iSheet := 1 to iSheetCount do
      begin           
        if not ExcelApp.WorkSheets[iSheet].Visible then Continue;                            
        sSheet := ExcelApp.WorkSheets[iSheet].Name;
        ExcelApp.WorkSheets[iSheet].Activate;
        sproj := Trim(sSheet);

        irow := 1;

        stitle1 := ExcelApp.Cells[irow, 1].Value;
        stitle2 := ExcelApp.Cells[irow, 2].Value;
        stitle3 := ExcelApp.Cells[irow, 3].Value;
        stitle4 := ExcelApp.Cells[irow, 4].Value;
        stitle5 := ExcelApp.Cells[irow, 5].Value;
        stitle6 := ExcelApp.Cells[irow, 6].Value;
        stitle7 := ExcelApp.Cells[irow, 7].Value;
        stitle8 := ExcelApp.Cells[irow, 8].Value;    
        stitle9 := ExcelApp.Cells[irow, 9].Value;

        stitle4x := stitle1 + stitle2 + stitle3 + stitle4;      
        stitle8x := stitle1 + stitle2 + stitle3 + stitle4 +
          stitle5 + stitle6 + stitle7 + stitle8;
          
        FHaveArea := False;

        if stitle4x = '制式物料编码颜色容量' then
        begin       
          iVer := 1;
          iNumber := 2;
          iColor := 3;
          iCap := 4;  
          icol := 5;
        end
        else if stitle8x = '项目整机/裸机包装标准制式制式物料编码颜色容量' then
        begin
          iVer := 5;
          iNumber := 6;
          iColor := 7;
          iCap := 8;
          icol := 9;
        end
        else if stitle4x = 'MRP区域制式物料编码颜色容量' then
        begin          
          FHaveArea := True;
          icolMRPArea := 1;
          iVer := 2;
          iNumber := 3;
          iColor := 4;
          iCap := 5;
          icol := 6;
        end
        else if stitle8x = '项目整机/裸机包装标准制式MRP区域制式物料编码颜色容量' then
        begin
          FHaveArea := True;
          icolMRPArea := 5;
          iVer := 6;
          iNumber := 7;
          iColor := 8;
          iCap := 9;
          icol := 10;
        end
        else
        begin
          Memo1.Lines.Add('Seet ' + sproj + ' 格式不符合');
          Continue;
        end;




//        sVer := ExcelApp.Cells[irow, iVer].Value;
//        sNumber := ExcelApp.Cells[irow, iNumber].Value;
//        sColor := ExcelApp.Cells[irow, iColor].Value;
//        sCap := ExcelApp.Cells[irow, iCap].Value;

//        aProjData := TProjData.Create(sproj); 
//        FProjs.AddObject(sproj, aProjData);
        slMonth := TStringList.Create;
            
        try
          //判断有多少个月 
          icol1 := icol;
          sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          while sWeek <> '' do
          begin
            if not IsWeek(ExcelApp, CITitleRow, icol, sWeek) then
            //if UpperCase(Copy(sWeek, 1, 1)) <> 'W' then
            begin
              aColRangePtr := New(PColRange);
              slMonth.AddObject(sWeek, TObject(aColRangePtr));
              aColRangePtr^.col1 := icol1;
              aColRangePtr^.col2 := icol - 1;
              icol1 := icol + 1;       
            end;

            icol := icol + 1;                       
            sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          end;
 
          slYearOfProj.Text := mmoYearOfProj.Text;
          dt1Prev := 0;

          // 读取日期
          for iMonth := 0 to slMonth.Count - 1 do
          begin
            Memo1.Lines.Add(' Month:  ' + slMonth[iMonth]);

            aColRangePtr := PColRange(slMonth.Objects[iMonth]);
            SetLength(aColRangePtr^.FDates1, aColRangePtr^.col2 - aColRangePtr^.col1 + 1);
            SetLength(aColRangePtr^.FDates2, aColRangePtr^.col2 - aColRangePtr^.col1 + 1);
            for icol := aColRangePtr^.col1 to aColRangePtr^.col2 do
            begin 
              if Copy(slMonth[iMonth], 1, 4) = '截止' then
              begin
                sDate := Copy(slMonth[iMonth], 5, Length(slMonth[iMonth]) - 4);  
                sDate1 := sDate;
                sDate2 := sDate;
              end
              else
              begin
                sDate := ExcelApp.Cells[CIDateRow, icol].Value;
                sDate1 := Copy(sDate, 1, Pos('-', sDate) - 1);
                sDate2 := Copy(sDate, Pos('-', sDate) + 1, Length(sDate) - Pos('-', sDate));
              end;
                  
              sDate1 := StringReplace(sDate1, '/', '-', [rfReplaceAll]);
              sDate2 := StringReplace(sDate2, '/', '-', [rfReplaceAll]);

               idxYear := IndexOfYearOfProj( slYearOfProj, sProj);
              if idxYear >= 0 then
              begin

                sDate1 := slYearOfProj.ValueFromIndex[idxYear] + '-' + StringReplace(sDate1, '/', '-', [rfReplaceAll]);
                sDate2 := slYearOfProj.ValueFromIndex[idxYear] + '-' + StringReplace(sDate2, '/', '-', [rfReplaceAll]);
                try
                  dt1 := myStrToDateTime(sDate1);
                  dt2 := myStrToDateTime(sDate2);
                except
                  on e: Exception do
                  begin
                    Memo1.Lines.Add('myStrToDateTime error ' + sDate1 + '  ' + sDate2);
                    raise e;
                  end;
                end;

                if (dt1Prev <> 0) and (dt1 < dt1Prev) then
                begin
                  slYearOfProj.ValueFromIndex[idxYear] := IntToStr( StrToInt(slYearOfProj.ValueFromIndex[idxYear]) + 1 );
                  dt1 := EncodeDate( StrToInt(slYearOfProj.ValueFromIndex[idxYear]), MonthOf(dt1), DayOf(dt1) );
                  dt2 := EncodeDate( StrToInt(slYearOfProj.ValueFromIndex[idxYear]), MonthOf(dt2), DayOf(dt2) );
                end;

                dt1Prev := dt1;
              end
              else
              begin
                sDate1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sDate1, '/', '-', [rfReplaceAll]);
                sDate2 := IntToStr(YearOf(Now)) + '-' + StringReplace(sDate2, '/', '-', [rfReplaceAll]);
                try
                  dt1 := myStrToDateTime(sDate1);
                  dt2 := myStrToDateTime(sDate2);
                except
                  on e: Exception do
                  begin
                    Memo1.Lines.Add('myStrToDateTime error 222 ' + sDate1 + '  ' + sDate2);
                    raise e;
                  end;
                end;
              end;
              aColRangePtr^.FDates1[icol - aColRangePtr^.col1] := dt1;
              aColRangePtr^.FDates2[icol - aColRangePtr^.col1] := dt2;
            end;

          end;
          

          irow := CITitleRow + 2; 
          while True do
          begin
            sVer := ExcelApp.Cells[irow, iVer].Value;

            if sVer = '' then Break; //读取完了
            if CellMerged(ExcelApp, irow, iNumber, irow, iColor) then Break; //读取完了

//            aVerData := aProjData.AddVer(sVer, slMonth);

            sMRPArea := ExcelApp.Cells[irow, icolMRPArea].Value;
            sNumber := ExcelApp.Cells[irow, iNumber].Value;
            sColor := ExcelApp.Cells[irow, iColor].Value;
            sCap := ExcelApp.Cells[irow, iCap].Value;

 
            irow1 := irow;
            irow2 := irow1 + 1;

            while CellMerged(ExcelApp, irow1, iVer, irow2, iVer) do
            begin
              irow2 := irow2 + 1;
            end;
            irow2 := irow2 - 1;

            for irow := irow1 to irow2 do
            begin                                                  
              sMRPArea := ExcelApp.Cells[irow, icolMRPArea].Value;
              sNumber := ExcelApp.Cells[irow, iNumber].Value;
              sColor := ExcelApp.Cells[irow, iColor].Value;
              sCap := ExcelApp.Cells[irow, iCap].Value;

              //按月读取数据
              for iMonth := 0 to slMonth.Count - 1 do
              begin
                aColRangePtr := PColRange(slMonth.Objects[iMonth]);
                for icol := aColRangePtr^.col1 to aColRangePtr^.col2 do
                begin

                  dt1 := aColRangePtr^.FDates1[icol - aColRangePtr^.col1];  
                  dt2 := aColRangePtr^.FDates1[icol - aColRangePtr^.col1]; 
 
                  try
                    v := ExcelApp.Cells[irow, icol].Value;
                    if not VarIsEmpty(v) and not VarIsNumeric(v) then
                    begin
                      s := 'Sheet ' + sSheet + ' 行' + IntToStr(irow) + '列' + GetRef(icol) + '格式不对';
                      MessageBox(Handle, PChar(s), '错误', 0);
                      raise Exception.Create(s);
                    end;
                    iQty := v;
                    if dt1 >= dtStartxx then // 加个过滤，大于等于设定日期的才写入
                    begin
                      aSOPWriter.SOPWriteNumber(sproj, sVer, sNumber, sColor, sCap, '整机', dt1, dt2, iQty, sMRPArea);
                    end;
                  except
                    on e: Exception do
                    begin
                      raise e;
                    end;
                  end;
                end;
              end;
            end;

            irow := irow2 + 1;
          end;
        finally
          for iMonth := 0 to slMonth.Count - 1 do
          begin
            aColRangePtr := PColRange(slMonth.Objects[iMonth]);   
            SetLength(aColRangePtr^.FDates1, 0);
            SetLength(aColRangePtr^.FDates2, 0);
            Dispose(aColRangePtr);
          end;
          slMonth.Free;
        end;
      end;
    finally
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
      WorkBook.Close;
    end;
  finally
    ExcelApp.Visible := True;
    ExcelApp.Quit;        
    
    slYearOfProj.Free; 
  end;
end;

procedure TfrmSOP2PPOrder.tbOEMClick(Sender: TObject);
var
  sfile: string;
  aSOPWriter: TSOPWriter;
begin
  Memo1.Lines.Add('-------------  OEM  ----------------------------');
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'OEM' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;
  sfile := SaveDialog1.FileName;

//  sfile := 'C:\Users\qiujinbo\Desktop\SOP-PC\odmxxx.xlsx';

  aSOPWriter := TSOPWriter.Create(Handle);
  try
    aSOPWriter.BeginWrite(sfile, '产品预测单');
    try
      OpenOEMDemand(leSOP.Text, aSOPWriter);


      MessageBox(Handle, '完成', '金蝶提示', 0);
    finally
      aSOPWriter.EndWrite;
    end;


  finally
    aSOPWriter.Free;
  end;
end;

procedure TfrmSOP2PPOrder.tbODMSumClick(Sender: TObject);
var
  sfile: string;
begin
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  SaveDialog1.FileName := 'ODM-SOP-Sum-' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;
  sfile := SaveDialog1.FileName;

  Clear;


  aFMonths := TStringList.Create;

//  sfile := 'C:\Users\qiujinbo\Desktop\SOP-PC\sop.xlsx';

  OpenODMDemand4Sum(leSOP.Text);

  MakeODMDemand4Sum(sfile);

  aFMonths.Free;

  MessageBox(Handle, '完成', '金蝶提示', 0);
end;

procedure TfrmSOP2PPOrder.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

end.


