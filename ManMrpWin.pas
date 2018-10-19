unit ManMrpWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ImgList, ComCtrls, ToolWin, ComObj, DateUtils,
  Buttons, IniFiles, CommUtils, SOPReaderUnit,
  SAPMaterialReader, SAPMaterialReader2;

type
  TNumberInfo = packed record
//    sProj: string;
    Number: string; //加工材料长代码
    Name: string; //加工材料名称
//    Color: string;//颜色
//    Cap: string; //容量
//    Ver: string;//制式
//    FG: string;//整机/裸机
//    Pkg: string;//豪华装
  end;
  PNumberInfo = ^TNumberInfo;

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

  TSOPWriter = class
  private
    ffile: string;
    fh: HWND;
    frow: Integer;
    ExcelApp, WorkBook: Variant;
  public
    constructor Create(h: HWND);
    procedure BeginWrite(const sfile, sSheetName: string);
    procedure SOPWriteNumber(const sSheet, sVer, sNumber, sColor, sCap, sFG: string;
      dt1, dt2: TDateTime; iQty: Integer; const sProj, sName, sPkg: string);
    procedure EndWrite;
  end;

  /////////////////////////////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////////////////////////////
  /////////////////////////////////////////////////////////////////////////////////////////////////////////

  TMPSInfo = packed record
    Number: string; // = 1;   //产品编码
    Name: string; // = 2;     //产品名称
    Date: TDateTime; // = 3;     //需求日期
    Qty: Double; // = 4;     	//数量

    Proj: string;
    Color: string;//颜色
    Cap: string; //容量
    Ver: string;//制式
    FG: string;//整机/裸机
    Pkg: string;//豪华装
  end;
  PMPSInfo = ^TMPSInfo;

  TDateSorter = class
  private
    FDateStart, FDateEnd: TDateTime;
    FQty: Double;
  public
    constructor Create(const dt1, dt2: TDateTime);
    destructor Destroy; override;
  end;

  TNumberSorter = class
  private
    FNumber: string;
    FName: string;
    FDates: TList;
    procedure Clear;
    procedure BuildDateList(slDates: TStringList);
    function GetDate(const dt: TDateTime): TDateSorter;
  public
    constructor Create(const snumber, sname: string; slDates: TStringList);
    destructor Destroy; override;
    function Add(p: PMPSInfo): Integer; overload;
    function Add(dt: TDateTime; dQty: Double): Integer; overload;
  end;

  TPkgSorter = class
  private
    FPkg: string;
    FNumbers: TList;
    procedure Clear;
  public
    constructor Create(const sPkg: string);
    destructor Destroy; override;
    function GetNumber(const snumber: string): TNumberSorter;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
  end;

  TFGSorter = class
  private
    FFG: string;
    FPkgs: TList;
    procedure Clear;
  public
    constructor Create(const sFG: string);
    destructor Destroy; override;
    function GetPkg(const sPkg: string): TPkgSorter;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
  end;

  TCapSorter = class
  private
    FCap: string;
    FFGs: TList;
    procedure Clear;
  public
    constructor Create(const sCap: string);
    destructor Destroy; override;
    function GetFG(const sFG: string): TFGSorter;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
  end;

  TColorSorter = class
  private
    FColor: string;
    FCaps: TList;
    procedure Clear;   
  public
    constructor Create(const sColor: string);
    destructor Destroy; override;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
    function GetCap(const sCap: string): TCapSorter;
  end;

  TVerSorter = class
  private
    FVer: string;
    FColors: TList;
    procedure Clear;
  public
    constructor Create(const sVer: string);
    destructor Destroy; override;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
    function GetColor(const sColor: string): TColorSorter;
  end;

  TMPSSorter = class
  private
    FVers: TList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(p: PMPSInfo; slDate: TStringList): Integer;
    function GetVer(const sVer: string): TVerSorter;
  end;

  TProjInfo = class  //xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
  private
    FName: string;
    FList: TList;
    FDates: TStringList;
    procedure Clear;
  public
    FIndex: Integer;
    constructor Create(const sproj: string);
    destructor Destroy; override;
    function Add(aMPSInfoPtr: PMPSInfo; const sdt: string): Integer;
  end;
      
  TProjInfos = class
  private
    FList: TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
    function GetProjInfo(const sproj: string; iIndex: Integer): TProjInfo;
    procedure SortDate;
  end;
          
  TKeySumer = class
  private
    FList: TStringList;
    procedure Clear;
  public
    constructor Create;
    destructor Destroy; override;
    function Add(const skey: string; aDateSorter: TDateSorter;
      slDate: TStringList): Integer;
    function GetItem(const skey: string): TObject;
  end;

  TfrmManMrp = class(TForm)
    leSOP: TLabeledEdit;
    btnSOP: TButton;
    OpenDialog1: TOpenDialog;
    ToolBar1: TToolBar;
    ImageList1: TImageList;
    Memo1: TMemo;
    SaveDialog1: TSaveDialog;
    leNumberList: TLabeledEdit;
    btnNumberList: TButton;
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    lvMPS: TListView;
    ProgressBar1: TProgressBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    Label1: TLabel;
    dtpStart: TDateTimePicker;
    mmoYearOfProj: TMemo;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    Label2: TLabel;
    dtpFilterDate: TDateTimePicker;
    Label3: TLabel;
    Label4: TLabel;
    dtpStart2: TDateTimePicker;
    Label5: TLabel;
    ToolButton5: TToolButton;
    tbQuit: TToolButton;
    tbOEM2: TToolButton;
    ToolButton7: TToolButton;
    procedure btnSOPClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnNumberListClick(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure dtpStartChange(Sender: TObject);
    procedure tbQuitClick(Sender: TObject);
    procedure tbOEM2Click(Sender: TObject);
  private
    { Private declarations }
    FProjs: TStringList;
    FMPSProjs: TProjInfos;
//    FNumberInfoList: TList;
    FSAPMaterialReader: TSAPMaterialReader2;

    procedure AddFile(lv: TListView);
    procedure RemoveFile(lv: TListView);
    
    procedure Clear;

//    function GetNumberInfo(const sNumber: string): PNumberInfo;
    procedure OpenOEMDemand(const sfile: string; aSOPWriter: TSOPWriter);
    
    //procedure ReadNumberList(const sFile: string);    
    procedure SortMPS(aProjInfo: TProjInfo; aMPSSorter: TMPSSorter);
    procedure ReadMPS(const sFile: string);     
    procedure SaveMPS(ExcelApp: Variant; aMPSSorter: TMPSSorter;
      slDate: TStringList);                
    function FileNameExists(const sfile: string; lv: TListView): Boolean;
    procedure OnLogEvent(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

   
implementation

{$R *.dfm}
                
const
  xlCenter = -4108;
  
var
  bMan: Boolean = True;

class procedure TfrmManMrp.ShowForm;  
var
  frmManMrp: TfrmManMrp;
begin
  frmManMrp := TfrmManMrp.Create(nil);
  try
    frmManMrp.ShowModal;
  finally
    frmManMrp.Free;
  end;
end;
     
// 冒泡排序
procedure SortDates(aDates: TStringList);
var
  i, j: Integer; 
  dti, dtj: TDateTime;
  s: string;
begin
  for i := 0 to aDates.Count - 2 do
  begin
    for j := i + 1 to aDates.Count -1  do
    begin
      dti := myStrToDateTime(aDates[i]);
      dtj := myStrToDateTime(aDates[j]);
      if dti > dtj then
      begin
        s := aDates[i];
        aDates[i] := aDates[j];
        aDates[j] := s;
      end;
    end;
  end;
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
    
{ TKeySumer }

constructor TKeySumer.Create;
begin
  FList := TStringList.Create;
end;

destructor TKeySumer.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

procedure TKeySumer.Clear;
var
  i: Integer;
  aNumberSorter: TNumberSorter;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aNumberSorter := TNumberSorter(FList.Objects[i]);
    aNumberSorter.Free;
  end;
  FList.Clear;
end;

function TKeySumer.Add(const skey: string; aDateSorter: TDateSorter;
  slDate: TStringList): Integer;
var
  idx: Integer;
  aNumberSorter: TNumberSorter;
begin
  idx := FList.IndexOf(skey);
  if idx < 0 then
  begin
    aNumberSorter := TNumberSorter.Create('', '', slDate);
    FList.AddObject(skey, aNumberSorter);
  end
  else
  begin
    aNumberSorter := TNumberSorter(FList.Objects[idx]);
  end;
  Result := aNumberSorter.Add(aDateSorter.FDateStart, aDateSorter.FQty);
end;

function TKeySumer.GetItem(const skey: string): TObject;
var
  idx: Integer;        
begin
  Result := nil;
  idx := FList.IndexOf(skey);
  if idx >= 0 then
  begin
    Result := FList.Objects[idx];
  end;
end;

{ TProjInfos }

constructor TProjInfos.Create;
begin
  FList := TStringList.Create;
end;

destructor TProjInfos.Destroy;
begin
  Clear;
  FList.Clear;
  inherited;
end;

procedure TProjInfos.Clear;
var
  i: Integer;
  aProjInfo: TProjInfo;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjInfo := TProjInfo(FList.Objects[i]);
    aProjInfo.Free;
  end;
  FList.Clear;
end;

function TProjInfos.GetProjInfo(const sproj: string; iIndex: Integer): TProjInfo;
var
  idx: Integer;
  aProjInfo: TProjInfo;
begin
  idx := FList.IndexOf(sproj);
  if idx < 0 then
  begin
    aProjInfo := TProjInfo.Create(sproj);
    aProjInfo.FIndex := iIndex;
    idx := FList.AddObject(sproj, aProjInfo);
  end;
  Result := TProjInfo(FList.Objects[idx]);
end;

procedure TProjInfos.SortDate;
var
  i: Integer;   
  aProjInfo: TProjInfo;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aProjInfo := TProjInfo(FList.Objects[i]);
    SortDates(aProjInfo.FDates);
  end;
end;

{ TDateSorter }

constructor TDateSorter.Create(const dt1, dt2: TDateTime);
begin
  FDateStart := dt1;
  FDateEnd := dt2;
  FQty := 0;
end;

destructor TDateSorter.Destroy;
begin
  inherited;
end;
   
{ TNumberSorter }

constructor TNumberSorter.Create(const snumber, sname: string;
  slDates: TStringList);
begin
  FNumber := snumber;
  FName := sname;
  FDates := TList.Create;
  BuildDateList(slDates);
end;

destructor TNumberSorter.Destroy;
begin
  Clear;
  FDates.Free;
  inherited;
end;

procedure TNumberSorter.Clear;
var
  i: Integer;
  aDateSorter: TDateSorter;
begin
  for i := 0 to FDates.Count - 1 do
  begin
    aDateSorter := TDateSorter(FDates[i]);
    aDateSorter.Free;
  end;
end;

function TNumberSorter.GetDate(const dt: TDateTime): TDateSorter;
var
  i: Integer;
  aDateSorter: TDateSorter;
begin
  Result := nil;
  for i := 0 to FDates.Count - 1 do
  begin
    aDateSorter := TDateSorter(FDates[i]);
    if (dt >= aDateSorter.FDateStart) and (dt < aDateSorter.FDateEnd) then
    begin
      Result := aDateSorter;
      Break;
    end;
  end;
end;

procedure TNumberSorter.BuildDateList(slDates: TStringList);
var
  aDateSorter: TDateSorter;
  i: Integer;
  s: string;
  dt, dt2: TDateTime;
begin
  if slDates.Count = 0 then Exit;
 
  for i := 0 to slDates.Count - 2 do
  begin
    s := slDates[i];
    dt := myStrToDateTime(s);

    s := slDates[i + 1];
    dt2 := myStrToDateTime(s);

    aDateSorter := TDateSorter.Create(dt, dt2);
    FDates.Add(aDateSorter); 
  end;

  s := slDates[slDates.Count - 1];
  dt := myStrToDateTime(s);
  aDateSorter := TDateSorter.Create(dt, dt + 7);
  FDates.Add(aDateSorter); 
end;

function TNumberSorter.Add(p: PMPSInfo): Integer;
var
  aDateSorter: TDateSorter;
begin
  Result := 0;

  aDateSorter := GetDate(p^.Date);
 
  aDateSorter.FQty  := aDateSorter.FQty + p^.Qty;
end;

function TNumberSorter.Add(dt: TDateTime; dQty: Double): Integer;
var
  aDateSorter: TDateSorter;
begin
  Result := 0;
  aDateSorter := GetDate(dt);
  aDateSorter.FQty  := aDateSorter.FQty + dQty;
end;

{ TPkgSorter }

constructor TPkgSorter.Create(const sPkg: string);
begin
  FPkg := sPkg;
  FNumbers := TList.Create;
end;

destructor TPkgSorter.Destroy;
begin
  Clear;
  FNumbers.Free;
  inherited;
end;

procedure TPkgSorter.Clear;
var
  i: Integer;
  aNumberSorter: TNumberSorter;
begin
  for i := 0 to FNumbers.Count - 1 do
  begin
    aNumberSorter := TNumberSorter(FNumbers[i]);
    aNumberSorter.Free;
  end;
  FNumbers.Clear;
end;

function TPkgSorter.GetNumber(const snumber: string): TNumberSorter;
var
  i: Integer;
  aNumberSorter: TNumberSorter;
begin
  Result := nil;
  for i := 0 to FNumbers.Count - 1 do
  begin
    aNumberSorter := TNumberSorter(FNumbers[i]);
    if aNumberSorter.FNumber = snumber then
    begin
      Result := aNumberSorter;
      Break;
    end;
  end;
end;

function TPkgSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aNumberSorter: TNumberSorter;
begin
  Result := 0;
  aNumberSorter := GetNumber(p^.Number);
  if aNumberSorter = nil then
  begin
    aNumberSorter := TNumberSorter.Create(p^.Number, p^.Name, slDate);
    FNumbers.Add(aNumberSorter);
  end;
  aNumberSorter.Add(p);
end;

{ TFGSorter }

constructor TFGSorter.Create(const sFG: string);
begin
  FFG := sFG;
  FPkgs := TList.Create;
end;

destructor TFGSorter.Destroy;
begin
  Clear;
  FPkgs.Free;
  inherited;
end;

procedure TFGSorter.Clear;
var
  i: Integer;
  aPkgSorter: TPkgSorter;
begin
  for i := 0 to FPkgs.Count - 1 do
  begin
    aPkgSorter := TPkgSorter(FPkgs[i]);
    aPkgSorter.Free;
  end;
  FPkgs.Clear;
end;

function TFGSorter.GetPkg(const sPkg: string): TPkgSorter;
var
  i: Integer;
  aPkgSorter: TPkgSorter;
begin
  Result := nil;
  for i := 0 to FPkgs.Count - 1 do
  begin
    aPkgSorter := TPkgSorter(FPkgs[i]);
    if aPkgSorter.FPkg = sPkg then
    begin
      Result := aPkgSorter;
      Break;
    end;
  end;
end;

function TFGSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aPkgSorter: TPkgSorter;
begin
  Result := 0;
  aPkgSorter := GetPkg(p^.Pkg);
  if aPkgSorter = nil then
  begin
    aPkgSorter := TPkgSorter.Create(p^.Pkg);
    FPkgs.Add(aPkgSorter);
  end;
  aPkgSorter.Add(p, slDate);
end;

{ TCapSorter }

constructor TCapSorter.Create(const sCap: string);
begin
  FCap := sCap;
  FFGs := TList.Create;
end;

destructor TCapSorter.Destroy;
begin
  Clear;
  FFGs.Free;
  inherited;
end;

procedure TCapSorter.Clear;
var
  i: Integer;
  aFGSorter: TFGSorter;
begin
  for i := 0 to FFGs.Count - 1 do
  begin
    aFGSorter := TFGSorter(FFGs[i]);
    aFGSorter.Free;
  end;
  FFGs.Clear;
end;

function TCapSorter.GetFG(const sFG: string): TFGSorter;
var
  i: Integer;
  aFGSorter: TFGSorter;
begin
  Result := nil;
  for i := 0 to FFGs.Count - 1 do
  begin
    aFGSorter := TFGSorter(FFGs[i]);
    if aFGSorter.FFG = sFG then
    begin
      Result := aFGSorter;
      Break;
    end;
  end;
end;

function TCapSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aFGSorter: TFGSorter;
begin
  Result := 0;
  aFGSorter := GetFG(p^.FG);
  if aFGSorter = nil then
  begin
    aFGSorter := TFGSorter.Create(p^.FG);
    FFGs.Add(aFGSorter);
  end;
  aFGSorter.Add(p, slDate);
end;
    
{ TColorSorter }

constructor TColorSorter.Create(const sColor: string);
begin
  FColor := sColor;
  FCaps := TList.Create;
end;

destructor TColorSorter.Destroy;
begin
  Clear;
  FCaps.Free;
  inherited;
end;

procedure TColorSorter.Clear;
var
  i: Integer;
  aCapSorter: TCapSorter;
begin
  for i := 0 to FCaps.Count - 1 do
  begin
    aCapSorter := TCapSorter(FCaps[i]);
    aCapSorter.Free;
  end;
  FCaps.Clear;
end;

function TColorSorter.GetCap(const sCap: string): TCapSorter;
var
  i: Integer;
  aCapSorter: TCapSorter;
begin
  Result := nil;
  for i := 0 to FCaps.Count - 1 do
  begin
    aCapSorter := TCapSorter(FCaps[i]);
    if aCapSorter.FCap = sCap then
    begin
      Result := aCapSorter;
      Break;
    end;  
  end;
end;

function TColorSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aCapSorter: TCapSorter;
begin
  Result := 0;
  aCapSorter := GetCap(p^.Cap);
  if aCapSorter = nil then
  begin
    aCapSorter := TCapSorter.Create(p^.Cap);
    FCaps.Add(aCapSorter);
  end;
  aCapSorter.Add(p, slDate);
end;

{ TVerSorter }

constructor TVerSorter.Create(const sVer: string);
begin
  FVer := sVer;
  FColors := TList.Create;
end;

destructor TVerSorter.Destroy;
begin
  Clear;
  FColors.Free;
  inherited;
end;

procedure TVerSorter.Clear;
var
  i: Integer;
  aColorSorter: TColorSorter;
begin
  for i := 0 to FColors.Count - 1 do
  begin
    aColorSorter := TColorSorter(FColors[i]);
    aColorSorter.Free;
  end;
  FColors.Clear;
end;

function TVerSorter.GetColor(const sColor: string): TColorSorter;
var
  i: Integer;
  aColorSorter: TColorSorter;
begin
  Result := nil;
  for i := 0 to FColors.Count - 1 do
  begin
    aColorSorter := TColorSorter(FColors[i]);
    if aColorSorter.FColor = sColor then
    begin
      Result := aColorSorter;
      Break;
    end;
  end;
end;

function TVerSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aColorSorter: TColorSorter;
begin
  Result := 0;
  aColorSorter := GetColor(p^.Color);
  if aColorSorter = nil then
  begin
    aColorSorter := TColorSorter.Create(p^.Color);
    FColors.Add(aColorSorter);
  end;
  aColorSorter.Add(p, slDate);
end;

{ TMPSSorter }

constructor TMPSSorter.Create;
begin
  FVers := TList.Create;
end;

destructor TMPSSorter.Destroy;
begin
  Clear;
  FVers.Free;
  inherited;
end;

procedure TMPSSorter.Clear;
var
  i: Integer;
  aVerSorter: TVerSorter;
begin
  for i := 0 to FVers.Count - 1 do
  begin
    aVerSorter := TVerSorter(FVers[i]);
    aVerSorter.Free;
  end;
  FVers.Clear;
end;

function TMPSSorter.GetVer(const sVer: string): TVerSorter;
var
  i: Integer;
  aVerSorter: TVerSorter;
begin
  Result := nil;
  for i := 0 to FVers.Count - 1 do
  begin
    aVerSorter := TVerSorter(FVers[i]);
    if aVerSorter.FVer = sVer then
    begin
      Result := aVerSorter;
      Break;
    end;
  end;
end;

function TMPSSorter.Add(p: PMPSInfo; slDate: TStringList): Integer;
var
  aVerSorter: TVerSorter;
begin
  Result := 0;
  aVerSorter := GetVer(p^.Ver);
  if aVerSorter = nil then
  begin
    aVerSorter := TVerSorter.Create(p^.Ver);
    FVers.Add(aVerSorter);
  end;
  aVerSorter.Add(p, slDate);
end;
  
{ TProjInfo }

constructor TProjInfo.Create(const sproj: string);
begin
  FName := sproj;
  FDates := TStringList.Create;
  FList := TList.Create;
end;

destructor TProjInfo.Destroy; 
begin
  Clear;
  FList.Free;
  FDates.Free;
  inherited;
end;

function TProjInfo.Add(aMPSInfoPtr: PMPSInfo; const sdt: string): Integer;
begin
  Result := FList.Add(aMPSInfoPtr);
  if FDates.IndexOf(sdt) < 0 then
  begin
    FDates.Add(sdt);
  end;
end;

procedure TProjInfo.Clear;
var
  i: Integer;
  aMPSInfo: PMPSInfo;
begin
  for i := 0 to FList.Count - 1 do
  begin
    aMPSInfo := PMPSInfo(FLIst[i]);
    Dispose(aMPSInfo);
  end;
  FList.Clear;

  FDates.Clear;
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
  
  ExcelApp.Cells[frow, 14].Value := '项目';    
  ExcelApp.Cells[frow, 15].Value := '物料名称';
  
  ExcelApp.Cells[frow, 16].Value := '颜色';
  ExcelApp.Cells[frow, 17].Value := '容量';
  ExcelApp.Cells[frow, 18].Value := '版本';
  ExcelApp.Cells[frow, 19].Value := '包装';
    
  ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 19]].HorizontalAlignment := xlCenter; //居中

  frow := frow + 1;

end;

procedure TSOPWriter.SOPWriteNumber(const sSheet,  sVer, sNumber, sColor, sCap, sFG: string;
  dt1, dt2: TDateTime; iQty: Integer; const sProj, sName, sPkg: string);
begin
  //if iQty = 0 then Exit;

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
    ExcelApp.Cells[frow, 3].Value := sSheet + ',' + sVer + ',' + sColor + ',' + sCap;
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
  ExcelApp.Cells[frow, 12].Value := sSheet;
  ExcelApp.Cells[frow, 13].Value := sVer;
  ExcelApp.Cells[frow, 14].Value := sProj;     
  ExcelApp.Cells[frow, 15].Value := sName;
  ExcelApp.Cells[frow, 16].Value := sColor;
  ExcelApp.Cells[frow, 17].Value := sCap;
  ExcelApp.Cells[frow, 18].Value := sFG;
  ExcelApp.Cells[frow, 19].Value := sPkg;

  frow := frow + 1;
end;

procedure TSOPWriter.EndWrite;
begin

  ExcelApp.Columns[1].ColumnWidth := 16;
  ExcelApp.Columns[3].ColumnWidth := 16;
  ExcelApp.Columns[6].ColumnWidth := 13;
  ExcelApp.Columns[7].ColumnWidth := 13;
  ExcelApp.Columns[12].ColumnWidth := 16;
  ExcelApp.Columns[13].ColumnWidth := 16;
  ExcelApp.Columns[14].ColumnWidth := 16;   
  ExcelApp.Columns[15].ColumnWidth := 16;
  ExcelApp.Columns[16].ColumnWidth := 16;
  ExcelApp.Columns[17].ColumnWidth := 16;
  ExcelApp.Columns[18].ColumnWidth := 16;     
  ExcelApp.Columns[19].ColumnWidth := 16;

  
  ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[frow-1, 19]].Borders.LineStyle := 1; //加边框

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

procedure TfrmManMrp.FormCreate(Sender: TObject);
var
//  dt: TDateTime;
  ini: TIniFile;
  s: string;
begin 
  Memo1.Clear;

  FProjs := TStringList.Create;
  FMPSProjs := TProjInfos.Create;
//  FNumberInfoList := TList.Create;

//  leSOP.Text := 'E:\erp\mrp\mrpcode\S&OP转产品预测单模板（MC）\OEM MPS总表0107.xlsx';
//  leNumberList.Text := 'E:\erp\mrp\mrpcode\S&OP转产品预测单模板（MC）\成品料号清单 2016-12-26.xlsx';

//  dt := Now;
//  dt := Trunc(dt);
//  dtpStart.DateTime := dt;

  ini := TIniFile.Create(AppIni);
  try
    mmoYearOfProj.Text := StringReplace( ini.ReadString(self.ClassName, mmoYearOfProj.Name, ''), ';', #13#10, [rfReplaceAll] );
    leNumberList.Text := ini.ReadString(self.ClassName, leNumberList.Name, '');
    leSOP.Text := ini.ReadString(self.ClassName, leSOP.Name, '');

    s := ini.ReadString(self.ClassName, dtpStart.Name, FormatDateTime('yyyy-MM-dd', Now));
    dtpStart.DateTime := myStrToDateTime(s);
    s := ini.ReadString(self.ClassName, dtpStart2.Name, FormatDateTime('yyyy-MM-dd', Now));
    dtpStart2.DateTime := myStrToDateTime(s);
    s := ini.ReadString(self.ClassName, dtpFilterDate.Name, FormatDateTime('yyyy-MM-dd', Now));
    dtpFilterDate.DateTime := myStrToDateTime(s);
  finally
    ini.Free;
  end;

  dtpStartChange(nil);
end;
   
procedure TfrmManMrp.btnSOPClick(Sender: TObject);
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
 
procedure TfrmManMrp.Clear;
var
  i: Integer;
  aProjData: TProjData;
  aNumberInfoPtr: PNumberInfo;
begin
  for i := 0 to FProjs.Count - 1 do
  begin
    aProjData := TProjData(FProjs.Objects[i]);
    aProjData.Free;
  end;
  FProjs.Clear;

//  for i := 0 to FNumberInfoList.Count - 1 do
//  begin
//    aNumberInfoPtr := PNumberInfo(FNumberInfoList[i]);
//    Dispose(aNumberInfoPtr);
//  end;
//  FNumberInfoList.Clear;

  FMPSProjs.Clear;
end;

procedure TfrmManMrp.FormDestroy(Sender: TObject);
var
  ini: TIniFile;
begin
  Clear;
  FProjs.Free;    
  FMPSProjs.Free;
//  FNumberInfoList.Free;

  ini := TIniFile.Create(AppIni);
  try
    ini.WriteString(self.ClassName, mmoYearOfProj.Name,StringReplace(mmoYearOfProj.Text, #13#10, ';', [rfReplaceAll] )); 
    ini.WriteString(self.ClassName, leNumberList.Name, leNumberList.Text);
    ini.WriteString(self.ClassName, leSOP.Name, leSOP.Text);

    ini.WriteString(self.ClassName, dtpStart.Name, FormatDateTime('yyyy-MM-dd', dtpStart.DateTime) );
    ini.WriteString(self.ClassName, dtpStart2.Name, FormatDateTime('yyyy-MM-dd', dtpStart2.DateTime) ); 
    ini.WriteString(self.ClassName, dtpFilterDate.Name, FormatDateTime('yyyy-MM-dd', dtpFilterDate.DateTime) );  
  finally
    ini.Free;
  end;
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

//function TfrmManMrp.GetNumberInfo(const sNumber: string): PNumberInfo;
//var
//  i: Integer;
//  aNumberInfoPtr: PNumberInfo;
//begin
//  Result := nil;
//  for i := 0 to FNumberInfoList.Count - 1 do
//  begin
//    aNumberInfoPtr := PNumberInfo(FNumberInfoList[i]);
//    if aNumberInfoPtr^.Number = sNumber then
//    begin
//      Result := aNumberInfoPtr;
//      Break;
//    end;
//  end;
//end;

procedure TfrmManMrp.OpenOEMDemand(const sfile: string; aSOPWriter: TSOPWriter);
const
  CIProj = 1;
  CIFG = 2;
  CIPkg = 3;
  CIStdVer = 4;

  CIVer = 5;
  CINumber = 6;
  CIColor = 7;
  CICap = 8;

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
  sSheet: string;
  sName: string;

  sVer: string;
  sNumber: string;
  sColor: string;
  sCap: string;
 
  slMonth: TStringList;
  iMonth: Integer; 

  aColRangePtr: PColRange;
 
  iQty: Integer;

  dt1, dt2: TDateTime;
  sDate, sDate1, sDate2: string;
  sNumberFormatlocal: string;     
  dtStart: TDateTime;
  
//  aNumberInfoPtr: PNumberInfo;
  aSAPMaterialRecordPtr: PSAPMaterialRecord;

  sProj: string;
  sFG: string;
  sPkg: string;
  sStdVer: string;

  slYearOfProj: TStringList;
  dt1Prev: TDateTime;

  v: Variant;
  s: string;
begin
  slYearOfProj := TStringList.Create;

  dtStart := myStrToDateTime(FormatDateTime('yyyy-MM-dd', dtpFilterDate.DateTime));

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

        Memo1.Lines.Add('Sheet ' + sSheet);

        irow := 1;
        sVer := ExcelApp.Cells[irow, CIVer].Value;
        sNumber := ExcelApp.Cells[irow, CINumber].Value;
        sColor := ExcelApp.Cells[irow, CIColor].Value;
        sCap := ExcelApp.Cells[irow, CICap].Value;

        if (sVer <> CSVer) or (sNumber <> CSNumber) or (sColor <> CSColor) or (sCap <> CSCap) then
        begin            
          Memo1.Lines.Add('Seet ' + sSheet + ' 格式不符合');
          Continue;
        end;
 
        slMonth := TStringList.Create;

        try
          
          Memo1.Lines.Add(' 判断有多少个月 ');

          //判断有多少个月
          icol := CICap + 1;
          sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          if Copy(sWeek, 1, 4) = '截止' then
          begin
            aColRangePtr := New(PColRange);
            slMonth.AddObject(sWeek, TObject(aColRangePtr));
            aColRangePtr^.col1 := icol;
            aColRangePtr^.col2 := icol;
            icol := icol + 1;
          end;

          icol1 := icol;
          sWeek := ExcelApp.Cells[CITitleRow, icol].Value;
          while sWeek <> '' do
          begin
            if not IsWeek(ExcelApp, CITitleRow, icol, sWeek) then
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
             
          sProj := ExcelApp.Cells[3, CIProj].Value;
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

              if slYearOfProj.IndexOfName(sProj) >= 0 then
              begin

                sDate1 := slYearOfProj.Values[sProj] + '-' + StringReplace(sDate1, '/', '-', [rfReplaceAll]);
                sDate2 := slYearOfProj.Values[sProj] + '-' + StringReplace(sDate2, '/', '-', [rfReplaceAll]);
                dt1 := myStrToDateTime(sDate1);
                dt2 := myStrToDateTime(sDate2);

                if (dt1Prev <> 0) and (dt1 < dt1Prev) then
                begin
                  slYearOfProj.Values[sProj] := IntToStr( StrToInt(slYearOfProj.Values[sProj]) + 1 );
                  dt1 := EncodeDate( StrToInt(slYearOfProj.Values[sProj]), MonthOf(dt1), DayOf(dt1) );
                  dt2 := EncodeDate( StrToInt(slYearOfProj.Values[sProj]), MonthOf(dt2), DayOf(dt2) );
                end;

                dt1Prev := dt1;
              end
              else
              begin
                sDate1 := IntToStr(YearOf(Now)) + '-' + StringReplace(sDate1, '/', '-', [rfReplaceAll]);
                sDate2 := IntToStr(YearOf(Now)) + '-' + StringReplace(sDate2, '/', '-', [rfReplaceAll]);
                dt1 := myStrToDateTime(sDate1);
                dt2 := myStrToDateTime(sDate2);
              end;
              aColRangePtr^.FDates1[icol - aColRangePtr^.col1] := dt1;
              aColRangePtr^.FDates2[icol - aColRangePtr^.col1] := dt2;
            end;

          end;
          
             
          irow := CITitleRow + 2; 
          while True do
          begin
            sVer := ExcelApp.Cells[irow, CIVer].Value;
                
            Memo1.Lines.Add(' sVer:  ' + sVer);
        
            if sVer = '' then Break; //读取完了
            if CellMerged(ExcelApp, irow, CINumber, irow, CIColor) then Break; //读取完了
 
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
              Memo1.Lines.Add(' irow:  ' + IntToStr(irow));
        
              sNumber := ExcelApp.Cells[irow, CINumber].Value;   
              sColor := ExcelApp.Cells[irow, CIColor].Value;
              sCap := ExcelApp.Cells[irow, CICap].Value;


              sProj := ExcelApp.Cells[irow, CIProj].Value;
              sFG := ExcelApp.Cells[irow, CIFG].Value;
              sPkg := ExcelApp.Cells[irow, CIPkg].Value;
              sStdVer := ExcelApp.Cells[irow, CIStdVer].Value;


              aSAPMaterialRecordPtr := FSAPMaterialReader.GetSAPMaterialRecord(sNumber);
//              aNumberInfoPtr := GetNumberInfo(sNumber);

              //按月读取数据
              for iMonth := 0 to slMonth.Count - 1 do
              begin
                Memo1.Lines.Add(' Month:  ' + slMonth[iMonth]);

                aColRangePtr := PColRange(slMonth.Objects[iMonth]);
                for icol := aColRangePtr^.col1 to aColRangePtr^.col2 do
                begin
                  dt1 := aColRangePtr^.FDates1[icol - aColRangePtr^.col1];  
                  dt2 := aColRangePtr^.FDates1[icol - aColRangePtr^.col1]; 
                  try
                    sName := '';
                    if aSAPMaterialRecordPtr <> nil then
                    begin 
                      sName := aSAPMaterialRecordPtr^.sName;
                    end;
                    v := ExcelApp.Cells[irow, icol].Value;
                    if not VarIsEmpty(v) and not VarIsNumeric(v) then
                    begin
                      s := v;
                      if Trim(s) = '' then
                      begin
                        v := 0;
                      end
                      else
                      begin
                        s := 'Sheet ' + sSheet + ' 行' + IntToStr(irow) + '列' + GetRef(icol) + '格式不对';
                        MessageBox(Handle, PChar(s), '错误', 0);
                        raise Exception.Create(s);
                      end;
                    end;
                    iQty := v;
                    if dt1 >= dtStart then // 加个过滤，大于等于设定日期的才写入
                    begin                    
                      aSOPWriter.SOPWriteNumber(sSheet, sStdVer, sNumber, sColor, sCap, sFG, dt1, dt2, iQty, sProj, sName, sPkg);
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
    
//procedure TfrmManMrp.ReadNumberList(const sFile: string);
//const
//  CINumber = 1; //加工材料长代码
//  CIName = 2; //加工材料名称
////  CIColor = 3; //颜色
////  CICap = 4; //容量
////  CIVer = 5; //制式
////  CIFG = 6; //整机/裸机
////  CIPkg = 7; //豪华装
//
//var
//  ExcelApp, WorkBook: Variant;
//  snumber: string;
//  irow: Integer;
//  iSheetCount, iSheet: Integer;
//  aNumberInfoPtr: PNumberInfo;
//  sProj: string;
//begin
//  ExcelApp := CreateOleObject('Excel.Application' );
//  ExcelApp.Visible := False;
//  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
//
//  try
//    try
//      WorkBook := ExcelApp.WorkBooks.Open(sFile);
//    except
//      Memo1.Lines.Add('打开文件失败: ' + sfile);
//      Exit;
//    end;
//    try
//      iSheetCount := ExcelApp.WorkSheets.Count;
//      for iSheet := 1 to iSheetCount do
//      begin
//        ExcelApp.WorkSheets[iSheet].Activate;
//        sProj := ExcelApp.WorkSheets[iSheet].Name;
//
//        irow := 2;
//        snumber := ExcelApp.Cells[irow, CINumber].Value;
//        while snumber <> '' do
//        begin
//          aNumberInfoPtr := New(PNumberInfo);
////          aNumberInfoPtr^.sProj := sProj;
//          aNumberInfoPtr^.Number := snumber;
//          aNumberInfoPtr^.Name := ExcelApp.Cells[irow, CIName].Value;
////          aNumberInfoPtr^.Color := ExcelApp.Cells[irow, CIColor].Value;
////          aNumberInfoPtr^.Cap := ExcelApp.Cells[irow, CICap].Value;
////          aNumberInfoPtr^.Ver := ExcelApp.Cells[irow, CIVer].Value;
////          aNumberInfoPtr^.FG := ExcelApp.Cells[irow, CIFG].Value;
////          aNumberInfoPtr^.Pkg := ExcelApp.Cells[irow, CIPkg].Value;
//          FNumberInfoList.Add(aNumberInfoPtr);
//        
//          irow := irow + 1;
//          snumber := ExcelApp.Cells[irow, CINumber].Value;
//        end;
//      end;
//    finally
//      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
//      WorkBook.Close;
//    end;
//  finally
//    ExcelApp.Visible := True;
//    ExcelApp.Quit; 
//  end;
//end;

procedure TfrmManMrp.btnNumberListClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options - [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  leNumberList.Text := OpenDialog1.FileName;
end;
     
procedure TfrmManMrp.SortMPS(aProjInfo: TProjInfo; aMPSSorter: TMPSSorter);
var
  iMPS: Integer;
  aMPSInfoPtr: PMPSInfo;
begin
  for iMPS := 0 to aProjInfo.FList.Count - 1 do
  begin
    aMPSInfoPtr := aProjInfo.FList[iMPS];
    aMPSSorter.Add(aMPSInfoPtr, aProjInfo.FDates);
  end;
end;

function StringListSortCompare_proj(List: TStringList; Index1, Index2: Integer): Integer;
var
  aProjInfo1, aProjInfo2: TProjInfo;
begin
  aProjInfo1 := TProjInfo(List.Objects[Index1]);
  aProjInfo2 := TProjInfo(List.Objects[Index2]);
  if aProjInfo1.FIndex > aProjInfo2.FIndex then
    Result := -1
  else if aProjInfo1.FIndex = aProjInfo2.FIndex then
    Result := 0
  else // <<<<
    Result := 1;
end;

procedure TfrmManMrp.ToolButton2Click(Sender: TObject);
var
  i: Integer;
  aProjInfo: TProjInfo;
  aMPSSorter: TMPSSorter;
  ExcelApp, WorkBook: Variant;
  iProj: Integer;
  dt: TDateTime;
begin         
  bMan := (Sender as TToolButton).Caption = '汇总手工';

  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';
  dt := Now;
  dt := dt - DayOfWeek(dt) - 1;
  if bMan then
  begin
    SaveDialog1.FileName := '物料需求进度表(' + FormatDateTime('yyyy-MM-dd', dt) + ' week' + IntToStr(WeekOf(dt)) + ').xlsx';
  end
  else
  begin
    SaveDialog1.FileName := 'MPS明细表(' + FormatDateTime('yyyy-MM-dd', dt) + ' week' + IntToStr(WeekOf(dt)) + ').xlsx';
  end;


  //SaveDialog1.FileName := 'MPSSum' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;

  Clear;

  for i := 0 to lvMPS.Items.Count - 1 do
  begin
    ReadMPS( lvMPS.Items[i].Caption );
  end;

  FMPSProjs.SortDate;

  FMPSProjs.FList.CustomSort(StringListSortCompare_proj);
 
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

  while ExcelApp.Sheets.Count > 1 do
  begin
    ExcelApp.Sheets[2].Delete;
  end;

  try
    ProgressBar1.Max := FMPSProjs.FList.Count;
    ProgressBar1.Position := 1;
    
    for iProj := 0 to FMPSProjs.FList.Count - 1 do
    begin
      aProjInfo := TProjInfo(FMPSProjs.FList.Objects[iProj]);

      if iProj > 0 then
      begin
        WorkBook.Sheets.Add;
      end;
      ExcelApp.Sheets[1].Activate;
      ExcelApp.Sheets[1].Name := aProjInfo.FName;


      aMPSSorter := TMPSSorter.Create;

      SortMPS(aProjInfo, aMPSSorter);
      SaveMPS(ExcelApp, aMPSSorter, aProjInfo.FDates);

      aMPSSorter.Free;

      ProgressBar1.Position := ProgressBar1.Position + 1;
    end;

    WorkBook.SaveAs(SaveDialog1.FileName);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
 
  finally
    WorkBook.Close;
    ExcelApp.Quit;
  end;

  MessageBox(Handle, '完成', '金蝶提示', 0);
end;
                        
function TfrmManMrp.FileNameExists(const sfile: string;
  lv: TListView): Boolean;
var
  i: Integer;
  item: TListItem;
begin
  Result := False;
  for i := 0 to lv.Items.Count - 1 do
  begin
    item := lv.Items[i];
    if item.Caption = sfile then
    begin
      Result := True;
      Break;
    end;
  end;
end;

procedure TfrmManMrp.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

procedure TfrmManMrp.AddFile(lv: TListView);
var
  ifile: Integer;
  li: TListItem;
begin
  OpenDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  OpenDialog1.FilterIndex := 0;
  OpenDialog1.DefaultExt := '.xlsx';
  OpenDialog1.Options := OpenDialog1.Options + [ofAllowMultiSelect];
  if not OpenDialog1.Execute then Exit;
  for ifile := 0 to OpenDialog1.Files.Count - 1 do
  begin
    if FileNameExists(OpenDialog1.Files[ifile], lv) then Continue;
    li := lv.Items.Add;
    li.Caption := OpenDialog1.Files[ifile];
  end; 
end;

procedure TfrmManMrp.RemoveFile(lv: TListView);
begin
  if lv.Selected = nil then Exit;
  lv.DeleteSelected;
end;

procedure TfrmManMrp.SpeedButton1Click(Sender: TObject);
begin
  AddFile(lvMPS);
end;

procedure TfrmManMrp.SpeedButton2Click(Sender: TObject);
begin
  RemoveFile(lvMPS);
end;
   
procedure TfrmManMrp.ReadMPS(const sFile: string);
const
  CINumber = 3;   //产品编码
  CIName = 15;     //产品名称
  CIDate = 6;     //需求日期
  CIQty = 5;     	//数量
  CIProj = 14;     //项目

  CIColor = 16;
  CICap = 17;
  CIVer = 13;
  CIFG = 18;
  CIPKG = 19;
            
var
  ExcelApp, WorkBook: Variant;
  snumber: string;   
  sdt: string; 
  aMPSInfoPtr: PMPSInfo;
  irow: Integer;
  aProjInfo: TProjInfo;
  iSheetCount, iSheet: Integer;
  iIndex: Integer; 
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
    
        irow := 2;
        snumber := ExcelApp.Cells[irow, CINumber].Value;
        while snumber <> '' do
        begin
          aMPSInfoPtr := New(PMPSInfo);

          aMPSInfoPtr^.Number := snumber;     
          aMPSInfoPtr^.Name := ExcelApp.Cells[irow, CIName].Value;
          sdt := ExcelApp.Cells[irow, CIDate].Value;
          sdt := StringReplace(sdt, '/', '-', [rfReplaceAll]);
          aMPSInfoPtr^.Date := StrToDate(sdt);
          aMPSInfoPtr^.Qty := ExcelApp.Cells[irow, CIQty].Value;

          aMPSInfoPtr^.Proj := ExcelApp.Cells[irow, CIProj].Value;
          aMPSInfoPtr^.Color := ExcelApp.Cells[irow, CIColor].Value;
          aMPSInfoPtr^.Cap := ExcelApp.Cells[irow, CICap].Value;
          aMPSInfoPtr^.Ver := ExcelApp.Cells[irow, CIVer].Value;
          aMPSInfoPtr^.FG := ExcelApp.Cells[irow, CIFG].Value;
          aMPSInfoPtr^.Pkg := ExcelApp.Cells[irow, CIPKG].Value;

          iIndex := mmoYearOfProj.Lines.IndexOfName(aMPSInfoPtr^.Proj);
          aProjInfo := FMPSProjs.GetProjInfo(aMPSInfoPtr^.Proj, iIndex);
          aProjInfo.Add(aMPSInfoPtr, sdt); 
        
          irow := irow + 1;
          snumber := ExcelApp.Cells[irow, CINumber].Value;

          if aMPSInfoPtr^.Proj = '' then
          begin
            Memo1.Lines.Add('项目 不能为空  irow: ' + IntToStr(irow) + '   icol: ' + IntToStr(CIProj));          
            raise Exception.Create('项目 不能为空  irow: ' + IntToStr(irow) + '   icol: ' + IntToStr(CIProj));
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

procedure SetOrderDateColor(ExcelApp: Variant; irow, icol: Integer; dtOrder: TDateTime);
var
  dt: TDateTime;
begin
  dt := ExcelApp.Cells[irow, icol].Value;
  if dt < dtOrder - 7 then
  begin
    ExcelApp.Cells[irow, iCol].Interior.Color := $EEEEEE;
  end
  else if dt <= dtOrder then
  begin
    ExcelApp.Cells[irow, iCol].Interior.Color := $00C0FF;
  end;
end;

function ExtractCap(const scap: string): Integer;
var
  s: string;
begin
  s := scap;
  s := UpperCase(s);
  if Pos('+', s) > 0 then
  begin
    s := Copy(s, Pos('+', s) + 1, Length(s));
  end;
  if Pos(' ', s) > 0 then
  begin
    s := Copy(s, Pos(' ', s) + 1, Length(s));
  end;
//  s := StringReplace(s, );
end;

function StringListSortCompare_Cap(List: TStringList; Index1, Index2: Integer): Integer;
var
  scap1, scap2: string;
  icap1, icap2: Integer;
begin
  scap1 := List[Index1];
  scap2 := List[Index2];

  icap1 := ExtractCap(scap1);
  icap2 := ExtractCap(scap2);
end;

procedure TfrmManMrp.SaveMPS(ExcelApp: Variant; aMPSSorter: TMPSSorter;
  slDate: TStringList);
var
  iVer: Integer;
  iColor: Integer;
  iCap: Integer;
  iFG: Integer;
  iPkg: Integer;
  iNumber: Integer;
  iDate: Integer;
  aVerSorter: TVerSorter;
  aColorSorter: TColorSorter;
  aCapSorter: TCapSorter;
  aFGSorter: TFGSorter;
  aPkgSorter: TPkgSorter;
  aNumberSorter: TNumberSorter;
  aDateSorter: TDateSorter;
  irow, icol: Integer;
  irowSum: Integer;
  irow1, irow2: Integer;
  irow_e: Integer;
  KeySumer_Color: TKeySumer;
  KeySumer_Ver: TKeySumer;
  KeySumer_Cap: TKeySumer;
  KeySumer_FG: TKeySumer;
  KeySumer_Pkg: TKeySumer;
  i: Integer;
  ssum: string; 
  iFreezeCol, iFreezeRow: Integer;
  scell: string;
begin
  irow_e := 0;
  
  KeySumer_Color := TKeySumer.Create;
  KeySumer_Ver := TKeySumer.Create;
  KeySumer_Cap := TKeySumer.Create;
  KeySumer_FG := TKeySumer.Create;
  KeySumer_Pkg := TKeySumer.Create;

  irow := 1;

  for iVer := 0 to aMPSSorter.FVers.Count - 1 do
  begin
    aVerSorter := TVerSorter(aMPSSorter.FVers[iVer]);
    for iColor := 0 to aVerSorter.FColors.Count - 1 do
    begin
      aColorSorter := TColorSorter(aVerSorter.FColors[iColor]);
      for iCap := 0 to aColorSorter.FCaps.Count - 1 do
      begin
        aCapSorter := TCapSorter(aColorSorter.FCaps[iCap]);
        for iFG := 0 to aCapSorter.FFGs.Count - 1 do
        begin
          aFGSorter := TFGSorter(aCapSorter.FFGs[iFG]);
          for iPkg := 0 to aFGSorter.FPkgs.Count - 1 do
          begin
            aPkgSorter := TPkgSorter(aFGSorter.FPkgs[iPkg]);
            for iNumber := 0 to aPkgSorter.FNumbers.Count - 1 do
            begin
              aNumberSorter := TNumberSorter(aPkgSorter.FNumbers[iNumber]);

              if irow = 1 then
              begin
                ExcelApp.Cells[irow, 1].Value := '制式';  
                ExcelApp.Columns[1].ColumnWidth := 10;
                ExcelApp.Cells[irow, 2].Value := '物料长代码';
                ExcelApp.Columns[2].ColumnWidth := 16;
                ExcelApp.Cells[irow, 3].Value := '物料名称';    
                ExcelApp.Columns[3].ColumnWidth := 40;
                ExcelApp.Cells[irow, 4].Value := '颜色';
                ExcelApp.Cells[irow, 5].Value := '容量';
                ExcelApp.Cells[irow, 6].Value := '整机/裸机';
                ExcelApp.Cells[irow, 7].Value := '豪华装';      
                ExcelApp.Columns[7].ColumnWidth := 17;
                icol := 8;

                for iDate := 0 to aNumberSorter.FDates.Count - 1 do
                begin
                  aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
                  ExcelApp.Cells[irow, icol + iDate].Value := FormatDateTime('yyyy-MM-DD', aDateSorter.FDateStart);
                  //slDate.Add(FormatDateTime('yyyy-MM-DD', aDateSorter.FDateStart));

                  ExcelApp.Columns[icol + iDate].ColumnWidth := 8;
                end;            
                icol := icol + aNumberSorter.FDates.Count;     
                ExcelApp.Cells[irow, icol].Value := '合计';    
                ExcelApp.Cells[irow, icol].HorizontalAlignment := xlCenter;

                ExcelApp.Range[ExcelApp.Cells[irow, 1], ExcelApp.Cells[irow, icol] ].Interior.Color := $B4D5FC;

                irow := irow + 1;
              end;
                
              ExcelApp.Cells[irow, 1].Value := aVerSorter.FVer;
              ExcelApp.Cells[irow, 2].Value := aNumberSorter.FNumber;
              ExcelApp.Cells[irow, 3].Value := aNumberSorter.FName;
              ExcelApp.Cells[irow, 4].Value := aColorSorter.FColor;
              ExcelApp.Cells[irow, 5].Value := aCapSorter.FCap;
              ExcelApp.Cells[irow, 6].Value := aFGSorter.FFG;
              ExcelApp.Cells[irow, 7].Value := aPkgSorter.FPkg;  

              ssum := '=0';
              icol := 8;
              for iDate := 0 to aNumberSorter.FDates.Count - 1 do
              begin
                aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
                ExcelApp.Cells[irow, icol + iDate].Value := aDateSorter.FQty;     

                KeySumer_Color.Add(aColorSorter.FColor, aDateSorter, slDate);
                KeySumer_Ver.Add(aVerSorter.FVer, aDateSorter, slDate);
                KeySumer_Cap.Add(aCapSorter.FCap, aDateSorter, slDate);
                KeySumer_FG.Add(aFGSorter.FFG, aDateSorter, slDate);
                KeySumer_Pkg.Add(aPkgSorter.FPkg, aDateSorter, slDate);

                ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow);
              end;                   
              icol := icol + aNumberSorter.FDates.Count;
              ExcelApp.Cells[irow, icol].Value := ssum;
                
              irow := irow + 1;
            end;
          end;
        end;
      end;
    end;
  end;                                                                    
  ExcelApp.Range[ ExcelApp.Cells[2, 8], ExcelApp.Cells[irow-1, slDate.Count + 8] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';   
  ExcelApp.Range[ ExcelApp.Cells[1, 8], ExcelApp.Cells[irow-1, slDate.Count + 8] ].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, slDate.Count + 8] ].Borders.LineStyle := 1; //加边框

  irow := irow + 1;
  irowSum := irow;
  ExcelApp.Cells[irow, 6].Value := '市场预测';       
  ExcelApp.Cells[irow, 7].Value := '需求日期'; 

  if bMan then
  begin
    ExcelApp.Cells[irow+1, 6].Value := '物料需求';
    ExcelApp.Range[ExcelApp.Cells[irow+1, 6], ExcelApp.Cells[irow+2, 6]].MergeCells := True;

    ExcelApp.Cells[irow+1, 7].Value := '电子件到料时间';
    ExcelApp.Cells[irow+2, 7].Value := '结构件到料时间';


    irow_e := irow + 1;

    iCol := 8;
    for iDate := 0 to slDate.Count -1  do
    begin
      ExcelApp.Cells[irow, iCol + iDate].Value := slDate[iDate];
      ExcelApp.Cells[irow+1, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow) + '-10';
      ExcelApp.Cells[irow+2, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow) + '-8';
    end;
    icol := icol + slDate.Count;
    ExcelApp.Cells[irow, iCol].Value := '合计';
    ExcelApp.Cells[irow, icol].Value := '合计';
    ExcelApp.Cells[irow, icol].HorizontalAlignment := xlCenter;
    ExcelApp.Range[ExcelApp.Cells[irow, icol], ExcelApp.Cells[irow + 2, icol] ].MergeCells := True;

    ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow, icol] ].Interior.Color := $B4D5FC;

    ExcelApp.Range[ExcelApp.Cells[irow+1, 6], ExcelApp.Cells[irow+2, icol] ].Interior.Color := $00FFFF;
          
    irow := irow + 2;
  end
  else
  begin
    iCol := 8;
    for iDate := 0 to slDate.Count -1  do
    begin
      ExcelApp.Cells[irow, iCol + iDate].Value := slDate[iDate];
    end;
    icol := icol + slDate.Count;
    ExcelApp.Cells[irow, iCol].Value := '合计';
    ExcelApp.Cells[irow, icol].Value := '合计';
    ExcelApp.Cells[irow, icol].HorizontalAlignment := xlCenter;

    ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow, icol] ].Interior.Color := $B4D5FC;
  end;

  irow := irow + 1;

  ExcelApp.Cells[irow, 6].Value := '颜色';
  ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_Color.FList.Count - 1, 6]].MergeCells := True;
  for i := 0 to KeySumer_Color.FList.Count - 1 do
  begin       
    ssum := '=0';
    ExcelApp.Cells[irow + i, 7].Value := KeySumer_Color.FList[i];
    aNumberSorter := TNumberSorter(KeySumer_Color.FList.Objects[i]);
    iCol := 8;
    for iDate := 0 to aNumberSorter.FDates.Count - 1 do
    begin
      aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
      ExcelApp.Cells[irow + i, iCol + iDate].Value := aDateSorter.FQty;    
      ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow + i);
    end;      
    icol := icol + aNumberSorter.FDates.Count;
    ExcelApp.Cells[irow + i, iCol].Value := ssum;
  end;

  irow := irow + KeySumer_Color.FList.Count;

  ExcelApp.Cells[irow, 6].Value := '制式';
  ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_Ver.FList.Count - 1, 6]].MergeCells := True;      
  for i := 0 to KeySumer_Ver.FList.Count - 1 do
  begin    
    ssum := '=0';
    ExcelApp.Cells[irow + i, 7].Value := KeySumer_Ver.FList[i];
    aNumberSorter := TNumberSorter(KeySumer_Ver.FList.Objects[i]);
    iCol := 8;
    for iDate := 0 to aNumberSorter.FDates.Count - 1 do
    begin
      aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
      ExcelApp.Cells[irow + i, iCol + iDate].Value := aDateSorter.FQty;
      ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow + i);
    end;
    icol := icol + aNumberSorter.FDates.Count;
    ExcelApp.Cells[irow + i, iCol].Value := ssum;

    if i = 0 then
    begin
      ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_Ver.FList.Count - 1, aNumberSorter.FDates.Count + 7]].Interior.Color := $EEEEEE;
    end;
  end;

  irow := irow + KeySumer_Ver.FList.Count;


//  KeySumer_Cap.FList.CustomSort(StringListSortCompare_Cap);

  ExcelApp.Cells[irow, 6].Value := '容量';
  ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_Cap.FList.Count - 1, 6]].MergeCells := True;
  for i := 0 to KeySumer_Cap.FList.Count - 1 do
  begin                     
    ssum := '=0';
    ExcelApp.Cells[irow + i, 7].Value := KeySumer_Cap.FList[i];
    aNumberSorter := TNumberSorter(KeySumer_Cap.FList.Objects[i]);
    iCol := 8;
    for iDate := 0 to aNumberSorter.FDates.Count - 1 do
    begin
      aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
      ExcelApp.Cells[irow + i, iCol + iDate].Value := aDateSorter.FQty;  
      ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow + i);
    end;
    icol := icol + aNumberSorter.FDates.Count;
    ExcelApp.Cells[irow + i, iCol].Value := ssum;
  end;

  irow := irow + KeySumer_Cap.FList.Count;

  ExcelApp.Cells[irow, 6].Value := '整机/裸机';
  ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_FG.FList.Count - 1, 6]].MergeCells := True;
  for i := 0 to KeySumer_FG.FList.Count - 1 do
  begin     
    ssum := '=0';
    ExcelApp.Cells[irow + i, 7].Value := KeySumer_FG.FList[i];
    aNumberSorter := TNumberSorter(KeySumer_FG.FList.Objects[i]);
    iCol := 8;
    for iDate := 0 to aNumberSorter.FDates.Count - 1 do
    begin
      aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
      ExcelApp.Cells[irow + i, iCol + iDate].Value := aDateSorter.FQty;   
      ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow + i);
    end;                      
    icol := icol + aNumberSorter.FDates.Count;
    ExcelApp.Cells[irow + i, iCol].Value := ssum;

    if i = 0 then
    begin
      ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_FG.FList.Count - 1, aNumberSorter.FDates.Count + 7]].Interior.Color := $EEEEEE;
    end;
  end;
          
  irow := irow + KeySumer_FG.FList.Count;

  irow1 := irow;

  ExcelApp.Cells[irow, 6].Value := '礼盒装';
  ExcelApp.Range[ExcelApp.Cells[irow, 6], ExcelApp.Cells[irow + KeySumer_Pkg.FList.Count - 1, 6]].MergeCells := True;
  for i := 0 to KeySumer_Pkg.FList.Count - 1 do
  begin      
    ssum := '=0';
    ExcelApp.Cells[irow + i, 7].Value := KeySumer_Pkg.FList[i];
    aNumberSorter := TNumberSorter(KeySumer_Pkg.FList.Objects[i]);
    iCol := 8;
    for iDate := 0 to aNumberSorter.FDates.Count - 1 do
    begin
      aDateSorter := TDateSorter(aNumberSorter.FDates[iDate]);
      ExcelApp.Cells[irow + i, iCol + iDate].Value := aDateSorter.FQty;   
      ssum := ssum + '+' + GetRef(icol + iDate) + IntToStr(irow + i);
    end;    
    icol := icol + aNumberSorter.FDates.Count;
    ExcelApp.Cells[irow + i, iCol].Value := ssum;
  end;

  irow := irow + KeySumer_Pkg.FList.Count;

  ExcelApp.Range[ ExcelApp.Cells[irowSum + 3, 8], ExcelApp.Cells[irow, slDate.Count + 8] ].NumberFormatLocal := '_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""??_ ;_ @_ ';

  ExcelApp.Range[ ExcelApp.Cells[irowSum, 6], ExcelApp.Cells[irow-1, slDate.Count + 8] ].Borders.LineStyle := 1; //加边框
        
  irow2 := irow - 1;


  if bMan then
  begin                                    
    ExcelApp.Cells[irow, 7].Value := '物料需求';
    ExcelApp.Cells[irow + 1, 7].Value := '累计物料需求';
    
    ExcelApp.Cells[irow + 2, 7].Value := '下需求时间(L/T=2W)';
    ExcelApp.Cells[irow + 3, 7].Value := '下需求时间(L/T=3W)';
    ExcelApp.Cells[irow + 4, 7].Value := '下需求时间(L/T=4W)';
    ExcelApp.Cells[irow + 5, 7].Value := '下需求时间(L/T=5W)';
    ExcelApp.Cells[irow + 6, 7].Value := '下需求时间(L/T=6W)';
    ExcelApp.Cells[irow + 7, 7].Value := '下需求时间(L/T=7W)';
    ExcelApp.Cells[irow + 8, 7].Value := '下需求时间(L/T=8W)';
    ExcelApp.Cells[irow + 9, 7].Value := '下需求时间(L/T=9W)';
    ExcelApp.Cells[irow + 10, 7].Value := '下需求时间(L/T=10W)';
  end
  else
  begin
    ExcelApp.Cells[irow, 7].Value := 'MPS';
    ExcelApp.Cells[irow + 1, 7].Value := '累计MPS';
  end;

  icol := icol - 1;
  ExcelApp.Range[ExcelApp.Cells[irow + 1, 7], ExcelApp.Cells[irow + 1, icol] ].Interior.Color := $00FFFF;
  if bMan then
  begin
    ExcelApp.Range[ExcelApp.Cells[irow + 6, 7], ExcelApp.Cells[irow + 6, icol] ].Interior.Color := $00FFFF;
  end;


  iCol := 8;
  for iDate := 0 to slDate.Count -1  do
  begin
    ExcelApp.Cells[irow, iCol + iDate].Value := '=SUM(' + GetRef(icol + iDate) + IntToStr(irow1) + ':' + GetRef(icol + iDate) + IntToStr(irow2) + ')';      
    if iDate = 0 then
    begin
      ExcelApp.Cells[irow + 1, iCol + iDate].Value := '=SUM(' + GetRef(icol + iDate) + IntToStr(irow1) + ':' + GetRef(icol + iDate) + IntToStr(irow2) + ')';
    end
    else
    begin
      ExcelApp.Cells[irow + 1, iCol + iDate].Value := '=' + GetRef(iCol + iDate - 1) + IntToStr(irow + 1) + '+SUM(' + GetRef(icol + iDate) + IntToStr(irow1) + ':' + GetRef(icol + iDate) + IntToStr(irow2) + ')';
    end;
    if bMan then
    begin
      ExcelApp.Cells[irow + 2, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-2*7-3';
      ExcelApp.Cells[irow + 3, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-3*7-3';
      ExcelApp.Cells[irow + 4, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-4*7-3';
      ExcelApp.Cells[irow + 5, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-5*7-3';
      ExcelApp.Cells[irow + 6, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-6*7-3';
      ExcelApp.Cells[irow + 7, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-7*7-3';
      ExcelApp.Cells[irow + 8, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-8*7-3';
      ExcelApp.Cells[irow + 9, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-9*7-3';
      ExcelApp.Cells[irow + 10, iCol + iDate].Value := '=' + GetRef(icol + iDate) + IntToStr(irow_e) + '-10*7-3';

      SetOrderDateColor(ExcelApp, irow + 2, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 3, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 4, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 5, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 6, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 7, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 8, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 9, icol + iDate, dtpStart.DateTime);
      SetOrderDateColor(ExcelApp, irow + 10, icol + iDate, dtpStart.DateTime);
    end;
  end;

  if bMan then
  begin
    ExcelApp.Range[ ExcelApp.Cells[irow + 2, 8], ExcelApp.Cells[irow + 10, slDate.Count + 7] ].NumberFormat := 'yyyy-mm-dd;@';

    ExcelApp.Range[ ExcelApp.Cells[irow, 7], ExcelApp.Cells[irow + 10, slDate.Count + 7] ].Borders.LineStyle := 1; //加边框

    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow + 10, slDate.Count + 8] ].Font.Size := 9;
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow + 10, slDate.Count + 8] ].Font.Name := 'Calibri';
  end
  else
  begin
    ExcelApp.Range[ ExcelApp.Cells[irow, 7], ExcelApp.Cells[irow + 1, slDate.Count + 7] ].Borders.LineStyle := 1; //加边框

    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow + 1, slDate.Count + 8] ].Font.Size := 9;
    ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[irow + 1, slDate.Count + 8] ].Font.Name := 'Calibri';
  end;

  
  KeySumer_Color.Free;
  KeySumer_Ver.Free;
  KeySumer_Cap.Free;
  KeySumer_FG.Free;
  KeySumer_Pkg.Free;
                   
//  ExcelApp.Range[ ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 1] ].Select;
//  iFreezeCol := 9;      
//  iFreezeRow := 3; 
//  ExcelApp.Range[ ExcelApp.Cells[iFreezeRow, iFreezeCol], ExcelApp.Cells[iFreezeRow, iFreezeCol] ].Select; 
//  ExcelApp.ActiveWindow.FreezePanes := True

end;

procedure TfrmManMrp.dtpStartChange(Sender: TObject);
begin
  dtpStart2.DateTime := dtpStart.DateTime - 7;
end;

procedure TfrmManMrp.tbQuitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmManMrp.tbOEM2Click(Sender: TObject);
var
  sfile: string;
  ExcelApp, WorkBook: Variant; 
  aSOPReader: TSOPReader;
  irow: Integer;
  iProj: Integer;
  aSOPProj: TSOPProj;
  iLine: Integer;    
  aSOPLine: SOPReaderUnit.TSOPLine;
  iDate: Integer;
  aSOPCol: TSOPCol;
  //aNumberInfoPtr: PNumberInfo;
  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  sPkgArea: string;
begin
  Memo1.Lines.Add('-------------  BEGIN  ----------------------------');
  SaveDialog1.Filter := 'Excel Files|*.xls;*.xlsx';
  SaveDialog1.FilterIndex := 0;
  SaveDialog1.DefaultExt := '.xlsx';

  SaveDialog1.FileName := 'Row' + FormatDateTime('yyyyMMddhhmmss', Now) + '.xlsx';
  if not SaveDialog1.Execute then Exit;
  sfile := SaveDialog1.FileName;


  FSAPMaterialReader := TSAPMaterialReader2.Create(leNumberList.Text, nil);
  //ReadNumberList(leNumberList.Text);


  aSOPReader := TSOPReader.Create(TStringList( mmoYearOfProj.Lines ),
    leSOP.Text, OnLogEvent);

  
  try
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

    while WorkBook.Sheets.Count > 1 do
    begin
      WorkBook.Sheets[2].Delete;
    end;

    WorkBook.Sheets[1].Name := '产品预测单';
  

    irow := 1;


    ExcelApp.Cells[irow, 1].Value := '编号*';
    ExcelApp.Cells[irow, 2].Value := '需求类型*';
    ExcelApp.Cells[irow, 3].Value := '物料长编码*';
    ExcelApp.Cells[irow, 4].Value := '单位*';
    ExcelApp.Cells[irow, 5].Value := '数量*';
    ExcelApp.Cells[irow, 6].Value := '预测开始日期*';
    ExcelApp.Cells[irow, 7].Value := '预测截止日期*';
                                                    
    ExcelApp.Cells[irow, 8].Value := '均化周期类型*';

    ExcelApp.Cells[irow, 9].Value := '源单类型*';
    ExcelApp.Cells[irow, 10].Value := '源单号*';
    ExcelApp.Cells[irow, 11].Value := '源单行号*';
    ExcelApp.Cells[irow, 12].Value := '备注*';
    ExcelApp.Cells[irow, 13].Value := '备注2';
  
    ExcelApp.Cells[irow, 14].Value := '项目';    
    ExcelApp.Cells[irow, 15].Value := '物料名称';
  
    ExcelApp.Cells[irow, 16].Value := '颜色';
    ExcelApp.Cells[irow, 17].Value := '容量';
    ExcelApp.Cells[irow, 18].Value := '版本';
    ExcelApp.Cells[irow, 19].Value := '包装';
    
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 19]].HorizontalAlignment := xlCenter; //居中

    irow := irow + 1;
  

    for iProj := 0 to aSOPReader.ProjCount - 1 do
    begin
      aSOPProj := aSOPReader.Projs[iProj];
      for iLine := 0 to aSOPProj.LineCount - 1 do
      begin
        aSOPLine := aSOPProj.Lines[iLine];

        //aNumberInfoPtr := GetNumberInfo(aSOPLine.sNumber);
        aSAPMaterialRecordPtr := FSAPMaterialReader.GetSAPMaterialRecord(aSOPLine.sNumber);
        
        for iDate := 0 to aSOPLine.DateCount - 1 do
        begin
          aSOPCol := aSOPLine.Dates[iDate];
          sPkgArea := aSOPLine.sPkg;
          if aSOPLine.sMRPArea <> '' then
          begin
            sPkgArea := sPkgArea + Copy(aSOPLine.sMRPArea, 1, 2);
          end;

          ExcelApp.Cells[irow, 1].Value := '';
          if Pos('售后', aSOPLine.sStdVer) > 0 then
          begin
            ExcelApp.Cells[irow, 2].Value := '售后';
          end
          else if Pos('海外', aSOPLine.sStdVer) > 0 then
          begin
            ExcelApp.Cells[irow, 2].Value := '海外量产';
          end
          else
          begin
            ExcelApp.Cells[irow, 2].Value := '国内量产';
          end;

          if aSOPLine.sNumber = '' then
          begin
            ExcelApp.Cells[irow, 3].Value := aSOPProj.FName + ',' + aSOPLine.sStdVer + ',' + aSOPLine.sColor + ',' + aSOPLine.sCap;
          end
          else
          begin
            ExcelApp.Cells[irow, 3].Value := aSOPLine.sNumber;  
          end;
          ExcelApp.Cells[irow, 4].Value := 'PCS';
          ExcelApp.Cells[irow, 5].Value := aSOPCol.iQty;
          ExcelApp.Cells[irow, 6].Value := aSOPCol.dt1;
          ExcelApp.Cells[irow, 7].Value := aSOPCol.dt2;

          ExcelApp.Cells[irow, 8].Value := '不均化';

          ExcelApp.Cells[irow, 9].Value := '';
          ExcelApp.Cells[irow, 10].Value := '';
          ExcelApp.Cells[irow, 11].Value := '';
          ExcelApp.Cells[irow, 12].Value := aSOPProj.FName;
          ExcelApp.Cells[irow, 13].Value := aSOPLine.sStdVer;
          ExcelApp.Cells[irow, 14].Value := aSOPProj.FName;
          if aSAPMaterialRecordPtr <> nil then
          begin
            ExcelApp.Cells[irow, 15].Value := aSAPMaterialRecordPtr.sName;
          end;
          ExcelApp.Cells[irow, 16].Value := aSOPLine.sColor;
          ExcelApp.Cells[irow, 17].Value := aSOPLine.sCap;
          ExcelApp.Cells[irow, 18].Value := aSOPLine.sFG;
          ExcelApp.Cells[irow, 19].Value := sPkgArea;

          irow := irow + 1;
        end;;
      end;
    end;


    ExcelApp.Columns[1].ColumnWidth := 16;
    ExcelApp.Columns[3].ColumnWidth := 16;
    ExcelApp.Columns[6].ColumnWidth := 13;
    ExcelApp.Columns[7].ColumnWidth := 13;
    ExcelApp.Columns[12].ColumnWidth := 16;
    ExcelApp.Columns[13].ColumnWidth := 16;
    ExcelApp.Columns[14].ColumnWidth := 16;   
    ExcelApp.Columns[15].ColumnWidth := 16;
    ExcelApp.Columns[16].ColumnWidth := 16;
    ExcelApp.Columns[17].ColumnWidth := 16;
    ExcelApp.Columns[18].ColumnWidth := 16;     
    ExcelApp.Columns[19].ColumnWidth := 16;

  
    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[irow-1, 19]].Borders.LineStyle := 1; //加边框

    WorkBook.SaveAs(sfile);
    ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    WorkBook.Close;
    ExcelApp.Quit;     

    lvMPS.Clear;
    lvMPS.Items.Add.Caption := sfile;
                    

    MessageBox(Handle, '完成', '金蝶提示', 0);
  finally
//    aSOPWriter.Free;
    aSOPReader.Free;
    FSAPMaterialReader.Free;
  end;
end;

end.


