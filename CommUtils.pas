unit CommUtils;

interface

uses
  Classes, Messages, Windows, Dialogs, Forms, SysUtils, ExtCtrls;


const
  xlCenter = -4108;

    
type
  
  TLogEvent = procedure(const s: string) of object;

  TOEMSOPvsDemandSet_OEM = (
    saDemandChange_OEM, //'市场需求变化',
    saDemandOutOfSupply_OEM, //'市场需求脱离实际供应能力',
    saFacCap_OEM, //'代工厂产能',
    saMatE_OEM,   //'物料供应-电子件',
    saMatS_OEM,   //'物料供应-机构件',
    saMatI_OEM,   //'物料供应-间接物料'
    saPC_OEM,
    saDesignAndFlymeECN_OEM     //'开发设计/Flyme ECN变更'
  );

  TOEMACTvsDemandSet_OEM = (
    sbDemandChange_OEM, //'市场需求变化',
    sbFacCap_OEM, //'代工厂产能',
    sbFacMan_OEM, //'代工厂执行力',
    sbMatE_OEM,   //'物料供应-电子件',
    sbMatS_OEM,   //'物料供应-机构件',
    sbMatI_OEM,   //'物料供应-间接物料',
    sbMPlanE_OEM, //'物料计划异常',
    sbPPlanE_OEM, //'生产计划异常',
    sbArtE_OEM,   //'制程品质异常',
    sbwwArtE_OEM, //'委外工艺',
    sbDesignAndFlymeECN_OEM     //'开发设计及ECN变更'
  );

  TOEMACTvsSchSet_OEM = (
    scDemandChange_OEM, //'市场需求变化',
    scFacCap_OEM, //'代工厂产能',
    scFacMan_OEM, //'代工厂执行力',
    scMatE_OEM, //'物料供应-电子件',
    scMatS_OEM, //'物料供应-机构件',
    scMatI_OEM, //'物料供应-间接物料',
    scMPlanE_OEM, //'物料计划异常',
    scPPlanE_OEM, //'生产计划异常',
    scArtE_OEM, //'制程品质异常',
    scwwArtE_OEM, //'委外工艺',
    scDesignAndFlymeECN_OEM     //'开发设计及ECN变更'
  );
     
  TOEMSOPvsDemandSet_ODM = (
    saDemandChange_ODM, //'魅族市场需求变化',
    saMZFirm_ODM, //'魅族固件',
    saMZMat_ODM,  //'魅族客供料交付未达成',
    saMZMatQ_ODM, //'魅族客供料品质异常',
    saFacCap_ODM, //'代工厂产能',
    saFacMan_ODM, //'代工厂计划执行力',
    saMatE_ODM,   //'代工厂物料-电子料',
    saMatS_ODM,   //'代工厂物料-结构料',
    saMatI_ODM,   //'代工厂物料-包材',
    saArtE_ODM    //'代工厂制程品质异常'
  );

  TOEMACTvsDemandSet_ODM = (
    sbDemandChange_ODM, //'魅族市场需求变化',
    sbMZFirm_ODM, //'魅族固件',
    sbMZMat_ODM,  //'魅族客供料交付未达成',
    sbMZMatQ_ODM, //'魅族客供料品质异常',
    sbFacCap_ODM, //'代工厂产能',
    sbFacMan_ODM, //'代工厂计划执行力',
    sbMatE_ODM,   //'代工厂物料-电子料',
    sbMatS_ODM,   //'代工厂物料-结构料',
    sbMatI_ODM,   //'代工厂物料-包材',
    sbArtE_ODM    //'代工厂制程品质异常'
  );

  TOEMACTvsSchSet_ODM = (
    scDemandChange_ODM, //'魅族市场需求变化',
    scMZFirm_ODM, //'魅族固件',
    scMZMat_ODM,  //'魅族客供料交付未达成',
    scMZMatQ_ODM, //'魅族客供料品质异常',
    scFacCap_ODM, //'代工厂产能',
    scFacMan_ODM, //'代工厂计划执行力',
    scMatE_ODM,   //'代工厂物料-电子料',
    scMatS_ODM,   //'代工厂物料-结构料',
    scMatI_ODM,   //'代工厂物料-包材',
    scArtE_ODM    //'代工厂制程品质异常'
  );
  
const
  CSOEMSOPvsDemand_OEM: array[TOEMSOPvsDemandSet_OEM] of string = (
    '市场需求变化',
    '市场需求脱离实际供应能力',
    '代工厂产能',
    '物料供应-电子件',
    '物料供应-机构件',
    '物料供应-间接物料',
    'SOP调整',
    '开发设计/Flyme ECN变更'
  );

  CSOEMACTvsDemand_OEM: array[TOEMACTvsDemandSet_OEM] of string = (
    '市场需求变化',
    '代工厂产能',
    '代工厂执行能力',
    '物料供应-电子件',
    '物料供应-机构件',
    '物料供应-间接物料',
    '物料计划异常',
    '生产计划异常',
    '制程、物料品质异常',
    '委外工艺',
    '开发设计/Flyme ECN变更'
  );

  CSOEMACTvsSch_OEM: array[TOEMACTvsSchSet_OEM] of string = (
    '市场需求变化',
    '代工厂产能',
    '代工厂执行能力',
    '物料供应-电子件',
    '物料供应-机构件',
    '物料供应-间接物料',
    '物料计划异常',
    '生产计划异常',
    '制程、物料品质异常',
    '委外工艺',
    '开发设计/Flyme ECN变更'
  );

  CSOEMSOPvsDemand_ODM: array[TOEMSOPvsDemandSet_OdM] of string = (
    '魅族市场需求变化',
    '魅族固件',
    '魅族客供料交付未达成',
    '魅族客供料品质异常',
    '代工厂产能',
    '代工厂计划执行力',
    '代工厂物料-电子料',
    '代工厂物料-结构料',
    '代工厂物料-包材',
    '代工厂制程品质异常'
  );
  
  CSOEMACTvsDemand_ODM: array[TOEMACTvsDemandSet_ODM] of string = (
    '魅族市场需求变化',
    '魅族固件',
    '魅族客供料交付未达成',
    '魅族客供料品质异常',
    '代工厂产能',
    '代工厂计划执行力',
    '代工厂物料-电子料',
    '代工厂物料-结构料',
    '代工厂物料-包材',
    '代工厂制程品质异常'
  );

  CSOEMACTvsSch_ODM: array[TOEMACTvsSchSet_ODM] of string = (
    '魅族市场需求变化',
    '魅族固件',
    '魅族客供料交付未达成',
    '魅族客供料品质异常',
    '代工厂产能',
    '代工厂计划执行力',
    '代工厂物料-电子料',
    '代工厂物料-结构料',
    '代工厂物料-包材',
    '代工厂制程品质异常'
  );

                                                    
function IsCellMerged(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer): Boolean;
procedure MergeCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);   
procedure CenterBoldCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
procedure CenterCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
function ExcelOpenDialog(var sfile: string): Boolean;
function ExcelOpenDialogs(var sfile: string): Boolean;
function ExcelSaveDialog(var sfile: string): Boolean;   
procedure ExcelSaveDialogBtnClick(f: TForm; sender: TObject);
function myGetFileVersion(FileName: string): string; 
                                 
function GetRef(const X:Integer):string;

procedure AddBorder(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
procedure AddColor(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer; dwColor: DWORD); overload;
procedure AddColor(ExcelApp: Variant; irow1, icol1: Integer; dwColor: DWORD); overload;
procedure AddHorizontalAlignment(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer; al: Integer);

function AppIni: string;
function myStrToDateTime(const s: string): TDateTime;
                                           
function DoubleG(d1, d2: Double): Boolean;
function DoubleGE(d1, d2: Double): Boolean;    
function DoubleL(d1, d2: Double): Boolean;
function DoubleLE(d1, d2: Double): Boolean;
function DoubleE(d1, d2: Double): Boolean;

function IndexOfCol(ExcelApp: Variant; irow: Integer; const sname: string): Integer;

function IsVerHW(const sver: string): Boolean;      
function IsNameHW(const snumber, sname: string): Boolean;

procedure savelogtoexe(const s: string);

var
  gserver: string;
  guser: string;
  gpwd: string;

implementation


function IsCellMerged(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer): Boolean;
var
  vma1, vma2: Variant;
begin
  vma1 := ExcelApp.Cells[irow1, icol1].MergeArea;
  vma2 := ExcelApp.Cells[irow2, icol2].MergeArea;
  Result := vma1.Address = vma2.Address;
end;

procedure MergeCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].MergeCells := True;
end;
    
procedure CenterBoldCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].HorizontalAlignment := xlCenter;
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].Font.Bold  := True;
end;
     
procedure CenterCells(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].HorizontalAlignment := xlCenter;
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

function myGetFileVersion(FileName: string): string;
type
 PVerInfo = ^TVS_FIXEDFILEINFO; 
 TVS_FIXEDFILEINFO = record
   dwSignature: longint; 
   dwStrucVersion: longint; 
   dwFileVersionMS: longint; 
   dwFileVersionLS: longint; 
   dwFileFlagsMask: longint; 
   dwFileFlags: longint; 
   dwFileOS: longint; 
   dwFileType: longint; 
   dwFileSubtype: longint; 
   dwFileDateMS: longint; 
   dwFileDateLS: longint; 
 end; 
var 
 ExeNames: array[0..255] of char;  
 VerInfo: PVerInfo; 
 Buf: pointer; 
 Sz: word; 
 L, Len: Cardinal; 
begin 
 StrPCopy(ExeNames, FileName); 
 Sz := GetFileVersionInfoSize(ExeNames, L); 
 if Sz=0 then 
 begin 
   Result:=''; 
   Exit; 
 end; 

 try
   GetMem(Buf, Sz); 
   try 
     GetFileVersionInfo(ExeNames, 0, Sz, Buf); 
     if VerQueryValue(Buf, '\', Pointer(VerInfo), Len) then 
     begin 
       Result := IntToStr(HIWORD(VerInfo.dwFileVersionMS)) + '.' + 
                 IntToStr(LOWORD(VerInfo.dwFileVersionMS)) + '.' + 
                 IntToStr(HIWORD(VerInfo.dwFileVersionLS)) + '.' + 
                 IntToStr(LOWORD(VerInfo.dwFileVersionLS)); 

     end; 
   finally 
     FreeMem(Buf); 
   end; 
 except 
   Result := '-1'; 
 end; 
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

function ExcelOpenDialogs(var sfile: string): Boolean;
var
  i: Integer;
begin
  with TOpenDialog.Create(nil) do
  try
    Filter := 'Excel Files|*.xls;*.xlsx';
    FilterIndex := 0;
    DefaultExt := '.xlsx';
    Options := Options + [ofAllowMultiSelect];
    Result := Execute;
    if Result then
    begin
      sfile := Files[0];
      for i := 1 to Files.Count - 1 do
      begin
        sfile := sfile + ';' + Files[i];
      end;
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
    Filter := 'Excel Files|*.xlsx;*.xls';
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

procedure ExcelSaveDialogBtnClick(f: TForm; sender: TObject);
var
  sfile: string;
  s: string;
  c: TComponent;
  le: TLabeledEdit;
begin
  if not ExcelSaveDialog(sfile) then Exit;
  s := (Sender as TComponent).Name;
  s := 'le' + Copy(s, 4, Length(s));
  c := f.FindComponent(s);
  le := c as TLabeledEdit;
  le.Text := sfile;
end;

procedure AddBorder(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].Borders.LineStyle := 1; //加边框
end;

procedure AddColor(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer; dwColor: DWORD);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].Interior.Color := dwColor //加边框
end;      

procedure AddColor(ExcelApp: Variant; irow1, icol1: Integer; dwColor: DWORD);
begin
  ExcelApp.Cells[irow1, icol1].Interior.Color := dwColor //加边框
end;
          
procedure AddHorizontalAlignment(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer; al: Integer);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].HorizontalAlignment := al;
end;

function AppIni: string;
begin
  Result := ChangeFileExt(Application.ExeName, '.ini');
end;
   
function myStrToDateTime(const s: string): TDateTime;
var
  aFormatSettings: TFormatSettings;
  sdt: string;
begin
  sdt := StringReplace(s, '/', '-', [rfReplaceAll]);
  GetLocaleFormatSettings(0, aFormatSettings);
  aFormatSettings.DateSeparator := '-';
  Result := StrToDateTime(sdt, aFormatSettings);
end;

procedure savelogtoexe(const s: string);
const
 msg_log = wm_user + 123;
var
  hw: hwnd;
  ha: ATOM;
begin
  hw := FindWindow('TfrmLog', nil);
  ha := GlobalAddAtom(PChar(s));
  SendMessage(hw, msg_log, ha, 0);
end;

const
  CDZero = 0.000000001;
       
function DoubleG(d1, d2: Double): Boolean;
begin
  Result := (d1 > d2) and (Abs(d1 - d2) > CDZero);
end;
  
function DoubleGE(d1, d2: Double): Boolean;
begin
  Result := (d1 > d2) or (Abs(d1 - d2) < CDZero);
end;
      
function DoubleL(d1, d2: Double): Boolean;
begin
  Result := (d1 < d2) and (Abs(d1 - d2) > CDZero);
end;

function DoubleLE(d1, d2: Double): Boolean;
begin
  Result := (d1 < d2) or (Abs(d1 - d2) < CDZero);
end;

function DoubleE(d1, d2: Double): Boolean;
begin
  Result := Abs(d1 - d2) < CDZero;
end;

function IndexOfCol(ExcelApp: Variant; irow: Integer; const sname: string): Integer;
var
  icol: Integer;
begin
  Result := 0;
  for icol := 1 to 100 do
  begin
    if ExcelApp.Cells[irow, icol].Value = sname then
    begin
      Result := icol;
      Break;
    end;
  end;
end;
    
function IsVerHW(const sver: string): Boolean;
begin
  Result := (Pos('海外', sver) > 0) or (Pos('SKD', sver) > 0) or (Pos('CKD', sver) > 0);
end;

function IsNameHW(const snumber, sname: string): Boolean;
begin
  if Pos('国内', sname) > 0 then     // 闻泰命名
  begin
    Result := False;
    Exit;
  end;              

  if Pos('海外', sname) > 0 then     // 闻泰命名
  begin
    Result := True;
    Exit;
  end;

  // 03.53 M1792L 整个项目算作国内
  Result := (Pos('H', sname) > 0)
          or ((Pos('L',sname) > 0) and (Copy(snumber, 1, 5) <> '03.53'))
end;

end.
