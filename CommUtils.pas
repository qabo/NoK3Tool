unit CommUtils;

interface

uses
  Classes, Messages, Windows, Dialogs, Forms, SysUtils, ExtCtrls;


const
  xlCenter = -4108;

    
type
  
  TLogEvent = procedure(const s: string) of object;

  TOEMSOPvsDemandSet_OEM = (
    saDemandChange_OEM, //'�г�����仯',
    saDemandOutOfSupply_OEM, //'�г���������ʵ�ʹ�Ӧ����',
    saFacCap_OEM, //'����������',
    saMatE_OEM,   //'���Ϲ�Ӧ-���Ӽ�',
    saMatS_OEM,   //'���Ϲ�Ӧ-������',
    saMatI_OEM,   //'���Ϲ�Ӧ-�������'
    saPC_OEM,
    saDesignAndFlymeECN_OEM     //'�������/Flyme ECN���'
  );

  TOEMACTvsDemandSet_OEM = (
    sbDemandChange_OEM, //'�г�����仯',
    sbFacCap_OEM, //'����������',
    sbFacMan_OEM, //'������ִ����',
    sbMatE_OEM,   //'���Ϲ�Ӧ-���Ӽ�',
    sbMatS_OEM,   //'���Ϲ�Ӧ-������',
    sbMatI_OEM,   //'���Ϲ�Ӧ-�������',
    sbMPlanE_OEM, //'���ϼƻ��쳣',
    sbPPlanE_OEM, //'�����ƻ��쳣',
    sbArtE_OEM,   //'�Ƴ�Ʒ���쳣',
    sbwwArtE_OEM, //'ί�⹤��',
    sbDesignAndFlymeECN_OEM     //'������Ƽ�ECN���'
  );

  TOEMACTvsSchSet_OEM = (
    scDemandChange_OEM, //'�г�����仯',
    scFacCap_OEM, //'����������',
    scFacMan_OEM, //'������ִ����',
    scMatE_OEM, //'���Ϲ�Ӧ-���Ӽ�',
    scMatS_OEM, //'���Ϲ�Ӧ-������',
    scMatI_OEM, //'���Ϲ�Ӧ-�������',
    scMPlanE_OEM, //'���ϼƻ��쳣',
    scPPlanE_OEM, //'�����ƻ��쳣',
    scArtE_OEM, //'�Ƴ�Ʒ���쳣',
    scwwArtE_OEM, //'ί�⹤��',
    scDesignAndFlymeECN_OEM     //'������Ƽ�ECN���'
  );
     
  TOEMSOPvsDemandSet_ODM = (
    saDemandChange_ODM, //'�����г�����仯',
    saMZFirm_ODM, //'����̼�',
    saMZMat_ODM,  //'����͹��Ͻ���δ���',
    saMZMatQ_ODM, //'����͹���Ʒ���쳣',
    saFacCap_ODM, //'����������',
    saFacMan_ODM, //'�������ƻ�ִ����',
    saMatE_ODM,   //'����������-������',
    saMatS_ODM,   //'����������-�ṹ��',
    saMatI_ODM,   //'����������-����',
    saArtE_ODM    //'�������Ƴ�Ʒ���쳣'
  );

  TOEMACTvsDemandSet_ODM = (
    sbDemandChange_ODM, //'�����г�����仯',
    sbMZFirm_ODM, //'����̼�',
    sbMZMat_ODM,  //'����͹��Ͻ���δ���',
    sbMZMatQ_ODM, //'����͹���Ʒ���쳣',
    sbFacCap_ODM, //'����������',
    sbFacMan_ODM, //'�������ƻ�ִ����',
    sbMatE_ODM,   //'����������-������',
    sbMatS_ODM,   //'����������-�ṹ��',
    sbMatI_ODM,   //'����������-����',
    sbArtE_ODM    //'�������Ƴ�Ʒ���쳣'
  );

  TOEMACTvsSchSet_ODM = (
    scDemandChange_ODM, //'�����г�����仯',
    scMZFirm_ODM, //'����̼�',
    scMZMat_ODM,  //'����͹��Ͻ���δ���',
    scMZMatQ_ODM, //'����͹���Ʒ���쳣',
    scFacCap_ODM, //'����������',
    scFacMan_ODM, //'�������ƻ�ִ����',
    scMatE_ODM,   //'����������-������',
    scMatS_ODM,   //'����������-�ṹ��',
    scMatI_ODM,   //'����������-����',
    scArtE_ODM    //'�������Ƴ�Ʒ���쳣'
  );
  
const
  CSOEMSOPvsDemand_OEM: array[TOEMSOPvsDemandSet_OEM] of string = (
    '�г�����仯',
    '�г���������ʵ�ʹ�Ӧ����',
    '����������',
    '���Ϲ�Ӧ-���Ӽ�',
    '���Ϲ�Ӧ-������',
    '���Ϲ�Ӧ-�������',
    'SOP����',
    '�������/Flyme ECN���'
  );

  CSOEMACTvsDemand_OEM: array[TOEMACTvsDemandSet_OEM] of string = (
    '�г�����仯',
    '����������',
    '������ִ������',
    '���Ϲ�Ӧ-���Ӽ�',
    '���Ϲ�Ӧ-������',
    '���Ϲ�Ӧ-�������',
    '���ϼƻ��쳣',
    '�����ƻ��쳣',
    '�Ƴ̡�����Ʒ���쳣',
    'ί�⹤��',
    '�������/Flyme ECN���'
  );

  CSOEMACTvsSch_OEM: array[TOEMACTvsSchSet_OEM] of string = (
    '�г�����仯',
    '����������',
    '������ִ������',
    '���Ϲ�Ӧ-���Ӽ�',
    '���Ϲ�Ӧ-������',
    '���Ϲ�Ӧ-�������',
    '���ϼƻ��쳣',
    '�����ƻ��쳣',
    '�Ƴ̡�����Ʒ���쳣',
    'ί�⹤��',
    '�������/Flyme ECN���'
  );

  CSOEMSOPvsDemand_ODM: array[TOEMSOPvsDemandSet_OdM] of string = (
    '�����г�����仯',
    '����̼�',
    '����͹��Ͻ���δ���',
    '����͹���Ʒ���쳣',
    '����������',
    '�������ƻ�ִ����',
    '����������-������',
    '����������-�ṹ��',
    '����������-����',
    '�������Ƴ�Ʒ���쳣'
  );
  
  CSOEMACTvsDemand_ODM: array[TOEMACTvsDemandSet_ODM] of string = (
    '�����г�����仯',
    '����̼�',
    '����͹��Ͻ���δ���',
    '����͹���Ʒ���쳣',
    '����������',
    '�������ƻ�ִ����',
    '����������-������',
    '����������-�ṹ��',
    '����������-����',
    '�������Ƴ�Ʒ���쳣'
  );

  CSOEMACTvsSch_ODM: array[TOEMACTvsSchSet_ODM] of string = (
    '�����г�����仯',
    '����̼�',
    '����͹��Ͻ���δ���',
    '����͹���Ʒ���쳣',
    '����������',
    '�������ƻ�ִ����',
    '����������-������',
    '����������-�ṹ��',
    '����������-����',
    '�������Ƴ�Ʒ���쳣'
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
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].Borders.LineStyle := 1; //�ӱ߿�
end;

procedure AddColor(ExcelApp: Variant; irow1, icol1, irow2, icol2: Integer; dwColor: DWORD);
begin
  ExcelApp.Range[ ExcelApp.Cells[irow1, icol1], ExcelApp.Cells[irow2, icol2] ].Interior.Color := dwColor //�ӱ߿�
end;      

procedure AddColor(ExcelApp: Variant; irow1, icol1: Integer; dwColor: DWORD);
begin
  ExcelApp.Cells[irow1, icol1].Interior.Color := dwColor //�ӱ߿�
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
  Result := (Pos('����', sver) > 0) or (Pos('SKD', sver) > 0) or (Pos('CKD', sver) > 0);
end;

function IsNameHW(const snumber, sname: string): Boolean;
begin
  if Pos('����', sname) > 0 then     // ��̩����
  begin
    Result := False;
    Exit;
  end;              

  if Pos('����', sname) > 0 then     // ��̩����
  begin
    Result := True;
    Exit;
  end;

  // 03.53 M1792L ������Ŀ��������
  Result := (Pos('H', sname) > 0)
          or ((Pos('L',sname) > 0) and (Copy(snumber, 1, 5) <> '03.53'))
end;

end.
