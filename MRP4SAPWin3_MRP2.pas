unit MRP4SAPWin3_MRP2;
 
(*
注意事项：
  1、SAP引出的BOM Excel文件，有些只有父项物料，没有子项物料，需从Excel文件删除掉。否则MRP计算程序会找不到BOM
  2、SAP引出的库存文件，注意只包括所有参与MRP计算的仓库
  3、Excel表的数据都要用Sheet1
*)

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, StdCtrls, ExtCtrls, IniFiles, CommUtils,
  DB, ADODB, Provider, DBClient, SOPReaderUnit, SAPBom2SBomWin, SAPOPOReader2, 
  DateUtils, ComObj, ExcelConsts, MrpMPSReader, NewSKUReader, LTP_CMS2MRPSimWin,
  KeyICItemSupplyReader, SBomReader, SOPSimReader, FGPriorityReader, Clipbrd,
  DOSReader, MRPWinReader, jpeg, ImgList, LTPCMSConfirmReader, SAPMaterialReader,
  SAPMaterialReader2, SAPWhereUseReader,
  SAPStockReader2, SAPBomReader, SAPBomReader2, SAPS618Reader, SAPMrpAreaStockReader;

type
  TDeltaRecord = packed record
    sproj: string;
    snumber: string;
    sdate: string;
    iqty: Double;
    iqty_org: Double;
  end;
  PDeltaRecord = ^TDeltaRecord;

  TEORecord = packed record
    snumber: string;
    sname: string;
    sMrpAreaNo: string;
    sMrpAreaName: string;
    dQtyDemand: Double;
    dQtyDemand17: Double;  
    dQtyDemand28: Double;  
    dQtyDemand60: Double;
    dQtyStock: Double;
    dQtyOPO: Double;     
    sMRPType: string;
  end;
  PEORecord = ^TEORecord;
                                                    
  TfrmMRP4SAP3_MRP2 = class(TForm)
    ToolBar1: TToolBar;
    tbClose: TToolButton;
    Memo1: TMemo;
    StatusBar1: TStatusBar;
    ProgressBar1: TProgressBar;
    ToolButton1: TToolButton;
    ImageList1: TImageList;
    leSAPStock: TLabeledEdit;
    btnSAPStock3: TButton;
    leSAPBom: TLabeledEdit;
    btnSAPBom3: TButton;
    leSAPPIR: TLabeledEdit;
    btnDemand3: TButton;
    btnMRP: TButton;
    leSAPOPO: TLabeledEdit;
    btnSAPOPO: TButton;
    leMaterial: TLabeledEdit;
    btnMaterial: TButton;
    leSAPMrpAreaStock: TLabeledEdit;
    btnSAPMrpAreaStock: TButton;
    GroupBox1: TGroupBox;
    mmoAreaStock: TMemo;
    leWhereUse: TLabeledEdit;
    btnWhereUse: TButton;
    procedure FormCreate(Sender: TObject); 
    procedure FormDestroy(Sender: TObject);
    procedure tbCloseClick(Sender: TObject);
    procedure btnMRPClick(Sender: TObject);
    procedure btnSAPStock3Click(Sender: TObject);
    procedure btnSAPBom3Click(Sender: TObject);
    procedure btnDemand3Click(Sender: TObject);
    procedure btnSAPOPOClick(Sender: TObject);
    procedure btnMaterialClick(Sender: TObject);
    procedure btnSAPMrpAreaStockClick(Sender: TObject);
    procedure btnWhereUseClick(Sender: TObject);
  private
    { Private declarations }  
    procedure OnLogEvent(const s: string);
  public
    { Public declarations }
    class procedure ShowForm;
  end;

implementation
 
{$R *.dfm}

type
  TNumberInfo = packed record
    scategory: string; //物料类别
    sg: string; //通用性
    snumber: string; //物料编码(主料)
    sname: string; //物料描述
    slt: string; //LT
    smc: string; //MC
    sproj: string; //项目
  end;
    
  TOrderByConditions = packed record
    isnew: Boolean;
    dos: Double;
    demand: Double;
  end;
  POrderByConditions = ^TOrderByConditions;

  // 简化的MRP计算，不考虑低位码

  TMRPUnit = class
  public
    id: Integer;
    pid: Integer;
    snumber: string;
    sname: string;
    spnumber: string;
    srnumber: string;
    dt: TDateTime;
    dQty: Double;
    dQtyStock: Double;
    dQtyStock2: Double;
    dQtyOPO: Double;
    bExpend: Boolean;
    bCalc: Boolean;
    aBom: TSapBom;
    aParentBom: TSapBom;
    aSAPMaterialRecordPtr: PSAPMaterialRecord;
    sDemandType: string;
    iSubstituteNo: Integer; // 替代组编号
    spriority: string;
    dPer: Double;
    sMrpArea: string;
    iAltCount: Integer; //通用关系数量，没通用关系为1
    sGroupNumbers: string; // 通用关系的物料编码拼接，用以判断完全相同的通用关系
    dQtyNetSPQ: Double; // 净需求，按SPQ
    constructor Create;
    destructor Destroy; override; 
  end;

  function StringListSortCompare(List: TStringList; Index1, Index2: Integer): Integer;
  var
    p1, p2: POrderByConditions;
  begin
    Result := 0;
    p1 := POrderByConditions(List.Objects[Index1]);
    p2 := POrderByConditions(List.Objects[Index2]);

    if p1^.isnew <> p2^.isnew then
    begin
      if p1^.isnew then
        Result := -1
      else Result := 1;
    end
    else
    begin
      if p1^.dos <> p2.dos then
      begin
        if p1^.dos > p2^.dos then
          Result := 1
        else Result := -1;
      end
      else
      begin
        if p1^.demand <> p2^.demand then
        begin
          if p1^.demand > p2^.demand then
            Result := -1
          else Result := 1;
        end;
      end;
    end; 
  end;

  function ListSortCompare_pid(Item1, Item2: Pointer): Integer;
  begin
    Result := TMRPUnit(Item1).pid - TMRPUnit(Item2).pid;
  end;
    
  function ListSortCompare_number_date(Item1, Item2: Pointer): Integer;
  var
    u1, u2: TMRPUnit;
  begin
    u1 := TMRPUnit(Item1);
    u2 := TMRPUnit(Item2);
    if u1.snumber > u2.snumber then
      Result := 1
    else if u1.snumber < u2.snumber then
      Result := -1
    else
    begin
      if u1.sMrpArea < u2.sMrpArea then
        Result := 1
      else if u1.sMrpArea > u2.sMrpArea then
        Result := -1
      else
      begin
        if u1.dt < u2.dt then
          Result := -1
        else if u1.dt > u2.dt then
          Result := 1
        else Result := 0;
      end;
    end;
  end;
           
  function ListSortCompare_Alt(Item1, Item2: Pointer): Integer;
  var
    u1, u2: TMRPUnit;
  begin
    u1 := TMRPUnit(Item1);
    u2 := TMRPUnit(Item2);
    if u1.iSubstituteNo > u2.iSubstituteNo then
      Result := 1
    else if u1.iSubstituteNo < u2.iSubstituteNo then
      Result := -1
    else
    begin
      if u1.sMrpArea < u2.sMrpArea then
        Result := 1
      else if u1.sMrpArea > u2.sMrpArea then
        Result := -1
      else
      begin
        if u1.dt < u2.dt then
          Result := -1
        else if u1.dt > u2.dt then
          Result := 1
        else Result := 0;
      end;
    end;
  end;





class procedure TfrmMRP4SAP3_MRP2.ShowForm;
var
  frmMRP4SAP3: TfrmMRP4SAP3_MRP2;
begin
  frmMRP4SAP3 := TfrmMRP4SAP3_MRP2.Create(nil);
  try
    frmMRP4SAP3.ShowModal;
  finally
    frmMRP4SAP3.Free;
  end;
end;
   
procedure TfrmMRP4SAP3_MRP2.FormCreate(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
  s: string;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
 
  leSAPStock.Text := ini.ReadString(self.ClassName, leSAPStock.Name, '');
  leSAPBom.Text := ini.ReadString(self.ClassName, leSAPBom.Name, '');
  leSAPPIR.Text := ini.ReadString(self.ClassName, leSAPPIR.Name, '');
  leSAPOPO.Text := ini.ReadString(self.ClassName, leSAPOPO.Name, '');   
  leMaterial.Text := ini.ReadString(self.ClassName, leMaterial.Name, '');
  leSAPMrpAreaStock.Text := ini.ReadString(Self.ClassName, leSAPMrpAreaStock.Name, '');

  s := ini.ReadString(self.ClassName, mmoAreaStock.Name, 'FIH01=AF0A||HQ001=AH0A||ML001=AM0A||WT001=AW0A||YD001=AY0A');
  mmoAreaStock.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

  leWhereUse.Text := ini.ReadString(self.ClassName, leWhereUse.Name, '');
        
 
  ini.Free;
end;

procedure TfrmMRP4SAP3_MRP2.FormDestroy(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
  s: string;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
 

  ini.WriteString(self.ClassName, leSAPStock.Name, leSAPStock.Text);
  ini.WriteString(self.ClassName, leSAPBom.Name, leSAPBom.Text);
  ini.WriteString(self.ClassName, leSAPPIR.Name, leSAPPIR.Text);  
  ini.WriteString(self.ClassName, leSAPOPO.Name, leSAPOPO.Text);
  ini.WriteString(self.ClassName, leMaterial.Name, leMaterial.Text);
  ini.WriteString(self.ClassName, leSAPMrpAreaStock.Name, leSAPMrpAreaStock.Text);

  s := StringReplace(mmoAreaStock.Text, #13#10, '||', [rfReplaceAll]);
  ini.WriteString(self.ClassName, mmoAreaStock.Name, s);

  ini.WriteString(self.ClassName, leWhereUse.Name, leWhereUse.Text);

  ini.Free;
end;
 
function ExtractSOPDate(const sdate: string): TDateTime;
var
  s: string;
begin
  try
    s := sdate;
    if Pos('/', s) > 0 then
    begin
      if Pos('-', s) > 0 then
      begin
        s := Copy(s, 1, Pos('-', s) - 1);
      end;
      s := '2017-' + StringReplace(s, '/', '-', [rfReplaceAll]);
    end;
    Result := myStrToDateTime(s);
  except
    Result := 0;
  end;
end;  

function GetNearestMPS(aMPSs: TList; aDeltaPtr: PDeltaRecord): PDeltaRecord;
var
  iDist: Double;
  i: Integer;
  p: PDeltaRecord;
  dt1, dt2: TDateTime;
begin
  Result := nil;
  iDist := MaxInt;
  for i := 0 to aMPSs.Count - 1 do
  begin
    p := PDeltaRecord(aMPSs[i]);
    if p^.iqty = 0 then Continue;
    if p^.snumber = aDeltaPtr^.snumber then
    begin
      dt1 := ExtractSOPDate(p^.sdate);
      dt2 := ExtractSOPDate(aDeltaPtr^.sdate);
      if iDist > Abs(dt1 - dt2) then
      begin
        iDist := Abs(dt1 - dt2);
        Result := p;
      end;
    end;
  end;  
end;
 
procedure TfrmMRP4SAP3_MRP2.tbCloseClick(Sender: TObject);
begin
  Close;
end;

constructor TMRPUnit.Create;
begin
  inherited;
end;

destructor TMRPUnit.Destroy;
begin

end;

    
    function ListSortCompare_Number_DateTime(Item1, Item2: Pointer): Integer;
    var
      p1, p2: TMRPUnit;
    begin
      p1 := TMRPUnit(Item1);
      p2 := TMRPUnit(Item2);

      if p1.snumber > p2.snumber then
      begin
        Result := 1;
      end
      else if p1.snumber < p2.snumber then
      begin
        Result := -1;
      end
      else
      begin
        if DoubleG(p1.dt, p2.dt) then
          Result := 1
        else if DoubleL(p1.dt, p2.dt) then
          Result := -1
        else Result := 0;
      end;
    end;
             
    function ListSortCompare_DateTime(Item1, Item2: Pointer): Integer;
    var
      p1, p2: TMRPUnit;
    begin
      p1 := TMRPUnit(Item1);
      p2 := TMRPUnit(Item2);
      
      if DoubleG(p1.dt, p2.dt) then
        Result := 1
      else if DoubleL(p1.dt, p2.dt) then
        Result := -1
      else
      begin  // 时间相同， 通用关系少的，放在前面
        if p1.iAltCount > p2.iAltCount then
          Result := 1
        else if p1.iAltCount < p2.iAltCount then
          Result := -1
        else
          Result := 0;
      end;
    end;
             
    function ListSortCompare_ForSPQ(Item1, Item2: Pointer): Integer;
    var
      p1, p2: TMRPUnit;
    begin
      p1 := TMRPUnit(Item1);
      p2 := TMRPUnit(Item2);

      if p1.sGroupNumbers > p2.sGroupNumbers then
        Result := 1
      else if p1.sGroupNumbers < p2.sGroupNumbers then
        Result := -1
      else
      begin
        if DoubleG(p1.dt, p2.dt) then
          Result := 1
        else if DoubleL(p1.dt, p2.dt) then
          Result := -1
        else
        begin  // 时间相同， 通用关系少的，放在前面
          if p1.snumber > p2.snumber then
            Result := 1
          else if p1.snumber < p2.snumber then
            Result := -1
          else
            Result := 0;
        end;
      end;
    end;

function GetDemand(lstDemand: TStringList; const snumber: string; dt1, dt2: TDateTime): Double;
var
  i: Integer;
  aMRPUnitPtr: TMRPUnit;
begin
  Result := 0;
  for i := 0 to lstDemand.Count - 1 do
  begin
    if lstDemand[i] <> snumber then Continue;
    aMRPUnitPtr := TMRPUnit(lstDemand.Objects[i]);
    if (aMRPUnitPtr.dt >= dt1) and (aMRPUnitPtr.dt < dt2) then
    begin
      Result := Result + aMRPUnitPtr.dQty;
    end;
  end;
end;

 
procedure TfrmMRP4SAP3_MRP2.btnSAPStock3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;
     
procedure TfrmMRP4SAP3_MRP2.btnSAPOPOClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPOPO.Text := sfile;
end;
      
procedure TfrmMRP4SAP3_MRP2.btnMaterialClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMaterial.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP2.btnSAPBom3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP2.btnDemand3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPPIR.Text := sfile;
end;
    
procedure TfrmMRP4SAP3_MRP2.btnSAPMrpAreaStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPMrpAreaStock.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP2.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

function ListSortCompare_priority(Item1, Item2: Pointer): Integer;
var
  aMRPUnitPtr1, aMRPUnitPtr2: TMRPUnit;
  iPriority1, iPriority2: Integer;
begin
  aMRPUnitPtr1 := Item1;
  aMRPUnitPtr2 := Item2;
  iPriority1 := StrToIntDef(aMRPUnitPtr1.spriority, 1);
  iPriority2 := StrToIntDef(aMRPUnitPtr2.spriority, 1);
  if iPriority1 > iPriority2 then
  begin
    Result := 1;
  end
  else if iPriority1 = iPriority2 then
  begin
    Result := 0;
  end
  else
  begin
    Result := -1;
  end;
end;

{*
查找指定的值是否在当前数组中(数组已经是有序的)
*}
function QuickSearchMrpUnit(lst: TList; pid: longint ): Integer; overload;
var
  idMid: integer;
  idLow, idHigh: integer;
begin
  idLow := 0;
  idHigh := lst.Count - 1;


  while ( idLow <= idHigh ) do
  begin
    if idLow = idHigh then
    begin
      if TMRPUnit(lst[ idLow ]).pid = pid then
      begin
        Result := idLow;
      end
      else
      begin
        Result := -1;
      end;
      Exit;
    end;
    idMid := ( idLow + idHigh ) div 2;
    if TMRPUnit(lst[ idMid ]).pid = pid then
    begin
      Result := idMid;
      Exit;
    end;

    if TMRPUnit(lst[ idMid ]).pid > pid then idHigh := idMid - 1;
    if TMRPUnit(lst[ idMid ]).pid < pid then idLow := idMid + 1;
  end;

  Result := -1;
end;

{*
查找指定的值是否在当前数组中(数组已经是有序的)
*}
function QuickSearchMrpUnit(lst: TList; const snumber: string ): Integer; overload;
var
  idMid: integer;
  idLow, idHigh: integer;
begin
  idLow := 0;
  idHigh := lst.Count - 1;


  while ( idLow <= idHigh ) do
  begin
    if idLow = idHigh then
    begin
      if TMRPUnit(lst[ idLow ]).snumber = snumber then
      begin
        Result := idLow;
      end
      else
      begin
        Result := -1;
      end;
      Exit;
    end;
    idMid := ( idLow + idHigh ) div 2;
    if TMRPUnit(lst[ idMid ]).snumber = snumber then
    begin
      Result := idMid;
      Exit;
    end;

    if TMRPUnit(lst[ idMid ]).snumber > snumber then idHigh := idMid - 1;
    if TMRPUnit(lst[ idMid ]).snumber < snumber then idLow := idMid + 1;
  end;

  Result := -1;
end;

{*
查找指定的值是否在当前数组中(数组已经是有序的)
*}
function QuickSearchMrpUnitAlt(lst: TList; const iAlt: Integer): Integer;
var
  idMid: integer;
  idLow, idHigh: integer;
begin
  idLow := 0;
  idHigh := lst.Count - 1;


  while ( idLow <= idHigh ) do
  begin
    if idLow = idHigh then
    begin
      if TMRPUnit(lst[ idLow ]).iSubstituteNo = iAlt then
      begin
        Result := idLow;
      end
      else
      begin
        Result := -1;
      end;
      Exit;
    end;
    idMid := ( idLow + idHigh ) div 2;
    if TMRPUnit(lst[ idMid ]).iSubstituteNo = iAlt then
    begin
      Result := idMid;
      Exit;
    end;

    if TMRPUnit(lst[ idMid ]).iSubstituteNo > iAlt then idHigh := idMid - 1;
    if TMRPUnit(lst[ idMid ]).iSubstituteNo < iAlt then idLow := idMid + 1;
  end;

  Result := -1;
end;

procedure QuickSearchMrpUnitAlts(lst: TList; const iAlt: Integer; res: TList );
var
  idx: Integer;
begin
  res.Clear;
  idx := QuickSearchMrpUnitAlt(lst, iAlt);
  if idx <= 0 then Exit;
  while TMRPUnit(lst[ idx ]).iSubstituteNo = iAlt do
  begin
    idx := idx - 1;
    if idx < 0 then Break;
  end;
  idx := idx + 1;
  while TMRPUnit(lst[ idx ]).iSubstituteNo = iAlt do
  begin
    res.Add(lst[ idx ]);
    idx := idx + 1;
    if idx >= lst.Count then Break;
  end;
end;

procedure QuickSearchMrpUnitNumbers(lst: TList; const snumber: string; res: TList);
var
  idx: Integer;
begin
  res.Clear;
  idx := QuickSearchMrpUnit(lst, snumber);
  if idx <= 0 then Exit;
  while TMRPUnit(lst[ idx ]).snumber = snumber do
  begin
    idx := idx - 1;
    if idx < 0 then Break;
  end;
  idx := idx + 1;
  while TMRPUnit(lst[ idx ]).snumber = snumber do
  begin
    res.Add(lst[ idx ]);
    idx := idx + 1;
    if idx >= lst.Count then Break;
  end;
end;

procedure WriteLine_MrpLog(ExcelApp: Variant; var irow: Integer;
  p: TMRPUnit; lst: TList; slNumber: TStringList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader;
  aSAPMaterialReader: TSAPMaterialReader2; sl: TStringList);
var
  idx: Integer;
  aMRPUnitPtr: TMRPUnit;
  smrparea: string;
  aEORecordPtr: PEORecord;
  sline: string;
begin
//  idx := GetChildMrpLog(lst, p);
  idx := QuickSearchMrpUnit(lst, p.id);
  while idx >= 0 do
  begin
    aMRPUnitPtr := TMRPUnit(lst[idx]);
    lst.Delete(idx);

    // E&O /////////////////////////////////////////////
    smrparea := aMRPUnitPtr.sMrpArea;

    idx := slNumber.IndexOf(aMRPUnitPtr.snumber);
    if idx < 0 then
    begin
      aEORecordPtr := New(PEORecord);
      aEORecordPtr^.snumber := aMRPUnitPtr.snumber;
      aEORecordPtr^.sname := aMRPUnitPtr.sname;
      aEORecordPtr^.sMrpAreaNo := smrparea;
      aEORecordPtr^.sMrpAreaName := aSAPMrpAreaStockReader.MrpAreaNo2Name(smrparea);
      aEORecordPtr^.dQtyDemand := 0;
      aEORecordPtr^.dQtyDemand17 := 0;
      aEORecordPtr^.dQtyDemand28 := 0;
      aEORecordPtr^.dQtyDemand60 := 0;
      aEORecordPtr^.dQtyStock := 0;
      aEORecordPtr^.dQtyOPO := 0;
      aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aMRPUnitPtr.snumber );
      slNumber.AddObject(aMRPUnitPtr.snumber, TObject(aEORecordPtr));
    end
    else
    begin
      aEORecordPtr := PEORecord(slNumber.Objects[idx]);
    end;

    aEORecordPtr^.dQtyDemand := aEORecordPtr^.dQtyDemand + aMRPUnitPtr.dQty;
        
    if DoubleL(aMRPUnitPtr.dt, today + 17) then
    begin
      aEORecordPtr^.dQtyDemand17 := aEORecordPtr^.dQtyDemand17 + aMRPUnitPtr.dQty;
    end;
    if DoubleL(aMRPUnitPtr.dt, today + 28) then
    begin
      aEORecordPtr^.dQtyDemand28 := aEORecordPtr^.dQtyDemand28 + aMRPUnitPtr.dQty;
    end;
    if DoubleL(aMRPUnitPtr.dt, today + 60) then
    begin
      aEORecordPtr^.dQtyDemand60 := aEORecordPtr^.dQtyDemand60 + aMRPUnitPtr.dQty;
    end; 

    //////////////////////////////////////////////////////////////////////


    sline := IntToStr(aMRPUnitPtr.id)+#9+           //1
      IntToStr(aMRPUnitPtr.pid)+#9+                 //2
      aMRPUnitPtr.snumber+#9+                       //3
      aMRPUnitPtr.sname+#9+                         //4
      FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt)+#9;  //5
    if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'F' then  // 外购  ////////////////////////////////////////
    begin
      sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD) + #9;      //6
    end
    else                                                         // 自制  ////////////////////////////////////////
    begin
      sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0) + #9;      //6
    end;

    sline := sline + Format('%0.0f', [aMRPUnitPtr.dqty])+#9+      //7
      Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2])+#9+   //8
      Format('%0.0f', [aMRPUnitPtr.dQtyOPO])+#9;   //9
    sline := sline + '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9+  //10
      Format('%0.0f', [aMRPUnitPtr.dQtyNetSPQ]) + #9 +    // 11
      IntToStr(aMRPUnitPtr.iSubstituteNo)+#9+   //12
      aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPer + #9 +   //13
      aMRPUnitPtr.aSAPMaterialRecordPtr.sBuyer +#9 +   //14
      aMRPUnitPtr.sMrpArea+#9+   //15
      aMRPUnitPtr.spnumber+#9+    //16
      aMRPUnitPtr.srnumber+#9;     //17
    if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType = 'F' then
    begin
      sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD]) + #9;  //18
    end
    else
    begin
      sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0]) + #9;  //18
    end;
    sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dSPQ]) + #9;   //19
    sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dMOQ]) + #9;   //20
    sline := sline + aMRPUnitPtr.sGroupNumbers + #9;   //21


    sl.Add(sline);

    (*
    ExcelApp.Cells[irow, 1].Value := aMRPUnitPtr.id;// 'ID';
    ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr.pid;// '父ID';
    ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr.snumber;// '物料';
    ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr.sname;// '物料名称';
    ExcelApp.Cells[irow, 5].Value := aMRPUnitPtr.dt;// '需求日期';

    //'建议下单日期'
    if aMRPUnitPtr.aSAPMaterialRecordPtr^.sMRPType = 'PD' then  // 外购  ////////////////////////////////////////
    begin
      ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD);
    end
    else                                                         // 自制  ////////////////////////////////////////
    begin
      ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0);
    end;
          
    ExcelApp.Cells[irow, 7].Value := Format('%0.0f', [aMRPUnitPtr.dqty]); //'需求数量';
    ExcelApp.Cells[irow, 8].Value := Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2]); //'可用库存';
    ExcelApp.Cells[irow, 9].Value := Format('%0.0f', [aMRPUnitPtr.dQtyOPO]); //'OPO';
    ExcelApp.Cells[irow, 10].Value := '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow); //'净需求';
    ExcelApp.Cells[irow, 11].Value := IntToStr(aMRPUnitPtr.iSubstituteNo); //'替代组';
    ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPer; //'MRP控制者';
    ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.sBuyer; //'采购员';
    ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr.sMrpArea; //'MRP区域';    
    ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr.spnumber;
    ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr.srnumber; 
    if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType = 'F' then
    begin
      ExcelApp.Cells[irow, 17].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD;
    end
    else
    begin
      ExcelApp.Cells[irow, 17].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0;
    end;
    *)
    irow := irow + 1;

    WriteLine_MrpLog(ExcelApp, irow, aMRPUnitPtr, lst, slNumber,
      aSAPMrpAreaStockReader, aSAPMaterialReader, sl);


    //idx := GetChildMrpLog(lst, p)    
    idx := QuickSearchMrpUnit(lst, p.id);
  end; 
end;

function FindNumber(lst: TList; const snumber: string): TMRPUnit;
var
  i: Integer;
  aMRPUnit: TMRPUnit;
begin
  Result := nil;
  for i := 0 to lst.Count - 1 do
  begin
     aMRPUnit := TMRPUnit(lst[i]);
     if aMRPUnit.snumber = snumber then
     begin
       Result := aMRPUnit;
       Break;
     end;
  end;
end;

function ShiftAlloc_stock(aMRPUnit: TMRPUnit; lstDemandNumber, lstDemandAlt: TList;
  dQtyE: Double; slNumberStack: TStringList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader): Double;
var
  slNumber: TStringList;
  iNumber: Integer;
  snumber: string;
  lstDemandShift: TList;
  aMRPUnitShiftDst: TMRPUnit;
  aMRPUnitShiftDstDo: TMRPUnit;
  aMRPUnitShiftSrc: TMRPUnit;
  iShift: Integer;
  iShiftDo: Integer;
  lstDemandShiftAlt: TList;
  dQtyS: Double;
  dQtyOk: Double;
begin
  Result := 0;
  
  lstDemandShift := TList.Create;
  lstDemandShiftAlt := TList.Create;
  slNumber := TStringList.Create;
  slNumber.Text := StringReplace(aMRPUnit.sGroupNumbers, '|', #13#10, [rfReplaceAll]);
  for iNumber := 0 to slNumber.Count - 1 do
  begin
    snumber := slNumber[iNumber];
    if slNumberStack.IndexOf(snumber) >= 0 then Continue; // 在栈里边
    QuickSearchMrpUnitNumbers(lstDemandNumber, snumber, lstDemandShift);
    for iShift := 0 to lstDemandShift.Count - 1 do
    begin
      aMRPUnitShiftDst := TMRPUnit(lstDemandShift[iShift]);
      QuickSearchMrpUnitAlts(lstDemandAlt, aMRPUnitShiftDst.iSubstituteNo, lstDemandShiftAlt);
      aMRPUnitShiftSrc := FindNumber(lstDemandShiftAlt, aMRPUnit.snumber);
      if aMRPUnitShiftSrc = nil then Continue; // 不通用  xxxxxxxxxxxxxxxxxxxxxxxxxx

      for iShiftDo := 0 to lstDemandShift.Count - 1 do
      begin
        aMRPUnitShiftDstDo := TMRPUnit(lstDemandShift[iShiftDo]);
        dQtyS := aMRPUnitShiftDstDo.dQty - aMRPUnitShiftDstDo.dQtyStock - aMRPUnitShiftDstDo.dQtyStock2;
        // 转移库存供给
        if DoubleG(dQtyS, 0) then
        begin
          if dQtyS > dQtyE then
          begin
            dQtyS := dQtyE;
          end;
          aMRPUnitShiftDstDo.dQtyStock := aMRPUnitShiftDstDo.dQtyStock + dQtyS;  // 增加库存供应
          aMRPUnitShiftDst.dQtyStock := aMRPUnitShiftDst.dQtyStock - dQtyS;      // 减少库存供应
          aMRPUnitShiftDst.dQty := aMRPUnitShiftDst.dQty - dQtyS;                // 减少需求
          aMRPUnitShiftSrc.dQty := aMRPUnitShiftSrc.dQty + dQtyS;                // 增加需求
          aMRPUnitShiftSrc.dQtyStock := aMRPUnitShiftSrc.dQtyStock +  
            aSAPMrpAreaStockReader.AllocStock(aMRPUnitShiftSrc.snumber, dQtyS, aMRPUnitShiftSrc.sMrpArea);  // 分配库存， 如果是间接的，可能无库存分配，需求会再次转移
          Result := Result + dQtyS;
          dQtyE := dQtyE - dQtyS;
          if DoubleE( dQtyE, 0 ) then Break;
        end;
      end;
                    
      if DoubleE(dQtyE, 0) then Break;

      dQtyS := aMRPUnitShiftDst.dQty;
      begin
        if dQtyS > dQtyE then
        begin
          dQtyS := dQtyE;
        end;

        slNumberStack.Add(aMRPUnitShiftDst.snumber);
        dQtyOk := ShiftAlloc_stock(aMRPUnitShiftDst, lstDemandNumber,
          lstDemandAlt, dQtyS, slNumberStack, aSAPMrpAreaStockReader);
        slNumberStack.Delete(slNumberStack.Count - 1);

        if dQtyOk > 0 then
        begin
          aMRPUnitShiftDst.dQty := aMRPUnitShiftDst.dQty - dQtyOk;                // 减少需求
          aMRPUnitShiftSrc.dQty := aMRPUnitShiftSrc.dQty + dQtyOk;                // 增加需求
          aMRPUnitShiftSrc.dQtyStock := aMRPUnitShiftSrc.dQtyStock +  
            aSAPMrpAreaStockReader.AllocStock(aMRPUnitShiftSrc.snumber, dQtyOk, aMRPUnitShiftSrc.sMrpArea);  // 分配库存， 如果是间接的，可能无库存分配，需求会再次转移

          Result := Result + dQtyOk;
        end;
        
        dQtyE := dQtyE - dQtyOk;    
        if DoubleE(dQtyE, 0) then Break;
      end;
     
    end;    
    if DoubleE(dQtyE, 0) then Break;
  end;
  slNumber.Free;
  lstDemandShift.Free;
  lstDemandShiftAlt.Free;
end;
     
function ShiftAlloc_po(aMRPUnit: TMRPUnit; lstDemandNumber, lstDemandAlt: TList;
  dQtyE: Double; slNumberStack: TStringList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader): Double;
var
  slNumber: TStringList;
  iNumber: Integer;
  snumber: string;
  lstDemandShift: TList;
  aMRPUnitShiftDst: TMRPUnit;
  aMRPUnitShiftDstDo: TMRPUnit;
  aMRPUnitShiftSrc: TMRPUnit;
  iShift: Integer;
  iShiftDo: Integer;
  lstDemandShiftAlt: TList;
  dQtyS: Double;
  dQtyOk: Double;
begin
  Result := 0;
  
  lstDemandShift := TList.Create;
  lstDemandShiftAlt := TList.Create;
  slNumber := TStringList.Create;
  slNumber.Text := StringReplace(aMRPUnit.sGroupNumbers, '|', #13#10, [rfReplaceAll]);
  for iNumber := 0 to slNumber.Count - 1 do
  begin
    snumber := slNumber[iNumber];
    if slNumberStack.IndexOf(snumber) >= 0 then Continue; // 在栈里边
    QuickSearchMrpUnitNumbers(lstDemandNumber, snumber, lstDemandShift);
    for iShift := 0 to lstDemandShift.Count - 1 do
    begin
      aMRPUnitShiftDst := TMRPUnit(lstDemandShift[iShift]);
      QuickSearchMrpUnitAlts(lstDemandAlt, aMRPUnitShiftDst.iSubstituteNo, lstDemandShiftAlt);
      aMRPUnitShiftSrc := FindNumber(lstDemandShiftAlt, aMRPUnit.snumber);
      if aMRPUnitShiftSrc = nil then Continue; // 不通用  xxxxxxxxxxxxxxxxxxxxxxxxxx

      for iShiftDo := 0 to lstDemandShift.Count - 1 do
      begin
        aMRPUnitShiftDstDo := TMRPUnit(lstDemandShift[iShiftDo]);
        dQtyS := aMRPUnitShiftDstDo.dQty - aMRPUnitShiftDstDo.dQtyStock - aMRPUnitShiftDstDo.dQtyStock2;
        // 转移库存供给
        if DoubleG(dQtyS, 0) then
        begin
          if dQtyS > dQtyE then
          begin
            dQtyS := dQtyE;
          end;
          aMRPUnitShiftDstDo.dQtyStock := aMRPUnitShiftDstDo.dQtyStock + dQtyS;  // 增加库存供应
          aMRPUnitShiftDst.dQtyStock := aMRPUnitShiftDst.dQtyStock - dQtyS;      // 减少库存供应
          aMRPUnitShiftDst.dQty := aMRPUnitShiftDst.dQty - dQtyS;                // 减少需求
          aMRPUnitShiftSrc.dQty := aMRPUnitShiftSrc.dQty + dQtyS;                // 增加需求
          aMRPUnitShiftSrc.dQtyStock := aMRPUnitShiftSrc.dQtyStock +  
            aSAPMrpAreaStockReader.AllocStock(aMRPUnitShiftSrc.snumber, dQtyS, aMRPUnitShiftSrc.sMrpArea);  // 分配库存， 如果是间接的，可能无库存分配，需求会再次转移
          Result := Result + dQtyS;
          dQtyE := dQtyE - dQtyS;
          if DoubleE( dQtyE, 0 ) then Break;
        end;
      end;
                    
      if DoubleE(dQtyE, 0) then Break;

      dQtyS := aMRPUnitShiftDst.dQty;
      begin
        if dQtyS > dQtyE then
        begin
          dQtyS := dQtyE;
        end;

        slNumberStack.Add(aMRPUnitShiftDst.snumber);
        dQtyOk := ShiftAlloc_stock(aMRPUnitShiftDst, lstDemandNumber,
          lstDemandAlt, dQtyS, slNumberStack, aSAPMrpAreaStockReader);
        slNumberStack.Delete(slNumberStack.Count - 1);

        if dQtyOk > 0 then
        begin
          aMRPUnitShiftDst.dQty := aMRPUnitShiftDst.dQty - dQtyOk;                // 减少需求
          aMRPUnitShiftSrc.dQty := aMRPUnitShiftSrc.dQty + dQtyOk;                // 增加需求
          aMRPUnitShiftSrc.dQtyStock := aMRPUnitShiftSrc.dQtyStock +  
            aSAPMrpAreaStockReader.AllocStock(aMRPUnitShiftSrc.snumber, dQtyOk, aMRPUnitShiftSrc.sMrpArea);  // 分配库存， 如果是间接的，可能无库存分配，需求会再次转移

          Result := Result + dQtyOk;
        end;
        
        dQtyE := dQtyE - dQtyOk;    
        if DoubleE(dQtyE, 0) then Break;
      end;
     
    end;    
    if DoubleE(dQtyE, 0) then Break;
  end;
  slNumber.Free;
  lstDemandShift.Free;
  lstDemandShiftAlt.Free;
end;

// 替代料优化计算
procedure AlternativeOptimization_stock(lstDemand: TList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader);
var
  iarea: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  inumber: Integer;
  aSAPStockRecordPtr: PSAPStockRecord;
  dQtyE: Double;
  iMrpUnit: Integer;
  aMRPUnit: TMRPUnit;
  lst: TList;
  slNumberStack: TStringList;
  lstDemandAlt: TList;
begin
  lstDemandAlt := TList.Create;
  for iMrpUnit := 0 to lstDemand.Count - 1 do
  begin
    lstDemandAlt.Add(lstDemand.Items[iMrpUnit]);
  end;
  lstDemandAlt.Sort(ListSortCompare_Alt);

  lst := TList.Create;
  slNumberStack := TStringList.Create;

  //排序，按料号、日期
  lstDemand.Sort(ListSortCompare_number_date);
  for iarea := 0 to aSAPMrpAreaStockReader.Count - 1 do
  begin
    aSAPMrpAREA := aSAPMrpAreaStockReader.Items[iarea];
    for inumber := 0 to aSAPMrpAREA.FSAPStockSum.FList.Count - 1 do
    begin
      aSAPStockRecordPtr := PSAPStockRecord(aSAPMrpAREA.FSAPStockSum.FList.Objects[inumber]);
                                                                                            
      dQtyE := aSAPStockRecordPtr^.dqty - aSAPStockRecordPtr^.dQty_Alloc;
      
      // 库存无呆滞
      if DoubleE( dQtyE, 0 ) then Continue;

      // 无需求
      iMrpUnit := QuickSearchMrpUnit(lstDemand, aSAPStockRecordPtr^.snumber);
      if iMrpUnit < 0 then Continue;

      while (iMrpUnit >= 0) and
        (TMRPUnit( lstDemand[iMrpUnit] ).snumber = aSAPStockRecordPtr^.snumber) do
      begin
        iMrpUnit := iMrpUnit - 1;
      end;
      iMrpUnit := iMrpUnit + 1;
      lst.Clear;

      // 把料号的需求都拿出来算
      while TMRPUnit( lstDemand[iMrpUnit] ).snumber = aSAPStockRecordPtr^.snumber do
      begin
        // 外购件、 有替代  才算   
        if (TMRPUnit( lstDemand[iMrpUnit] ).aSAPMaterialRecordPtr^.sPType = 'F')
          and (TMRPUnit( lstDemand[iMrpUnit] ).iSubstituteNo > 0) then
        begin
          lst.Add(lstDemand[iMrpUnit]);
        end;                           
        iMrpUnit := iMrpUnit + 1;
      end;

      //无需要计算的
      if lst.Count = 0 then Continue;

      for iMrpUnit := 0 to lst.Count - 1 do
      begin
        aMRPUnit := TMRPUnit( lst[iMrpUnit] );
        slNumberStack.Add(aMRPUnit.snumber);
        ShiftAlloc_stock(aMRPUnit, lstDemand, lstDemandAlt, dQtyE,
          slNumberStack, aSAPMrpAreaStockReader);     
        slNumberStack.Delete(slNumberStack.Count - 1);
        
        if DoubleE(dQtyE, 0) then Break; // 优化完了
      end;
    end;
  end;
  lst.Free;
  slNumberStack.Clear;
  lstDemandAlt.Free;
end;
           
// 替代料优化计算
procedure AlternativeOptimization_po(lstDemand: TList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader);
var
  iarea: Integer;
  aSAPMrpAREA: TSAPMrpAREA;
  inumber: Integer;
  aSAPStockRecordPtr: PSAPStockRecord;
  dQtyE: Double;
  iMrpUnit: Integer;
  aMRPUnit: TMRPUnit;
  lst: TList;
  slNumberStack: TStringList;
  lstDemandAlt: TList;
begin
  lstDemandAlt := TList.Create;
  for iMrpUnit := 0 to lstDemand.Count - 1 do
  begin
    lstDemandAlt.Add(lstDemand.Items[iMrpUnit]);
  end;
  lstDemandAlt.Sort(ListSortCompare_Alt);

  lst := TList.Create;
  slNumberStack := TStringList.Create;

  //排序，按料号、日期
  lstDemand.Sort(ListSortCompare_number_date);
  for iarea := 0 to aSAPMrpAreaStockReader.Count - 1 do
  begin
    aSAPMrpAREA := aSAPMrpAreaStockReader.Items[iarea];
    for inumber := 0 to aSAPMrpAREA.FSAPStockSum.FList.Count - 1 do
    begin
      aSAPStockRecordPtr := PSAPStockRecord(aSAPMrpAREA.FSAPStockSum.FList.Objects[inumber]);
                                                                                            
      dQtyE := aSAPStockRecordPtr^.dqty - aSAPStockRecordPtr^.dQty_Alloc;
      
      // 库存无呆滞
      if DoubleE( dQtyE, 0 ) then Continue;

      // 无需求
      iMrpUnit := QuickSearchMrpUnit(lstDemand, aSAPStockRecordPtr^.snumber);
      if iMrpUnit < 0 then Continue;

      while (iMrpUnit >= 0) and
        (TMRPUnit( lstDemand[iMrpUnit] ).snumber = aSAPStockRecordPtr^.snumber) do
      begin
        iMrpUnit := iMrpUnit - 1;
      end;
      iMrpUnit := iMrpUnit + 1;
      lst.Clear;

      // 把料号的需求都拿出来算
      while TMRPUnit( lstDemand[iMrpUnit] ).snumber = aSAPStockRecordPtr^.snumber do
      begin
        // 外购件、 有替代  才算   
        if (TMRPUnit( lstDemand[iMrpUnit] ).aSAPMaterialRecordPtr^.sPType = 'F')
          and (TMRPUnit( lstDemand[iMrpUnit] ).iSubstituteNo > 0) then
        begin
          lst.Add(lstDemand[iMrpUnit]);
        end;                           
        iMrpUnit := iMrpUnit + 1;
      end;

      //无需要计算的
      if lst.Count = 0 then Continue;

      for iMrpUnit := 0 to lst.Count - 1 do
      begin
        aMRPUnit := TMRPUnit( lst[iMrpUnit] );
        slNumberStack.Add(aMRPUnit.snumber);
        ShiftAlloc_po(aMRPUnit, lstDemand, lstDemandAlt, dQtyE,
          slNumberStack, aSAPMrpAreaStockReader);     
        slNumberStack.Delete(slNumberStack.Count - 1);
        
        if DoubleE(dQtyE, 0) then Break; // 优化完了
      end;
    end;
  end;
  lst.Free;
  slNumberStack.Clear;
  lstDemandAlt.Free;
end;
   
procedure TfrmMRP4SAP3_MRP2.btnMRPClick(Sender: TObject);
  function GetMRPUnit(lst: TList; const snumber: string): TMRPUnit;
  var
    i: Integer;
  begin
    Result := nil;
    for i := 0 to lst.Count - 1 do
    begin
      if TMRPUnit( lst[i] ).snumber = snumber then
      begin
        Result := TMRPUnit( lst[i] );
        Break;
      end;
    end;
  end;  
var                                                                                                       
  sfile: string;
  aSAPMaterialReader: TSAPMaterialReader2;
  aSAPBomReader: TSAPBomReader2;
  aSAPStockReader: TSAPStockReader2;
  //aSAPStockSum: TSAPStockSum;
  aSAPS618Reader: TSAPPIRReader;
  aSAPOPOReader2: TSAPOPOReader2;
  lstDemand: TList;
  iLine: Integer;
  iWeek: Integer;
  aSOPProj: TSOPProj;
  iDate: Integer;
  aMRPUnitPtr: TMRPUnit;
  aMRPUnitPtr_Dep: TMRPUnit; 
  iMrpUnit: Integer;

  iChild: Integer;
  iChildItem: Integer;
  aSapItemGroup: TSapItemGroup;

  aSapBomChild: TSapBom;
  
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  bLoop: BOOL;
  dQty: Double;
  dQty_Stock_a: Double; 
  iNumber: Integer; 
  dt: TDateTime;
  icol: Integer;
  icolMax: Integer;
  idx: Integer;
  dtToday: TDateTime;
  dtNextMonday: TDateTime;
  iSheet: Integer;

  ExcelApp2, WorkBook2: Variant;    
  iSheetCount2, iSheet2: Integer;
  sSheet: string;
  stitle1, stitle2, stitle3, stitle4, stitle5, stitle6, stitle7, stitle8, stitle9: string;
  stitle: string;

  aSAPS618: TSAPS618;
  aSAPS618ColPtr: PSAPS618Col;
  //lstMrpDetail: TList;
  iID: Integer;
  slNumber: TStringList;

  
  aSAPOPOLine: TSAPOPOLine;
  aSAPOPOAllocPtr: PSAPOPOAlloc;
  iLineAlloc: Integer;

  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  iLowestCode: Integer;
  iSubstituteNo: Integer;  // 替代组编号
  lstSubstituteDemand: TList;
  lstPOLine: TList;

  lstDemand_Count: Integer;

  iPer100: Integer;
  sl: TStringList;
  sline: string;

  aEORecordPtr: PEORecord;
  aSAPStockRecordPtr: PSAPStockRecord;
  today: TDateTime;

  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader;
  smrparea: string;

  slArea2BomFac: TStringList;

  lst: TList;

  aSAPWhereUseReader: TSAPWhereUseReader;

  aMRPUnitPtr0: TMRPUnit;
  dqty_pr: Double;
  dqty_pr_sum: Double;
  dqty_pr_sum_spq: Double;
  dtMon: TDateTime; 
  slAreaStock: TStringList;
  slGroupNumber: TStringList;
  sGroupNumbers: string;

  // for spq
  dQtyNet: Double;
  dQtyEx: Double;

  slPerErr: TStringList;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  slArea2BomFac := TStringList.Create;
  slArea2BomFac.Add('FIH01=FX');
  slArea2BomFac.Add('HQ001=HQ');
  slArea2BomFac.Add('ML001=ML');
  slArea2BomFac.Add('WT001=WT');
  slArea2BomFac.Add('YD001=YD');

  today := myStrToDateTime(FormatDateTime('yyyy-MM-dd', Now));
                                         
  Memo1.Lines.Add('开始读取 MRP区域仓库列表  ' + leSAPMrpAreaStock.Text);
  aSAPMrpAreaStockReader := TSAPMrpAreaStockReader.Create(leSAPMrpAreaStock.Text);
  
  Memo1.Lines.Add('开始读取 BOM  ' + leSAPBom.Text);
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 库存  ' + leSAPStock.Text);
  aSAPStockReader := TSAPStockReader2.Create(leSAPStock.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 OPO  ' + leSAPOPO.Text);
  aSAPOPOReader2 := TSAPOPOReader2.Create(leSAPOPO.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 物料  ' + leMaterial.Text);
  aSAPMaterialReader:= TSAPMaterialReader2.Create(leMaterial.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 PIR  ' + leSAPPIR.Text);
  aSAPS618Reader := TSAPPIRReader.Create(leSAPPIR.Text, OnLogEvent);



  Memo1.Lines.Add('开始读取 在产项目  ' + leWhereUse.Text);
  aSAPWhereUseReader := TSAPWhereUseReader.Create(leWhereUse.Text, OnLogEvent);


  aSAPMrpAreaStockReader.SetOPOList(aSAPOPOReader2);
  aSAPMrpAreaStockReader.SetStock(aSAPStockReader);

  //aSAPStockSum := TSAPStockSum.Create;
  //aSAPStockReader.SumTo(aSAPStockSum);

//  lstMrpDetail := TList.Create;

  lstDemand := TList.Create;

  slGroupNumber := TStringList.Create;

  slPerErr := TStringList.Create;
       
  sline := '物料编码'#9'子项物料编码'#9'描述';
  slPerErr.Add(sline);
                                     
  iID := 1;
 

  ////  计算低位码  ////////////////////////////////////////////////////////////
  for idx := 0 to aSAPMaterialReader.Count - 1 do
  begin
    aSAPMaterialRecordPtr := aSAPMaterialReader.Items[idx];
    aSAPBomReader.GetLowestCode(aSAPMaterialRecordPtr);
//    Memo1.Lines.Add(aSAPMaterialRecordPtr^.sNumber + '   ' + IntToStr(aSAPMaterialRecordPtr^.iLowestCode));
  end;


  Memo1.Lines.Add('整理销售计划需求');
  for iLine := 0 to aSAPS618Reader.Count - 1 do
  begin
    aSAPS618 := aSAPS618Reader.Items[iLine];
    for iWeek := 0 to aSAPS618.Count - 1 do
    begin
      aSAPS618ColPtr := aSAPS618.Items[iWeek];    
      if DoubleE( aSAPS618ColPtr^.dqty, 0 ) then Continue;
      
      aMRPUnitPtr := TMRPUnit.Create;
      aMRPUnitPtr.id := iid;
      iid := iid + 1;
      aMRPUnitPtr.pid := 0;
      aMRPUnitPtr.snumber := aSAPS618ColPtr^.sNumber;
      aMRPUnitPtr.sname := aSAPS618ColPtr^.sname;
      aMRPUnitPtr.spnumber := '';
      aMRPUnitPtr.srnumber := aSAPS618ColPtr^.snumber;
      aMRPUnitPtr.dt := aSAPS618ColPtr^.dt1;
      aMRPUnitPtr.dQty := aSAPS618ColPtr^.dQty;
      aMRPUnitPtr.dQtyStock := 0;
      aMRPUnitPtr.dQtyStock2 := 0;
//      aMRPUnitPtr^.dQtyStockParent := 0;
      aMRPUnitPtr.dQtyOPO := 0;
      aMRPUnitPtr.bExpend := False;
      aMRPUnitPtr.bCalc := False;
      aMRPUnitPtr.aBom := nil;
      aMRPUnitPtr.aParentBom := nil;
      aMRPUnitPtr.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSAPS618ColPtr^.sNumber);
      aMRPUnitPtr.sDemandType := aSAPS618.sDemandType;
      aMRPUnitPtr.iSubstituteNo := 0;
      aMRPUnitPtr.spriority := '';
      aMRPUnitPtr.sMrpArea := aSAPS618.FMrpArea;
      aMRPUnitPtr.iAltCount := 1;
      lstDemand.Add(aMRPUnitPtr);
    end; 
  end;
                            
  Memo1.Lines.Add('开始模拟MRP计算');
  try
    iSubstituteNo := 1;  // 从
    iLowestCode := 0;
    bLoop := True;
    while bLoop do
    begin
      bLoop := False;
                      
      //排序，按日期
      lstDemand.Sort(ListSortCompare_DateTime);

      lstDemand_Count := lstDemand.Count;
      for iMrpUnit := 0 to lstDemand_Count - 1 do
      begin
        aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);

        if aMRPUnitPtr.bCalc then Continue; // 计算过的，不算

        // 低位码小于等于当前计算低位码，才计算
        if aMRPUnitPtr.aSAPMaterialRecordPtr^.iLowestCode > iLowestCode then
        begin
          bLoop := True;  // 有需求还没计算，需继续循环
          Continue;
        end;


        ////  自制件 展开BOM  //////////////////////////////////////////////
        if (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'E')
         or (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'X') then
        begin
          // 无无无无 替代料
          if aMRPUnitPtr.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr.bCalc := True;
            aMRPUnitPtr.bExpend := True;

            if aMRPUnitPtr.aParentBom = nil then // 根节点，需查找BOM
            begin
              aMRPUnitPtr.aBom := aSAPBomReader.GetSapBom(aMRPUnitPtr.snumber, slArea2BomFac.Values[aMRPUnitPtr.sMrpArea]);
            end;

            if aMRPUnitPtr.sDemandType <> 'BSF' then  //  LSF考虑库存， BSF 不考虑库存
            begin
              aMRPUnitPtr.dQtyStock := aSAPMrpAreaStockReader.AllocStock(
                aMRPUnitPtr.snumber, aMRPUnitPtr.dQty, aMRPUnitPtr.sMrpArea);
            end;

            // 分配库存后，需求满足了，无无无无需往下展开
            if DoubleLE( aMRPUnitPtr.dQty, aMRPUnitPtr.dQtyStock ) then
            begin
              Continue;
            end;
          
            if aMRPUnitPtr.aBom = nil then  // 有需求，但是 没有BOM，异常，记录日志
            begin
              Memo1.Lines.Add(aMRPUnitPtr.snumber + ' 无BOM'); 
              Continue;
            end;

            aSapBomChild := nil;

            //展开需求到下层
            for iChild := 0 to aMRPUnitPtr.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr.aBom.Childs[iChild];
              iPer100 := 0;

              slGroupNumber.Clear;
              for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
              begin
                aSapBomChild := aSapItemGroup.Items[iChildItem];
                slGroupNumber.Add(aSapBomChild.FNumber);
              end;
              slGroupNumber.Sort;
              sGroupNumbers := StringReplace(slGroupNumber.Text, #13#10, '|', [rfReplaceAll]);

              for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
              begin
                aSapBomChild := aSapItemGroup.Items[iChildItem];

                aMRPUnitPtr_Dep := TMRPUnit.Create;
                aMRPUnitPtr_Dep.id := iid;
                iid := iid + 1;

                aMRPUnitPtr_Dep.srnumber := aMRPUnitPtr.srnumber;
                aMRPUnitPtr_Dep.spnumber := aMRPUnitPtr.snumber;
                aMRPUnitPtr_Dep.sMrpArea := aMRPUnitPtr.sMrpArea;
                aMRPUnitPtr_Dep.pid := aMRPUnitPtr.id;
                aMRPUnitPtr_Dep.snumber := aSapBomChild.FNumber;
                aMRPUnitPtr_Dep.sname := aSapBomChild.FName;
                aMRPUnitPtr_Dep.dt := aMRPUnitPtr.dt - aMRPUnitPtr.aBom.lt;
                if aSapBomChild.sgroup = '' then
                begin
                  aMRPUnitPtr_Dep.iAltCount := 1;
                  aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage;
                  iPer100 := 100;
                end
                else
                begin
                  aMRPUnitPtr_Dep.iAltCount := aSapItemGroup.ItemCount;
                  // 有替代料，按配比分
                  aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                  iPer100 := iPer100 + Round(aSapBomChild.dPer);
                end;
                aMRPUnitPtr_Dep.sGroupNumbers := sGroupNumbers;
                
                aMRPUnitPtr_Dep.dQtyStock := 0;
                aMRPUnitPtr_Dep.dQtyStock2 := 0;
                aMRPUnitPtr_Dep.dQtyOPO := 0;
                aMRPUnitPtr_Dep.bExpend := False;
                aMRPUnitPtr_Dep.bCalc := False;
                aMRPUnitPtr_Dep.aBom := aSapBomChild;
                aMRPUnitPtr_Dep.aParentBom := aMRPUnitPtr.aBom;
                aMRPUnitPtr_Dep.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                if (aSapBomChild.sptype = 'E') or (aSapBomChild.sptype = 'X') then
                begin
                  aMRPUnitPtr_Dep.aSAPMaterialRecordPtr.sMRPType := 'M0';
                end
                else
                begin
                  aMRPUnitPtr_Dep.aSAPMaterialRecordPtr.sMRPType := 'PD';
                end;
                aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // 优先级
                aMRPUnitPtr_Dep.dPer := aSapBomChild.dPer;

                if aSapItemGroup.ItemCount = 1 then
                begin
                  aMRPUnitPtr_Dep.iSubstituteNo := 0; // 没有替代组
                end
                else
                begin
                  aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                end;
                lstDemand.Add(aMRPUnitPtr_Dep);
              end;
              iSubstituteNo := iSubstituteNo + 1; //  替代组编号 + 1，确保唯一

              // 配比总和不为 100
              if iPer100 <> 100 then
              begin
                Memo1.Lines.Add('配比总和不为 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr.snumber);
                sline := aSapBomChild.FNumber + #9 + aMRPUnitPtr.snumber + #9'配比总和不为 100(' + IntToStr(iPer100) + ')';
                slPerErr.Add(sline);
              end;

            end;
            bLoop := True;  //  展开有新的需求，需继续循环
          end
          else //  有替代料 // 半成品的替代料  /////////////////////////////////
          begin 
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            for idx := 0 to lstDemand.Count - 1 do //  查找替代组的所有踢掉料
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
              begin
                dQty := dQty + aMRPUnitPtr_Dep.dQty;  // 汇总替代料的需求
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // 替代料消耗优先级
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea);   // 如果某个替代料库存全满足了，剩余数量为0， 自动会为下面的替代料分配0的库存和需求
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep.dQtyStock;
              
              aMRPUnitPtr_Dep.bCalc := True;
            end;

            // 需求没有全满足， 按比例分配需求给替代料
            // 没需求也继续往下分配，把过程全部展现出来
//            if dQty > 0 then
//            begin
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr := lstSubstituteDemand[idx];  
                aMRPUnitPtr.bCalc := True;
                aMRPUnitPtr.bExpend := True;
                aMRPUnitPtr.dQty := aMRPUnitPtr.dQty + dQty * aMRPUnitPtr.dPer / 100;
//                if DoubleE( aMRPUnitPtr.dQty, 0) then Continue;

                // 分配了需求就往下展需求
                for iChild := 0 to aMRPUnitPtr.aBom.ChildCount - 1 do
                begin
                  aSapItemGroup := aMRPUnitPtr.aBom.Childs[iChild];


                  slGroupNumber.Clear;
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem];
                    slGroupNumber.Add(aSapBomChild.FNumber);
                  end;
                  slGroupNumber.Sort;
                  sGroupNumbers := StringReplace(slGroupNumber.Text, #13#10, '|', [rfReplaceAll]);


                  iPer100 := 0;
                  aSapBomChild := nil;
                  //展开需求到下层
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem];


                    aMRPUnitPtr_Dep := TMRPUnit.Create;
                    aMRPUnitPtr_Dep.id := iid;
                    iid := iid + 1;

                    aMRPUnitPtr_Dep.sMrpArea := aMRPUnitPtr.sMrpArea;
                    aMRPUnitPtr_Dep.pid := aMRPUnitPtr.id;
                    aMRPUnitPtr_Dep.snumber := aSapBomChild.FNumber;
                    aMRPUnitPtr_Dep.sname := aSapBomChild.FName;     
                    aMRPUnitPtr_Dep.srnumber := aMRPUnitPtr.srnumber;
                    aMRPUnitPtr_Dep.spnumber := aMRPUnitPtr.snumber;
                    aMRPUnitPtr_Dep.dt := aMRPUnitPtr.dt - aMRPUnitPtr.aBom.lt;
                    if aSapBomChild.sgroup = '' then
                    begin
                      aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage;
                      iPer100 := 100;
                    end
                    else
                    begin
                      // 有替代料，按配比分
                      aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                      iPer100 := iPer100 + Round(aSapBomChild.dPer);
                    end;                                    
                    aMRPUnitPtr_Dep.sGroupNumbers := sGroupNumbers;

                    aMRPUnitPtr_Dep.dQtyStock := 0;
                    aMRPUnitPtr_Dep.dQtyStock2 := 0;
                    aMRPUnitPtr_Dep.dQtyOPO := 0;
                    aMRPUnitPtr_Dep.bExpend := False;
                    aMRPUnitPtr_Dep.bCalc := False;
                    aMRPUnitPtr_Dep.aBom := aSapBomChild;
                    aMRPUnitPtr_Dep.aParentBom := aMRPUnitPtr.aBom;
                    aMRPUnitPtr_Dep.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                    if (aSapBomChild.sptype = 'E') or (aSapBomChild.sptype = 'X') then
                    begin
                      aMRPUnitPtr_Dep.aSAPMaterialRecordPtr^.sMRPType := 'M0';
                    end
                    else
                    begin
                      aMRPUnitPtr_Dep.aSAPMaterialRecordPtr^.sMRPType := 'PD';
                    end;
                    aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // 优先级
                    aMRPUnitPtr_Dep.dPer := aSapBomChild.dPer;

                    if aSapItemGroup.ItemCount = 1 then
                    begin
                      aMRPUnitPtr_Dep.iSubstituteNo := 0; // 没有替代组
                    end
                    else
                    begin
                      aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                    end;
                    lstDemand.Add(aMRPUnitPtr_Dep);
                  end;
                  iSubstituteNo := iSubstituteNo + 1; //  替代组编号 + 1，确保唯一

                  // 配比总和不为 100
                  if iPer100 <> 100 then
                  begin
                    Memo1.Lines.Add('配比总和不为 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr.snumber);    
                    sline := aSapBomChild.FNumber + #9 + aMRPUnitPtr.snumber + #9'配比总和不为 100(' + IntToStr(iPer100) + ')';
                    slPerErr.Add(sline);
                  end;
                  
                end;
                       

                
              end;                       // 之前分配 库存的时候已经分配了一部分Qty，所以要加          
              bLoop := True;  //  展开有新的需求，需继续循环
//            end;

            lstSubstituteDemand.Free;
          end;
        end
        else  //// 外购件，不展开BOM  ///////  PD  ////////////////////////////////
        begin
          if aMRPUnitPtr.bCalc then
          begin
            Continue;  // 已计算
          end;

          // 无无无无 替代料
          if aMRPUnitPtr.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr.snumber, aMRPUnitPtr.dQty, aMRPUnitPtr.sMrpArea);
            aMRPUnitPtr.bCalc := True;
          end
          else
          begin
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            for idx := 0 to lstDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
              begin
                dQty := dQty + aMRPUnitPtr_Dep.dQty;
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // 替代料消耗优先级 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea);   // 如果某个替代料库存全满足了，剩余数量为0， 自动会为下面的替代料分配0的库存和需求
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep.dQtyStock;
              
              aMRPUnitPtr_Dep.bCalc := True;
            end;

            // 需求没有全满足， 按比例分配需求给替代料
            if dQty > 0 then
            begin
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
                aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty * aMRPUnitPtr_Dep.dPer / 100;
              end;                       // 之前分配 库存的时候已经分配了一部分Qty，所以要加
            end; 

            lstSubstituteDemand.Free;
          end;
        end;        

      end;

      iLowestCode := iLowestCode + 1; 
    end;

    ////////////////////////////////////////////////////////////////////////////
    // 替代料优化计算  库存
    Memo1.Lines.Add('开始替代料优化计算  库存');
    AlternativeOptimization_stock(lstDemand, aSAPMrpAreaStockReader);



    
                                                 
    Memo1.Lines.Add('开始计算  OPO');

    //排序，按日期
    lstDemand.Sort(ListSortCompare_DateTime);

    (*
    //计算别的代工厂库存////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    lstDemand_Count := lstDemand.Count;
    for iMrpUnit := 0 to lstDemand_Count - 1 do
    begin
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);    
      aMRPUnitPtr.bCalc := False;
      aSAPMrpAreaStockReader.AccDemand(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea); 
    end;

    lstDemand_Count := lstDemand.Count;
    for iMrpUnit := 0 to lstDemand_Count - 1 do
    begin
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
      if aMRPUnitPtr.bCalc then Continue; // 替代料可能已经算过了

      // 无无无无 替代料
      if aMRPUnitPtr.iSubstituteNo = 0 then
      begin
        dQty := aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock; // 减去已分配过的库存
        if DoubleG(dQty, 0) then
        begin
          aMRPUnitPtr.dQtyStock2 := aSAPMrpAreaStockReader.AllocStock_area(aMRPUnitPtr.snumber, dQty, aMRPUnitPtr.sMrpArea);
        end;
        aMRPUnitPtr.bCalc := True;
      end
      else
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
          begin
            dQty := dQty + aMRPUnitPtr_Dep.dQty - aMRPUnitPtr_Dep.dQtyStock; // 减去已分配过的库存
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
          end;
        end;
          
        if DoubleG(dQty, 0) then
        begin
          // 替代料消耗优先级
          lstSubstituteDemand.Sort(ListSortCompare_priority);

          for idx := 0 to lstSubstituteDemand.Count - 1 do
          begin       
            aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
            aMRPUnitPtr_Dep.dQtyStock2 := aSAPMrpAreaStockReader.AllocStock_area(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea);   // 如果某个替代料库存全满足了，剩余数量为0， 自动会为下面的替代料分配0的库存和需求
            aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock + aMRPUnitPtr_Dep.dQtyStock2;
            dQty := dQty  - aMRPUnitPtr_Dep.dQtyStock2;      
            aMRPUnitPtr_Dep.bCalc := True;
          end;

          // 需求没有全满足， 按比例分配需求给替代料
          if dQty > 0 then
          begin
            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty * aMRPUnitPtr_Dep.dPer / 100;
            end;                       // 之前分配 库存的时候已经分配了一部分Qty，所以要加
          end; 
        end;
        lstSubstituteDemand.Free;
      end;      
    end;
    *)
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////        




     
    ////  计算 PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
      aMRPUnitPtr.bCalc := False;
    end;

    ////  计算 PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);

   
      if aMRPUnitPtr.snumber = '81.03.83100195H0' then
      begin
        Sleep(1);
      end;
      
      // 非外购件不计算 PO
      if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType <> 'F' then Continue;

      if aMRPUnitPtr.bCalc then Continue; //计算过了 

      slNumber := TStringList.Create;
      // 无无无 替代料
      if aMRPUnitPtr.iSubstituteNo = 0 then
      begin
        aMRPUnitPtr.bCalc := True;
                            
        if DoubleE( aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock - aMRPUnitPtr.dQtyStock2,  0) then Continue; // 没有需求

        slNumber.Add(aMRPUnitPtr.snumber);

        aMRPUnitPtr.dQtyOPO := aSAPMrpAreaStockReader.Alloc(slNumber, aMRPUnitPtr.dt,
          aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock - aMRPUnitPtr.dQtyStock2, aMRPUnitPtr.sMrpArea);

      end
      else  //  有有有有  替代料
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
          begin
            dQty := dQty + aMRPUnitPtr_Dep.dQty - aMRPUnitPtr_Dep.dQtyStock - aMRPUnitPtr_Dep.dQtyStock2; // 减去已分配库存
            aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock + aMRPUnitPtr_Dep.dQtyStock2;         // 分配的库存是固定的，不要清掉
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
            slNumber.Add(aMRPUnitPtr_Dep.snumber);


            aMRPUnitPtr_Dep.bCalc := True;
          end;
        end;

        if DoubleE(dQty, 0) then
        begin
          Continue; // 库存分配已满足需求
        end;

        // 替代料消耗优先级 
        lstSubstituteDemand.Sort(ListSortCompare_priority);

        // 先分配交期早的订单
        lstPOLine := TList.Create;
        //  内部已经排过序的，不需要再排序
        aSAPMrpAreaStockReader.GetOPOs(slNumber, lstPOLine, aMRPUnitPtr.sMrpArea); // 找到所有替代料的可用采购订单

        for idx := 0 to lstPOLine.Count - 1 do
        begin
          aSAPOPOLine := TSAPOPOLine(lstPOLine[idx]);
          aMRPUnitPtr_Dep := GetMRPUnit(lstSubstituteDemand, aSAPOPOLine.FNumber);

          // 分配的OPO 累加
          aMRPUnitPtr_Dep.dQtyOPO := aMRPUnitPtr_Dep.dQtyOPO + aSAPOPOLine.Alloc(aMRPUnitPtr_Dep.dt, dQty, aMRPUnitPtr_Dep.sMrpArea);
          aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock + aMRPUnitPtr_Dep.dQtyStock2 + aMRPUnitPtr_Dep.dQtyOPO;

          if DoubleE( dQty, 0) then  // 需求满足了
          begin
            Break;
          end;
        end;

        lstPOLine.Free;
         
        // 需求没有全满足， 按比例分配需求给替代料
        if dQty > 0 then
        begin
          for idx := 0 to lstSubstituteDemand.Count - 1 do
          begin       
            aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
            aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty * aMRPUnitPtr_Dep.dPer / 100; 
          end;
        end; 

        lstSubstituteDemand.Free;
      end;  


      slNumber.Free;
    end;

  finally

    aSAPBomReader.Free;

  end;


  ////////////////////////////////////////////////////////////////////////////
  // 替代料优化计算  库存
  //AlternativeOptimization_po(lstDemand, aSAPMrpAreaStockReader);

           
  //排序，替代料号， 日期， 料号
  lstDemand.Sort(ListSortCompare_ForSPQ);
  dQtyEx := 0;
  sGroupNumbers := '';
  for iMrpUnit := 0 to lstDemand.Count - 1 do
  begin        
    aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
    
    if (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'E') or
     (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'X') then Continue;      // 自制件不计算
    if DoubleE( aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ, 0 ) then Continue;// SPQ 为 0 或 1 不计算
    if DoubleE( aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ, 1 ) then Continue;

    dQtyNet := aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock -
      aMRPUnitPtr.dQtyStock2 - aMRPUnitPtr.dQtyOPO;
    if DoubleLE( dQtyNet , 0) then Continue;
    if (sGroupNumbers = '') or (sGroupNumbers <> aMRPUnitPtr.sGroupNumbers) then // 第一行 或者 换一个替代组 （ 没替代的，把自己当替代组 ）
    begin
      aMRPUnitPtr.dQtyNetSPQ := Round((dQtyNet / aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ) + 0.5) * aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ;
      dQtyEx := aMRPUnitPtr.dQtyNetSPQ - dQtyNet;
      sGroupNumbers := aMRPUnitPtr.sGroupNumbers;
    end
    else
    begin
      if DoubleLE( dQtyNet , dQtyEx ) then // 余量够满足
      begin
        dQtyEx := dQtyEx - dQtyNet;
        aMRPUnitPtr.dQtyNetSPQ := 0;
      end
      else
      begin
        aMRPUnitPtr.dQtyNetSPQ := Round(((dQtyNet - dQtyEx) / aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ) + 0.5) * aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ;
        dQtyEx := dQtyEx + aMRPUnitPtr.dQtyNetSPQ - dQtyNet;  // aMRPUnitPtr.dQtyNetSPQ - dQtyNet  可正可负
      end;
    end;
  end;
   
    
  slNumber := TStringList.Create;


  for iLine := 0 to aSAPStockReader.Count - 1 do
  begin
    aSAPStockRecordPtr := aSAPStockReader.Items[ iLine ];
    smrparea := aSAPMrpAreaStockReader.MrpAreaOfStockNo(aSAPStockRecordPtr^.sstock);
    if smrparea = '' then
    begin
      Memo1.Lines.Add('仓库 ' + aSAPStockRecordPtr^.sstock + ' 没有对应MRP区域');
    end;

    idx := slNumber.IndexOf(aSAPStockRecordPtr^.snumber);
    if idx < 0 then
    begin
      aEORecordPtr := New(PEORecord);
      aEORecordPtr^.snumber := aSAPStockRecordPtr^.snumber;
      aEORecordPtr^.sname := aSAPStockRecordPtr^.sname;
      aEORecordPtr^.sMrpAreaNo := smrparea;
      aEORecordPtr^.sMrpAreaName := aSAPMrpAreaStockReader.MrpAreaNo2Name(smrparea);
      aEORecordPtr^.dQtyDemand := 0;
      aEORecordPtr^.dQtyStock := 0;
      aEORecordPtr^.dQtyOPO := 0;      
      aEORecordPtr^.dQtyDemand17 := 0;
      aEORecordPtr^.dQtyDemand28 := 0;
      aEORecordPtr^.dQtyDemand60 := 0;
      aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aSAPStockRecordPtr^.snumber );
      slNumber.AddObject(aSAPStockRecordPtr^.snumber, TObject(aEORecordPtr));
    end
    else
    begin
      aEORecordPtr := PEORecord(slNumber.Objects[idx]);
    end;

    aEORecordPtr^.dQtyStock := aEORecordPtr^.dQtyStock + aSAPStockRecordPtr^.dqty;
  end;

          
  aSAPStockReader.Free;


  
  try
    // 保存 //////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('开始保存模拟MRP计算结果');

    Memo1.Lines.Add('保存独立需求');

    // 开始保存 Excel
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
    ExcelApp.DisplayAlerts := False;

    while ExcelApp.Sheets.Count > 1 do
    begin
      ExcelApp.Sheets[2].Delete;
    end;

    iSheet := 1;
    ExcelApp.Sheets[iSheet].Activate; 
    ExcelApp.Sheets[iSheet].Name := 'FCST';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '产品编码';
    ExcelApp.Cells[irow, 2].Value := '产品名称';
    ExcelApp.Cells[irow, 3].Value := '日期';
    ExcelApp.Cells[irow, 4].Value := '数量';    
    ExcelApp.Cells[irow, 5].Value := '需求类型';     
    ExcelApp.Cells[irow, 6].Value := 'MRP Area';
    ExcelApp.Cells[irow, 7].Value := '所用项目';

    irow := 2;

    for iLine := 0 to aSAPS618Reader.Count - 1 do
    begin
      aSAPS618 := aSAPS618Reader.Items[iLine];
      for iWeek := 0 to aSAPS618.Count - 1 do
      begin
        aSAPS618ColPtr := aSAPS618.Items[iWeek];    
        if DoubleE( aSAPS618ColPtr^.dqty, 0 ) then Continue;
      
        ExcelApp.Cells[irow, 1].Value := aSAPS618ColPtr^.sNumber;
        ExcelApp.Cells[irow, 2].Value := aSAPS618ColPtr^.sname;
        ExcelApp.Cells[irow, 3].Value := aSAPS618ColPtr^.dt1;
        ExcelApp.Cells[irow, 4].Value := aSAPS618ColPtr^.dQty;
        ExcelApp.Cells[irow, 5].Value := aSAPS618.sDemandType;
        ExcelApp.Cells[irow, 6].Value := aSAPS618.FMrpArea;    
        ExcelApp.Cells[irow, 7].Value := aSAPWhereUseReader.GetWhereUse(aSAPS618ColPtr^.sNumber);

        irow := irow + 1;
      end; 
    end;
            
    aSAPS618Reader.Free;
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('MRP Log');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'MRP Log';

    irow := 1;
//    ExcelApp.Cells[irow, 1].Value := 'ID';
//    ExcelApp.Cells[irow, 2].Value := '父ID';
//    ExcelApp.Cells[irow, 3].Value := '物料';
//    ExcelApp.Cells[irow, 4].Value := '物料名称';
//    ExcelApp.Cells[irow, 5].Value := '需求日期';
//    ExcelApp.Cells[irow, 6].Value := '建议下单日期';
//    ExcelApp.Cells[irow, 7].Value := '需求数量';
//    ExcelApp.Cells[irow, 8].Value := '可用库存';
//    ExcelApp.Cells[irow, 9].Value := 'OPO';
//    ExcelApp.Cells[irow, 10].Value := '净需求';
//    ExcelApp.Cells[irow, 11].Value := '替代组';
//    ExcelApp.Cells[irow, 12].Value := 'MRP控制者';
//    ExcelApp.Cells[irow, 13].Value := '采购员';
//    ExcelApp.Cells[irow, 14].Value := 'MRP区域';  
//    ExcelApp.Cells[irow, 15].Value := '上层料号';
//    ExcelApp.Cells[irow, 16].Value := '根料号';
//    ExcelApp.Cells[irow, 17].Value := 'L/T';

    lst := TList.Create;
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      lst.Add(lstDemand[iMrpUnit]);
    end;

    lst.Sort(ListSortCompare_pid);

    sl := TStringList.Create;
    try        //1      2       3         4            5            6             7           8         9     10          11        12          13         14         15         16       17           18     19     20      21
      sline := 'ID'#9'父ID'#9'物料'#9'物料名称'#9'需求日期'#9'建议下单日期'#9'需求数量'#9'可用库存'#9'OPO'#9'净需求'#9'SPQ净需求'#9'替代组'#9'MRP控制者'#9'采购员'#9'MRP区域'#9'上层料号'#9'根料号'#9'L/T'#9'SPQ'#9'MOQ'#9'替代物料组';
      sl.Add(sline);
 
      irow := irow + 1;
      while lst.Count > 0 do
      begin
        for iMrpUnit := 0 to lst.Count - 1 do
        begin
          aMRPUnitPtr := TMRPUnit(lst[iMrpUnit]);
          if aMRPUnitPtr.pid = 0 then
          begin    
            lst.Delete( iMrpUnit );
            // E&O /////////////////////////////////////////////     
            smrparea := aMRPUnitPtr.sMrpArea;

            idx := slNumber.IndexOf(aMRPUnitPtr.snumber);
            if idx < 0 then
            begin
              aEORecordPtr := New(PEORecord);
              aEORecordPtr^.snumber := aMRPUnitPtr.snumber;
              aEORecordPtr^.sname := aMRPUnitPtr.sname;
              aEORecordPtr^.sMrpAreaNo := smrparea;
              aEORecordPtr^.sMrpAreaName := aSAPMrpAreaStockReader.MrpAreaNo2Name(smrparea);
              aEORecordPtr^.dQtyDemand := 0;
              aEORecordPtr^.dQtyDemand17 := 0;
              aEORecordPtr^.dQtyDemand28 := 0;
              aEORecordPtr^.dQtyDemand60 := 0;
              aEORecordPtr^.dQtyStock := 0;
              aEORecordPtr^.dQtyOPO := 0;
              aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aMRPUnitPtr.snumber );
              slNumber.AddObject(aMRPUnitPtr.snumber, TObject(aEORecordPtr));
            end
            else
            begin
              aEORecordPtr := PEORecord(slNumber.Objects[idx]);
            end;

            aEORecordPtr^.dQtyDemand := aEORecordPtr^.dQtyDemand + aMRPUnitPtr.dQty;
        
            if DoubleL(aMRPUnitPtr.dt, today + 17) then
            begin
              aEORecordPtr^.dQtyDemand17 := aEORecordPtr^.dQtyDemand17 + aMRPUnitPtr.dQty;
            end;
            if DoubleL(aMRPUnitPtr.dt, today + 28) then
            begin
              aEORecordPtr^.dQtyDemand28 := aEORecordPtr^.dQtyDemand28 + aMRPUnitPtr.dQty;
            end;
            if DoubleL(aMRPUnitPtr.dt, today + 60) then
            begin
              aEORecordPtr^.dQtyDemand60 := aEORecordPtr^.dQtyDemand60 + aMRPUnitPtr.dQty;
            end; 

            //////////////////////////////////////////////////////////////////////

            sline := IntToStr(aMRPUnitPtr.id)+#9+           //1
              IntToStr(aMRPUnitPtr.pid)+#9+                 //2
              aMRPUnitPtr.snumber+#9+                       //3
              aMRPUnitPtr.sname+#9+                         //4
              FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt)+#9;  //5
            if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'F' then  // 外购  ////////////////////////////////////////
            begin
              sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD) + #9;    //6
            end
            else                                                         // 自制  ////////////////////////////////////////
            begin
              sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0) + #9;    //6
            end;

            sline := sline + Format('%0.0f', [aMRPUnitPtr.dqty])+#9+     //7
              Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2])+#9+   //8
              Format('%0.0f', [aMRPUnitPtr.dQtyOPO])+#9;    //9
            sline := sline + '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9+  //10      
              Format('%0.0f', [aMRPUnitPtr.dQtyNetSPQ]) + #9 +   //11
              IntToStr(aMRPUnitPtr.iSubstituteNo)+#9+    //12
              aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPer + #9 +   //13
              aMRPUnitPtr.aSAPMaterialRecordPtr.sBuyer +#9 +    //14
              aMRPUnitPtr.sMrpArea+#9+                          //15
              aMRPUnitPtr.spnumber+#9+                          //16
              aMRPUnitPtr.srnumber+#9;                          //17
              if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType = 'F' then
              begin
                sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD]) + #9;   //18
              end
              else
              begin
                sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0]) + #9;   //18
              end;
              sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dSPQ]) + #9;   //19
              sline := sline + Format('%0.0f', [aMRPUnitPtr.aSAPMaterialRecordPtr.dMOQ]) + #9;   //20
              sline := sline + aMRPUnitPtr.sGroupNumbers + #9;   //21

              
//            ExcelApp.Cells[irow, 1].Value := aMRPUnitPtr.id;// 'ID';
//            ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr.pid;// '父ID';
//            ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr.snumber;// '物料';
//            ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr.sname;// '物料名称';
//            ExcelApp.Cells[irow, 5].Value := aMRPUnitPtr.dt;// '需求日期';

            //'建议下单日期'
//            if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'F' then  // 外购  ////////////////////////////////////////
//            begin
//              ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD);
//            end
//            else                                                         // 自制  ////////////////////////////////////////
//            begin
//              ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0);
//            end;

//            ExcelApp.Cells[irow, 7].Value := Format('%0.0f', [aMRPUnitPtr.dqty]); //'需求数量';
//            ExcelApp.Cells[irow, 8].Value := Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2]); //'可用库存';
//            ExcelApp.Cells[irow, 9].Value := Format('%0.0f', [aMRPUnitPtr.dQtyOPO]); //'OPO';
//            ExcelApp.Cells[irow, 10].Value := '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow); //'净需求';
//            ExcelApp.Cells[irow, 11].Value := IntToStr(aMRPUnitPtr.iSubstituteNo); //'替代组';
//            ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPer; //'MRP控制者';
//            ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.sBuyer; //'采购员';
//            ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr.sMrpArea;
//            ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr.spnumber;
//            ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr.srnumber; 
//            if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType = 'F' then
//            begin
//              ExcelApp.Cells[irow, 17].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD;
//            end
//            else
//            begin
//              ExcelApp.Cells[irow, 17].Value := aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0;
//            end;

            sl.Add(sline);
            irow := irow + 1;

            WriteLine_MrpLog(ExcelApp, irow, aMRPUnitPtr, lst, slNumber,
              aSAPMrpAreaStockReader, aSAPMaterialReader, sl);

            Break;
          end;
        end;
      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.Cells[1, 1].Select;
      ExcelApp.ActiveSheet.Paste;
          
    finally
      sl.Free;
    end;

 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('保存  PR Sum');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'PR Sum';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '序号';
    ExcelApp.Cells[irow, 2].Value := '采购申请单类型';
    ExcelApp.Cells[irow, 3].Value := '抬头文本';
    ExcelApp.Cells[irow, 4].Value := '物料编码';
    ExcelApp.Cells[irow, 5].Value := '采购申请单数量';
    ExcelApp.Cells[irow, 6].Value := '交货日期';
    ExcelApp.Cells[irow, 7].Value := '工厂';
    ExcelApp.Cells[irow, 8].Value := '库存地点';
    ExcelApp.Cells[irow, 9].Value := '项目文本';      // 所用项目
    ExcelApp.Cells[irow, 10].Value := '建议下单日期';
    ExcelApp.Cells[irow, 11].Value := 'MRP控制者';
    ExcelApp.Cells[irow, 12].Value := '采购员';
    ExcelApp.Cells[irow, 13].Value := 'L/T';
    ExcelApp.Cells[irow, 14].Value := 'SPQ';
    ExcelApp.Cells[irow, 15].Value := 'MOQ';
    ExcelApp.Cells[irow, 16].Value := '采购类型';   
    ExcelApp.Cells[irow, 17].Value := '采购申请单数量SPQ';
    ExcelApp.Cells[irow, 18].Value := '物料名称';
    ExcelApp.Cells[irow, 19].Value := 'ABC';


    AddColor(ExcelApp, 1, 10, 1, 19, $FFFF);

    irow := irow + 1;
    
    iLine := 1;
    aMRPUnitPtr0 := nil;
    dqty_pr_sum := 0;
    dqty_pr_sum_spq := 0;
    slAreaStock := TStringList( mmoAreaStock.Lines );

    lst.Clear;
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      lst.Add(lstDemand[iMrpUnit]);
    end;
    lst.Sort(ListSortCompare_number_date);

    for iMrpUnit := 0 to lst.Count - 1 do
    begin
      aMRPUnitPtr := TMRPUnit(lst[iMrpUnit]);

      // 不是外购  ////////////////////////////////////////
      if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType <> 'F' then Continue;

      // 没需求
      dqty_pr := aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock -
        aMRPUnitPtr.dQtyStock2 - aMRPUnitPtr.dQtyOPO;
      if DoubleE(dqty_pr, 0) then Continue;

      if aMRPUnitPtr0 = nil then  // 第一行MrpLog
      begin
        aMRPUnitPtr0 := aMRPUnitPtr;
        dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
        dqty_pr_sum := dqty_pr;
        dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;
      end
      else
      begin
        if (aMRPUnitPtr0.snumber <> aMRPUnitPtr.snumber) or
          (aMRPUnitPtr0.sMrpArea <> aMRPUnitPtr.sMrpArea) then // 料号变了，生成一行PR
        begin
          // 判断是不是同一周的

          ExcelApp.Cells[irow, 1].Value := '''' + Copy( IntToStr(1000 + iLine), 2, 3 );
          ExcelApp.Cells[irow, 2].Value := 'NB';
          ExcelApp.Cells[irow, 3].Value := '';
          ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr0.snumber;
          ExcelApp.Cells[irow, 5].Value := Format('%0.0f', [dqty_pr_sum]);
          ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyyMMdd', aMRPUnitPtr0.dt);
          ExcelApp.Cells[irow, 7].Value := '1001';
          ExcelApp.Cells[irow, 8].Value := slAreaStock.Values[ aMRPUnitPtr0.sMrpArea ];
          ExcelApp.Cells[irow, 9].Value := aSAPWhereUseReader.GetWhereUse(aMRPUnitPtr0.snumber);
          ExcelApp.Cells[irow, 10].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr0.dt - aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD);
          ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPer;
          ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
          ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
          ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
          ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
          ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;
          ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);


          ExcelApp.Cells[irow, 18].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sName;
          ExcelApp.Cells[irow, 19].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sAbc;

          

          aMRPUnitPtr0 := aMRPUnitPtr;
          dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
          dqty_pr_sum := dqty_pr;  
          dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;

          iLine := iLine + 1;   // 分不同PR
          
          irow := irow + 1;
        end
        else if dtMon + 7 < aMRPUnitPtr.dt then // 下一周
        begin

          ExcelApp.Cells[irow, 1].Value := '''' + Copy( IntToStr(1000 + iLine), 2, 3 );
          ExcelApp.Cells[irow, 2].Value := 'NB';
          ExcelApp.Cells[irow, 3].Value := '';
          ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr0.snumber;
          ExcelApp.Cells[irow, 5].Value := Format('%0.0f', [dqty_pr_sum]);
          ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyyMMdd', aMRPUnitPtr0.dt);
          ExcelApp.Cells[irow, 7].Value := '1001';
          ExcelApp.Cells[irow, 8].Value := slAreaStock.Values[ aMRPUnitPtr0.sMrpArea ];
          ExcelApp.Cells[irow, 9].Value := aSAPWhereUseReader.GetWhereUse(aMRPUnitPtr0.snumber);
          ExcelApp.Cells[irow, 10].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr0.dt - aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD);
          ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPer;
          ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
          ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
          ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
          ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
          ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;
          ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);
                                                  
          ExcelApp.Cells[irow, 18].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sName;
          ExcelApp.Cells[irow, 19].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sAbc;

          
          aMRPUnitPtr0 := aMRPUnitPtr;
          dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
          dqty_pr_sum := dqty_pr;        
          dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;
            
          irow := irow + 1; 
        end
        else // 同料号， 同区域， 同周，合并数量
        begin
          dqty_pr_sum := dqty_pr_sum + dqty_pr;
          dqty_pr_sum_spq := dqty_pr_sum_spq + aMRPUnitPtr.dQtyNetSPQ;
        end;
      end;

    end;

    if aMRPUnitPtr0 <> nil then
    begin
      ExcelApp.Cells[irow, 1].Value := '''' + Copy( IntToStr(1000 + iLine), 2, 3 );
      ExcelApp.Cells[irow, 2].Value := 'NB';
      ExcelApp.Cells[irow, 3].Value := '';
      ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr0.snumber;
      ExcelApp.Cells[irow, 5].Value := Format('%0.0f', [dqty_pr_sum]);
      ExcelApp.Cells[irow, 6].Value := FormatDateTime('yyyyMMdd', aMRPUnitPtr0.dt);
      ExcelApp.Cells[irow, 7].Value := '1001';
      ExcelApp.Cells[irow, 8].Value := slAreaStock.Values[ aMRPUnitPtr0.sMrpArea ];
      ExcelApp.Cells[irow, 9].Value := aSAPWhereUseReader.GetWhereUse(aMRPUnitPtr0.snumber);
      ExcelApp.Cells[irow, 10].Value := FormatDateTime('yyyy-MM-dd', aMRPUnitPtr0.dt - aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD);
      ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPer;
      ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
      ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
      ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
      ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
      ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;   
      ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);     
      ExcelApp.Cells[irow, 18].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sName;
      ExcelApp.Cells[irow, 19].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sAbc;
    end;

    lst.Free;
 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('保存  PO Action');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'PO Action';


    sl := TStringList.Create;
    try
      sline := '采购凭证'#9'行号'#9'物料'#9'物料名称'#9'建议数量'#9'建议到料日期'#9'订单交期'#9'建议'#9'MRP Area'#9'Mrp Area No'#9'MC'#9'采购员'#9'LT'#9'SPQ'#9'MOQ'#9'凭证日期'#9'供应商';
      sl.Add(sline);
 
      for iLine := 0 to aSAPOPOReader2.Count - 1 do
      begin
        aSAPOPOLine := aSAPOPOReader2.Items[iLine];
        aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSAPOPOLine.FNumber);
        for iLineAlloc := 0 to aSAPOPOLine.Count - 1 do
        begin
          aSAPOPOAllocPtr := aSAPOPOLine.Items[iLineAlloc];

          sline := aSAPOPOLine.FBillNo + #9 +
            aSAPOPOLine.FLine + #9 +
            aSAPOPOLine.FNumber + #9 +
            aSAPOPOLine.FName + #9 +
            Format('%0.0f', [aSAPOPOAllocPtr^.dQty]) + #9 +
            FormatDatetime('yyyy-MM-dd', aSAPOPOAllocPtr^.dt) + #9 +
            FormatDateTime('yyyy-MM-dd', aSAPOPOLine.FDate);

          if DoubleE( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // 准时交货
          begin
            sline := sline + #9 + 'OTD';
          end
          else if DoubleG( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // 订单交期， 晚于需求日期，Push In
          begin
            sline := sline + #9 + 'Push In';
          end
          else            // 订单日期日期，早于需求日期, Push Out
          begin
            sline := sline + #9 + 'Push Out';
          end;

          sline := sline + #9 + aSAPMrpAreaStockReader.MrpAreaNo2Name(aSAPOPOAllocPtr^.sMrpAreaNo);
          sline := sline + #9 + aSAPOPOAllocPtr^.sMrpAreaNo;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPer;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_PD]); // 外购件肯定 是计划交货时间
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dSPQ]);
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dMOQ]);  
          sline := sline + #9 + FormatDatetime('yyyy-MM-dd', aSAPOPOLine.FBillDate);
          sline := sline + #9 + aSAPOPOLine.FSupplier;

          sl.Add(sline);
        end;


        // 没有分配的
        if DoubleG( aSAPOPOLine.FQty, aSAPOPOLine.FQtyAlloc ) then
        begin
          sline := aSAPOPOLine.FBillNo + #9 +
            aSAPOPOLine.FLine + #9 +
            aSAPOPOLine.FNumber + #9 +
            aSAPOPOLine.FName + #9 +
            Format('%0.0f', [aSAPOPOLine.FQty - aSAPOPOLine.FQtyAlloc]) + #9 +
            '' + #9 +
            FormatDatetime('yyyy-MM-dd', aSAPOPOLine.FDate) + #9 +
            'Cancel' + #9 +
            '' + #9 +
            '';
          sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPer;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;     
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_PD]); // 外购件肯定 是计划交货时间
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dSPQ]);
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dMOQ]);   
          sline := sline + #9 + FormatDatetime('yyyy-MM-dd', aSAPOPOLine.FBillDate);
          sline := sline + #9 + aSAPOPOLine.FSupplier;
          
          sl.Add(sline);
        end;

        // E&O ////////            
        smrparea := aSAPMrpAreaStockReader.MrpAreaOfStockNo(aSAPOPOLine.FStock);

        idx := slNumber.IndexOf(aSAPOPOLine.FNumber);
        if idx < 0 then
        begin
          aEORecordPtr := New(PEORecord);
          aEORecordPtr^.snumber := aSAPOPOLine.FNumber;   
          aEORecordPtr^.sname := aSAPOPOLine.FName;    
          aEORecordPtr^.sMrpAreaNo := smrparea;   
          aEORecordPtr^.sMrpAreaName := aSAPMrpAreaStockReader.MrpAreaNo2Name(smrparea);
          aEORecordPtr^.dQtyDemand := 0;
          aEORecordPtr^.dQtyDemand17 := 0;
          aEORecordPtr^.dQtyDemand28 := 0;
          aEORecordPtr^.dQtyDemand60 := 0;
          aEORecordPtr^.dQtyStock := 0;
          aEORecordPtr^.dQtyOPO := 0;   
          aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aSAPOPOLine.Fnumber );
          slNumber.AddObject(aSAPOPOLine.FNumber, TObject(aEORecordPtr));
        end
        else
        begin
          aEORecordPtr := PEORecord(slNumber.Objects[idx]);
        end;

        aEORecordPtr^.dQtyOPO := aEORecordPtr^.dQtyOPO + aSAPOPOLine.FQty;

      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;      
          
    finally
      sl.Free;
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    
    aSAPOPOReader2.Free;

                        
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('保存 E&o');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'E&O';


    sl := TStringList.Create;
    try
      sline := '物料'#9'物料名称'#9'总需求'#9'总库存'#9'总订单'#9'库存呆滞'#9'订单呆滞'#9'总呆滞'#9'E/O'#9'17天需求'#9'28天需求'#9'60天需求'#9'MRP类型'#9'MRP区域'#9'MC'#9'采购员'#9'LT'#9'SPQ'#9'MOQ'#9'采购类型'#9'MRP组';
      sl.Add(sline);

      irow := 2;
      for iLine := 0 to slNumber.Count - 1 do
      begin
        aEORecordPtr := PEORecord(slNumber.Objects[iLine]);    
        aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aEORecordPtr^.snumber);
 
        sline := aEORecordPtr^.snumber + #9 +
          aEORecordPtr^.sname + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyStock]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyOPO]) + #9 +
          '=IF(D' + IntToStr(irow) + '-C' + IntToStr(irow) + '>0,D' + IntToStr(irow) + '-C' + IntToStr(irow) + ',0)'#9 +
          '=IF(D' + IntToStr(irow) + '>=C' + IntToStr(irow) + ',E' + IntToStr(irow) + ',IF(E' + IntToStr(irow) + '-(C' + IntToStr(irow) + '-D' + IntToStr(irow) + ')>0,E' + IntToStr(irow) + '-(C' + IntToStr(irow) + '-D' + IntToStr(irow) + '),0))'#9 +
          '=IF(D' + IntToStr(irow) + '+E' + IntToStr(irow) + '-C' + IntToStr(irow) + ' > 0, D' + IntToStr(irow) + '+E' + IntToStr(irow) + '-C' + IntToStr(irow) + ', 0)' + #9 +          
          '=IF(E' + IntToStr(irow) + '+D' + IntToStr(irow) + '>C' + IntToStr(irow) + ',IF(C' + IntToStr(irow) + '>0,"Excess","Obslete"),"")' + #9 +          
          Format('%0.0f', [aEORecordPtr^.dQtyDemand17]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand28]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand60]) + #9 +
          aEORecordPtr^.sMRPType + #9 +
          aEORecordPtr^.sMrpAreaName;        
        sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPer;
        sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;
        if aSAPMaterialRecordPtr^.sPType = 'F' then
        begin
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_PD]);
        end
        else
        begin
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_M0]);
        end;
        sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dSPQ]);
        sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dMOQ]);    
        sline := sline + #9 + aSAPMaterialRecordPtr^.sPType;
        sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPGroup;

        sl.Add(sline);
        irow := irow + 1;

      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;
          
    finally
      sl.Free;
    end;

       
    FreeAndNil(aSAPMaterialReader);


    for iLine := 0 to slNumber.Count - 1 do
    begin
      aEORecordPtr := PEORecord(slNumber.Objects[iLine]);
      Dispose(aEORecordPtr);
    end;
    slNumber.Free;

        
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('库存');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'OH';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leSAPStock.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;
          stitle5 := ExcelApp2.Cells[irow, 5].Value;
          stitle6 := ExcelApp2.Cells[irow, 6].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;

          if stitle <> '工厂库存地点仓储地点的描述物料物料描述非限制使用的库存' then
          begin
            Memo1.Lines.Add(sSheet +'  不是  SAP导出库存  格式');
            Continue;
          end;

          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste;   
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;
 
          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;         

        
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('OPO');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'OPO';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leSAPOPO.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;
          stitle5 := ExcelApp2.Cells[irow, 5].Value;
          stitle6 := ExcelApp2.Cells[irow, 6].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4 + stitle5 + stitle6;

          if stitle <> '采购凭证项目计划行采购凭证类型采购凭证类别采购组' then
          begin
            Memo1.Lines.Add(sSheet +'  不是  SAP导出OPO  格式');
            Continue;
          end;

          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste;  
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;
 
          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;
             
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('BOM');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'BOM';
          

    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := '应用程序调用 Microsoft Excel';
    try

      WorkBook2 := ExcelApp2.WorkBooks.Open(leSAPBom.Text);


      try
        iSheetCount2 := ExcelApp2.Sheets.Count;
        for iSheet2 := 1 to iSheetCount2 do
        begin
          if not ExcelApp2.Sheets[iSheet2].Visible then Continue;

          ExcelApp2.Sheets[iSheet2].Activate;

          sSheet := ExcelApp2.Sheets[iSheet2].Name;
          Memo1.Lines.Add(sSheet);

          irow := 1;
          stitle1 := ExcelApp2.Cells[irow, 1].Value;
          stitle2 := ExcelApp2.Cells[irow, 2].Value;
          stitle3 := ExcelApp2.Cells[irow, 3].Value;
          stitle4 := ExcelApp2.Cells[irow, 4].Value;
          stitle := stitle1 + stitle2 + stitle3 + stitle4;

          if stitle <> '母件物料编码母件物料描述工厂用途' then
          begin         
            Memo1.Lines.Add(sSheet +'  不是SAP导出BOM格式');
            Continue;
          end;
          
          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste; 
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;     

          ExcelApp2.CutCopyMode := False;


          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //新加的,设置已经保存
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;

   
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('配比不为100');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'BOM配比';
          

    Clipboard.SetTextBuf(PChar(slPerErr.Text));
    ExcelApp.ActiveSheet.Paste;

    slPerErr.Free;

    

    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    
  finally 
    Clipboard.SetTextBuf('');

    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
      aMRPUnitPtr.Free;
    end;
    lstDemand.Free;

    aSAPMrpAreaStockReader.Free;
  end;

  slArea2BomFac.Free;
  aSAPWhereUseReader.Free;
  
  slGroupNumber.Free;

  MessageBox(Handle, '完成', '提示', 0);
end;

procedure TfrmMRP4SAP3_MRP2.btnWhereUseClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leWhereUse.Text := sfile;
end;

end.
