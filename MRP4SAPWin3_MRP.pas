unit MRP4SAPWin3_MRP;
 
(*
ע�����
  1��SAP������BOM Excel�ļ�����Щֻ�и������ϣ�û���������ϣ����Excel�ļ�ɾ����������MRP���������Ҳ���BOM
  2��SAP�����Ŀ���ļ���ע��ֻ�������в���MRP����Ĳֿ�
  3��Excel������ݶ�Ҫ��Sheet1
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
  SAPStockReader2, SAPBomReader, SAPBomReader3, SAPS618Reader, SAPMrpAreaStockReader;

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
                                                    
  TfrmMRP4SAP3_MRP = class(TForm)
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
    scategory: string; //�������
    sg: string; //ͨ����
    snumber: string; //���ϱ���(����)
    sname: string; //��������
    slt: string; //LT
    smc: string; //MC
    sproj: string; //��Ŀ
  end;
    
  TOrderByConditions = packed record
    isnew: Boolean;
    dos: Double;
    demand: Double;
  end;
  POrderByConditions = ^TOrderByConditions;

  // �򻯵�MRP���㣬�����ǵ�λ��

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
    //dQty_old: Double;
    dQtyStock: Double;
    dQtyStock2: Double;
    dQtyOPO: Double;
    bExpend: Boolean;
    bCalc: Boolean;
    aBom: TSapBom;
    aParentBom: TSapBom;
    aSAPMaterialRecordPtr: PSAPMaterialRecord;
    sDemandType: string;
    iSubstituteNo: Integer; // �������
    spriority: string;
    dPer: Double;
    sMrpArea: string;
    iAltCount: Integer; //ͨ�ù�ϵ������ûͨ�ù�ϵΪ1
    sGroupNumbers: string; // ͨ�ù�ϵ�����ϱ���ƴ�ӣ������ж���ȫ��ͬ��ͨ�ù�ϵ
    dQtyNetSPQ: Double; // �����󣬰�SPQ
    sNote: string;
    bSupp: Boolean; // �Ƿ���Ȳ������� ����ǣ� ����������䣬�ɹ�PO����
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
    if u1.snumber < u2.snumber then
      Result := 1
    else if u1.snumber > u2.snumber then
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





class procedure TfrmMRP4SAP3_MRP.ShowForm;
var
  frmMRP4SAP3: TfrmMRP4SAP3_MRP;
begin
  frmMRP4SAP3 := TfrmMRP4SAP3_MRP.Create(nil);
  try
    frmMRP4SAP3.ShowModal;
  finally
    frmMRP4SAP3.Free;
  end;
end;
   
procedure TfrmMRP4SAP3_MRP.FormCreate(Sender: TObject);
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

 
  s := ini.ReadString(self.ClassName, mmoAreaStock.Name, '');
  mmoAreaStock.Text := StringReplace(s, '||', #13#10, [rfReplaceAll]);

  leWhereUse.Text := ini.ReadString(self.ClassName, leWhereUse.Name, '');
        
 
  ini.Free;
end;

procedure TfrmMRP4SAP3_MRP.FormDestroy(Sender: TObject);
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
 
procedure TfrmMRP4SAP3_MRP.tbCloseClick(Sender: TObject);
begin
  Close;
end;

constructor TMRPUnit.Create;
begin
  inherited;
  bSupp := False;
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
      begin  // ʱ����ͬ�� ͨ�ù�ϵ�ٵģ�����ǰ��
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
        begin  // ʱ����ͬ�� ͨ�ù�ϵ�ٵģ�����ǰ��
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

 
procedure TfrmMRP4SAP3_MRP.btnSAPStock3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;
     
procedure TfrmMRP4SAP3_MRP.btnSAPOPOClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPOPO.Text := sfile;
end;
      
procedure TfrmMRP4SAP3_MRP.btnMaterialClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMaterial.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP.btnSAPBom3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP.btnDemand3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPPIR.Text := sfile;
end;
    
procedure TfrmMRP4SAP3_MRP.btnSAPMrpAreaStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPMrpAreaStock.Text := sfile;
end;

procedure TfrmMRP4SAP3_MRP.OnLogEvent(const s: string);
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
����ָ����ֵ�Ƿ��ڵ�ǰ������(�����Ѿ��������)
*}
function SearchData(lst: TList; pid: longint ): Integer;
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

//function GetChildMrpLog(lst: TList; aMRPUnitPtr: TMRPUnit): Integer;
//var
//  idx: Integer;
//  a: TMRPUnit;
//begin 
//  Result := -1;
//  for idx := 0 to lst.Count - 1 do
//  begin
//    a := TMRPUnit(lst[idx]);
//    if a.pid = aMRPUnitPtr.id then
//    begin
//      Result := idx;
//      Break;
//    end;
//  end;
//end;

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
  idx := SearchData(lst, p.id);
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
    if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'F' then  // �⹺  ////////////////////////////////////////
    begin
      sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD) + #9;      //6
    end
    else                                                         // ����  ////////////////////////////////////////
    begin
      sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0) + #9;      //6
    end;

    sline := sline + Format('%0.0f', [aMRPUnitPtr.dqty])+#9+      //7
      Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2])+#9+   //8
      Format('%0.0f', [aMRPUnitPtr.dQtyOPO])+#9;   //9
    sline := sline + '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9+  //10
      Format('%0.0f', [aMRPUnitPtr.dQtyNetSPQ]) + #9 +    // 11
      IntToStr(aMRPUnitPtr.iSubstituteNo)+#9+   //12
      aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPerDesc + #9 +   //13
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
    sline := sline + aMRPUnitPtr.sNote + #9;   //22


    sl.Add(sline);
 
    irow := irow + 1;

    WriteLine_MrpLog(ExcelApp, irow, aMRPUnitPtr, lst, slNumber,
      aSAPMrpAreaStockReader, aSAPMaterialReader, sl);


    //idx := GetChildMrpLog(lst, p)    
    idx := SearchData(lst, p.id);
  end; 
end;

   
procedure TfrmMRP4SAP3_MRP.btnMRPClick(Sender: TObject);
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
  aSAPBomReader: TSAPBomReader3;
  aSAPStockReader: TSAPStockReader2;
  //aSAPStockSum: TSAPStockSum;
  aSAPS618Reader: TSAPPIRReader;
  aSAPOPOReader2: TSAPOPOReader2;
  lstDemand: TList;
  lstDemand_tmp: TList;
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
  dPerSum: Double;
  dQty_delta: Double;
  dQty_left: Double;
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
  iSubstituteNo: Integer;  // �������
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
                                         
  Memo1.Lines.Add('��ʼ��ȡ MRP����ֿ��б�  ' + leSAPMrpAreaStock.Text);
  aSAPMrpAreaStockReader := TSAPMrpAreaStockReader.Create(leSAPMrpAreaStock.Text);
  
  Memo1.Lines.Add('��ʼ��ȡ BOM  ' + leSAPBom.Text);
  aSAPBomReader := TSAPBomReader3.Create(leSAPBom.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ ���  ' + leSAPStock.Text);
  aSAPStockReader := TSAPStockReader2.Create(leSAPStock.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ OPO  ' + leSAPOPO.Text);
  aSAPOPOReader2 := TSAPOPOReader2.Create(leSAPOPO.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ ����  ' + leMaterial.Text);
  aSAPMaterialReader:= TSAPMaterialReader2.Create(leMaterial.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ PIR  ' + leSAPPIR.Text);
  aSAPS618Reader := TSAPPIRReader.Create(leSAPPIR.Text, OnLogEvent);



  Memo1.Lines.Add('��ʼ��ȡ �ڲ���Ŀ  ' + leWhereUse.Text);
  aSAPWhereUseReader := TSAPWhereUseReader.Create(leWhereUse.Text, OnLogEvent);


  aSAPMrpAreaStockReader.SetOPOList(aSAPOPOReader2);
  aSAPMrpAreaStockReader.SetStock(aSAPStockReader);

  //aSAPStockSum := TSAPStockSum.Create;
  //aSAPStockReader.SumTo(aSAPStockSum);

//  lstMrpDetail := TList.Create;

  lstDemand := TList.Create;

  slGroupNumber := TStringList.Create;

  slPerErr := TStringList.Create;
       
  sline := '���ϱ���'#9'����'#9'�������ϱ���'#9'��������'#9'���'#9'˵��';
  slPerErr.Add(sline);
                                     
  iID := 1;
 

  ////  �����λ��  ////////////////////////////////////////////////////////////
  for idx := 0 to aSAPMaterialReader.Count - 1 do
  begin
    aSAPMaterialRecordPtr := aSAPMaterialReader.Items[idx];
    aSAPBomReader.GetLowestCode(aSAPMaterialRecordPtr);
//    Memo1.Lines.Add(aSAPMaterialRecordPtr^.sNumber + '   ' + IntToStr(aSAPMaterialRecordPtr^.iLowestCode));
  end;


  Memo1.Lines.Add('�������ۼƻ�����');
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
                            
  Memo1.Lines.Add('��ʼģ��MRP����');
  try
    iSubstituteNo := 1;  // ��
    iLowestCode := 0;
    bLoop := True;
    while bLoop do
    begin
      bLoop := False;
                      
      //���򣬰�����
      lstDemand.Sort(ListSortCompare_DateTime);

      lstDemand_Count := lstDemand.Count;
      for iMrpUnit := 0 to lstDemand_Count - 1 do
      begin
        aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);

        if aMRPUnitPtr.bCalc then Continue; // ������ģ�����

        // ��λ��С�ڵ��ڵ�ǰ�����λ�룬�ż���
        if aMRPUnitPtr.aSAPMaterialRecordPtr^.iLowestCode > iLowestCode then
        begin
          bLoop := True;  // ������û���㣬�����ѭ��
          Continue;
        end;


        ////  ���Ƽ� չ��BOM  //////////////////////////////////////////////
        if (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'E') or
          (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'X') then
        begin
          // �������� �����
          if aMRPUnitPtr.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr.bCalc := True;
            aMRPUnitPtr.bExpend := True;

            if aMRPUnitPtr.aParentBom = nil then // ���ڵ㣬�����BOM
            begin
              aMRPUnitPtr.aBom := aSAPBomReader.GetSapBom(aMRPUnitPtr.snumber, slArea2BomFac.Values[aMRPUnitPtr.sMrpArea]);
            end;
                                                          // ��Ȳ������󲻷�����
            if ( aMRPUnitPtr.sDemandType <> 'BSF' ) and (not aMRPUnitPtr.bSupp) then  //  LSF���ǿ�棬 BSF �����ǿ��
            begin
              aMRPUnitPtr.dQtyStock := aSAPMrpAreaStockReader.AllocStock(
                aMRPUnitPtr.snumber, aMRPUnitPtr.dQty, aMRPUnitPtr.sMrpArea);
            end;

            // ����������������ˣ���������������չ��
            if DoubleLE( aMRPUnitPtr.dQty, aMRPUnitPtr.dQtyStock ) then
            begin
              Continue;
            end;
          
            if aMRPUnitPtr.aBom = nil then  // �����󣬵��� û��BOM���쳣����¼��־
            begin
              Memo1.Lines.Add(aMRPUnitPtr.snumber + ' ��BOM'); 
              Continue;
            end;

//            aSapBomChild := nil;

            //չ�������²�
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
                  // ������ϣ�����ȷ�
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
                aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // ���ȼ�
                aMRPUnitPtr_Dep.dPer := aSapBomChild.dPer;

                if aSapItemGroup.ItemCount = 1 then
                begin
                  aMRPUnitPtr_Dep.iSubstituteNo := 0; // û�������
                end
                else
                begin
                  aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                end;
                aMRPUnitPtr_Dep.bSupp := aMRPUnitPtr.bSupp;   // �̳и�����
                aMRPUnitPtr_Dep.bCalc := aMRPUnitPtr_Dep.bSupp;
                
                lstDemand.Add(aMRPUnitPtr_Dep);
              end;
              // ����ܺͲ�Ϊ 100
              if iPer100 <> 100 then
              begin
                // �������󣬱�����
                if aSapItemGroup.ItemCount > 0 then
                begin
                  aSapBomChild := aSapItemGroup.Items[0];

                  if aSapBomChild.FNumber = '01.04.1330012000A0' then
                  begin
                    Sleep(1);
                  end;

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

                  aMRPUnitPtr_Dep.iAltCount := aSapItemGroup.ItemCount;
                  // ������ϣ�����ȷ�
                  aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage * (100 - iPer100) / 100;


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
                  aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // ���ȼ�
                  aMRPUnitPtr_Dep.dPer := (100 - iPer100);

                  aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                  aMRPUnitPtr_Dep.sNote := '�����1';
                  aMRPUnitPtr_Dep.bSupp := True;
                  aMRPUnitPtr_Dep.bCalc := aMRPUnitPtr_Dep.bSupp;

                  lstDemand.Add(aMRPUnitPtr_Dep);                  
                end;

                if slPerErr.IndexOf(aMRPUnitPtr.snumber) < 0 then
                begin
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem];
                    Memo1.Lines.Add('����ܺͲ�Ϊ 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr.snumber);
                    sline := aMRPUnitPtr.snumber + #9 + aMRPUnitPtr.sname + #9 + aSapBomChild.FNumber + #9 + aSapBomChild.FName + #9 + FloatToStr(aSapBomChild.dPer) + #9'����ܺͲ�Ϊ 100';
                    slPerErr.Add(sline);
                  end;
                end;
              end;

              iSubstituteNo := iSubstituteNo + 1; //  ������� + 1��ȷ��Ψһ
            end;
            bLoop := True;  //  չ�����µ����������ѭ��
          end
          else //  ������� // ���Ʒ�������  /////////////////////////////////
          begin 
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            dPerSum := 0;
            for idx := 0 to lstDemand.Count - 1 do //  ���������������ߵ���
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep.bSupp then Continue;   // ��Ȳ����������·�������
              if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
              begin                                               
                dPerSum := dPerSum + aMRPUnitPtr_Dep.dPer;
                dQty := dQty + aMRPUnitPtr_Dep.dQty;  // ��������ϵ�����
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // ������������ȼ�
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx]; 
              aMRPUnitPtr_Dep.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep.dQtyStock;

              aMRPUnitPtr_Dep.bCalc := True;
            end;

            // ����û��ȫ���㣬 ��������������������
            // û����Ҳ�������·��䣬�ѹ���ȫ��չ�ֳ���

            dQty_left := dQty;
            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr := lstSubstituteDemand[idx];  
              aMRPUnitPtr.bCalc := True;
              aMRPUnitPtr.bExpend := True;
              //aMRPUnitPtr.dQty := aMRPUnitPtr.dQty + dQty * aMRPUnitPtr.dPer / 100;

              if dPerSum > 0 then
              begin
                if idx = lstSubstituteDemand.Count - 1 then
                begin
                  aMRPUnitPtr.dQty := aMRPUnitPtr.dQty + dQty_left;
                end
                else
                begin
                  dQty_delta := dQty * (aMRPUnitPtr.dPer / dPerSum);
                  aMRPUnitPtr.dQty := aMRPUnitPtr.dQty + dQty_delta;
                  dQty_left := dQty_left - dQty_delta;
                end; 
              end;
                
 
              // ���������������չ����
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
//                  aSapBomChild := nil;
                //չ�������²�
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
                    // ������ϣ�����ȷ�
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
                  aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // ���ȼ�
                  aMRPUnitPtr_Dep.dPer := aSapBomChild.dPer;

                  if aSapItemGroup.ItemCount = 1 then
                  begin
                    aMRPUnitPtr_Dep.iSubstituteNo := 0; // û�������
                  end
                  else
                  begin
                    aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                  end;
                  aMRPUnitPtr_Dep.bSupp := aMRPUnitPtr.bSupp;
                  aMRPUnitPtr_Dep.bCalc := aMRPUnitPtr_Dep.bSupp;
                  lstDemand.Add(aMRPUnitPtr_Dep);
                end;

                // ����ܺͲ�Ϊ 100
                if iPer100 <> 100 then
                begin

                  // �������󣬱�����
                  if aSapItemGroup.ItemCount > 0 then
                  begin
                    aSapBomChild := aSapItemGroup.Items[0];
                           
                    if aSapBomChild.FNumber = '01.04.1330012000A0' then
                    begin
                      Sleep(1);
                    end;

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

                    aMRPUnitPtr_Dep.iAltCount := aSapItemGroup.ItemCount;
                    // ������ϣ�����ȷ�
                    aMRPUnitPtr_Dep.dQty := (aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock) * aSapBomChild.dusage * (100 - iPer100) / 100;


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
                    aMRPUnitPtr_Dep.spriority := aSapBomChild.spriority; // ���ȼ�
                    aMRPUnitPtr_Dep.dPer := (100 - iPer100);

                    aMRPUnitPtr_Dep.iSubstituteNo := iSubstituteNo;
                    aMRPUnitPtr_Dep.sNote := '�����2';
                    aMRPUnitPtr_Dep.bSupp := True;
                    aMRPUnitPtr_Dep.bCalc := aMRPUnitPtr_Dep.bSupp;

                    lstDemand.Add(aMRPUnitPtr_Dep);                  
                  end;

                  
                  if slPerErr.IndexOf(aMRPUnitPtr.snumber) < 0 then
                  begin
                    for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                    begin
                      aSapBomChild := aSapItemGroup.Items[iChildItem];
                      Memo1.Lines.Add('����ܺͲ�Ϊ 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr.snumber);
                      sline := aMRPUnitPtr.snumber + #9 + aMRPUnitPtr.sname + #9 + aSapBomChild.FNumber + #9 + aSapBomChild.FName + #9 + FloatToStr(aSapBomChild.dPer) + #9'����ܺͲ�Ϊ 100';
                      slPerErr.Add(sline);
                    end;
                  end;
//                    Memo1.Lines.Add('����ܺͲ�Ϊ 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr.snumber);
//                    sline := aSapBomChild.FNumber + #9 + aMRPUnitPtr.snumber + #9'����ܺͲ�Ϊ 100';
//                    slPerErr.Add(sline);
                end;
                            
                iSubstituteNo := iSubstituteNo + 1; //  ������� + 1��ȷ��Ψһz
              end;
                       

                
            end;                       // ֮ǰ���� ����ʱ���Ѿ�������һ����Qty������Ҫ��          
            bLoop := True;  //  չ�����µ����������ѭ��
 

            lstSubstituteDemand.Free;
          end;
        end
        else  //// �⹺������չ��BOM  ///////  PD  ////////////////////////////////
        begin
          if aMRPUnitPtr.bCalc then
          begin
            Continue;  // �Ѽ���
          end;

          // �������� �����
          if aMRPUnitPtr.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr.snumber, aMRPUnitPtr.dQty, aMRPUnitPtr.sMrpArea);
            aMRPUnitPtr.bCalc := True;
          end
          else
          begin
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            dPerSum := 0;
            for idx := 0 to lstDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep.bSupp then Continue;
              if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
              begin                                          
                dPerSum := dPerSum + aMRPUnitPtr_Dep.dPer;
                dQty := dQty + aMRPUnitPtr_Dep.dQty;
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // ������������ȼ� 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr_Dep.snumber, dQty, aMRPUnitPtr_Dep.sMrpArea);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep.dQtyStock;
              
              aMRPUnitPtr_Dep.bCalc := True;
            end;

            // ����û��ȫ���㣬 ��������������������
            if dPerSum > 0 then    // dQty > 0�� ��Ȼ dQty_old > 0
            begin
              dQty_left := dQty;
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
                //aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty * aMRPUnitPtr_Dep.dPer / 100;

                if idx = lstSubstituteDemand.Count - 1 then // ���һ��
                begin    
                  aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty_left;
                end
                else
                begin       
                  dQty_delta := dQty * (aMRPUnitPtr_Dep.dPer / dPerSum);
                  aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty_delta;
                  dQty_left := dQty_left - dQty_delta;
                end; 

              end;                       // ֮ǰ���� ����ʱ���Ѿ�������һ����Qty������Ҫ��
            end; 

            lstSubstituteDemand.Free;
          end;
        end;        

      end;

      iLowestCode := iLowestCode + 1; 
    end;
               
    //���򣬰�����
    lstDemand.Sort(ListSortCompare_DateTime);
 

     
    ////  ���� PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
      aMRPUnitPtr.bCalc := False;
    end;

    ////  ���� PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);

   
      if aMRPUnitPtr.snumber = '81.03.83100195H0' then
      begin
        Sleep(1);
      end;
      
      // ���⹺�������� PO
      if aMRPUnitPtr.aSAPMaterialRecordPtr.sPType <> 'F' then Continue;

      if aMRPUnitPtr.bCalc then Continue; //������� 

      slNumber := TStringList.Create;
      // ������ �����
      if aMRPUnitPtr.iSubstituteNo = 0 then
      begin
        aMRPUnitPtr.bCalc := True;
                            
        if DoubleE( aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock - aMRPUnitPtr.dQtyStock2,  0) then Continue; // û������

        slNumber.Add(aMRPUnitPtr.snumber);

        aMRPUnitPtr.dQtyOPO := aSAPMrpAreaStockReader.Alloc(slNumber, aMRPUnitPtr.dt,
          aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock - aMRPUnitPtr.dQtyStock2, aMRPUnitPtr.sMrpArea);

      end
      else  //  ��������  �����
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        dPerSum := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep.bSupp then Continue;
          if aMRPUnitPtr_Dep.iSubstituteNo = aMRPUnitPtr.iSubstituteNo then
          begin
            dPerSum := dPerSum + aMRPUnitPtr_Dep.dPer;
            dQty := dQty + aMRPUnitPtr_Dep.dQty - aMRPUnitPtr_Dep.dQtyStock - aMRPUnitPtr_Dep.dQtyStock2; // ��ȥ�ѷ�����
            aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock + aMRPUnitPtr_Dep.dQtyStock2;         // ����Ŀ���ǹ̶��ģ���Ҫ���
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
            slNumber.Add(aMRPUnitPtr_Dep.snumber);


            aMRPUnitPtr_Dep.bCalc := True;
          end;
        end;

        if DoubleE(dQty, 0) then
        begin
          Continue; // ����������������
        end;


        // ������������ȼ� 
        lstSubstituteDemand.Sort(ListSortCompare_priority);

        // �ȷ��佻����Ķ���
        lstPOLine := TList.Create;
        //  �ڲ��Ѿ��Ź���ģ�����Ҫ������
        aSAPMrpAreaStockReader.GetOPOs(slNumber, lstPOLine, aMRPUnitPtr.sMrpArea); // �ҵ���������ϵĿ��òɹ�����

        for idx := 0 to lstPOLine.Count - 1 do
        begin
          aSAPOPOLine := TSAPOPOLine(lstPOLine[idx]);
          aMRPUnitPtr_Dep := GetMRPUnit(lstSubstituteDemand, aSAPOPOLine.FNumber);

          // �����OPO �ۼ�
          aMRPUnitPtr_Dep.dQtyOPO := aMRPUnitPtr_Dep.dQtyOPO + aSAPOPOLine.Alloc(aMRPUnitPtr_Dep.dt, dQty, aMRPUnitPtr_Dep.sMrpArea);
          aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQtyStock + aMRPUnitPtr_Dep.dQtyStock2 + aMRPUnitPtr_Dep.dQtyOPO;

          if DoubleE( dQty, 0) then  // ����������
          begin
            Break;
          end;
        end;

        lstPOLine.Free;
         
        // ����û��ȫ���㣬 ��������������������
        if dPerSum > 0 then  // dQty > 0�� ��Ȼ  dQty_old > 0
        begin
          dQty_left := dQty;
          for idx := 0 to lstSubstituteDemand.Count - 1 do
          begin                                  
            aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
            if idx = lstSubstituteDemand.Count - 1 then
            begin
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty_left;
            end
            else
            begin
              dQty_delta := dQty * aMRPUnitPtr_Dep.dPer / dPerSum;
              aMRPUnitPtr_Dep.dQty := aMRPUnitPtr_Dep.dQty + dQty_delta;
              dQty_left := dQty_left - dQty_delta;
            end;
          end;
        end; 

        lstSubstituteDemand.Free;
      end;  


      slNumber.Free;
    end;

  finally

    aSAPBomReader.Free;

  end;

               
  //��������Ϻţ� ���ڣ� �Ϻ�
  lstDemand.Sort(ListSortCompare_ForSPQ);
  dQtyEx := 0;
  sGroupNumbers := '';
  for iMrpUnit := 0 to lstDemand.Count - 1 do
  begin        
    aMRPUnitPtr := TMRPUnit(lstDemand[iMrpUnit]);
    
    if (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'E') or
      (aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'X') then Continue;      // ���Ƽ�������
    if DoubleE( aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ, 0 ) then Continue;// SPQ Ϊ 0 �� 1 ������
    if DoubleE( aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ, 1 ) then Continue;

    dQtyNet := aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock -
      aMRPUnitPtr.dQtyStock2 - aMRPUnitPtr.dQtyOPO;
    if DoubleLE( dQtyNet , 0) then Continue;
    if (sGroupNumbers = '') or (sGroupNumbers <> aMRPUnitPtr.sGroupNumbers) then // ��һ�� ���� ��һ������� �� û����ģ����Լ�������� ��
    begin
      aMRPUnitPtr.dQtyNetSPQ := Round((dQtyNet / aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ) + 0.5) * aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ;
      dQtyEx := aMRPUnitPtr.dQtyNetSPQ - dQtyNet;
      sGroupNumbers := aMRPUnitPtr.sGroupNumbers;
    end
    else
    begin
      if DoubleLE( dQtyNet , dQtyEx ) then // ����������
      begin
        dQtyEx := dQtyEx - dQtyNet;
        aMRPUnitPtr.dQtyNetSPQ := 0;
      end
      else
      begin
        aMRPUnitPtr.dQtyNetSPQ := Round(((dQtyNet - dQtyEx) / aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ) + 0.5) * aMRPUnitPtr.aSAPMaterialRecordPtr^.dSPQ;
        dQtyEx := dQtyEx + aMRPUnitPtr.dQtyNetSPQ - dQtyNet;  // aMRPUnitPtr.dQtyNetSPQ - dQtyNet  �����ɸ�
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
      Memo1.Lines.Add('�ֿ� ' + aSAPStockRecordPtr^.sstock + ' û�ж�ӦMRP����');
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
    // ���� //////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('��ʼ����ģ��MRP������');

    Memo1.Lines.Add('�����������');

    // ��ʼ���� Excel
    try
      ExcelApp := CreateOleObject('Excel.Application' );
      ExcelApp.Visible := False;
      ExcelApp.Caption := 'Ӧ�ó������ Microsoft Excel';
    except
      on e: Exception do
      begin
        MessageBox(Handle, PChar(e.Message), '�����ʾ', 0);
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
    ExcelApp.Cells[irow, 1].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 2].Value := '��Ʒ����';
    ExcelApp.Cells[irow, 3].Value := '����';
    ExcelApp.Cells[irow, 4].Value := '����';    
    ExcelApp.Cells[irow, 5].Value := '��������';     
    ExcelApp.Cells[irow, 6].Value := 'MRP Area';
    ExcelApp.Cells[irow, 7].Value := '������Ŀ';

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
//    ExcelApp.Cells[irow, 2].Value := '��ID';
//    ExcelApp.Cells[irow, 3].Value := '����';
//    ExcelApp.Cells[irow, 4].Value := '��������';
//    ExcelApp.Cells[irow, 5].Value := '��������';
//    ExcelApp.Cells[irow, 6].Value := '�����µ�����';
//    ExcelApp.Cells[irow, 7].Value := '��������';
//    ExcelApp.Cells[irow, 8].Value := '���ÿ��';
//    ExcelApp.Cells[irow, 9].Value := 'OPO';
//    ExcelApp.Cells[irow, 10].Value := '������';
//    ExcelApp.Cells[irow, 11].Value := '�����';
//    ExcelApp.Cells[irow, 12].Value := 'MRP������';
//    ExcelApp.Cells[irow, 13].Value := '�ɹ�Ա';
//    ExcelApp.Cells[irow, 14].Value := 'MRP����';  
//    ExcelApp.Cells[irow, 15].Value := '�ϲ��Ϻ�';
//    ExcelApp.Cells[irow, 16].Value := '���Ϻ�';
//    ExcelApp.Cells[irow, 17].Value := 'L/T';

    lst := TList.Create;
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      lst.Add(lstDemand[iMrpUnit]);
    end;

    lst.Sort(ListSortCompare_pid);

    sl := TStringList.Create;
    try        //1      2       3         4            5            6             7           8         9     10          11        12          13         14         15         16       17           18     19     20      21          22
      sline := 'ID'#9'��ID'#9'����'#9'��������'#9'��������'#9'�����µ�����'#9'��������'#9'���ÿ��'#9'OPO'#9'������'#9'SPQ������'#9'�����'#9'MRP������'#9'�ɹ�Ա'#9'MRP����'#9'�ϲ��Ϻ�'#9'���Ϻ�'#9'L/T'#9'SPQ'#9'MOQ'#9'���������'#9'��ע';
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
            if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType = 'F' then  // �⹺  ////////////////////////////////////////
            begin
              sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_PD) + #9;    //6
            end
            else                                                         // ����  ////////////////////////////////////////
            begin
              sline := sline +  FormatDateTime('yyyy-MM-dd', aMRPUnitPtr.dt - aMRPUnitPtr.aSAPMaterialRecordPtr.dLT_M0) + #9;    //6
            end;

            sline := sline + Format('%0.0f', [aMRPUnitPtr.dqty])+#9+     //7
              Format('%0.0f', [aMRPUnitPtr.dqtystock + aMRPUnitPtr.dqtystock2])+#9+   //8
              Format('%0.0f', [aMRPUnitPtr.dQtyOPO])+#9;    //9
            sline := sline + '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9+  //10      
              Format('%0.0f', [aMRPUnitPtr.dQtyNetSPQ]) + #9 +   //11
              IntToStr(aMRPUnitPtr.iSubstituteNo)+#9+    //12
              aMRPUnitPtr.aSAPMaterialRecordPtr.sMRPerDesc + #9 +   //13
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
              sline := sline + aMRPUnitPtr.sNote + #9;   //22

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

    Memo1.Lines.Add('����  PR Sum');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'PR Sum';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '���';
    ExcelApp.Cells[irow, 2].Value := '�ɹ����뵥����';
    ExcelApp.Cells[irow, 3].Value := '̧ͷ�ı�';
    ExcelApp.Cells[irow, 4].Value := '���ϱ���';
    ExcelApp.Cells[irow, 5].Value := '�ɹ����뵥����';
    ExcelApp.Cells[irow, 6].Value := '��������';
    ExcelApp.Cells[irow, 7].Value := '����';
    ExcelApp.Cells[irow, 8].Value := '���ص�';
    ExcelApp.Cells[irow, 9].Value := '��Ŀ�ı�';      // ������Ŀ
    ExcelApp.Cells[irow, 10].Value := '�����µ�����';
    ExcelApp.Cells[irow, 11].Value := 'MRP������';
    ExcelApp.Cells[irow, 12].Value := '�ɹ�Ա';
    ExcelApp.Cells[irow, 13].Value := 'L/T';
    ExcelApp.Cells[irow, 14].Value := 'SPQ';
    ExcelApp.Cells[irow, 15].Value := 'MOQ';
    ExcelApp.Cells[irow, 16].Value := '�ɹ�����';   
    ExcelApp.Cells[irow, 17].Value := '�ɹ����뵥����SPQ';

    ExcelApp.Cells[irow, 19].Value := '��������';
    ExcelApp.Cells[irow, 20].Value := 'ABC';


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

      // �����⹺  ////////////////////////////////////////
      if aMRPUnitPtr.aSAPMaterialRecordPtr^.sPType <> 'F' then Continue;

      if aMRPUnitPtr.bSupp then Continue;

      // û����
      dqty_pr := aMRPUnitPtr.dQty - aMRPUnitPtr.dQtyStock -
        aMRPUnitPtr.dQtyStock2 - aMRPUnitPtr.dQtyOPO;
      if DoubleE(dqty_pr, 0) then Continue;

      if aMRPUnitPtr0 = nil then  // ��һ��MrpLog
      begin
        aMRPUnitPtr0 := aMRPUnitPtr;
        dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
        dqty_pr_sum := dqty_pr;
        dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;
      end
      else
      begin
        if (aMRPUnitPtr0.snumber <> aMRPUnitPtr.snumber) or
          (aMRPUnitPtr0.sMrpArea <> aMRPUnitPtr.sMrpArea) then // �Ϻű��ˣ�����һ��PR
        begin
          // �ж��ǲ���ͬһ�ܵ�

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
          ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPerDesc;
          ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
          ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
          ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
          ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
          ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;
          ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);


          ExcelApp.Cells[irow, 19].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sName;
          ExcelApp.Cells[irow, 20].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sAbc;

          

          aMRPUnitPtr0 := aMRPUnitPtr;
          dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
          dqty_pr_sum := dqty_pr;  
          dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;

          iLine := iLine + 1;   // �ֲ�ͬPR
          
          irow := irow + 1;
        end
        else if dtMon + 7 < aMRPUnitPtr.dt then // ��һ��
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
          ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPerDesc;
          ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
          ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
          ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
          ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
          ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;
          ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);
                                                  
          ExcelApp.Cells[irow, 19].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sName;
          ExcelApp.Cells[irow, 20].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr^.sAbc;

          
          aMRPUnitPtr0 := aMRPUnitPtr;
          dtMon := aMRPUnitPtr0.dt + 2 - DayOfWeek(aMRPUnitPtr0.dt);
          dqty_pr_sum := dqty_pr;        
          dqty_pr_sum_spq := aMRPUnitPtr0.dQtyNetSPQ;
            
          irow := irow + 1; 
        end
        else // ͬ�Ϻţ� ͬ���� ͬ�ܣ��ϲ�����
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
      ExcelApp.Cells[irow, 11].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sMRPerDesc;
      ExcelApp.Cells[irow, 12].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sBuyer;
      ExcelApp.Cells[irow, 13].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dLT_PD;
      ExcelApp.Cells[irow, 14].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dSPQ;
      ExcelApp.Cells[irow, 15].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.dMOQ;
      ExcelApp.Cells[irow, 16].Value := aMRPUnitPtr0.aSAPMaterialRecordPtr.sPType;   
      ExcelApp.Cells[irow, 17].Value := Format('%0.0f', [dqty_pr_sum_spq]);
    end;

    lst.Free;
 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('����  PO Action');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'PO Action';


    sl := TStringList.Create;
    try
      sline := '�ɹ�ƾ֤'#9'�к�'#9'����'#9'��������'#9'��������'#9'���鵽������'#9'��������'#9'����'#9'MRP Area'#9'Mrp Area No'#9'MC'#9'�ɹ�Ա'#9'LT'#9'SPQ'#9'MOQ'#9'ƾ֤����'#9'��Ӧ��';
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

          if DoubleE( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // ׼ʱ����
          begin
            sline := sline + #9 + 'OTD';
          end
          else if DoubleG( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // �������ڣ� �����������ڣ�Push In
          begin
            sline := sline + #9 + 'Push In';
          end
          else            // �����������ڣ�������������, Push Out
          begin
            sline := sline + #9 + 'Push Out';
          end;

          sline := sline + #9 + aSAPMrpAreaStockReader.MrpAreaNo2Name(aSAPOPOAllocPtr^.sMrpAreaNo);
          sline := sline + #9 + aSAPOPOAllocPtr^.sMrpAreaNo;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPerDesc;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_PD]); // �⹺���϶� �Ǽƻ�����ʱ��
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dSPQ]);
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dMOQ]);  
          sline := sline + #9 + FormatDatetime('yyyy-MM-dd', aSAPOPOLine.FBillDate);
          sline := sline + #9 + aSAPOPOLine.FSupplier;

          sl.Add(sline);
        end;


        // û�з����
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
          sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPerDesc;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;     
          sline := sline + #9 + Format('%0.0f', [aSAPMaterialRecordPtr^.dLT_PD]); // �⹺���϶� �Ǽƻ�����ʱ��
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

    Memo1.Lines.Add('���� E&o');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'E&O';


    sl := TStringList.Create;
    try
      sline := '����'#9'��������'#9'������'#9'�ܿ��'#9'�ܶ���'#9'������'#9'��������'#9'�ܴ���'#9'E/O'#9'17������'#9'28������'#9'60������'#9'MRP����'#9'MRP����'#9'MC'#9'�ɹ�Ա'#9'LT'#9'SPQ'#9'MOQ'#9'�ɹ�����'#9'MRP��';
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
        sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPerDesc;
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
    (*
    Memo1.Lines.Add('���');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'OH';


    ExcelApp2 := CreateOleObject('Excel.Application' );
    ExcelApp2.Visible := False;
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
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

          if stitle <> '�������ص�ִ��ص������������������������ʹ�õĿ��' then
          begin
            Memo1.Lines.Add(sSheet +'  ����  SAP�������  ��ʽ');
            Continue;
          end;

          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste;   
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;
 
          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
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
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
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

          if stitle <> '�ɹ�ƾ֤��Ŀ�ƻ��вɹ�ƾ֤���Ͳɹ�ƾ֤���ɹ���' then
          begin
            Memo1.Lines.Add(sSheet +'  ����  SAP����OPO  ��ʽ');
            Continue;
          end;

          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste;  
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;
 
          ExcelApp2.CutCopyMode := False;

          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
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
    ExcelApp2.Caption := 'Ӧ�ó������ Microsoft Excel';
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

          if stitle <> 'ĸ�����ϱ���ĸ����������������;' then
          begin         
            Memo1.Lines.Add(sSheet +'  ����SAP����BOM��ʽ');
            Continue;
          end;
          
          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste; 
          ExcelApp2.ActiveSheet.Cells[1,1].Copy;     

          ExcelApp2.CutCopyMode := False;


          break;
        end;
      finally
        ExcelApp2.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����
        WorkBook2.Close;
      end;

    finally
      ExcelApp2.Visible := True;
      ExcelApp2.Quit;
    end;

    *)

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('��Ȳ�Ϊ100');
    
    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'BOM���';
          

    Clipboard.SetTextBuf(PChar(slPerErr.Text));
    ExcelApp.ActiveSheet.Paste;

    slPerErr.Free;

    

    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //�¼ӵ�,�����Ѿ�����

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

  MessageBox(Handle, '���', '��ʾ', 0);
end;

procedure TfrmMRP4SAP3_MRP.btnWhereUseClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leWhereUse.Text := sfile;
end;

end.
