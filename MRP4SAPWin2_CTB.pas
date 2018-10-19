unit MRP4SAPWin2_CTB;
 
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
  SAPMaterialReader2,
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
                                                             
  TfrmMRP4SAP2_CTB = class(TForm)
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
    leSAPOPO: TLabeledEdit;
    btnSAPOPO: TButton;
    leMaterial: TLabeledEdit;
    btnMaterial: TButton;
    leSAPMrpAreaStock: TLabeledEdit;
    btnSAPMrpAreaStock: TButton;
    cbUseOtherFacSemiOH: TCheckBox;
    tbSave: TToolButton;
    ToolButton3: TToolButton;
    procedure FormCreate(Sender: TObject); 
    procedure FormDestroy(Sender: TObject);
    procedure tbCloseClick(Sender: TObject);
    procedure btnStockClick(Sender: TObject);
    procedure btnSAPStock3Click(Sender: TObject);
    procedure btnSAPBom3Click(Sender: TObject);
    procedure btnDemand3Click(Sender: TObject);
    procedure btnSAPOPOClick(Sender: TObject);
    procedure btnMaterialClick(Sender: TObject);
    procedure btnSAPMrpAreaStockClick(Sender: TObject);
    procedure tbSaveClick(Sender: TObject);
  private
    { Private declarations }  
    procedure OnLogEvent(const s: string);
    procedure Alloc2(dQtyUnalloc: Double; slGroup: TStringList;
      aSAPMrpAreaStockReader: TSAPMrpAreaStockReader; bUsage: Boolean);
    function GetStockPO(const snumber: string;
      aSAPMrpAreaStockReader: TSAPMrpAreaStockReader): Double;
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
  PNumberInfo = ^TNumberInfo;
    
  TOrderByConditions = packed record
    isnew: Boolean;
    dos: Double;
    demand: Double;
  end;
  POrderByConditions = ^TOrderByConditions;

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

class procedure TfrmMRP4SAP2_CTB.ShowForm;
var
  frmMRP4SAP2_CTB: TfrmMRP4SAP2_CTB;
begin
  frmMRP4SAP2_CTB := TfrmMRP4SAP2_CTB.Create(nil);
  try
    frmMRP4SAP2_CTB.ShowModal;
  finally
    frmMRP4SAP2_CTB.Free;
  end;
end;
   
procedure TfrmMRP4SAP2_CTB.FormCreate(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
 
  leSAPStock.Text := ini.ReadString(self.ClassName, leSAPStock.Name, '');
  leSAPBom.Text := ini.ReadString(self.ClassName, leSAPBom.Name, '');
  leSAPPIR.Text := ini.ReadString(self.ClassName, leSAPPIR.Name, '');
  leSAPOPO.Text := ini.ReadString(self.ClassName, leSAPOPO.Name, '');   
  leMaterial.Text := ini.ReadString(self.ClassName, leMaterial.Name, '');
  leSAPMrpAreaStock.Text := ini.ReadString(Self.ClassName, leSAPMrpAreaStock.Name, '');
  cbUseOtherFacSemiOH.Checked := ini.ReadBool(Self.ClassName, cbUseOtherFacSemiOH.Name, False);
 
  ini.Free;
end;

procedure TfrmMRP4SAP2_CTB.FormDestroy(Sender: TObject);
var
  sfile: string;
  ini: TIniFile;
begin
  sfile := AppIni;
  ini := TIniFile.Create(sfile);
 

  ini.WriteString(self.ClassName, leSAPStock.Name, leSAPStock.Text);
  ini.WriteString(self.ClassName, leSAPBom.Name, leSAPBom.Text);
  ini.WriteString(self.ClassName, leSAPPIR.Name, leSAPPIR.Text);  
  ini.WriteString(self.ClassName, leSAPOPO.Name, leSAPOPO.Text);
  ini.WriteString(self.ClassName, leMaterial.Name, leMaterial.Text);
  ini.WriteString(self.ClassName, leSAPMrpAreaStock.Name, leSAPMrpAreaStock.Text);
  ini.WriteBool(self.ClassName, cbUseOtherFacSemiOH.Name, cbUseOtherFacSemiOH.Checked);

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
 
procedure TfrmMRP4SAP2_CTB.tbCloseClick(Sender: TObject);
begin
  Close;
end;
 
procedure TfrmMRP4SAP2_CTB.btnStockClick(Sender: TObject);
begin
end;

  // �򻯵�MRP���㣬�����ǵ�λ��
  type
    TMRPUnit = packed record
      id: Integer;
      pid: Integer;
      snumber: string;
      sname: string;
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
      iSubstituteNo: Integer; // �������
      spriority: string;
      dPer: Double;
      sMrpArea: string;
      iAltCount: Integer;
    end;
    PMRPUnit = ^TMRPUnit;
    
    function ListSortCompare_Number_DateTime(Item1, Item2: Pointer): Integer;
    var
      p1, p2: PMRPUnit;
    begin
      p1 := PMRPUnit(Item1);
      p2 := PMRPUnit(Item2);

      if p1^.snumber > p2^.snumber then
      begin
        Result := 1;
      end
      else if p1^.snumber < p2^.snumber then
      begin
        Result := -1;
      end
      else
      begin
        if DoubleG(p1^.dt, p2^.dt) then
          Result := 1
        else if DoubleL(p1^.dt, p2^.dt) then
          Result := -1
        else Result := 0;
      end;
    end;
             
    function ListSortCompare_DateTime(Item1, Item2: Pointer): Integer;
    var
      p1, p2: PMRPUnit;
    begin
      p1 := PMRPUnit(Item1);
      p2 := PMRPUnit(Item2);
      
      if DoubleG(p1^.dt, p2^.dt) then
        Result := 1
      else if DoubleL(p1^.dt, p2^.dt) then
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

function GetDemand(lstDemand: TStringList; const snumber: string; dt1, dt2: TDateTime): Double;
var
  i: Integer;
  aMRPUnitPtr: PMRPUnit;
begin
  Result := 0;
  for i := 0 to lstDemand.Count - 1 do
  begin
    if lstDemand[i] <> snumber then Continue;
    aMRPUnitPtr := PMRPUnit(lstDemand.Objects[i]);
    if (aMRPUnitPtr^.dt >= dt1) and (aMRPUnitPtr^.dt < dt2) then
    begin
      Result := Result + aMRPUnitPtr^.dQty;
    end;
  end;
end;
    
procedure TfrmMRP4SAP2_CTB.btnSAPStock3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;
     
procedure TfrmMRP4SAP2_CTB.btnSAPOPOClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPOPO.Text := sfile;
end;
      
procedure TfrmMRP4SAP2_CTB.btnMaterialClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMaterial.Text := sfile;
end;

procedure TfrmMRP4SAP2_CTB.btnSAPBom3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmMRP4SAP2_CTB.btnDemand3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPPIR.Text := sfile;
end;
    
procedure TfrmMRP4SAP2_CTB.btnSAPMrpAreaStockClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPMrpAreaStock.Text := sfile;
end;

procedure TfrmMRP4SAP2_CTB.OnLogEvent(const s: string);
begin
  Memo1.Lines.Add(s);
end;

function ListSortCompare_priority(Item1, Item2: Pointer): Integer;
var
  aMRPUnitPtr1, aMRPUnitPtr2: PMRPUnit;
  iPriority1, iPriority2: Integer;
begin
  aMRPUnitPtr1 := Item1;
  aMRPUnitPtr2 := Item2;
  iPriority1 := StrToIntDef(aMRPUnitPtr1^.spriority, 1);
  iPriority2 := StrToIntDef(aMRPUnitPtr2^.spriority, 1);
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

procedure TfrmMRP4SAP2_CTB.Alloc2(dQtyUnalloc: Double; slGroup: TStringList;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader; bUsage: Boolean);
var
  dSum: Double;
  dStockPO: Double;
  i: Integer;
  aMRPUnitPtr: PMRPUnit;
begin
  if slGroup.Count = 0 then Exit;
  
  dSum := 0;
  for i := 0 to slGroup.Count - 1 do
  begin
    aMRPUnitPtr := PMRPUnit(slGroup.Objects[i]);
    dStockPO := GetStockPO(aMRPUnitPtr^.snumber, aSAPMrpAreaStockReader);
    slGroup[i] := IntToStr(Round(dStockPO));
    dSum := dSum + dStockPO;
  end;

  if dSum = 0 then
  begin
    aMRPUnitPtr := PMRPUnit(slGroup.Objects[0]);
    if bUsage then
    begin
      aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + dQtyUnalloc* aMRPUnitPtr^.aBom.dusage;
    end
    else
    begin
      aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + dQtyUnalloc {* aMRPUnitPtr^.aBom.dusage};
    end;
  end
  else
  begin
    for i := 0 to slGroup.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(slGroup.Objects[i]);
      if bUsage then
      begin
        aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + (dQtyUnalloc * StrToInt(slGroup[i]) / dSum) * aMRPUnitPtr^.aBom.dusage;
      end
      else
      begin
        aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + (dQtyUnalloc * StrToInt(slGroup[i]) / dSum) {* aMRPUnitPtr^.aBom.dusage};
      end;  
    end;
  end;
end;  

function TfrmMRP4SAP2_CTB.GetStockPO(const snumber: string;
  aSAPMrpAreaStockReader: TSAPMrpAreaStockReader): Double;
begin
  Result := aSAPMrpAreaStockReader.GetStockSum(snumber) +
    aSAPMrpAreaStockReader.GetOPOSum(snumber);
end;  

procedure TfrmMRP4SAP2_CTB.tbSaveClick(Sender: TObject);
  function GetMRPUnit(lst: TList; const snumber: string): PMRPUnit;
  var
    i: Integer;
  begin
    Result := nil;
    for i := 0 to lst.Count - 1 do
    begin
      if PMRPUnit( lst[i] )^.snumber = snumber then
      begin
        Result := PMRPUnit( lst[i] );
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
  lstDemand_tmp: TList;
  iLine: Integer;
  iWeek: Integer;
  aSOPProj: TSOPProj;
  iDate: Integer;
  aMRPUnitPtr: PMRPUnit;
  aMRPUnitPtr_Dep: PMRPUnit;
  iMrpUnit: Integer;

  iChild: Integer;
  iChildItem: Integer;
  aSapItemGroup: TSapItemGroup;

  aSapBomChild: TSapBom;
  
  ExcelApp, WorkBook: Variant;
  irow: Integer;
  bLoop: BOOL;
  dQty: Double;
  dQty0: Double;
  dQty_Alloc: Double;
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

  slTransfer: TList;
  aTransferRecordPtr: PTransferRecord;
  slGroup: TStringList;
  dQtyUnalloc: Double;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  slArea2BomFac := TStringList.Create;
  slArea2BomFac.Add('FIH01=FX');
  slArea2BomFac.Add('HQ001=HQ');
  slArea2BomFac.Add('ML001=ML');
  slArea2BomFac.Add('WT001=WT');
  slArea2BomFac.Add('YD001=YD');

  today := myStrToDateTime(FormatDateTime('yyyy-MM-dd', Now));
      
  Memo1.Lines.Add('��ʼ��ȡ PIR  ' + leSAPPIR.Text);
  aSAPS618Reader := TSAPPIRReader.Create(leSAPPIR.Text, OnLogEvent);
                                   
  Memo1.Lines.Add('��ʼ��ȡ MRP����ֿ��б�  ' + leSAPMrpAreaStock.Text);
  aSAPMrpAreaStockReader := TSAPMrpAreaStockReader.Create(leSAPMrpAreaStock.Text);
  
  Memo1.Lines.Add('��ʼ��ȡ BOM  ' + leSAPBom.Text);
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ ���  ' + leSAPStock.Text);
  aSAPStockReader := TSAPStockReader2.Create(leSAPStock.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ OPO  ' + leSAPOPO.Text);
  aSAPOPOReader2 := TSAPOPOReader2.Create(leSAPOPO.Text, OnLogEvent);

  Memo1.Lines.Add('��ʼ��ȡ ����  ' + leMaterial.Text);
  aSAPMaterialReader:= TSAPMaterialReader2.Create(leMaterial.Text, OnLogEvent);

  aSAPMrpAreaStockReader.SetOPOList(aSAPOPOReader2);                       
  aSAPMrpAreaStockReader.SetStock(aSAPStockReader);

  //aSAPStockSum := TSAPStockSum.Create;
  //aSAPStockReader.SumTo(aSAPStockSum);

//  lstMrpDetail := TList.Create;

  lstDemand := TList.Create;
  slGroup := TStringList.Create;


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
      
      aMRPUnitPtr := New(PMRPUnit);
      aMRPUnitPtr^.id := iid;   
      iid := iid + 1;
      aMRPUnitPtr^.pid := 0;
      aMRPUnitPtr^.snumber := aSAPS618ColPtr^.sNumber;
      aMRPUnitPtr^.sname := aSAPS618ColPtr^.sname;
      aMRPUnitPtr^.dt := aSAPS618ColPtr^.dt1;
      aMRPUnitPtr^.dQty := aSAPS618ColPtr^.dQty;
      aMRPUnitPtr^.dQtyStock := 0;
      aMRPUnitPtr^.dQtyStock2 := 0;
//      aMRPUnitPtr^.dQtyStockParent := 0;
      aMRPUnitPtr^.dQtyOPO := 0;
      aMRPUnitPtr^.bExpend := False;
      aMRPUnitPtr^.bCalc := False;
      aMRPUnitPtr^.aBom := nil;
      aMRPUnitPtr^.aParentBom := nil;
      aMRPUnitPtr^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSAPS618ColPtr^.sNumber);
      aMRPUnitPtr^.sDemandType := aSAPS618.sDemandType;
      aMRPUnitPtr^.iSubstituteNo := 0;
      aMRPUnitPtr^.spriority := '';
      aMRPUnitPtr^.sMrpArea := aSAPS618.FMrpArea;
      aMRPUnitPtr^.iAltCount := 1;
      lstDemand.Add(aMRPUnitPtr);

      if aMRPUnitPtr^.snumber = '03.42.34212415-P' then
      begin
        Sleep(1);
      end;
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
        aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);


        if iMrpUnit = 565 then
        begin
          Sleep(1);
        end;
        if aMRPUnitPtr^.snumber = '82.06.82600007' then
        begin
          Sleep(1);
        end;
 
        if aMRPUnitPtr^.bCalc then Continue; // ������ģ�����

        // ��λ��С�ڵ��ڵ�ǰ�����λ�룬�ż���
        if aMRPUnitPtr^.aSAPMaterialRecordPtr^.iLowestCode > iLowestCode then
        begin
          bLoop := True;  // ������û���㣬�����ѭ��
          Continue;
        end;


        ////  ���Ƽ� չ��BOM  //////////////////////////////////////////////
        if (aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType = 'E')
          or (aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType = 'X') then
        begin
          // �������� �����
          if aMRPUnitPtr^.iSubstituteNo = 0 then
          begin        
            aMRPUnitPtr^.bCalc := True;
            aMRPUnitPtr^.bExpend := True;

            if aMRPUnitPtr^.aParentBom = nil then // ���ڵ㣬�����BOM
            begin
              aMRPUnitPtr^.aBom := aSAPBomReader.GetSapBom(aMRPUnitPtr^.snumber, slArea2BomFac.Values[aMRPUnitPtr^.sMrpArea]);
            end;

            if aMRPUnitPtr^.sDemandType <> 'BSF' then  //  LSF���ǿ�棬 BSF �����ǿ��
            begin
              if cbUseOtherFacSemiOH.Checked then
              begin
                aMRPUnitPtr^.dQtyStock := aSAPMrpAreaStockReader.AllocStock(
                  aMRPUnitPtr^.snumber, aMRPUnitPtr^.dQty, aMRPUnitPtr^.sMrpArea);
              end
              else
              begin
                aMRPUnitPtr^.dQtyStock := aSAPMrpAreaStockReader.AllocStock2(
                  aMRPUnitPtr^.snumber, aMRPUnitPtr^.dQty, aMRPUnitPtr^.sMrpArea);
              end;
            end;

            // ����������������ˣ���������������չ��
            if DoubleLE( aMRPUnitPtr^.dQty, aMRPUnitPtr^.dQtyStock ) then
            begin
              Continue;
            end;
          
            if aMRPUnitPtr^.aBom = nil then  // �����󣬵��� û��BOM���쳣����¼��־
            begin
              Memo1.Lines.Add(aMRPUnitPtr^.snumber + ' ��BOM'); 
              Continue;
            end;

            aSapBomChild := nil;

            //չ�������²�
            for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
              slGroup.Clear;
              iPer100 := 0;  
              for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
              begin
                aSapBomChild := aSapItemGroup.Items[iChildItem];
 
                aMRPUnitPtr_Dep := New(PMRPUnit);
                aMRPUnitPtr_Dep^.id := iid;
                iid := iid + 1;

                aMRPUnitPtr_Dep^.sMrpArea := aMRPUnitPtr^.sMrpArea;
                aMRPUnitPtr_Dep^.pid := aMRPUnitPtr^.id;
                aMRPUnitPtr_Dep^.snumber := aSapBomChild.FNumber;
                aMRPUnitPtr_Dep^.sname := aSapBomChild.FName;
                aMRPUnitPtr_Dep^.dt := aMRPUnitPtr^.dt - aMRPUnitPtr^.aBom.lt;
                if aSapBomChild.sgroup = '' then
                begin                                                  
                  aMRPUnitPtr_Dep^.iAltCount := 1;
                  aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage;
                  iPer100 := 100;
                end
                else
                begin                     
                  aMRPUnitPtr_Dep^.iAltCount := aSapItemGroup.ItemCount;
                  // ������ϣ�����ȷ�
                  aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                  iPer100 := iPer100 + Round(aSapBomChild.dPer);
                end;

                aMRPUnitPtr_Dep^.dQtyStock := 0;
                aMRPUnitPtr_Dep^.dQtyStock2 := 0;
                aMRPUnitPtr_Dep^.dQtyOPO := 0;
                aMRPUnitPtr_Dep^.bExpend := False;
                aMRPUnitPtr_Dep^.bCalc := False;
                aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                if (aSapBomChild.sptype = 'E') or (aSapBomChild.sptype = 'X') then
                begin
                  aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr^.sMRPType := 'M0';
                end
                else
                begin
                  aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr^.sMRPType := 'PD';
                end;
                aMRPUnitPtr_Dep^.spriority := aSapBomChild.spriority; // ���ȼ�
                aMRPUnitPtr_Dep^.dPer := aSapBomChild.dPer;

                if aSapItemGroup.ItemCount = 1 then
                begin
                  aMRPUnitPtr_Dep^.iSubstituteNo := 0; // û�������
                end
                else
                begin
                  aMRPUnitPtr_Dep^.iSubstituteNo := iSubstituteNo;
                end;
                lstDemand.Add(aMRPUnitPtr_Dep);

                slGroup.AddObject(IntToStr(Round(aSapBomChild.dPer)), TObject( aMRPUnitPtr_Dep ));
                
              end;
              iSubstituteNo := iSubstituteNo + 1; //  ������� + 1��ȷ��Ψһ
                
                if (aSapBomChild.FNumber = '01.03.1310035400A0') or (aSapBomChild.FNumber = '01.03.1310035400B0') then
                begin
                  Sleep(1);
                end;

              // ����ܺͲ�Ϊ 100
              if iPer100 <> 100 then
              begin
                Memo1.Lines.Add('����ܺͲ�Ϊ 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr^.snumber);

                dQtyUnalloc := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * (100 - iPer100) / 100; // δ��������
                Alloc2(dQtyUnalloc, slGroup, aSAPMrpAreaStockReader, True);
                
              end;

            end;
            bLoop := True;  //  չ�����µ����������ѭ��
          end
          else //  ������� // ���Ʒ�������  /////////////////////////////////
          begin 
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            for idx := 0 to lstDemand.Count - 1 do //  ���������������ߵ���
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
              begin
                dQty := dQty + aMRPUnitPtr_Dep^.dQty;  // ��������ϵ�����
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // ������������ȼ� 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              if cbUseOtherFacSemiOH.Checked then
              begin
                aMRPUnitPtr_Dep^.dQtyStock := aSAPMrpAreaStockReader.AllocStock(aMRPUnitPtr_Dep^.snumber, dQty, aMRPUnitPtr_Dep^.sMrpArea);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������
              end
              else
              begin
                aMRPUnitPtr_Dep^.dQtyStock := aSAPMrpAreaStockReader.AllocStock2(aMRPUnitPtr_Dep^.snumber, dQty, aMRPUnitPtr_Dep^.sMrpArea);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������              
              end;
              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep^.dQtyStock;
              
              aMRPUnitPtr_Dep^.bCalc := True;
            end;

            // ����û��ȫ���㣬 ��������������������
            if dQty > 0 then
            begin
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr := lstSubstituteDemand[idx];  
                aMRPUnitPtr^.bCalc := True;
                aMRPUnitPtr^.bExpend := True;
                aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + dQty * aMRPUnitPtr^.dPer / 100;
                if DoubleE( aMRPUnitPtr^.dQty, 0) then Continue; 

                // ���������������չ����
                for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
                begin
                  aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
                  slGroup.Clear;
                  iPer100 := 0;
                  aSapBomChild := nil;
                  //չ�������²�
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem];


                    aMRPUnitPtr_Dep := New(PMRPUnit);
                    aMRPUnitPtr_Dep^.id := iid;
                    iid := iid + 1;

                    aMRPUnitPtr_Dep^.sMrpArea := aMRPUnitPtr^.sMrpArea;
                    aMRPUnitPtr_Dep^.pid := aMRPUnitPtr^.id;
                    aMRPUnitPtr_Dep^.snumber := aSapBomChild.FNumber;
                    aMRPUnitPtr_Dep^.sname := aSapBomChild.FName;
                    aMRPUnitPtr_Dep^.dt := aMRPUnitPtr^.dt - aMRPUnitPtr^.aBom.lt;
                    if aSapBomChild.sgroup = '' then
                    begin
                      aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage;
                      iPer100 := 100;
                    end
                    else
                    begin
                      // ������ϣ�����ȷ�
                      aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                      iPer100 := iPer100 + Round(aSapBomChild.dPer);
                    end;

                    aMRPUnitPtr_Dep^.dQtyStock := 0; 
                    aMRPUnitPtr_Dep^.dQtyStock2 := 0;
                    aMRPUnitPtr_Dep^.dQtyOPO := 0;
                    aMRPUnitPtr_Dep^.bExpend := False;
                    aMRPUnitPtr_Dep^.bCalc := False;
                    aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                    aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                    aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);     
                    if (aSapBomChild.sptype = 'E') or (aSapBomChild.sptype = 'X') then
                    begin
                      aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr^.sMRPType := 'M0';
                    end
                    else
                    begin
                      aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr^.sMRPType := 'PD';
                    end;
                    aMRPUnitPtr_Dep^.spriority := aSapBomChild.spriority; // ���ȼ�
                    aMRPUnitPtr_Dep^.dPer := aSapBomChild.dPer;

                    if aSapItemGroup.ItemCount = 1 then
                    begin
                      aMRPUnitPtr_Dep^.iSubstituteNo := 0; // û�������
                    end
                    else
                    begin
                      aMRPUnitPtr_Dep^.iSubstituteNo := iSubstituteNo;
                    end;
                    lstDemand.Add(aMRPUnitPtr_Dep);

                    slGroup.AddObject(IntToStr(Round(aSapBomChild.dPer)), TObject( aMRPUnitPtr_Dep ));
                  end;
                  iSubstituteNo := iSubstituteNo + 1; //  ������� + 1��ȷ��Ψһ
                        
                    if (aSapBomChild.FNumber = '01.03.1310035400A0') or (aSapBomChild.FNumber = '01.03.1310035400B0') then
                    begin
                      Sleep(1);
                    end;

                  // ����ܺͲ�Ϊ 100
                  if iPer100 <> 100 then
                  begin
                    Memo1.Lines.Add('����ܺͲ�Ϊ 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr^.snumber);

                    dQtyUnalloc := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * (100 - iPer100) / 100; // δ��������
                    Alloc2(dQtyUnalloc, slGroup, aSAPMrpAreaStockReader, True);
                                    
                  end;
                  
                end;
                       

                
              end;                       // ֮ǰ���� ����ʱ���Ѿ�������һ����Qty������Ҫ��          
              bLoop := True;  //  չ�����µ����������ѭ��
            end;


 

            lstSubstituteDemand.Free;


          end;
        end
        else  //// �⹺������չ��BOM  ///////  PD  ////////////////////////////////
        begin
          if aMRPUnitPtr^.bCalc then
          begin
            Continue;  // �Ѽ���
          end;



          if (aMRPUnitPtr^.snumber = '01.03.1310035400A0') or
            (aMRPUnitPtr^.snumber = '01.03.1310035400B0') then
          begin
            Sleep(1);
          end;
          
          // �������� �����
          if aMRPUnitPtr^.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr^.dQtyStock := aSAPMrpAreaStockReader.AllocStock2(aMRPUnitPtr^.snumber, aMRPUnitPtr^.dQty, aMRPUnitPtr^.sMrpArea);
            aMRPUnitPtr^.bCalc := True;
          end
          else
          begin
            lstSubstituteDemand := TList.Create;
            dQty := 0;
            for idx := 0 to lstDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstDemand[idx];
              if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
              begin
                dQty := dQty + aMRPUnitPtr_Dep^.dQty;
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // ������������ȼ� 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep^.dQtyStock := aSAPMrpAreaStockReader.AllocStock2(aMRPUnitPtr_Dep^.snumber, dQty, aMRPUnitPtr_Dep^.sMrpArea);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������
              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep^.dQtyStock;
              
              aMRPUnitPtr_Dep^.bCalc := True;
            end;

            // ����û��ȫ���㣬 ��������������������
            if dQty > 0 then
            begin
              dQty0 := dQty;
              iPer100 := 0;
              slGroup.Clear;
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin
                aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
                dQty_Alloc := dQty0 * aMRPUnitPtr_Dep^.dPer / 100;
                aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty_Alloc;
                dQty := dQty - dQty_Alloc;
                iPer100 := iPer100 + Round(aMRPUnitPtr_Dep^.dPer);

                slGroup.AddObject(IntToStr(Round(aMRPUnitPtr_Dep^.dPer)), TObject( aMRPUnitPtr_Dep ));
              end;                       // ֮ǰ���� ����ʱ���Ѿ�������һ����Qty������Ҫ��

              if lstSubstituteDemand.Count > 0 then
              begin
//                aMRPUnitPtr_Dep := lstSubstituteDemand[0];
//                aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty;

                // ����ܺͲ�Ϊ 100
                if iPer100 <> 100 then
                begin 
                  dQtyUnalloc := dQty; // δ��������
                  Alloc2(dQtyUnalloc, slGroup, aSAPMrpAreaStockReader, False);
                end; 
              end;
            end;

            lstSubstituteDemand.Free;
          end;
        end;

      end;

      iLowestCode := iLowestCode + 1;
    end;

    //���򣬰�����
    lstDemand.Sort(ListSortCompare_DateTime);


    //�����Ĵ��������////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    slTransfer := TList.Create;

    lstDemand_Count := lstDemand.Count;
    for iMrpUnit := 0 to lstDemand_Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);    
      aMRPUnitPtr^.bCalc := False;
      aSAPMrpAreaStockReader.AccDemand(aMRPUnitPtr^.snumber, dQty, aMRPUnitPtr^.sMrpArea); 
    end;

    lstDemand_Count := lstDemand.Count;
    for iMrpUnit := 0 to lstDemand_Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      if aMRPUnitPtr^.bCalc then Continue; // ����Ͽ����Ѿ������

      if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType <> 'F' then
      begin            
        aMRPUnitPtr^.bCalc := True;
        Continue;
      end;
 
      // �������� �����
      if aMRPUnitPtr^.iSubstituteNo = 0 then
      begin                     
        dQty := aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock; // ��ȥ�ѷ�����Ŀ��
        if DoubleG(dQty, 0) then
        begin
          aMRPUnitPtr^.dQtyStock2 := aSAPMrpAreaStockReader.AllocStock_area(
            aMRPUnitPtr^.snumber,aMRPUnitPtr^.sname, aMRPUnitPtr^.dt, dQty,
            aMRPUnitPtr^.sMrpArea, slTransfer);
        end;
        aMRPUnitPtr^.bCalc := True;
      end
      else
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
          begin
            dQty := dQty + aMRPUnitPtr_Dep^.dQty - aMRPUnitPtr_Dep^.dQtyStock; // ��ȥ�ѷ�����Ŀ��
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
          end;
        end;
          
        if DoubleG(dQty, 0) then
        begin
          // ������������ȼ�
          lstSubstituteDemand.Sort(ListSortCompare_priority);

          for idx := 0 to lstSubstituteDemand.Count - 1 do
          begin       
            aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
            aMRPUnitPtr_Dep^.dQtyStock2 := aSAPMrpAreaStockReader.AllocStock_area(
              aMRPUnitPtr_Dep^.snumber, aMRPUnitPtr_Dep^.sname, aMRPUnitPtr_Dep^.dt,
              dQty, aMRPUnitPtr_Dep^.sMrpArea, slTransfer);   // ���ĳ������Ͽ��ȫ�����ˣ�ʣ������Ϊ0�� �Զ���Ϊ���������Ϸ���0�Ŀ�������
            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock + aMRPUnitPtr_Dep^.dQtyStock2;  
            dQty := dQty  - aMRPUnitPtr_Dep^.dQtyStock2;      
            aMRPUnitPtr_Dep^.bCalc := True;
          end;

          // ����û��ȫ���㣬 ��������������������
          if dQty > 0 then
          begin
            iPer100 := 0;
            slGroup.Clear;
            dQty0 := dQty;
            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              dQty_Alloc := dQty0 * aMRPUnitPtr_Dep^.dPer / 100;
              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty_Alloc;
              dQty := dQty - dQty_Alloc; 
              iPer100 := iPer100 + Round(aMRPUnitPtr_Dep^.dPer);                      

              slGroup.AddObject(IntToStr(Round(aMRPUnitPtr_Dep^.dPer)), TObject( aMRPUnitPtr_Dep ));
            end;                       // ֮ǰ���� ����ʱ���Ѿ�������һ����Qty������Ҫ��

            if lstSubstituteDemand.Count > 0 then
            begin
//              iPer100 := 100 - iPer100;
//              aMRPUnitPtr_Dep := lstSubstituteDemand[0];
//              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty * iPer100 / 100;

              // ����ܺͲ�Ϊ 100
              if iPer100 <> 100 then
              begin 
                dQtyUnalloc := dQty; // δ��������
                Alloc2(dQtyUnalloc, slGroup, aSAPMrpAreaStockReader, False);
              end;               
            end;
          end; 
        end;
        lstSubstituteDemand.Free;
      end;      
    end;

    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////        




     
    ////  ���� PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      aMRPUnitPtr^.bCalc := False;
    end;

    ////  ���� PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);

   
              
      // ���⹺�������� PO 
      if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType <> 'F' then Continue;

      if aMRPUnitPtr^.bCalc then Continue; //������� 

      slNumber := TStringList.Create;
      // ������ �����
      if aMRPUnitPtr^.iSubstituteNo = 0 then
      begin
        aMRPUnitPtr^.bCalc := True;
                            
        if DoubleE( aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock - aMRPUnitPtr^.dQtyStock2,  0) then Continue; // û������

        slNumber.Add(aMRPUnitPtr^.snumber);

        aMRPUnitPtr^.dQtyOPO := aSAPMrpAreaStockReader.Alloc(slNumber, aMRPUnitPtr^.dt,
          aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock - aMRPUnitPtr^.dQtyStock2, aMRPUnitPtr^.sMrpArea);

      end
      else  //  ��������  �����
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
          begin
            dQty := dQty + aMRPUnitPtr_Dep^.dQty - aMRPUnitPtr_Dep^.dQtyStock - aMRPUnitPtr_Dep^.dQtyStock2; // ��ȥ�ѷ�����  
            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock + aMRPUnitPtr_Dep^.dQtyStock2;         // ����Ŀ���ǹ̶��ģ���Ҫ���
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
            slNumber.Add(aMRPUnitPtr_Dep^.snumber);


            aMRPUnitPtr_Dep^.bCalc := True;
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
        aSAPMrpAreaStockReader.GetOPOs(slNumber, lstPOLine, aMRPUnitPtr^.sMrpArea); // �ҵ���������ϵĿ��òɹ�����

        for idx := 0 to lstPOLine.Count - 1 do
        begin
          aSAPOPOLine := TSAPOPOLine(lstPOLine[idx]);
          aMRPUnitPtr_Dep := GetMRPUnit(lstSubstituteDemand, aSAPOPOLine.FNumber);

          // �����OPO �ۼ�
          aMRPUnitPtr_Dep^.dQtyOPO := aMRPUnitPtr_Dep^.dQtyOPO + aSAPOPOLine.Alloc(aMRPUnitPtr_Dep^.dt, dQty, aMRPUnitPtr_Dep^.sMrpArea);
          aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock + aMRPUnitPtr_Dep^.dQtyStock2 + aMRPUnitPtr_Dep^.dQtyOPO;

          if DoubleE( dQty, 0) then  // ����������
          begin
            Break;
          end;
        end;

        lstPOLine.Free;
         
        // ����û��ȫ���㣬 ��������������������
        if dQty > 0 then
        begin
          iPer100 := 0;
          slGroup.Clear;
          dQty0 := dQty;
          for idx := 0 to lstSubstituteDemand.Count - 1 do
          begin 
            aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
            dQty_Alloc := dQty0 * aMRPUnitPtr_Dep^.dPer / 100;
            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty_Alloc;
            dQty := dQty - dQty_Alloc; 
            iPer100 := iPer100 + Round(aMRPUnitPtr_Dep^.dPer);   

            slGroup.AddObject(IntToStr(Round(aMRPUnitPtr_Dep^.dPer)), TObject( aMRPUnitPtr_Dep ));               
          end;

            
          if lstSubstituteDemand.Count > 0 then
          begin
//            iPer100 := 100 - iPer100;
//            aMRPUnitPtr_Dep := lstSubstituteDemand[0];
//            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty * iPer100 / 100;

            // ����ܺͲ�Ϊ 100
            if iPer100 <> 100 then
            begin
              dQtyUnalloc := dQty; // δ��������
              Alloc2(dQtyUnalloc, slGroup, aSAPMrpAreaStockReader, False);
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


    sl := TStringList.Create;
    try
      sline := 'ID'#9'��ID'#9'����'#9'��������'#9'��������'#9'�����µ�����'#9'��������'#9'���ÿ��'#9'���ÿ��2'#9'OPO'#9'������'#9'�����'#9'MRP������'#9'�ɹ�Ա'#9'MRP����';
      sl.Add(sline);

      irow := 2;
      for iMrpUnit := 0 to lstDemand.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);

        sline := IntToStr(aMRPUnitPtr^.id) + #9 +
          IntToStr(aMRPUnitPtr^.pid) + #9 +
          aMRPUnitPtr^.snumber + #9 +
          aMRPUnitPtr^.sname + #9 +
          FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt) + #9;

        if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType = 'F' then  // �⹺  ////////////////////////////////////////
        begin
          sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_PD) + #9;
        end
        else                                                         // ����  ////////////////////////////////////////
        begin
          sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_M0) + #9;
        end;  
          
        sline := sline + Format('%0.0f', [aMRPUnitPtr^.dqty]) + #9 +
          Format('%0.0f', [aMRPUnitPtr^.dqtystock]) + #9 +
          Format('%0.0f', [aMRPUnitPtr^.dqtystock2]) + #9 +
          Format('%0.0f', [aMRPUnitPtr^.dQtyOPO]) + #9 +
          '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + '-' + GetRef(10) + IntToStr(irow) + #9 +  // 9 = 6 - 7 - 8
          IntToStr(aMRPUnitPtr^.iSubstituteNo) + #9 +
          aMRPUnitPtr^.aSAPMaterialRecordPtr^.sMRPer + #9 +
          aMRPUnitPtr^.aSAPMaterialRecordPtr^.sBuyer + #9 +
          aMRPUnitPtr^.sMrpArea;

        irow := irow + 1;
        sl.Add(sline);
                         
        smrparea := aMRPUnitPtr^.sMrpArea;

        idx := slNumber.IndexOf(aMRPUnitPtr^.snumber);
        if idx < 0 then
        begin
          aEORecordPtr := New(PEORecord);
          aEORecordPtr^.snumber := aMRPUnitPtr^.snumber;  
          aEORecordPtr^.sname := aMRPUnitPtr^.sname;
          aEORecordPtr^.sMrpAreaNo := smrparea;
          aEORecordPtr^.sMrpAreaName := aSAPMrpAreaStockReader.MrpAreaNo2Name(smrparea);
          aEORecordPtr^.dQtyDemand := 0;      
          aEORecordPtr^.dQtyDemand17 := 0;
          aEORecordPtr^.dQtyDemand28 := 0;
          aEORecordPtr^.dQtyDemand60 := 0;
          aEORecordPtr^.dQtyStock := 0;
          aEORecordPtr^.dQtyOPO := 0;
          aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aMRPUnitPtr^.snumber );
          slNumber.AddObject(aMRPUnitPtr^.snumber, TObject(aEORecordPtr));
        end
        else
        begin
          aEORecordPtr := PEORecord(slNumber.Objects[idx]);
        end;

        aEORecordPtr^.dQtyDemand := aEORecordPtr^.dQtyDemand + aMRPUnitPtr^.dQty;
        
        if DoubleL(aMRPUnitPtr^.dt, today + 17) then
        begin
          aEORecordPtr^.dQtyDemand17 := aEORecordPtr^.dQtyDemand17 + aMRPUnitPtr^.dQty; 
        end;  
        if DoubleL(aMRPUnitPtr^.dt, today + 28) then
        begin
          aEORecordPtr^.dQtyDemand28 := aEORecordPtr^.dQtyDemand28 + aMRPUnitPtr^.dQty;
        end;
        if DoubleL(aMRPUnitPtr^.dt, today + 60) then
        begin
          aEORecordPtr^.dQtyDemand60 := aEORecordPtr^.dQtyDemand60 + aMRPUnitPtr^.dQty;
        end; 

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

    Memo1.Lines.Add('����  ��������');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '��������';


    sl := TStringList.Create;
    try
      sline := '���ϱ���'#9'��������'#9'����'#9'��������'#9'FROM'#9'TO';
      sl.Add(sline);
 
      for iLine := 0 to slTransfer.Count - 1 do
      begin
        aTransferRecordPtr := slTransfer.Items[iLine];
        sline := aTransferRecordPtr^.snumber + #9 +
          aTransferRecordPtr^.sname + #9 +
          Format('%0.0f', [aTransferRecordPtr^.dqty]) + #9 +
          FormatDateTime('yyyy-MM-dd', aTransferRecordPtr^.dt) + #9 +
          aTransferRecordPtr^.sfrom + #9 +
          aTransferRecordPtr^.sto;

        sl.Add(sline);
      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;      
          
    finally
      sl.Free;
    end;

    for iLine := 0 to slTransfer.Count - 1 do
    begin
      aTransferRecordPtr := PTransferRecord(slTransfer[iLine]);
      Dispose(aTransferRecordPtr);
    end;
    slTransfer.Clear;

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
      sline := '�ɹ�ƾ֤'#9'�к�'#9'����'#9'��������'#9'��������'#9'���鵽������'#9'��������'#9'����'#9'MRP Area'#9'Mrp Area No'#9'MC'#9'�ɹ�Ա';
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
          sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPer;
          sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;

          sl.Add(sline);
        end;

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
          
          sl.Add(sline);
        end;

                    
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
      sline := '����'#9'��������'#9'������'#9'�ܿ��'#9'�ܶ���'#9'������'#9'��������'#9'E/O'#9'17������'#9'28������'#9'60������'#9'MRP����'#9'MRP����'#9'MC'#9'�ɹ�Ա';
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
          '=IF(E' + IntToStr(irow) + '+D' + IntToStr(irow) + '>C' + IntToStr(irow) + ',IF(C' + IntToStr(irow) + '>0,"Excess","Obslete"),"")' + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand17]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand28]) + #9 +
          Format('%0.0f', [aEORecordPtr^.dQtyDemand60]) + #9 +
          aEORecordPtr^.sMRPType + #9 +
          aEORecordPtr^.sMrpAreaName;        
        sline := sline + #9 + aSAPMaterialRecordPtr^.sMRPer;
        sline := sline + #9 + aSAPMaterialRecordPtr^.sBuyer;

        sl.Add(sline);
        irow := irow + 1;

      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;
          
    finally
      sl.Free;
    end;

       
    aSAPMaterialReader.Free;
    aSAPMaterialReader := nil;


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
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstDemand.Free;

    slGroup.Free;

    aSAPMrpAreaStockReader.Free;
  end;

  slArea2BomFac.Free;

  MessageBox(Handle, '���', '��ʾ', 0);
end;

end.
