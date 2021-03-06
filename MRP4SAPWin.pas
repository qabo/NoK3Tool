unit MRP4SAPWin;
 
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
  DOSReader, MRPWinReader, jpeg, ImgList, LTPCMSConfirmReader, SAPMaterialReader2,
  SAPMaterialReader,
  SAPStockReader2, SAPBomReader, SAPBomReader2, SAPS618Reader;

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
    dQtyDemand: Double;
    dQtyDemand17: Double;  
    dQtyDemand28: Double;  
    dQtyDemand60: Double;
    dQtyStock: Double;
    dQtyOPO: Double;     
    sMRPType: string;
  end;
  PEORecord = ^TEORecord;
                                                             
  TfrmMRP4SAP = class(TForm)
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
    procedure FormCreate(Sender: TObject); 
    procedure FormDestroy(Sender: TObject);
    procedure tbCloseClick(Sender: TObject);
    procedure btnStockClick(Sender: TObject);
    procedure btnMRPClick(Sender: TObject);
    procedure btnSAPStock3Click(Sender: TObject);
    procedure btnSAPBom3Click(Sender: TObject);
    procedure btnDemand3Click(Sender: TObject);
    procedure btnSAPOPOClick(Sender: TObject);
    procedure btnMaterialClick(Sender: TObject);
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

class procedure TfrmMRP4SAP.ShowForm;
var
  frmMRPSimulation: TfrmMRP4SAP;
begin
  frmMRPSimulation := TfrmMRP4SAP.Create(nil);
  try
    frmMRPSimulation.ShowModal;
  finally
    frmMRPSimulation.Free;
  end;
end;
   
procedure TfrmMRP4SAP.FormCreate(Sender: TObject);
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
 
  ini.Free;
end;

procedure TfrmMRP4SAP.FormDestroy(Sender: TObject);
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
 
procedure TfrmMRP4SAP.tbCloseClick(Sender: TObject);
begin
  Close;
end;
 
procedure TfrmMRP4SAP.btnStockClick(Sender: TObject);
begin
end;

  // 简化的MRP计算，不考虑低位码
  type
    TMRPUnit = packed record
      id: Integer;
      pid: Integer;
      snumber: string;
      sname: string;
      dt: TDateTime;
      dQty: Double;
      dQtyStock: Double;
//      dQtyStockParent: Double;
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
      else Result := 0;
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
    
procedure TfrmMRP4SAP.btnSAPStock3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPStock.Text := sfile;
end;
     
procedure TfrmMRP4SAP.btnSAPOPOClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPOPO.Text := sfile;
end;
      
procedure TfrmMRP4SAP.btnMaterialClick(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leMaterial.Text := sfile;
end;

procedure TfrmMRP4SAP.btnSAPBom3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPBom.Text := sfile;
end;

procedure TfrmMRP4SAP.btnDemand3Click(Sender: TObject);
var
  sfile: string;
begin
  if not ExcelOpenDialog(sfile) then Exit;
  leSAPPIR.Text := sfile;
end;

procedure TfrmMRP4SAP.OnLogEvent(const s: string);
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

procedure TfrmMRP4SAP.btnMRPClick(Sender: TObject);
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

  function GetByNumberDate(lst: TList; const snumber: string; dt: TDateTime): PMRPUnit;
  var
    i: Integer;
  begin
    Result := nil;
    for i := 0 to lst.Count - 1 do
    begin
      if (PMRPUnit( lst[i] )^.snumber = snumber) and DoubleE(PMRPUnit( lst[i] )^.dt, dt) then
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
  aSAPStockSum: TSAPStockSum;
  aSAPS618Reader: TSAPPIRReader;
  aSAPOPOReader: TSAPOPOReader2;
  lstDemand: TList;
  lstDemand_tmp: TList;
  iLine: Integer;
  iWeek: Integer;
  aSOPProj: TSOPProj;
  iDate: Integer;
  aMRPUnitPtr: PMRPUnit;
  aMRPUnitPtr_Dep: PMRPUnit;
  aMRPUnitPtr_Sum: PMRPUnit;
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

  lstPRSum: TList;
begin
  if not ExcelSaveDialog(sfile) then Exit;

  today := myStrToDateTime(FormatDateTime('yyyy-MM-dd', Now));

  Memo1.Lines.Add('开始读取 BOM  ' + leSAPBom.Text);
  aSAPBomReader := TSAPBomReader2.Create(leSAPBom.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 库存  ' + leSAPStock.Text);
  aSAPStockReader := TSAPStockReader2.Create(leSAPStock.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 OPO  ' + leSAPOPO.Text);
  aSAPOPOReader := TSAPOPOReader2.Create(leSAPOPO.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 物料  ' + leMaterial.Text);
  aSAPMaterialReader:= TSAPMaterialReader2.Create(leMaterial.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 PIR  ' + leSAPPIR.Text);
  aSAPS618Reader := TSAPPIRReader.Create(leSAPPIR.Text, OnLogEvent);


  aSAPStockSum := TSAPStockSum.Create;
  aSAPStockReader.SumTo(aSAPStockSum);

//  lstMrpDetail := TList.Create;

  lstDemand := TList.Create; 
                                     
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
      
      aMRPUnitPtr := New(PMRPUnit);
      aMRPUnitPtr^.id := iid;   
      iid := iid + 1;
      aMRPUnitPtr^.pid := 0;
      aMRPUnitPtr^.snumber := aSAPS618ColPtr^.sNumber;
      aMRPUnitPtr^.sname := aSAPS618ColPtr^.sname;
      aMRPUnitPtr^.dt := aSAPS618ColPtr^.dt1;
      aMRPUnitPtr^.dQty := aSAPS618ColPtr^.dQty;
      aMRPUnitPtr^.dQtyStock := 0;
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
        aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
 
        if aMRPUnitPtr^.bCalc then Continue; // 计算过的，不算

        // 低位码小于等于当前计算低位码，才计算
        if aMRPUnitPtr^.aSAPMaterialRecordPtr^.iLowestCode > iLowestCode then
        begin
          bLoop := True;  // 有需求还没计算，需继续循环
          Continue;
        end;


        ////  自制件 展开BOM  //////////////////////////////////////////////
        if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sMRPType = 'M0' then
        begin
          // 无无无无 替代料
          if aMRPUnitPtr^.iSubstituteNo = 0 then
          begin        
            aMRPUnitPtr^.bCalc := True;
            aMRPUnitPtr^.bExpend := True;

            if aMRPUnitPtr^.aParentBom = nil then // 根节点，需查找BOM
            begin
              aMRPUnitPtr^.aBom := aSAPBomReader.GetSapBom(aMRPUnitPtr^.snumber, '');
            end;

            if aMRPUnitPtr^.sDemandType <> 'BSF' then  //  LSF考虑库存， BSF 不考虑库存
            begin
              aMRPUnitPtr^.dQtyStock := aSAPStockSum.Alloc2('', aMRPUnitPtr^.snumber, aMRPUnitPtr^.dQty);
            end;

            // 分配库存后，需求满足了，无无无无需往下展开
            if DoubleLE( aMRPUnitPtr^.dQty, aMRPUnitPtr^.dQtyStock ) then
            begin
              Continue;
            end;
          
            if aMRPUnitPtr^.aBom = nil then  // 有需求，但是 没有BOM，异常，记录日志
            begin
              Memo1.Lines.Add(aMRPUnitPtr^.snumber + ' 无BOM'); 
              Continue;
            end;

            aSapBomChild := nil;

            //展开需求到下层
            for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
              iPer100 := 0;  
              for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
              begin
                aSapBomChild := aSapItemGroup.Items[iChildItem];
 
                aMRPUnitPtr_Dep := New(PMRPUnit);
                aMRPUnitPtr_Dep^.id := iid;
                iid := iid + 1;
              
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
                  // 有替代料，按配比分
                  aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                  iPer100 := iPer100 + Round(aSapBomChild.dPer);
                end;
                
                aMRPUnitPtr_Dep^.dQtyStock := 0;
                aMRPUnitPtr_Dep^.dQtyOPO := 0;
                aMRPUnitPtr_Dep^.bExpend := False;
                aMRPUnitPtr_Dep^.bCalc := False;
                aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                aMRPUnitPtr_Dep^.spriority := aSapBomChild.spriority; // 优先级
                aMRPUnitPtr_Dep^.dPer := aSapBomChild.dPer;

                if aSapItemGroup.ItemCount = 1 then
                begin
                  aMRPUnitPtr_Dep^.iSubstituteNo := 0; // 没有替代组
                end
                else
                begin
                  aMRPUnitPtr_Dep^.iSubstituteNo := iSubstituteNo;
                end;
                lstDemand.Add(aMRPUnitPtr_Dep);
              end;
              iSubstituteNo := iSubstituteNo + 1; //  替代组编号 + 1，确保唯一

              // 配比总和不为 100
              if iPer100 <> 100 then
              begin
                Memo1.Lines.Add('配比总和不为 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr^.snumber);
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
              if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
              begin
                dQty := dQty + aMRPUnitPtr_Dep^.dQty;  // 汇总替代料的需求
                lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
              end;
            end;

            // 替代料消耗优先级 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep^.dQtyStock := aSAPStockSum.Alloc2('', aMRPUnitPtr_Dep^.snumber, dQty);   // 如果某个替代料库存全满足了，剩余数量为0， 自动会为下面的替代料分配0的库存和需求
              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep^.dQtyStock;
              
              aMRPUnitPtr_Dep^.bCalc := True;
            end;

            // 需求没有全满足， 按比例分配需求给替代料
            if dQty > 0 then
            begin
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr := lstSubstituteDemand[idx];  
                aMRPUnitPtr^.bCalc := True;
                aMRPUnitPtr^.bExpend := True;
                aMRPUnitPtr^.dQty := aMRPUnitPtr^.dQty + dQty * aMRPUnitPtr^.dPer / 100;
                if DoubleE( aMRPUnitPtr^.dQty, 0) then Continue; 

                // 分配了需求就往下展需求
                for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
                begin
                  aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
                  iPer100 := 0;
                  aSapBomChild := nil;
                  //展开需求到下层
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem];


                    aMRPUnitPtr_Dep := New(PMRPUnit);
                    aMRPUnitPtr_Dep^.id := iid;
                    iid := iid + 1;
              
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
                      // 有替代料，按配比分
                      aMRPUnitPtr_Dep^.dQty := (aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock) * aSapBomChild.dusage * aSapBomChild.dPer / 100;
                      iPer100 := iPer100 + Round(aSapBomChild.dPer);
                    end;
                
                    aMRPUnitPtr_Dep^.dQtyStock := 0;
                    aMRPUnitPtr_Dep^.dQtyOPO := 0;
                    aMRPUnitPtr_Dep^.bExpend := False;
                    aMRPUnitPtr_Dep^.bCalc := False;
                    aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                    aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;
                    aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                    aMRPUnitPtr_Dep^.spriority := aSapBomChild.spriority; // 优先级
                    aMRPUnitPtr_Dep^.dPer := aSapBomChild.dPer;

                    if aSapItemGroup.ItemCount = 1 then
                    begin
                      aMRPUnitPtr_Dep^.iSubstituteNo := 0; // 没有替代组
                    end
                    else
                    begin
                      aMRPUnitPtr_Dep^.iSubstituteNo := iSubstituteNo;
                    end;
                    lstDemand.Add(aMRPUnitPtr_Dep);
                  end;
                  iSubstituteNo := iSubstituteNo + 1; //  替代组编号 + 1，确保唯一

                  // 配比总和不为 100
                  if iPer100 <> 100 then
                  begin
                    Memo1.Lines.Add('配比总和不为 100  ' + aSapBomChild.FNumber + ' ' + aMRPUnitPtr^.snumber);
                  end;
                  
                end;
                       

                
              end;                       // 之前分配 库存的时候已经分配了一部分Qty，所以要加          
              bLoop := True;  //  展开有新的需求，需继续循环
            end;


 

            lstSubstituteDemand.Free;


          end;
        end
        else  //// 外购件，不展开BOM  ///////  PD  ////////////////////////////////
        begin
          if aMRPUnitPtr^.bCalc then
          begin
            Continue;  // 已计算
          end;

          // 无无无无 替代料
          if aMRPUnitPtr^.iSubstituteNo = 0 then
          begin
            aMRPUnitPtr^.dQtyStock := aSAPStockSum.Alloc2('', aMRPUnitPtr^.snumber, aMRPUnitPtr^.dQty);
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

            // 替代料消耗优先级 
            lstSubstituteDemand.Sort(ListSortCompare_priority);

            for idx := 0 to lstSubstituteDemand.Count - 1 do
            begin       
              aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
              aMRPUnitPtr_Dep^.dQtyStock := aSAPStockSum.Alloc2('', aMRPUnitPtr_Dep^.snumber, dQty);   // 如果某个替代料库存全满足了，剩余数量为0， 自动会为下面的替代料分配0的库存和需求
              aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock;
              dQty := dQty  - aMRPUnitPtr_Dep^.dQtyStock;
              
              aMRPUnitPtr_Dep^.bCalc := True;
            end;

            // 需求没有全满足， 按比例分配需求给替代料
            if dQty > 0 then
            begin
              for idx := 0 to lstSubstituteDemand.Count - 1 do
              begin       
                aMRPUnitPtr_Dep := lstSubstituteDemand[idx];
                aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty * aMRPUnitPtr_Dep^.dPer / 100;
              end;                       // 之前分配 库存的时候已经分配了一部分Qty，所以要加
            end; 

            lstSubstituteDemand.Free;
          end;
        end;        

      end;

      iLowestCode := iLowestCode + 1; 
    end;
               
    //排序，按日期
    lstDemand.Sort(ListSortCompare_DateTime);

     
    ////  计算 PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      aMRPUnitPtr^.bCalc := False;
    end;

    ////  计算 PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin                                 
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);

  
              
      // 非外购件不计算 PO 
      if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType <> 'F' then Continue;

      if aMRPUnitPtr^.bCalc then Continue; //计算过了 

      slNumber := TStringList.Create;
      // 无无无 替代料
      if aMRPUnitPtr^.iSubstituteNo = 0 then
      begin
        aMRPUnitPtr^.bCalc := True;
                            
        if DoubleE( aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock,  0) then Continue; // 没有需求

        slNumber.Add(aMRPUnitPtr^.snumber);

        aMRPUnitPtr^.dQtyOPO := aSAPOPOReader.Alloc(slNumber, aMRPUnitPtr^.dt,
          aMRPUnitPtr^.dQty - aMRPUnitPtr^.dQtyStock);

      end
      else  //  有有有有  替代料
      begin
        lstSubstituteDemand := TList.Create;
        dQty := 0;
        for idx := 0 to lstDemand.Count - 1 do
        begin
          aMRPUnitPtr_Dep := lstDemand[idx];
          if aMRPUnitPtr_Dep^.iSubstituteNo = aMRPUnitPtr^.iSubstituteNo then
          begin
            dQty := dQty + aMRPUnitPtr_Dep^.dQty - aMRPUnitPtr_Dep^.dQtyStock; // 减去已分配库存  
            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock;         // 分配的库存是固定的，不要清掉
            lstSubstituteDemand.Add(aMRPUnitPtr_Dep);
            slNumber.Add(aMRPUnitPtr_Dep^.snumber);


            aMRPUnitPtr_Dep^.bCalc := True;
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
        aSAPOPOReader.GetOPOs(slNumber, lstPOLine); // 找到所有替代料的可用采购订单

        lstPOLine.Sort(ListSortCompare_DateTime_PO);  // 按日期排序

        for idx := 0 to lstPOLine.Count - 1 do
        begin
          aSAPOPOLine := TSAPOPOLine(lstPOLine[idx]);
          aMRPUnitPtr_Dep := GetMRPUnit(lstSubstituteDemand, aSAPOPOLine.FNumber);

          // 分配的OPO 累加
          aMRPUnitPtr_Dep^.dQtyOPO := aMRPUnitPtr_Dep^.dQtyOPO + aSAPOPOLine.Alloc(aMRPUnitPtr_Dep^.dt, dQty, '');
          aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQtyStock + aMRPUnitPtr_Dep^.dQtyOPO;

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
            aMRPUnitPtr_Dep^.dQty := aMRPUnitPtr_Dep^.dQty + dQty * aMRPUnitPtr_Dep^.dPer / 100; 
          end;
        end; 

        lstSubstituteDemand.Free;
      end;  


      slNumber.Free;
    end;

  finally

    aSAPBomReader.Free;
    aSAPStockSum.Free;

  end;
   
    
  slNumber := TStringList.Create;


  for iLine := 0 to aSAPStockReader.Count - 1 do
  begin
    aSAPStockRecordPtr := aSAPStockReader.Items[ iLine ];

      idx := slNumber.IndexOf(aSAPStockRecordPtr^.snumber);
      if idx < 0 then
      begin
        aEORecordPtr := New(PEORecord);
        aEORecordPtr^.snumber := aSAPStockRecordPtr^.snumber;
        aEORecordPtr^.sname := aSAPStockRecordPtr^.sname;
        aEORecordPtr^.dQtyDemand := 0;
        aEORecordPtr^.dQtyStock := 0;
        aEORecordPtr^.dQtyOPO := 0;      
        aEORecordPtr^.dQtyDemand17 := 0;
        aEORecordPtr^.dQtyDemand28 := 0;
        aEORecordPtr^.dQtyDemand60 := 0;
        aEORecordPtr^.sMRPType := aSAPMaterialReader.GetMrpType( aSAPStockRecordPtr^.snumber );
        if aEORecordPtr^.snumber = '01.01.1012978A' then
        begin
          Sleep(100);
        end;
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
    ExcelApp.Sheets[iSheet].Name := '独立需求';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '产品编码';
    ExcelApp.Cells[irow, 2].Value := '产品名称';
    ExcelApp.Cells[irow, 3].Value := '日期';
    ExcelApp.Cells[irow, 4].Value := '数量';    
    ExcelApp.Cells[irow, 5].Value := '需求类型';

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

    Memo1.Lines.Add('计算结果明细');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '计算结果明细';


    sl := TStringList.Create;
    try
      sline := 'ID'#9'父ID'#9'物料'#9'物料名称'#9'需求日期'#9'建议下单日期'#9'需求数量'#9'可用库存'#9'OPO'#9'净需求'#9'替代组'#9'MRP控制者'#9'采购员';
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

        if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType = 'F' then  // 外购  ////////////////////////////////////////
        begin
          sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_PD) + #9;
        end
        else                                                         // 自制  ////////////////////////////////////////
        begin
          sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_M0) + #9;
        end;  
          
        sline := sline + Format('%0.0f', [aMRPUnitPtr^.dqty]) + #9 +
          Format('%0.0f', [aMRPUnitPtr^.dqtystock]) + #9 +
          Format('%0.0f', [aMRPUnitPtr^.dQtyOPO]) + #9 +
          '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9 +  // 9 = 6 - 7 - 8
          IntToStr(aMRPUnitPtr^.iSubstituteNo) + #9 +
          aMRPUnitPtr^.aSAPMaterialRecordPtr^.sMRPer + #9 +
          aMRPUnitPtr^.aSAPMaterialRecordPtr^.sBuyer;

        irow := irow + 1;
        sl.Add(sline);

        idx := slNumber.IndexOf(aMRPUnitPtr^.snumber);
        if idx < 0 then
        begin
          aEORecordPtr := New(PEORecord);
          aEORecordPtr^.snumber := aMRPUnitPtr^.snumber;  
          aEORecordPtr^.sname := aMRPUnitPtr^.sname;
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

        if aEORecordPtr^.snumber = '01.01.1012978A' then
        begin
          Sleep(100);
        end;

        aEORecordPtr^.dQtyDemand := aEORecordPtr^.dQtyDemand + aMRPUnitPtr^.dQty;
        
        if DoubleL(aMRPUnitPtr^.dt, today + 17) then
        begin
          aEORecordPtr^.dQtyDemand17 := aEORecordPtr^.dQtyDemand17 + aMRPUnitPtr^.dQty;
        end
        else
        begin
          if aMRPUnitPtr^.snumber = '01.01.1012228' then
          begin
            Sleep(1);
          end;
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

    end;
            
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('计算结果汇总');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '计算结果汇总';  // 物料、日期

    lstPRSum := TList.Create;

    try
      for iMrpUnit := 0 to lstDemand.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);

        aMRPUnitPtr_Sum := GetByNumberDate(lstPRSum, aMRPUnitPtr^.snumber, aMRPUnitPtr^.dt);
        if aMRPUnitPtr_Sum = nil then
        begin
          aMRPUnitPtr_Sum := New(PMRPUnit);
          lstPRSum.Add(aMRPUnitPtr_Sum);
          aMRPUnitPtr_Sum^ := aMRPUnitPtr^;
        end
        else
        begin
          aMRPUnitPtr_Sum^.dQty := aMRPUnitPtr_Sum^.dQty + aMRPUnitPtr^.dQty; 
          aMRPUnitPtr_Sum^.dQtyStock := aMRPUnitPtr_Sum^.dQtyStock + aMRPUnitPtr^.dQtyStock; 
          aMRPUnitPtr_Sum^.dQtyOPO := aMRPUnitPtr_Sum^.dQtyOPO + aMRPUnitPtr^.dQtyOPO;
        end;
      end;

 
      sl := TStringList.Create;
      try
        sline := 'ID'#9'父ID'#9'物料'#9'物料名称'#9'需求日期'#9'建议下单日期'#9'需求数量'#9'可用库存'#9'OPO'#9'净需求'#9'替代组'#9'MRP控制者'#9'采购员';
        sl.Add(sline);

        irow := 2;
        for iMrpUnit := 0 to lstPRSum.Count - 1 do
        begin
          aMRPUnitPtr := PMRPUnit(lstPRSum[iMrpUnit]);

          sline := IntToStr(aMRPUnitPtr^.id) + #9 +
            IntToStr(aMRPUnitPtr^.pid) + #9 +
            aMRPUnitPtr^.snumber + #9 +
            aMRPUnitPtr^.sname + #9 +
            FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt) + #9;

          if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sPType = 'F' then  // 外购  ////////////////////////////////////////
          begin
            sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_PD) + #9;
          end
          else                                                         // 自制  ////////////////////////////////////////
          begin
            sline := sline + FormatDateTime('yyyy-MM-dd', aMRPUnitPtr^.dt - aMRPUnitPtr^.aSAPMaterialRecordPtr^.dLT_M0) + #9;
          end;  
          
          sline := sline + Format('%0.0f', [aMRPUnitPtr^.dqty]) + #9 +
            Format('%0.0f', [aMRPUnitPtr^.dqtystock]) + #9 +
            Format('%0.0f', [aMRPUnitPtr^.dQtyOPO]) + #9 +
            '=' + GetRef(7) + IntToStr(irow) + '-' + GetRef(8) + IntToStr(irow) + '-' + GetRef(9) + IntToStr(irow) + #9 +  // 9 = 6 - 7 - 8
            IntToStr(aMRPUnitPtr^.iSubstituteNo) + #9 +
            aMRPUnitPtr^.aSAPMaterialRecordPtr^.sMRPer + #9 +
            aMRPUnitPtr^.aSAPMaterialRecordPtr^.sBuyer;

          irow := irow + 1;
          sl.Add(sline);

          (*
          idx := slNumber.IndexOf(aMRPUnitPtr^.snumber);
          if idx < 0 then
          begin
            aEORecordPtr := New(PEORecord);
            aEORecordPtr^.snumber := aMRPUnitPtr^.snumber;  
            aEORecordPtr^.sname := aMRPUnitPtr^.sname;
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

          if aEORecordPtr^.snumber = '01.01.1012978A' then
          begin
            Sleep(100);
          end;
                    
          aEORecordPtr^.dQtyDemand := aEORecordPtr^.dQtyDemand + aMRPUnitPtr^.dQty;
        
          if DoubleL(aMRPUnitPtr^.dt, today + 17) then
          begin
            aEORecordPtr^.dQtyDemand17 := aEORecordPtr^.dQtyDemand17 + aMRPUnitPtr^.dQty;
          end
          else
          begin
            if aMRPUnitPtr^.snumber = '01.01.1012228' then
            begin
              Sleep(1);
            end;
          end;  
          if DoubleL(aMRPUnitPtr^.dt, today + 28) then
          begin
            aEORecordPtr^.dQtyDemand28 := aEORecordPtr^.dQtyDemand28 + aMRPUnitPtr^.dQty;
          end;
          if DoubleL(aMRPUnitPtr^.dt, today + 60) then
          begin
            aEORecordPtr^.dQtyDemand60 := aEORecordPtr^.dQtyDemand60 + aMRPUnitPtr^.dQty;
          end;
          *)

        end;   
            
        Clipboard.SetTextBuf(PChar(sl.Text));
        ExcelApp.ActiveSheet.Paste;      
         
      finally
        sl.Free;
      end;


      
    finally
      for iMrpUnit := 0 to lstPRSum.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstPRSum[iMrpUnit]);
        Dispose(aMRPUnitPtr);
      end;
      lstPRSum.Free;
    end;


         
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
      sline := '行号'#9'物料'#9'物料名称'#9'建议数量'#9'需求日期'#9'建议到料日期'#9'建议';
      sl.Add(sline);
 
      for iLine := 0 to aSAPOPOReader.Count - 1 do
      begin
        aSAPOPOLine := aSAPOPOReader.Items[iLine];
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
            'Cancel';
          
          sl.Add(sline);
        end;


        idx := slNumber.IndexOf(aSAPOPOLine.FNumber);
        if idx < 0 then
        begin
          aEORecordPtr := New(PEORecord);
          aEORecordPtr^.snumber := aSAPOPOLine.FNumber;   
          aEORecordPtr^.sname := aSAPOPOLine.FName;
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
        if aEORecordPtr^.snumber = '01.01.1012978A' then
        begin
          Sleep(100);
        end;
        aEORecordPtr^.dQtyOPO := aEORecordPtr^.dQtyOPO + aSAPOPOLine.FQty;

      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;      
          
    finally
      sl.Free;
    end;
             
    aSAPMaterialReader.Free;

    (*
    irow := 1;
    
    ExcelApp.Cells[irow, 1].Value := '订单号';
    ExcelApp.Cells[irow, 2].Value := '行号';
    ExcelApp.Cells[irow, 3].Value := '物料';
    ExcelApp.Cells[irow, 4].Value := '物料名称';
    ExcelApp.Cells[irow, 5].Value := '需求数量';
    ExcelApp.Cells[irow, 6].Value := '需求日期';
    ExcelApp.Cells[irow, 7].Value := '交货日期';
    ExcelApp.Cells[irow, 8].Value := '建议';

    irow := 2;

    for iLine := 0 to aSAPOPOReader.Count - 1 do
    begin
      aSAPOPOLine := aSAPOPOReader.Items[iLine];
      for iLineAlloc := 0 to aSAPOPOLine.Count - 1 do
      begin
        aSAPOPOAllocPtr := aSAPOPOLine.Items[iLineAlloc];
        ExcelApp.Cells[irow, 1].Value := aSAPOPOLine.FBillNo;
        ExcelApp.Cells[irow, 2].Value := aSAPOPOLine.FLine;
        ExcelApp.Cells[irow, 3].Value := aSAPOPOLine.FNumber;
        ExcelApp.Cells[irow, 4].Value := aSAPOPOLine.FName;
        ExcelApp.Cells[irow, 5].Value := aSAPOPOAllocPtr^.dQty;
        ExcelApp.Cells[irow, 6].Value := aSAPOPOAllocPtr^.dt;
        ExcelApp.Cells[irow, 7].Value := aSAPOPOLine.FDate;

        if DoubleE( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // 准时交货
        begin
          ExcelApp.Cells[irow, 8].Value := 'OTD';
        end
        else if DoubleG( aSAPOPOLine.FDate, aSAPOPOAllocPtr^.dt ) then // 订单交期， 晚于需求日期，Push In
        begin
          ExcelApp.Cells[irow, 8].Value := 'Push In';
        end
        else            // 订单日期日期，早于需求日期, Push Out
        begin
          ExcelApp.Cells[irow, 8].Value := 'Push Out';
        end;
        irow := irow + 1;
      end;

      if DoubleG( aSAPOPOLine.FQty, aSAPOPOLine.FQtyAlloc ) then
      begin
        ExcelApp.Cells[irow, 1].Value := aSAPOPOLine.FBillNo;
        ExcelApp.Cells[irow, 2].Value := aSAPOPOLine.FLine;
        ExcelApp.Cells[irow, 3].Value := aSAPOPOLine.FNumber;
        ExcelApp.Cells[irow, 4].Value := aSAPOPOLine.FName;
        ExcelApp.Cells[irow, 5].Value := aSAPOPOLine.FQty - aSAPOPOLine.FQtyAlloc;
        ExcelApp.Cells[irow, 6].Value := '';
        ExcelApp.Cells[irow, 7].Value := aSAPOPOLine.FDate;
        ExcelApp.Cells[irow, 8].Value := 'Cancel';   
        irow := irow + 1;
      end;

    end;

    *) 
    ////////////////////////////////////////////////////////////////////////////
                                                                                 
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    
    aSAPOPOReader.Free;

                        
    ////////////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('保存 E&o');

    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'E&O';


    sl := TStringList.Create;
    try
      sline := '物料'#9'物料名称'#9'总需求'#9'总库存'#9'总订单'#9'库存呆滞'#9'订单呆滞'#9'E/O'#9'17天需求'#9'28天需求'#9'60天需求'#9'MRP类型';
      sl.Add(sline);

      irow := 2;
      for iLine := 0 to slNumber.Count - 1 do
      begin
        aEORecordPtr := PEORecord(slNumber.Objects[iLine]);

        if aEORecordPtr^.snumber = '01.01.1012978A' then
        begin
          Sleep(100);
        end;
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
          aEORecordPtr^.sMRPType;
                  
        sl.Add(sline);
        irow := irow + 1;

      end;

      Clipboard.SetTextBuf(PChar(sl.Text));
      ExcelApp.ActiveSheet.Paste;
          
    finally
      sl.Free;
    end;

             

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
    ExcelApp.Sheets[iSheet].Name := '可用库存';


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

          if stitle <> '采购凭证项目物料短文本计划行采购凭证类型' then
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

          if stitle <> '工厂物料凭证物料凭证项目凭证日期' then
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
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstDemand.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
