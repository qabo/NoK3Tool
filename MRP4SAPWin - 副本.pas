unit MRP4SAPWin;

(*
1、 库存可以满足需求的物料，SAP ZPPR020 报表不会提取。Load物料供应的时候需要需要提取库存


*)

(*
系统外做模拟MRP与齐套模拟步骤：
  1、导出调整后的要货计划  ZPPR028
  2、导出BOM ZPPR021， 料号从调整后的要货计划取
  3、导出库存 MB52
  4、BOM转SBOM  无需带库存
  5、调整后的要货计划转S&OP格式，注意SKU不要重复
  6、SKU优先级计算， SKU优先级文件里的才会参与齐套分析
  7、模拟MRP
  8、齐套分析 *******8
*)

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, StdCtrls, ExtCtrls, IniFiles, CommUtils,
  DB, ADODB, Provider, DBClient, SOPReaderUnit, SAPBom2SBomWin, SAPOPOReader, 
  DateUtils, ComObj, ExcelConsts, MrpMPSReader, NewSKUReader, LTP_CMS2MRPSimWin,
  KeyICItemSupplyReader, SBomReader, SOPSimReader, FGPriorityReader,
  DOSReader, MRPWinReader, jpeg, ImgList, LTPCMSConfirmReader, SAPMaterialReader,
  SAPStockReader, SAPBomReader, SAPS618Reader;

type
  TDeltaRecord = packed record
    sproj: string;
    snumber: string;
    sdate: string;
    iqty: Double;
    iqty_org: Double;
  end;
  PDeltaRecord = ^TDeltaRecord;
                                                             
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
    Result := StrToDateTime(s);
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
      dQtyStockParent: Double;
      dQtyOPO: Double;
      bExpend: Boolean;
      aBom: TSapBom;
      aParentBom: TSapBom;
      aSAPMaterialRecordPtr: PSAPMaterialRecord;
      sDemandType: string;
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

procedure TfrmMRP4SAP.btnMRPClick(Sender: TObject);
var                                                                                                       
  sfile: string;
  aSAPMaterialReader: TSAPMaterialReader;
  aSAPBomReader: TSAPBomReader;
  aSAPStockReader: TSAPStockReader;
  aSAPStockSum: TSAPStockSum;
  aSAPS618Reader: TSAPPIRReader;
  aSAPOPOReader: TSAPOPOReader;
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
  lstMrpDetail: TList;
  iID: Integer;
  slNumber: TStringList;

  
  aSAPOPOLine: TSAPOPOLine;
  aSAPOPOAllocPtr: PSAPOPOAlloc;
  iLineAlloc: Integer;

  aSAPMaterialRecordPtr: PSAPMaterialRecord;
  iLowestCode: Integer;
begin
  if not ExcelSaveDialog(sfile) then Exit;


  Memo1.Lines.Add('开始读取 物料');
  aSAPMaterialReader:= TSAPMaterialReader.Create(leMaterial.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 库存');
  aSAPStockReader := TSAPStockReader.Create(leSAPStock.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 OPO');
  aSAPOPOReader := TSAPOPOReader.Create(leSAPOPO.Text, OnLogEvent);

  Memo1.Lines.Add('开始读取 BOM');
  aSAPBomReader := TSAPBomReader.Create(leSAPBom.Text, OnLogEvent);
  
  Memo1.Lines.Add('开始读取 PIR');
  aSAPS618Reader := TSAPPIRReader.Create(leSAPPIR.Text, OnLogEvent);


  aSAPStockSum := TSAPStockSum.Create;
  aSAPStockReader.SumTo(aSAPStockSum);

  lstMrpDetail := TList.Create;
  
  lstDemand := TList.Create; 
                                     
  iID := 1;
 

  ////  计算低位码  ////////////////////////////////////////////////////////////
  for idx := 0 to aSAPMaterialReader.Count - 1 do
  begin
    aSAPMaterialRecordPtr := aSAPMaterialReader.Items[idx];
    aSAPBomReader.GetLowestCode(aSAPMaterialRecordPtr);
    Memo1.Lines.Add(aSAPMaterialRecordPtr^.sNumber + '   ' + IntToStr(aSAPMaterialRecordPtr^.iLowestCode));
  end;


  Memo1.Lines.Add('整理要货计划需求');
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
      aMRPUnitPtr^.dQtyStockParent := 0;
      aMRPUnitPtr^.dQtyOPO := 0;
      aMRPUnitPtr^.bExpend := False;
      aMRPUnitPtr^.aBom := nil;     
      aMRPUnitPtr^.aParentBom := nil;    
      aMRPUnitPtr^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSAPS618ColPtr^.sNumber);
      aMRPUnitPtr^.sDemandType := aSAPS618.sDemandType;
      lstDemand.Add(aMRPUnitPtr);
    end; 
  end;
                            
  Memo1.Lines.Add('开始模拟MRP计算');
  try
    iLowestCode := 0;
    bLoop := True;
    while bLoop do
    begin
      bLoop := False;
      
      //备份需求
      lstDemand_tmp := TList.Create;
      for iMrpUnit := 0 to lstDemand.Count - 1 do
      begin
        lstDemand_tmp.Add(lstDemand[iMrpUnit]);
      end;
      lstDemand.Clear;

      //排序，按日期
      lstDemand_tmp.Sort(ListSortCompare_DateTime);
 
      for iMrpUnit := 0 to lstDemand_tmp.Count - 1 do
      begin
        aMRPUnitPtr := PMRPUnit(lstDemand_tmp[iMrpUnit]);
        if not aMRPUnitPtr^.bExpend then
        begin
          // 根节点  sDemandType = BSF 不考虑库存 ////////////////////////////////////////////////////
          if aMRPUnitPtr^.aParentBom = nil then
          begin
            ////  自制件 展开BOM  
            if aMRPUnitPtr^.aSAPMaterialRecordPtr^.sMRPType = 'PD' then
            begin
              aMRPUnitPtr^.bExpend := True;
              aMRPUnitPtr^.aBom := aSAPBomReader.GetSapBom(aMRPUnitPtr^.snumber);
              if aMRPUnitPtr^.aBom = nil then  // 没有BOM，异常，记录日志
              begin
                Memo1.Lines.Add(aMRPUnitPtr^.snumber + ' 无BOM');
                aMRPUnitPtr_Dep := New(PMRPUnit);
                aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
                lstDemand.Add(aMRPUnitPtr_Dep);
                Continue;
              end;
            end
            else  //// 外购件，不展开BOM
            begin
              aMRPUnitPtr_Dep := New(PMRPUnit);
              aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
              lstDemand.Add(aMRPUnitPtr_Dep);
              Continue;
            end;
            
            for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
            begin
              aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];
              // BOM 上层产生的需求  ， 需减去上层已分配库存  
              dQty := aMRPUnitPtr^.dQty * aSapItemGroup.Items[0].dusage;      // 每一层的需求，都是扣减了已分配库存的
              // 计算可用库存
              if dQty > 0 then
              begin
                for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                begin
                  aSapBomChild := aSapItemGroup.Items[iChildItem];
                  dQty_Stock_a := aSAPStockSum.GetAvailStock(aSapBomChild.FNumber);

                  if DoubleE(dQty_Stock_a, 0) then Continue;
                  if dQty <= dQty_Stock_a then
                  begin
                    aSapBomChild.FStock := dQty;
                    aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty);  
                    dQty := 0;
                    Break;
                  end
                  else
                  begin
                    aSapBomChild.FStock := dQty_Stock_a;           
                    aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty_Stock_a);
                    dQty := dQty - dQty_Stock_a;
                  end;
                end;
              end;
              
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
                  aMRPUnitPtr_Dep^.dQty := dQty;   
                end
                else
                begin
                  aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // 半成品替代料按配比分
                end;
                
                aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;   
                aSapBomChild.FStock := 0;  // 赋值后清0，否则遗留的数值会影响下一轮计算
                aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr^.dQtyStock + aMRPUnitPtr^.dQtyStockParent;  // 父项库存
                aMRPUnitPtr_Dep^.dQtyOPO := 0;
                aMRPUnitPtr_Dep^.bExpend := False;
                aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;   
                aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                lstDemand.Add(aMRPUnitPtr_Dep);
              end;
              bLoop := True;
            end;
          end
          else
          // 非根节点  不考虑库存 //////////////////////////////////////////////////      
          begin
            // 叶子节点， 无需往下展开//////////////////////////////////////////////
            if aMRPUnitPtr^.aBom.ChildCount = 0 then
            begin
//              aMRPUnitPtr_Dep := New(PMRPUnit);
//              aMRPUnitPtr_Dep^ := aMRPUnitPtr^;
//              lstDemand.Add(aMRPUnitPtr_Dep);
            end
            // 非叶子节点， 非根节点，即是半成品，展开需求//////////////////////////
            else    // 这里需考虑半成品库存 ////////////////////////////////////////
            begin
              aMRPUnitPtr^.bExpend := True;
 
              for iChild := 0 to aMRPUnitPtr^.aBom.ChildCount - 1 do
              begin
                aSapItemGroup := aMRPUnitPtr^.aBom.Childs[iChild];

                // BOM 上层产生的需求  ， 需减去上层已分配库存 
                dQty := aMRPUnitPtr^.dQty* aSapItemGroup.Items[0].dusage;  // 每一层的需求，都是扣减了已分配库存的

                // 计算可用库存
                if dQty > 0 then
                begin
                  for iChildItem := 0 to aSapItemGroup.ItemCount - 1 do
                  begin
                    aSapBomChild := aSapItemGroup.Items[iChildItem]; 
 
                    dQty_Stock_a := aSAPStockSum.GetAvailStock(aSapBomChild.FNumber);
                    if DoubleE(dQty_Stock_a, 0) then Continue;
                  
                    if dQty <= dQty_Stock_a then
                    begin
                      aSapBomChild.FStock := dQty;
                      aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty);
                      dQty := 0;
                      Break;
                    end
                    else
                    begin
                      aSapBomChild.FStock := dQty_Stock_a;           
                      aSAPStockSum.Alloc(aSapBomChild.FNumber, dQty_Stock_a);
                      dQty := dQty - dQty_Stock_a;
                    end;
                  end;
                end;

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
                    aMRPUnitPtr_Dep^.dQty := dQty;  // 半成品替代料按配比分
                  end
                  else
                  begin
                    aMRPUnitPtr_Dep^.dQty := dQty * aSapBomChild.dPer / 100;  // 半成品替代料按配比分
                  end;
                  aMRPUnitPtr_Dep^.dQtyStock := aSapBomChild.FStock;
                  aSapBomChild.FStock := 0;  // 赋值后清0，否则遗留的数值会影响下一轮计算

                  aMRPUnitPtr_Dep^.dQtyStockParent := aMRPUnitPtr^.dQtyStock + aMRPUnitPtr^.dQtyStockParent;
                  aMRPUnitPtr_Dep^.dQtyOPO := 0;
                  aMRPUnitPtr_Dep^.bExpend := False;
                  aMRPUnitPtr_Dep^.aBom := aSapBomChild;
                  aMRPUnitPtr_Dep^.aParentBom := aMRPUnitPtr^.aBom;     
                  aMRPUnitPtr_Dep^.aSAPMaterialRecordPtr := aSAPMaterialReader.GetSAPMaterialRecord(aSapBomChild.FNumber);
                  lstDemand.Add(aMRPUnitPtr_Dep);
                end;
                bLoop := True;
              end;          
            end;
          end;
        end;
        lstMrpDetail.Add(aMRPUnitPtr);
      end;
      lstDemand_tmp.Free;
      iLowestCode := 0;
    end;
               
    //排序，按日期
    lstMrpDetail.Sort(ListSortCompare_DateTime);

    ////  计算 PO  /////////////////////////////////////////////////////////////
    for iMrpUnit := 0 to lstMrpDetail.Count - 1 do
    begin                                 
      aMRPUnitPtr := PMRPUnit(lstMrpDetail[iMrpUnit]);
      if aMRPUnitPtr^.bExpend then Continue;
      if aMRPUnitPtr^.aBom = nil then Continue;
                                       
      if DoubleE( aMRPUnitPtr^.dQty, 0 ) then Continue; // 没有需求

      slNumber := TStringList.Create;
      for iChildItem := 0 to aMRPUnitPtr^.aBom.FParent.ItemCount - 1 do
      begin                                           
        aSapBomChild := aMRPUnitPtr^.aBom.FParent.Items[iChildItem];
        slNumber.Add(aSapBomChild.FNumber);
      end;
   
      aMRPUnitPtr^.dQty := aSAPOPOReader.Alloc(slNumber, aMRPUnitPtr^.dt,
        aMRPUnitPtr^.dQty, aMRPUnitPtr^.dQtyOPO);

      slNumber.Free;
    end;


    // 保存 //////////////////////////////////////////////////////////////////////

    Memo1.Lines.Add('开始保存模拟MRP计算结果，只保存关键物料');
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
    ExcelApp.Sheets[iSheet].Name := '成品要货计划';

    irow := 1;
    ExcelApp.Cells[irow, 1].Value := '产品编码';
    ExcelApp.Cells[irow, 2].Value := '产品名称';
    ExcelApp.Cells[irow, 3].Value := '日期';
    ExcelApp.Cells[irow, 4].Value := '数量';      

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
          
        irow := irow + 1;
      end; 
    end;
    
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////

    
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

          if stitle <> '采购凭证项目采购凭证类型采购凭证类别采购组制单人' then
          begin
            Memo1.Lines.Add(sSheet +'  不是  SAP导出库存  格式');
            Continue;
          end;

          ExcelApp2.ActiveSheet.Cells.Copy;

          ExcelApp.Sheets[iSheet].Paste;
 
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


    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := '计算过程日志';
    irow := 1;    
    ExcelApp.Cells[irow, 1].Value := 'ID';
    ExcelApp.Cells[irow, 2].Value := '父ID';
    ExcelApp.Cells[irow, 3].Value := '物料';
    ExcelApp.Cells[irow, 4].Value := '物料名称';
    ExcelApp.Cells[irow, 5].Value := '需求日期';
    ExcelApp.Cells[irow, 6].Value := '需求数量';
    ExcelApp.Cells[irow, 7].Value := '可用库存';
    ExcelApp.Cells[irow, 8].Value := '父项可用库存';      
    ExcelApp.Cells[irow, 9].Value := 'OPO';

    irow := 2;

    for iMrpUnit := 0 to lstMrpDetail.Count - 1 do
    begin 
      aMRPUnitPtr := PMRPUnit(lstMrpDetail[iMrpUnit]);
      ExcelApp.Cells[irow, 1].Value := aMRPUnitPtr^.id; 
      ExcelApp.Cells[irow, 2].Value := aMRPUnitPtr^.pid;
      ExcelApp.Cells[irow, 3].Value := aMRPUnitPtr^.snumber;
      ExcelApp.Cells[irow, 4].Value := aMRPUnitPtr^.sname;
      ExcelApp.Cells[irow, 5].Value := aMRPUnitPtr^.dt;
      ExcelApp.Cells[irow, 6].Value := aMRPUnitPtr^.dqty;
      ExcelApp.Cells[irow, 7].Value := aMRPUnitPtr^.dqtystock;
      ExcelApp.Cells[irow, 8].Value := aMRPUnitPtr^.dQtyStockParent;    
      ExcelApp.Cells[irow, 9].Value := aMRPUnitPtr^.dQtyOPO;

      irow := irow + 1;
    end;       
         
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////


    WorkBook.Sheets.Add(after:=WorkBook.Sheets[iSheet]);
    iSheet := iSheet + 1;
    ExcelApp.Sheets[iSheet].Activate;
    ExcelApp.Sheets[iSheet].Name := 'PO Action';
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
    ////////////////////////////////////////////////////////////////////////////

    try

      WorkBook.SaveAs(sfile);
      ExcelApp.ActiveWorkBook.Saved := True;   //新加的,设置已经保存

    finally
      WorkBook.Close;
      ExcelApp.Quit;
    end;
    
  finally
    aSAPMaterialReader.Free;
    aSAPBomReader.Free;
    aSAPStockReader.Free;
    aSAPOPOReader.Free;
    aSAPStockSum.Free;
    aSAPS618Reader.Free;

    for iMrpUnit := 0 to lstDemand.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstDemand[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstDemand.Free;
 
    for iMrpUnit := 0 to lstMrpDetail.Count - 1 do
    begin
      aMRPUnitPtr := PMRPUnit(lstMrpDetail[iMrpUnit]);
      Dispose(aMRPUnitPtr);
    end;
    lstMrpDetail.Free;
  end;

  MessageBox(Handle, '完成', '提示', 0);
end;

end.
