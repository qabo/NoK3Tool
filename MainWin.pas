unit MainWin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, FGDemandWin, CommUtils, MPACompareWin, FGTableReader,
  WaterfallWin2, ManMrpWin, DataIntMergeWin, DBConfigWin, ScheActExceptionWin,
  StdCtrls, LocalFGDemandWin, NewSKUReader, SopSimSumWin, SOP2PPOrderWin,
  ProjOfICItemWin, ComObj, FGStockReader, SAPBomReader, SAPStockReader,
  LTP_CMS2MRPSimWin, MRPSimulationWin, MRP4SAPWin, MRP4SAPWin3_MRP2, SOPVerCompareWin,
  SAPHW90Reader, SAPMrpAreaStockReader, FacAccountCheckWin, CPInAndStockWin,
  MRPAreaStockCheckWin, MergeS620ImportTemplateWin, XHBomLocation2RowWin,
  SAPBomMatrixWin, MRP4SAPWin2_CTB, MakeFGReportWin, SubstractMrpLogWin,
  SalePlanWin, FGAll2MZMBWin, MRP4SAPWin3_MRP, MergeBomWin, ComCtrls,
  MrpSimDemandWin, SalePlanWFWin, Excel2MrpBomWin, WhereUseWin, SOPSumWin,
  BomAllocCheckWin, FGPlanNumberWin, PCNumberWin, KittingWin;

const
  WM_My_ShowForm = WM_USER + 12;
  WP_TfrmSOPVSAct = 1;
  WP_TfrmMergePlansAnalysis = 2;
  WP_ShowForm = 3;

type
  TfrmMain = class(TForm)
    MainMenu1: TMainMenu;
    PC1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    MRP1: TMenuItem;
    MC1: TMenuItem;
    SOP1: TMenuItem;
    N3: TMenuItem;
    VS1: TMenuItem;
    BS1: TMenuItem;
    N4: TMenuItem;
    ODM1: TMenuItem;
    N5: TMenuItem;
    MPSWaterfall1: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    SOP2: TMenuItem;
    SOPVSMPS1: TMenuItem;
    N12: TMenuItem;
    SOP3: TMenuItem;
    N13: TMenuItem;
    SOPtoSAP1: TMenuItem;
    SimpleWaterfall1: TMenuItem;
    SAPtoSOP1: TMenuItem;
    nSAPBom2SBom: TMenuItem;
    nLTP_CMS2MRPSim: TMenuItem;
    SAPtoSOP2: TMenuItem;
    SOP4: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    test1: TMenuItem;
    MRP2: TMenuItem;
    SOP5: TMenuItem;
    mmiMRPAreaStockCheck: TMenuItem;
    MRP3: TMenuItem;
    N16: TMenuItem;
    asdf1: TMenuItem;
    N17: TMenuItem;
    S6201: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    SAPBomMatrix1: TMenuItem;
    N20: TMenuItem;
    MRPLog1: TMenuItem;
    N21: TMenuItem;
    MRP21: TMenuItem;
    ExcelBOM1: TMenuItem;
    StatusBar1: TStatusBar;
    miMrpSimDemand: TMenuItem;
    Waterfall1: TMenuItem;
    ExcelBomSAPBom1: TMenuItem;
    WhereUse1: TMenuItem;
    SOP6: TMenuItem;
    BOM1: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    PC2: TMenuItem;
    N24: TMenuItem;
    procedure N2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MRP1Click(Sender: TObject);
    procedure SOP1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure VS1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BS1Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure ODM1Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure MPSWaterfall1Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure SOP2Click(Sender: TObject);
    procedure SOPVSMPS1Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure SOP3Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure SOPtoSAP1Click(Sender: TObject);
    procedure SimpleWaterfall1Click(Sender: TObject);
    procedure SAPtoSOP1Click(Sender: TObject);
    procedure nSAPBom2SBomClick(Sender: TObject);
    procedure nLTP_CMS2MRPSimClick(Sender: TObject);
    procedure SOP4Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure MRP2Click(Sender: TObject);
    procedure SOP5Click(Sender: TObject);
    procedure mmiMRPAreaStockCheckClick(Sender: TObject);
    procedure MRP3Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure asdf1Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure S6201Click(Sender: TObject);
    procedure N19Click(Sender: TObject);
    procedure SAPBomMatrix1Click(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure MRPLog1Click(Sender: TObject);
    procedure test1Click(Sender: TObject);
    procedure N21Click(Sender: TObject);
    procedure MRP21Click(Sender: TObject);
    procedure ExcelBOM1Click(Sender: TObject);
    procedure miMrpSimDemandClick(Sender: TObject);
    procedure Waterfall1Click(Sender: TObject);
    procedure ExcelBomSAPBom1Click(Sender: TObject);
    procedure WhereUse1Click(Sender: TObject);
    procedure SOP6Click(Sender: TObject);
    procedure BOM1Click(Sender: TObject);
    procedure N23Click(Sender: TObject);
    procedure PC2Click(Sender: TObject);
    procedure N24Click(Sender: TObject);
  private
    { Private declarations }
    procedure OnMyShowForm(var message: TMessage); message WM_My_ShowForm;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation
   
uses
  HWCinSumWin, MrpPRMergeWin, SOPAchievementWin, MergePlansWin, SOPvsMPSWin,
  MergePlansAnalysisWin2, SOPVSActWin, BSFormatWin, HWPkgStuffWin, SOP2SAPWin,
  SWaterfall, SAP2SOPWin, SAPBom2SBomWin, MakeFGReportWin2, MergePlansAnalysisWin;

{$R *.dfm}

procedure TfrmMain.N2Click(Sender: TObject);
begin
  TfrmHWCinSum.ShowForm;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  Self.Caption := Self.Caption + '   ' + myGetFileVersion(Application.ExeName);

  {$ifdef __SAP}
  N14.Visible := False;    
  N9.Visible := False;
  MC1.Visible := False;
  PC1.Visible := False;
  {$ENDIF}

  {$ifdef _XiaoHai}
  PC1.Visible := False;
  MC1.Visible := False;
  N14.Visible := False;
  N9.Visible := False;
  SOPtoSAP1.Visible := False;
  SAPtoSOP2.Visible := False;
  N18.Visible := True;
  N19.Visible := True;
  N21.Visible := False;
  {$else}
    {$ifdef _ML}    
    PC1.Visible := False;
    MC1.Visible := False;
    N14.Visible := False;
    N9.Visible := False;
    SOPtoSAP1.Visible := False;
    SAPtoSOP2.Visible := False;
    N18.Visible := True;
    N19.Visible := False;
    N21.Visible := True;
    {$else}
    N18.Visible := False;
    {$endif}
  {$endif}
end;

procedure TfrmMain.MRP1Click(Sender: TObject);
begin
  TfrmMrpPRMerge.ShowForm;
end;

procedure TfrmMain.SOP1Click(Sender: TObject);
begin
  TfrmSOPAchievement.ShowForm;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  TfrmMergePlans.ShowForm;
end;
        
procedure TfrmMain.asdf1Click(Sender: TObject);
begin
  TfrmMergePlansAnalysis.ShowForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  TfrmMergePlansAnalysis2.ShowForm;
end;

procedure TfrmMain.N5Click(Sender: TObject);
begin
  TfrmMPACompare.ShowForm;
end;

procedure TfrmMain.VS1Click(Sender: TObject);
begin
  TfrmSOPVSAct.ShowForm;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
//  PostMessage(Self.Handle, WM_My_ShowForm, WP_TfrmSOPVSAct, 0);
//  PostMessage(Self.Handle, WM_My_ShowForm, WP_TfrmMergePlansAnalysis, 0);
  PostMessage(Self.Handle, WM_My_ShowForm, WP_ShowForm, 0);
end;

procedure TfrmMain.OnMyShowForm(var message: TMessage);
begin
  case message.WParam of
    WP_TfrmSOPVSAct:
    begin
      TfrmSOPVSAct.ShowForm;
    end;
    WP_TfrmMergePlansAnalysis:
    begin
      TfrmMergePlansAnalysis2.ShowForm;
    end;
    WP_ShowForm:
    begin
      TfrmKitting.ShowForm;   
    end;
  end;
end;

procedure TfrmMain.BS1Click(Sender: TObject);
begin
  TfrmBSFormat.ShowForm;
end;

procedure TfrmMain.N4Click(Sender: TObject);
begin
  TfrmFGDemand.ShowForm;
end;

procedure TfrmMain.ODM1Click(Sender: TObject);
begin
  TfrmScheActException.ShowForm; 
end;

procedure TfrmMain.MPSWaterfall1Click(Sender: TObject);
begin
  TfrmWaterfall2.ShowForm;
end;
     
procedure TfrmMain.SimpleWaterfall1Click(Sender: TObject);
begin
  TfrmSWaterfall.ShowForm;
end;

procedure TfrmMain.N6Click(Sender: TObject);
begin
  TfrmManMrp.ShowForm
end;

procedure TfrmMain.N7Click(Sender: TObject);
begin
  TfrmDataIntMerge.ShowForm;
end;

procedure TfrmMain.N10Click(Sender: TObject);
begin
  TfrmDBConfig.ShowForm;
end;

procedure TfrmMain.N11Click(Sender: TObject);
begin
  TfrmLocalFGDemand.ShowForm;
end;

procedure TfrmMain.SOP2Click(Sender: TObject);
begin
  TfrmSopSimSum.ShowForm;
end;

type
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
        Result := 1
      else Result := -1;
    end
    else
    begin
      if p1^.dos <> p2.dos then
      begin
        if p1^.dos > p2^.dos then
          Result := -1
        else Result := 1;
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

procedure TfrmMain.SOPVSMPS1Click(Sender: TObject);
begin
  TfrmSOPvsMPS.ShowForm;
end;

procedure TfrmMain.N12Click(Sender: TObject);
begin
  TfrmProjOfICItem.ShowForm;
end;

procedure TfrmMain.SOP3Click(Sender: TObject);
begin
  TfrmSOP2PPOrder.ShowForm;
end;

procedure TfrmMain.N13Click(Sender: TObject);
begin
  TfrmHWPkgStuff.ShowForm;   
end;

procedure TfrmMain.SOPtoSAP1Click(Sender: TObject);
begin
  TfrmSOP2SAP.ShowForm;
end;

procedure TfrmMain.SAPtoSOP1Click(Sender: TObject);
begin
  TfrmSAP2SOP.ShowForm;
end;

procedure TfrmMain.nSAPBom2SBomClick(Sender: TObject);
begin
  TfrmSAPBom2SBom.ShowForm;
end;

procedure TfrmMain.nLTP_CMS2MRPSimClick(Sender: TObject);
begin
  TfrmLTP_CMS2MRPSim.ShowForm;
end;

procedure TfrmMain.SOP4Click(Sender: TObject);
begin
  TfrmMRPSimulation.ShowForm;
end;

procedure TfrmMain.MRP2Click(Sender: TObject);
begin
  TfrmMRP4SAP2_CTB.ShowForm;
end;
      
procedure TfrmMain.MRP3Click(Sender: TObject);
begin
  TfrmMRP4SAP3_MRP.ShowForm;
end;

procedure TfrmMain.SOP5Click(Sender: TObject);
begin
  TfrmSOPVerCompare.ShowForm;
end;

procedure TfrmMain.mmiMRPAreaStockCheckClick(Sender: TObject);
begin
  TfrmMRPAreaStockCheck.ShowForm;
end;

procedure TfrmMain.N16Click(Sender: TObject);
begin
  TfrmFacAccountCheck.ShowForm;
end;

procedure TfrmMain.N17Click(Sender: TObject);
begin
  TfrmCPInAndStock.ShowForm;
end;

procedure TfrmMain.S6201Click(Sender: TObject);
begin
  TfrmMergeS620ImportTemplate.ShowForm;
end;

procedure TfrmMain.N19Click(Sender: TObject);
begin
  TfrmXHBomLocation2Row.ShowForm;
end;

procedure TfrmMain.SAPBomMatrix1Click(Sender: TObject);
begin
  TfrmSAPBomMatrix.ShowForm;
end;

procedure TfrmMain.N20Click(Sender: TObject);
begin
  TfrmMakeFGReport.ShowForm;
end;
    
procedure TfrmMain.N15Click(Sender: TObject);
begin
  TfrmMakeFGReport2.ShowForm;
end;

procedure TfrmMain.MRPLog1Click(Sender: TObject);
begin
  TfrmSubstractMrpLog.ShowForm;
end;

procedure TfrmMain.N21Click(Sender: TObject);
begin
  TfrmFGAll2MZMB.ShowForm;
end;

procedure TfrmMain.MRP21Click(Sender: TObject);
begin
  TfrmMRP4SAP3_MRP2.ShowForm;
end;

procedure TfrmMain.ExcelBOM1Click(Sender: TObject);
begin
  TfrmMergeBom.ShowForm('E');
end;

procedure TfrmMain.miMrpSimDemandClick(Sender: TObject);
begin
  TfrmMrpSimDemand.ShowForm;
end;
    
procedure TfrmMain.test1Click(Sender: TObject);
begin
  TfrmSalePlan.ShowForm;
end;

procedure TfrmMain.Waterfall1Click(Sender: TObject);
begin
  TfrmSalePlanWF.ShowForm;
end;

procedure TfrmMain.ExcelBomSAPBom1Click(Sender: TObject);
begin
  TfrmExcel2MrpBom.ShowForm;
end;

procedure TfrmMain.WhereUse1Click(Sender: TObject);
begin
  TfrmWhereUse.ShowForm;
end;

procedure TfrmMain.SOP6Click(Sender: TObject);
begin
  TfrmSOPSum.ShowForm;
end;

procedure TfrmMain.BOM1Click(Sender: TObject);
begin
  TfrmBomAllocCheck.ShowForm;
end;

procedure TfrmMain.N23Click(Sender: TObject);
begin
  TfrmFGPlanNumber.ShowForm;
end;

procedure TfrmMain.PC2Click(Sender: TObject);
begin
  TfrmPCNumber.ShowForm;
end;

procedure TfrmMain.N24Click(Sender: TObject);
begin
  TfrmKitting.ShowForm;   
end;

end.
